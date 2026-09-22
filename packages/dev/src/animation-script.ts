/**
 * Object animation playback, shared by the preview's presentation mode and the
 * editor panel's Play button.
 *
 * The script is a string so the preview page can inline it and the server can
 * serve it as a module (`/animation-player.js`) for the editor bundle. It
 * depends on nothing but the DOM and the steps `getSlideAnimations` reports, so
 * both callers drive the same cursor.
 *
 * The rule it is built around: **play only what the deck states.** A step this
 * library reads but cannot reproduce is never approximated — not by guessing a
 * moment for it, not by giving it a click of its own, and not by deciding for
 * itself whether the shape should be on the slide before its turn. Such a step
 * is reported as unsupported and everything it touches is left exactly as the
 * renderer drew it, which is the only state that cannot lose content.
 *
 * What that means in detail:
 *
 *   - Only the main sequence. An `interactiveSeq` runs off its own trigger —
 *     a click on some shape — and is not part of the slide's click order, so
 *     it neither plays nor hides anything here.
 *   - A `click` step opens a stop, wherever it appears: a click stop is stated
 *     by the tree, so a group of known effects after an unsupported one still
 *     plays. `withPrevious` runs from the start of the group in hand and
 *     `afterPrevious` from its end. The first step of a slide has no
 *     predecessor, so a leading `withPrevious` / `afterPrevious` starts as the
 *     slide appears rather than waiting for a click.
 *   - A step whose start, length, delay, target or preset we cannot read makes
 *     the rest of *its own group* unplaceable: an `afterPrevious` chained onto
 *     an effect of unknown length has no stated moment, and the
 *     `withPrevious` effects sharing that start have none either. That span is
 *     unsupported. The next click stop starts from a stated moment again.
 *   - A shape id that is drawn more than once is not a handle. The schema puts
 *     no uniqueness constraint on `<p:cNvPr id>`, and nothing says a timing
 *     tree that names it means every one of them, so the step is unsupported
 *     rather than animating all of them together.
 */
export const animationScript = `
const ANIMATION_INSTANT_EFFECTS=['appear','disappear'];
const ANIMATION_ENTRANCE_EFFECTS=['fadeIn','appear'];
const ANIMATION_EXIT_EFFECTS=['fadeOut','disappear'];

/** Whether a step reveals its target or hides it; null when we cannot tell. */
function animationKindOf(step){
  if(ANIMATION_ENTRANCE_EFFECTS.includes(step.effect))return 'entrance';
  if(ANIMATION_EXIT_EFFECTS.includes(step.effect))return 'exit';
  return null;
}

/**
 * How long an effect runs, or null when the deck does not say. Null is not
 * zero: it is the reason nothing can be chained onto the end of this effect.
 */
function animationDurationOf(step){
  if(Number.isFinite(step.durationMs)&&step.durationMs>=0)return step.durationMs;
  // 'appear' and 'disappear' are instantaneous by definition of the preset, so
  // a tree that states no duration for one is not leaving anything unsaid.
  if(ANIMATION_INSTANT_EFFECTS.includes(step.effect))return 0;
  return null;
}

/**
 * Groups the steps into the stops the viewer clicks through, with each step's
 * start measured from the beginning of its stop, and lists the steps that are
 * left out with the reason why.
 *
 * Reasons: 'notInMainSeq', 'notModelled' (the effect, its target or its delay
 * is one this library only reads) and 'unknownTiming' (the step's moment
 * depends on an effect whose length or place the tree does not state).
 */
function buildAnimationStops(steps){
  const stops=[];
  const unsupported=[];
  const skip=(step,reason)=>{unsupported.push({step:step,reason:reason});};
  let stop=null,groupStart=0,groupEnd=0,seen=false;
  for(const step of steps||[]){
    if(step.sequence!=='mainSeq'){skip(step,'notInMainSeq');continue;}
    const leading=!seen;
    seen=true;
    if(step.start==='unknown'){
      // Nothing states whether this opens a click stop or joins the one in
      // hand, so it is neither played nor allowed to place what follows it.
      groupStart=null;groupEnd=null;
      skip(step,'notModelled');
      continue;
    }
    if(step.start==='click'||leading){
      stop={items:[],auto:leading&&step.start!=='click'};
      stops.push(stop);
      groupStart=0;groupEnd=0;
    }else if(step.start==='afterPrevious'){
      groupStart=groupEnd;groupEnd=groupStart;
    }
    // 'withPrevious' joins the group in hand: same start, same end so far.
    const kind=step.playable?animationKindOf(step):null;
    const duration=kind===null?null:animationDurationOf(step);
    const delay=step.delayMs;
    if(kind===null||delay===null||groupStart===null){
      // Its place in the click order is still known — it keeps its stop — but
      // its moment, or what it does, is not. Nothing may be chained onto it.
      skip(step,kind===null||delay===null?'notModelled':'unknownTiming');
      groupEnd=null;
      continue;
    }
    const begin=groupStart+delay;
    stop.items.push({step:step,kind:kind,begin:begin,duration:duration});
    groupEnd=duration===null||groupEnd===null?null:Math.max(groupEnd,begin+duration);
  }
  return {stops:stops,unsupported:unsupported};
}

const animationShapeSelector=id=>'[data-pptx-shape-id="'+id+'"]';

/**
 * The elements one step animates: a whole object, or paragraphs of its text.
 * Empty when the step's target is one we do not act on.
 */
function animationTargets(root,step,blocked){
  const target=step.target;
  // 'unsupported' names a shape only as a hint about what an effect touches,
  // not as the thing to animate, so it is never resolved to an element.
  if(!target||(target.kind!=='shape'&&target.kind!=='paragraphs'))return [];
  if(!Number.isInteger(target.shapeId))return [];
  const id=String(target.shapeId);
  if(blocked&&blocked.has(id))return [];
  const shapes=root.querySelectorAll(animationShapeSelector(id));
  if(shapes.length!==1)return [];
  const shape=shapes[0];
  if(target.kind!=='paragraphs')return [shape];
  const out=[];
  // One pass over the markers that were actually drawn: a paragraph range may
  // name far more paragraphs than the shape has, and a group nested inside this
  // shape brings paragraphs of its own that belong to a different target. The
  // nearest shape of *any* id is the one a paragraph belongs to.
  for(const marker of shape.querySelectorAll('[data-pptx-paragraph]')){
    if(marker.closest('[data-pptx-shape-id]')!==shape)continue;
    const n=Number(marker.getAttribute('data-pptx-paragraph'));
    if(n>=target.firstParagraph&&n<=target.lastParagraph)out.push(marker);
  }
  return out;
}

/**
 * Plays a slide's animations against a rendered SVG.
 *
 * 'root' is the node the slide was drawn into (a shadow root is fine), or a
 * function returning it. It is asked again on every query because a transition
 * draws the slide into a layer of its own, and the layer beside it holds a
 * different slide whose shapes may carry the very same ids.
 * 'steps' is what 'getSlideAnimations' reported for that slide.
 * 'reducedMotion' is asked afresh on every advance, so a viewer who turns it
 * on mid-show gets the rest of the deck without motion — in the same order.
 * 'onChange' is called whenever the cursor moves and whenever the slide changes
 * by itself — a delayed effect arriving, an effect finishing — so a view
 * mirroring the slide never lags behind what the audience can see.
 */
function createAnimationPlayer(options){
  const root=()=>typeof options.root==='function'?options.root():options.root;
  const parsed=buildAnimationStops(options.steps);
  const stops=parsed.stops;
  const unsupported=parsed.unsupported.slice();
  const reduced=()=>Boolean(options.reducedMotion&&options.reducedMotion());
  const onChange=options.onChange||function(){};
  const animations=[];
  const timers=[];
  const touched=new Set();
  const nowMs=()=>(typeof performance==='object'?performance.now():Date.now());
  // Bumped by every call that sets where the slide stands, so a second view can
  // tell an explicit seek from a passing report — including a seek that lands
  // on the cursor it is already on and turns a running stop into a settled one.
  let cursor=0,disposed=false,startedAt=null,generation=0;

  // Every shape an unsupported step names — 'targetShapeIds' covers a composite
  // effect that drives several — is left the way the renderer drew it. We know
  // something animates it and cannot reproduce that, and of the two states we
  // could choose, only 'shown' is unable to hide content the deck does show.
  const blocked=new Set();
  for(const entry of unsupported)
    for(const shapeId of entry.step.targetShapeIds||[])blocked.add(String(shapeId));
  const drawn=new Map();
  for(const stop of stops)
    for(const item of stop.items){
      const id=String(item.step.target.shapeId);
      if(!drawn.has(id))drawn.set(id,root().querySelectorAll(animationShapeSelector(id)).length);
    }
  for(const stop of stops)
    for(const item of stop.items)
      if(drawn.get(String(item.step.target.shapeId))>1)
        unsupported.push({step:item.step,reason:'ambiguousTarget'});
  for(const [id,count] of drawn)if(count>1)blocked.add(id);

  const clearPending=()=>{
    for(const timer of timers)clearTimeout(timer);
    timers.length=0;
    for(const animation of animations)animation.cancel();
    animations.length=0;
  };
  const drop=(list,entry)=>{const at=list.indexOf(entry);if(at>=0)list.splice(at,1);};
  // Every moment the slide changes on its own, not just when the viewer clicks:
  // an effect waiting on a delay arrives while the ones beside it are still
  // running, and a presenter view mirroring the slide has to see that. It is
  // also how a self-advancing slide learns that its last effect has finished —
  // the stop stopped being 'pending' when it started, not when it ended.
  const notifyChanged=()=>{
    if(disposed)return;
    // Once the last effect of a stop has run, the slide is settled: saying so
    // is what stops a second view from resuming into a stop that has ended.
    if(timers.length===0&&animations.length===0)startedAt=null;
    onChange(player);
  };
  const show=(el)=>{touched.add(el);el.style.visibility='';el.style.opacity='';};
  const hide=(el)=>{touched.add(el);el.style.visibility='hidden';el.style.opacity='';};

  /**
   * Puts every animated element where it stands once 'played' stops have run,
   * with no motion. Idempotent, so back, jump and a cancelled advance all land
   * on the same picture as playing forward to that point.
   */
  const settle=(played)=>{
    const byElement=new Map();
    stops.forEach((stop,stopIndex)=>{
      for(const item of stop.items){
        for(const el of animationTargets(root(),item.step,blocked)){
          const list=byElement.get(el)||[];
          list.push({kind:item.kind,played:stopIndex<played});
          byElement.set(el,list);
        }
      }
    });
    for(const [el,list] of byElement){
      // An entrance that has not run yet keeps its target off the slide, even
      // when a later step would have shown it.
      if(list[0].kind==='entrance'&&!list[0].played){hide(el);continue;}
      let visible=true;
      for(const entry of list)if(entry.played)visible=entry.kind!=='exit';
      if(visible)show(el);else hide(el);
    }
  };

  // 'into' is how far through the effect already is — 0 for one starting now,
  // more for one a second view is joining part-way through.
  const play=(item,into)=>{
    const targets=animationTargets(root(),item.step,blocked);
    if(targets.length===0)return;
    const entering=item.kind==='entrance';
    const duration=item.duration;
    const from=Math.max(0,into||0);
    if(reduced()||duration===null||duration===0||from>=duration){
      for(const el of targets)entering?show(el):hide(el);
      return;
    }
    for(const el of targets){
      touched.add(el);
      el.style.visibility='';
      const animation=el.animate(
        entering?[{opacity:0},{opacity:1}]:[{opacity:1},{opacity:0}],
        {duration:duration,fill:'both'},
      );
      animation.currentTime=from;
      animations.push(animation);
      animation.finished.then(()=>{
        if(disposed)return;
        drop(animations,animation);
        animation.cancel();
        entering?show(el):hide(el);
        notifyChanged();
      },()=>{});
    }
  };

  const runStop=(stopIndex,elapsed)=>{
    const stop=stops[stopIndex];
    if(!stop)return;
    const since=Math.max(0,elapsed||0);
    startedAt=nowMs()-since;
    for(const item of stop.items){
      const wait=item.begin-since;
      if(reduced()||wait<=0){play(item,since-item.begin);continue;}
      const timer=setTimeout(()=>{drop(timers,timer);play(item,0);notifyChanged();},wait);
      timers.push(timer);
    }
  };

  const player={
    get cursor(){return cursor;},
    get stopCount(){return stops.length;},
    /** True while a click still has an effect to play on this slide. */
    get pending(){return cursor<stops.length;},
    /** True while an effect already started is still on its way. */
    get running(){return timers.length>0||animations.length>0;},
    /**
     * What the presenter view shows, and what the panel's Play reports.
     * 'elapsed' is how long the stop in hand has been running, which is what a
     * second view needs to join it part-way through rather than from the top.
     */
    get progress(){
      return {
        cursor:cursor,
        stops:stops.length,
        elapsed:startedAt===null?null:Math.round(nowMs()-startedAt),
        generation:generation,
      };
    },
    /**
     * The steps this player does not run, so a caller can say so rather than
     * let a slide look as though it has no animation. One entry per step, with
     * its reason and the shapes it names.
     */
    get unsupported(){
      return unsupported.map((entry)=>({
        reason:entry.reason,
        effect:entry.step.effect,
        shapeIds:(entry.step.targetShapeIds||[]).slice(),
      }));
    },
    reset(){
      generation+=1;clearPending();startedAt=null;cursor=0;settle(0);onChange(player);
      // A slide whose first effect runs with or after 'the previous one' has no
      // previous one: it starts as the slide appears, without a click.
      if(stops[0]&&stops[0].auto)player.advance();
    },
    /** Plays the next stop. False when the slide has none left. */
    advance(){
      if(disposed||cursor>=stops.length)return false;
      generation+=1;clearPending();
      settle(cursor);
      const stopIndex=cursor;
      cursor+=1;
      runStop(stopIndex);
      onChange(player);
      return true;
    },
    /** Takes one stop back, showing everything before it. False at the start. */
    back(){
      if(disposed||cursor<=0)return false;
      generation+=1;clearPending();startedAt=null;cursor-=1;settle(cursor);onChange(player);
      return true;
    },
    /** Moves straight to a point in the click order, without motion. */
    jumpTo(next){
      if(disposed)return;
      generation+=1;clearPending();startedAt=null;
      cursor=Math.max(0,Math.min(Math.round(next)||0,stops.length));
      settle(cursor);onChange(player);
    },
    /**
     * Shows the moment another view of the same slide is showing: the stops
     * before 'next' played out, and the one it is on started 'elapsed' ago —
     * so an effect waiting on a delay over there has not arrived here either,
     * and one half-way through a fade is half-way through it here.
     */
    resume(next,elapsed){
      if(disposed)return;
      generation+=1;clearPending();startedAt=null;
      cursor=Math.max(0,Math.min(Math.round(next)||0,stops.length));
      if(cursor>0&&Number.isFinite(elapsed)){settle(cursor-1);runStop(cursor-1,elapsed);}
      else settle(cursor);
      onChange(player);
    },
    /** Everything the slide ends with — what a still preview shows. */
    finish(){player.jumpTo(stops.length);},
    dispose(){
      if(disposed)return;
      disposed=true;clearPending();
      for(const el of touched){el.style.visibility='';el.style.opacity='';}
      touched.clear();
    },
  };
  return player;
}
`;
