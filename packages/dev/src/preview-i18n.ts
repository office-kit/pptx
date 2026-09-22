/** Labels for the preview shell; the editor sends its active locale via editor-focus. */
export const previewI18nScript = `
let previewLocale=navigator.language?.toLowerCase().startsWith('ja')?'ja':'en';
try{const stored=localStorage.getItem('ok-editor-locale');if(stored==='ja'||stored==='en')previewLocale=stored;}catch{}
let connectionLost=false;
const previewJapanese={
 'Preview':'プレビュー','Edit':'編集','Present':'プレゼンテーション','Presenter view':'発表者ビュー',
 'Download PPTX':'PPTXをダウンロード','Slides':'スライド','Slide viewer':'スライドビューアー','Presentation editor':'プレゼンテーション編集',
 'Slide chat':'スライドチャット','Chat width':'チャットの幅','Drag to resize · Double-click to reset':'ドラッグで幅を変更・ダブルクリックでリセット',
 'Previous slide':'前のスライド','Next slide':'次のスライド','Zoom':'ズーム','Fit':'画面に合わせる',
 'Exit · Esc':'終了 · Esc','Exit view · Esc':'表示を終了 · Esc','No slides':'スライドがありません','Whole project':'プロジェクト全体',
 'Waiting for slides…':'スライドを待っています…','Building…':'ビルド中…','Updating…':'更新中…','Reconnecting…':'再接続中…',
 'Build failed · showing last successful output':'ビルドに失敗しました・最後に成功した結果を表示中',
 'Changes appear automatically · Text can be selected and copied':'変更は自動的に反映されます・文字を選択してコピーできます',
 'Skipped during presentation':'プレゼンテーションでスキップされます',
 'Allow popups to open presenter view.':'発表者ビューを開くにはポップアップを許可してください。',
 'Office Kit — PowerPoint preview':'Office Kit — PowerPoint プレビュー'
};
const pt=key=>previewLocale==='ja'?(previewJapanese[key]??key):key;
const slideLabel=i=>previewLocale==='ja'?'スライド '+(i+1):'Slide '+(i+1);
const slideCount=()=>!state.slides.length?pt('No slides'):previewLocale==='ja'?'スライド '+(index+1)+' / '+state.slides.length:'Slide '+(index+1)+' of '+state.slides.length;
function updatePreviewStatus(){
 byId('status').textContent=connectionLost?pt('Reconnecting…'):state.error?pt('Build failed · showing last successful output'):state.building?pt('Updating…'):previewLocale==='ja'?state.slides.length+' スライド・ライブ':state.slides.length+' slides · Live';
}
function updatePreviewLabels(){
 document.documentElement.lang=previewLocale;document.title=pt('Office Kit — PowerPoint preview');
 const labels={'#present':'Present','#presenter':'Presenter view','.download':'Download PPTX','.filmstrip h2':'Slides','#empty':'Waiting for slides…','footer .hint':'Changes appear automatically · Text can be selected and copied','label[for="zoom"]':'Zoom','#zoom option[value="fit"]':'Fit'};
 for(const [selector,key] of Object.entries(labels))document.querySelector(selector).textContent=pt(key);
 for(const [selector,key] of Object.entries({'.filmstrip':'Slides','main':'Slide viewer','#chat':'Slide chat','#chat-resizer':'Chat width','#prev':'Previous slide','#present-prev':'Previous slide','#next':'Next slide','#present-next':'Next slide'}))document.querySelector(selector).setAttribute('aria-label',pt(key));
 byId('editor-frame').title=pt('Presentation editor');byId('chat-resizer').title=pt('Drag to resize · Double-click to reset');
 byId('exit-present').textContent=pt('Exit · Esc');
 byId('toggle-editor').textContent=pt(document.body.classList.contains('editing')?'Preview':'Edit');
 byId('count').textContent=slideCount();byId('present-count').textContent=slideCount();
 slide.setAttribute('aria-label',slideLabel(index));
 for(const [i,item] of Array.from(thumbnails.children).entries()){
  item.firstElementChild.setAttribute('aria-label',slideLabel(i));
  item.firstElementChild.title=state.hiddenSlides?.[i]?pt('Skipped during presentation'):'';
 }
 updateChatWidthAria();updatePreviewStatus();
}
`;
