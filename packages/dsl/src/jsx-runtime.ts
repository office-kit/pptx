import { node, visit, type Child, type Children, type Node } from './model.ts';

export function jsx<P>(component: (props: P) => Node, props: P, _key?: string | number): Node {
  if (typeof component !== 'function')
    throw new Error('Use imported PPTX components, not HTML tags.');
  return component(props);
}
export const jsxs = jsx;
export function jsxDEV<P>(
  component: (props: P) => Node,
  props: P,
  key?: string | number,
  _isStaticChildren?: boolean,
  source?: Node['source'],
): Node {
  const result = jsx(component, props, key);
  return source ? { ...result, source } : result;
}
export function Fragment({ children }: Children): Node {
  return node('Fragment', (context) => visit(children, context));
}
export namespace JSX {
  export type Element = Node;
  export interface ElementChildrenAttribute {
    children: Child;
  }
  export interface IntrinsicAttributes {
    key?: string | number;
  }
}
