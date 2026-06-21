import type {
  BlockNode,
  BlockEndNode,
  BlockStartNode,
  EachColumnNode,
  EachNode,
  HelperNode,
  GridNode,
  IfNode,
  ImageNode,
  PlaceholderNode,
  TemplateNode,
  TextNode,
} from '../ast/nodes.js';

export interface TemplateNodeVisitor<TResult> {
  visitText(node: TextNode): TResult;
  visitPlaceholder(node: PlaceholderNode): TResult;
  visitEach(node: EachNode): TResult;
  visitEachColumn(node: EachColumnNode): TResult;
  visitGrid(node: GridNode): TResult;
  visitIf(node: IfNode): TResult;
  visitHelper(node: HelperNode): TResult;
  visitImage(node: ImageNode): TResult;
  visitBlock(node: BlockNode): TResult;
  visitBlockStart(node: BlockStartNode): TResult;
  visitBlockEnd(node: BlockEndNode): TResult;
}

export function visitTemplateNode<TResult>(
  node: TemplateNode,
  visitor: TemplateNodeVisitor<TResult>,
): TResult {
  switch (node.kind) {
    case 'TextNode':
      return visitor.visitText(node);
    case 'PlaceholderNode':
      return visitor.visitPlaceholder(node);
    case 'EachNode':
      return visitor.visitEach(node);
    case 'EachColumnNode':
      return visitor.visitEachColumn(node);
    case 'GridNode':
      return visitor.visitGrid(node);
    case 'IfNode':
      return visitor.visitIf(node);
    case 'HelperNode':
      return visitor.visitHelper(node);
    case 'ImageNode':
      return visitor.visitImage(node);
    case 'BlockNode':
      return visitor.visitBlock(node);
    case 'BlockStartNode':
      return visitor.visitBlockStart(node);
    case 'BlockEndNode':
      return visitor.visitBlockEnd(node);
  }
}
