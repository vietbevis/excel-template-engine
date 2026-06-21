import type {
  BlockNode,
  BlockEndNode,
  BlockStartNode,
  EachColumnNode,
  EachNode,
  HelperArgument,
  HelperNode,
  GridNode,
  IfNode,
  ImageNode,
  PlaceholderNode,
  TemplateNode,
  TextNode,
} from '../ast/nodes.js';
import { ExpressionEvaluator } from '../evaluator/expression-evaluator.js';
import { GridContext } from '../grid/grid-context.js';
import type { TemplateNodeVisitor } from '../visitor/template-node-visitor.js';
import { visitTemplateNode } from '../visitor/template-node-visitor.js';
import { LoopContext } from './loop-context.js';
import type { RenderContext } from './render-context.js';

export class TemplateRenderVisitor implements TemplateNodeVisitor<Promise<string>> {
  private readonly evaluator = new ExpressionEvaluator();

  constructor(private readonly context: RenderContext) {}

  render(nodes: readonly TemplateNode[]): Promise<string> {
    return this.renderNodes(nodes, this.context);
  }

  visitText(node: TextNode): Promise<string> {
    return Promise.resolve(node.value);
  }

  async visitPlaceholder(node: PlaceholderNode): Promise<string> {
    return this.stringify(this.resolvePath(node.path, this.context));
  }

  async visitEach(node: EachNode): Promise<string> {
    return this.renderRepeated(node.path, node.children);
  }

  async visitEachColumn(node: EachColumnNode): Promise<string> {
    return this.renderRepeated(node.path, node.children);
  }

  async visitGrid(node: GridNode): Promise<string> {
    const rows = this.resolvePath(node.rowPath, this.context);
    const columns = this.resolvePath(node.columnPath, this.context);
    if (!Array.isArray(rows) || !Array.isArray(columns)) {
      return '';
    }

    const chunks: string[] = [];
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
      const row = rows[rowIndex];
      for (let columnIndex = 0; columnIndex < columns.length; columnIndex += 1) {
        chunks.push(await this.renderNodes(node.children, this.context.child(new GridContext({
          row,
          column: columns[columnIndex],
          rowIndex,
          columnIndex,
        }).toCurrentScope())));
      }
    }

    return chunks.join('');
  }

  async visitIf(node: IfNode): Promise<string> {
    const condition = this.resolvePath(node.conditionPath, this.context);
    return condition ? this.renderNodes(node.children, this.context) : '';
  }

  async visitHelper(node: HelperNode): Promise<string> {
    const args = node.args.map((arg) => this.resolveHelperArgument(arg, this.context));
    const value = await this.context.helpers.invoke(node.name, args, {
      data: this.context.root,
      root: this.context.root,
      current: this.context.current,
    });

    return this.stringify(value);
  }

  async visitImage(node: ImageNode): Promise<string> {
    return this.stringify(this.resolvePath(node.path, this.context));
  }

  async visitBlock(node: BlockNode): Promise<string> {
    return this.renderRepeated(node.path, node.children);
  }

  visitBlockStart(_node: BlockStartNode): Promise<string> {
    return Promise.resolve('');
  }

  visitBlockEnd(_node: BlockEndNode): Promise<string> {
    return Promise.resolve('');
  }

  private async renderRepeated(path: string, children: readonly TemplateNode[]): Promise<string> {
    const value = this.resolvePath(path, this.context);
    if (!Array.isArray(value)) {
      return '';
    }

    const chunks: string[] = [];
    for (let index = 0; index < value.length; index += 1) {
      const item = value[index];
      const loop = LoopContext.forItem(index, value.length, this.context.loop);
      chunks.push(await this.renderNodes(children, this.context.child(item, loop)));
    }

    return chunks.join('');
  }

  private async renderNodes(nodes: readonly TemplateNode[], context: RenderContext): Promise<string> {
    const visitor = context === this.context ? this : new TemplateRenderVisitor(context);
    const chunks: string[] = [];

    for (const node of nodes) {
      chunks.push(await visitTemplateNode(node, visitor));
    }

    return chunks.join('');
  }

  private resolveHelperArgument(arg: HelperArgument, context: RenderContext): unknown {
    if (arg.kind === 'literal') {
      return arg.value;
    }

    return this.resolvePath(arg.value, context);
  }

  private resolvePath(path: string, context: RenderContext): unknown {
    const resolutionContext = context.loop
      ? { root: context.root, current: context.current, loop: context.loop }
      : { root: context.root, current: context.current };

    return this.evaluator.evaluate(
      path,
      resolutionContext,
      {
        missingValue: context.missingValue,
      },
    );
  }

  private stringify(value: unknown): string {
    if (value == null) {
      return '';
    }

    if (value instanceof Date) {
      return value.toISOString();
    }

    return String(value);
  }

}
