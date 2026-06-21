# UML Và Sơ Đồ Phụ Thuộc

## Sơ Đồ Thành Phần

```mermaid
classDiagram
  class ExcelTemplateEngine {
    +registerHelper(name, helper)
    +render(input, data, options)
  }

  class TemplateEngineService {
    +render(request)
  }

  class TemplateLexer {
    +tokenize(input)
  }

  class TemplateParser {
    +parseCell(input)
    +parseWorkbook(source)
  }

  class RenderPlanner {
    +createPlan(ast, context)
  }

  class WorkbookRenderer {
    <<interface>>
    +load(input)
    +apply(plan)
    +write()
  }

  class ExcelJsWorkbookRenderer {
    +load(input)
    +apply(plan)
    +write()
  }

  class MergeManager
  class FormulaManager
  class StyleCloneManager
  class ImageManager

  ExcelTemplateEngine --> TemplateEngineService
  TemplateEngineService --> TemplateParser
  TemplateEngineService --> RenderPlanner
  TemplateEngineService --> WorkbookRenderer
  TemplateParser --> TemplateLexer
  WorkbookRenderer <|.. ExcelJsWorkbookRenderer
  ExcelJsWorkbookRenderer --> MergeManager
  ExcelJsWorkbookRenderer --> FormulaManager
  ExcelJsWorkbookRenderer --> StyleCloneManager
  ExcelJsWorkbookRenderer --> ImageManager
```

## Luồng Render

```mermaid
sequenceDiagram
  participant User as Người dùng
  participant Engine as ExcelTemplateEngine
  participant Parser as TemplateParser
  participant Planner as RenderPlanner
  participant Renderer as ExcelJsWorkbookRenderer

  User->>Engine: render(template, data)
  Engine->>Renderer: load(template)
  Engine->>Parser: parseWorkbook(scanned cells)
  Parser-->>Engine: WorkbookAST
  Engine->>Planner: createPlan(ast, data)
  Planner-->>Engine: RenderPlan
  Engine->>Renderer: apply(plan)
  Renderer-->>Engine: workbook output
  Engine-->>User: Uint8Array
```

## Sơ Đồ Core Engine

```mermaid
flowchart TD
  Engine["TemplateEngine"]
  Parser["TemplateParser"]
  Lexer["TemplateLexer"]
  AST["TemplateNode AST"]
  Context["RenderContext"]
  Visitor["TemplateRenderVisitor"]
  Resolver["JsonPathResolver"]
  Helpers["DefaultHelperRegistry"]

  Engine --> Parser
  Parser --> Lexer
  Parser --> AST
  Engine --> Context
  Engine --> Visitor
  Visitor --> AST
  Visitor --> Resolver
  Visitor --> Helpers
```

## Sequence EachNode Renderer

```mermaid
sequenceDiagram
  participant Visitor as TemplateRenderVisitor
  participant Evaluator as ExpressionEvaluator
  participant Loop as LoopContext
  participant Child as RenderContext child

  Visitor->>Evaluator: evaluate(each.path)
  Evaluator-->>Visitor: array
  loop item trong array
    Visitor->>Loop: forItem(index, length, parent)
    Visitor->>Child: create child context
    Visitor->>Evaluator: evaluate child placeholders
    Evaluator-->>Visitor: value/index/first/last
  end
```

## Sequence GridNode Renderer

```mermaid
sequenceDiagram
  participant Planner as RenderPlanner
  participant Cursor as RenderCursor
  participant GridContext as GridContext
  participant Renderer as ExcelJS Renderer

  Planner->>Planner: resolve students
  Planner->>Planner: resolve subjects
  Planner->>Planner: emit CloneRow/CloneColumn nếu cần
  loop row x column
    Planner->>GridContext: build cell scope
    Planner->>Cursor: compute target cell
    Planner->>Planner: emit SetCellValue
  end
  Renderer->>Renderer: expand columns
  Renderer->>Renderer: expand rows
  Renderer->>Renderer: fill intersection cells
```
