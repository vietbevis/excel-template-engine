import type { AssetResolverOptions, TemplateInput, WorkbookRenderer } from '../managers/ports.js';

export type { TemplateInput } from '../managers/ports.js';

export type JsonObject = Record<string, unknown>;

export interface EngineRenderOptions {
  readonly preserveFormulas?: boolean;
  readonly recalculateFormulas?: boolean;
  readonly missingValue?: 'empty-string' | 'null' | 'throw';
  readonly renderer?: WorkbookRenderer;
  readonly limits?: RenderLimits;
  readonly assetResolver?: AssetResolverOptions;
}

export interface RenderLimits {
  readonly maxTemplateBytes?: number;
  readonly maxWorksheets?: number;
  readonly maxRows?: number;
  readonly maxColumns?: number;
  readonly maxOperations?: number;
}

export interface EngineRenderRequest {
  readonly template: TemplateInput;
  readonly data: JsonObject;
  readonly options?: EngineRenderOptions;
}

export interface EngineRenderResult {
  readonly output: Uint8Array;
  readonly warnings: readonly string[];
}

export interface HelperContext {
  readonly data: JsonObject;
  readonly current: unknown;
  readonly root: JsonObject;
}

export type EngineHelper = (
  args: readonly unknown[],
  context: HelperContext,
) => unknown | Promise<unknown>;
