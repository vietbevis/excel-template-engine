import type ExcelJS from 'exceljs';
import type { AssetResolver, ImageInsertOptions, ImageManager } from '../../application/managers/ports.js';
import type { CellAddress } from '../../shared/address/address.js';
import { DefaultAssetResolver } from '../assets/default-asset-resolver.js';

export class ExcelJsImageManager implements ImageManager {
  constructor(
    private readonly worksheet: ExcelJS.Worksheet,
    private readonly assetResolver: AssetResolver = new DefaultAssetResolver(),
  ) {}

  async insertImage(source: unknown, target: CellAddress, options: ImageInsertOptions = {}): Promise<void> {
    if (source == null || source === '') {
      return;
    }

    const asset = await this.assetResolver.resolve(source);
    const imageId = this.worksheet.workbook.addImage({
      base64: Buffer.from(asset.bytes).toString('base64'),
      extension: asset.extension === 'jpg' ? 'jpeg' : asset.extension,
    });

    this.worksheet.addImage(imageId, this.toImageRange(target, options));
  }

  private toImageRange(target: CellAddress, options: ImageInsertOptions): ExcelJS.ImageRange | ExcelJS.ImagePosition {
    const topLeft = {
      col: target.column - 1,
      row: target.row - 1,
    };

    if (options.width || options.height) {
      return {
        tl: topLeft,
        ext: {
          width: options.width ?? this.worksheet.getColumn(target.column).width ?? 64,
          height: options.height ?? this.worksheet.getRow(target.row).height ?? 64,
        },
      };
    }

    return {
      tl: topLeft,
      br: {
        col: target.column,
        row: target.row,
      },
    };
  }
}
