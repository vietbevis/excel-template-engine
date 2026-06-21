import { readFile, stat } from 'node:fs/promises';
import { extname, isAbsolute, relative, resolve } from 'node:path';
import type { AssetResolver, AssetResolverOptions, ResolvedAsset } from '../../application/managers/ports.js';
import { LimitExceededError } from '../../shared/errors/engine-error.js';

export class DefaultAssetResolver implements AssetResolver {
  private readonly baseDir: string;
  private readonly allowAbsolutePaths: boolean;
  private readonly maxBytes: number;

  constructor(options: AssetResolverOptions | string = {}) {
    const normalizedOptions = typeof options === 'string' ? { baseDir: options } : options;
    this.baseDir = resolve(normalizedOptions.baseDir ?? process.cwd());
    this.allowAbsolutePaths = normalizedOptions.allowAbsolutePaths ?? false;
    this.maxBytes = normalizedOptions.maxBytes ?? 10 * 1024 * 1024;
  }

  async resolve(source: unknown): Promise<ResolvedAsset> {
    if (typeof source === 'string') {
      return this.resolveString(source);
    }

    if (Buffer.isBuffer(source)) {
      return this.resolveBytes(source);
    }

    if (source instanceof Uint8Array) {
      return this.resolveBytes(Buffer.from(source));
    }

    if (source instanceof ArrayBuffer) {
      return this.resolveBytes(Buffer.from(source));
    }

    throw new Error('Unsupported image source. Expected file path, base64, Buffer, Uint8Array, or ArrayBuffer.');
  }

  private async resolveString(source: string): Promise<ResolvedAsset> {
    const trimmed = source.trim();
    const dataUrl = /^data:image\/(?<type>png|jpe?g);base64,(?<payload>.+)$/i.exec(trimmed);
    if (dataUrl?.groups?.payload) {
      return this.resolveBytes(this.decodeBase64(dataUrl.groups.payload), dataUrl.groups.type);
    }

    if (this.looksLikeBase64(trimmed)) {
      return this.resolveBytes(this.decodeBase64(trimmed));
    }

    const filePath = this.resolveFilePath(trimmed);
    await this.assertFileSize(filePath);
    const bytes = await readFile(filePath);
    return this.resolveBytes(bytes, extname(filePath).slice(1));
  }

  private resolveBytes(bytes: Uint8Array, extensionHint?: string): ResolvedAsset {
    if (bytes.byteLength > this.maxBytes) {
      throw new LimitExceededError('maxImageBytes', bytes.byteLength, this.maxBytes);
    }

    const extension = this.normalizeExtension(extensionHint) ?? this.detectExtension(bytes);
    if (!extension) {
      throw new Error('Unsupported image bytes. Only png and jpg/jpeg are supported.');
    }

    return {
      bytes,
      extension,
      contentType: extension === 'png' ? 'image/png' : 'image/jpeg',
    };
  }

  private detectExtension(bytes: Uint8Array): ResolvedAsset['extension'] | undefined {
    if (bytes.length >= 8
      && bytes[0] === 0x89
      && bytes[1] === 0x50
      && bytes[2] === 0x4e
      && bytes[3] === 0x47
      && bytes[4] === 0x0d
      && bytes[5] === 0x0a
      && bytes[6] === 0x1a
      && bytes[7] === 0x0a) {
      return 'png';
    }

    if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
      return 'jpg';
    }

    return undefined;
  }

  private normalizeExtension(extension: string | undefined): ResolvedAsset['extension'] | undefined {
    const normalized = extension?.toLowerCase();
    if (normalized === 'png') {
      return 'png';
    }

    if (normalized === 'jpg' || normalized === 'jpeg') {
      return normalized;
    }

    return undefined;
  }

  private looksLikeBase64(value: string): boolean {
    if (!value || value.length % 4 !== 0) {
      return false;
    }

    return /^[A-Za-z0-9+/]+={0,2}$/.test(value);
  }

  private decodeBase64(value: string): Buffer {
    const estimatedBytes = Math.floor((value.replace(/=+$/, '').length * 3) / 4);
    if (estimatedBytes > this.maxBytes) {
      throw new LimitExceededError('maxImageBytes', estimatedBytes, this.maxBytes);
    }

    return Buffer.from(value, 'base64');
  }

  private resolveFilePath(source: string): string {
    if (isAbsolute(source)) {
      if (!this.allowAbsolutePaths) {
        throw new Error('Absolute image paths are disabled by default.');
      }
      return source;
    }

    const filePath = resolve(this.baseDir, source);
    const relativePath = relative(this.baseDir, filePath);
    if (relativePath.startsWith('..') || isAbsolute(relativePath)) {
      throw new Error('Image path escapes the configured baseDir.');
    }

    return filePath;
  }

  private async assertFileSize(filePath: string): Promise<void> {
    const fileStat = await stat(filePath);
    if (fileStat.size > this.maxBytes) {
      throw new LimitExceededError('maxImageBytes', fileStat.size, this.maxBytes, { filePath });
    }
  }
}
