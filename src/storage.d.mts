/** Frozen optional embedding surface: maintenance fixes only. The integrated app uses the MCP CLI. See docs/SDK.md. */
import type {StorageAdapter,StorageObject,StorageReceipt} from './sdk.mjs';
export class SQLiteObjectStore implements StorageAdapter {
  constructor(file: string,limits?: {maxObjectBytes?: number;maxTotalBytes?: number;maxObjects?: number});
  read(key: string): Promise<StorageObject>;
  putIfAbsent(input: {key: string;bytes: Buffer;operationId: string}): Promise<StorageReceipt>;
  close(): void;
}
