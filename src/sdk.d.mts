/** Frozen optional embedding surface: maintenance fixes only. The integrated app uses the MCP CLI. See docs/SDK.md. */
export interface StorageObject { bytes: Buffer; revision: string }
export interface StorageReceipt { key: string; revision: string; operationId: string; published: true }
/** putIfAbsent MUST atomically preserve existing keys and return the same result for identical operationId + key + bytes on retry, including after process restart. */
export interface StorageAdapter {
  read(key: string): Promise<StorageObject>;
  putIfAbsent(input: {key: string; bytes: Buffer; operationId: string}): Promise<StorageReceipt>;
}
export class OfficeClient {
  constructor(options: {directory: string; storage: StorageAdapter; limits?: Record<string,number>});
  manifest(): Array<{name: string; description: string; inputSchema: object}>;
  call(name: string,args?: Record<string,unknown>): Promise<{isError?: boolean; content?: Array<{type:"image";mimeType:string;data:string}>; payload: {status: string; [key: string]: unknown}}>;
  importSource(key: string,format: 'docx'|'xlsx'|'pptx'|'pdf'): Promise<{assetId: string;hash: string;size: number;format: string}>;
  exportAsset(input: {assetId: string;key: string;operationId: string}): Promise<StorageReceipt>;
  close(): void;
}
