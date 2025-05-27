/**
 * Copyright 2025 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import {inject, Injectable} from '@angular/core';
import {HttpClient, HttpErrorResponse, HttpHeaders, HttpResponse} from '@angular/common/http';
import {Observable, throwError, timer, of, firstValueFrom} from 'rxjs';
import {retry, catchError, map, tap} from 'rxjs/operators';
import {decode} from 'html-entities';
import {AuthService} from '../auth/auth.service';

// Configuration interface for the retry logic
export interface RetryConfig {
  maxRetries?: number;
  initialDelayMs?: number;
  backoffFactor?: number;
  retryableStatusCodes?: number[];
}

// Base interface for the item being processed in a batch operation.
export interface BaseProcessedItem {
  processingError?: {
    message: string;
    stage?: string;
    details?: any;
  };
  title?: string;
}

// Generic interface for a single operation within a batch request
export interface BatchOperation<P extends BaseProcessedItem, B> {
  id: string;
  method: 'POST';
  path: string;
  body: B;
  processedItem: P;
  retriesAttempted: number;
}

// Generic interface for the result of sending a single batch request
export interface SingleBatchRequestResult<P extends BaseProcessedItem, B> {
  processedOperationsThisAttempt: BatchOperation<P, B>[];
  operationsToRetryNext: BatchOperation<P, B>[];
}

// Type definition for the parser function
export type BatchResponseParser<P extends BaseProcessedItem> = (
  processedItem: P,
  responseJson: any,
  statusCode: number,
  statusText: string,
  operationId: string
) => void;


@Injectable({
  providedIn: 'root'
})
export class UtilitiesService {

  // --- API Endpoints ---
  public readonly DRIVE_API_UPLOAD_ENDPOINT = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,appProperties,webViewLink,parents';
  public readonly DRIVE_API_FILES_ENDPOINT = 'https://www.googleapis.com/drive/v3/files';
  public readonly FORMS_API_CREATE_ENDPOINT = 'https://forms.googleapis.com/v1/forms';
  public readonly FORMS_API_BATCHUPDATE_BASE_ENDPOINT = 'https://forms.googleapis.com/v1/forms/';
  public readonly FORMS_API_GET_BASE_ENDPOINT = 'https://forms.googleapis.com/v1/forms/';


  // --- Constants for specific services (can be used by those services when calling batch methods) ---
  public readonly GOOGLE_CLASSROOM_BATCH_ENDPOINT_URL = 'https://classroom.googleapis.com/batch';
  public readonly GOOGLE_CLASSROOM_MAX_OPERATIONS_PER_BATCH = 50;
  public readonly GOOGLE_DRIVE_BATCH_ENDPOINT_URL = 'https://www.googleapis.com/batch/drive/v3';
  public readonly GOOGLE_DRIVE_MAX_OPERATIONS_PER_BATCH = 100;

  private http = inject(HttpClient);
  private auth = inject(AuthService); // Inject AuthService

  constructor() { }

  /**
   * Helper function to escape single quotes and backslashes for Drive API query parameters.
   */
  public escapeQueryParam(value: string): string {
    if (!value) return '';
    return value.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  }

  /**
   * Wraps an Observable with retry logic.
   */
  public retryRequest<T>(
    request$: Observable<T>,
    config?: RetryConfig,
    operationName?: string
  ): Observable<T> {
    const defaults: Required<RetryConfig> = {
      maxRetries: 3,
      initialDelayMs: 1500,
      backoffFactor: 2,
      retryableStatusCodes: [500, 503, 504, 429]
    };
    const retryConfig: Required<RetryConfig> = {...defaults, ...config};
    const opName = operationName ? ` (${operationName})` : '';

    return request$.pipe(
      retry({
        count: retryConfig.maxRetries,
        delay: (error: HttpErrorResponse | Error, retryCount: number) => {
          const isRetryable = error instanceof HttpErrorResponse &&
            retryConfig.retryableStatusCodes.includes(error.status);
          if (isRetryable) {
            const delayTime = retryConfig.initialDelayMs * Math.pow(retryConfig.backoffFactor, retryCount - 1) + (Math.random() * 1000);
            console.warn(`UtilitiesService: Request${opName} failed (Attempt ${retryCount}/${retryConfig.maxRetries}) with status ${error instanceof HttpErrorResponse ? error.status : 'N/A'}. Retrying in ${delayTime}ms.`);
            return timer(delayTime);
          } else {
            console.error(`UtilitiesService: Request${opName} failed with non-retryable error:`, this.formatHttpError(error));
            return throwError(() => error);
          }
        }
      })
    );
  }

  /**
   * Orchestrates the execution of generic batch operations.
   * Fetches OAuth token internally.
   */
  public executeBatchOperations<P extends BaseProcessedItem, B>(
    initialOperations: BatchOperation<P, B>[],
    retryConfig: Required<RetryConfig>,
    batchUrl: string,
    maxOperationsPerBatch: number,
    responseParser: BatchResponseParser<P>
  ): Observable<P[]> {
    const serviceCallId = `utils-execBatch-${Date.now()}`;

    const authToken = this.auth.getGoogleAccessToken(); // Fetch token internally
    if (!authToken) {
      const errorMessage = `[UtilitiesService][${serviceCallId}] executeBatchOperations: Auth token missing. Cannot proceed.`;
      console.error(errorMessage);
      // Mark all items as failed due to auth token issue before returning an error observable
      initialOperations.forEach(op => {
        if (!op.processedItem.processingError) {
          op.processedItem.processingError = {
            message: "Authentication token missing for batch operation.",
            stage: "Batch Auth (UtilitiesService)"
          };
        }
      });
      // Return an observable that emits the modified items and then errors,
      // or just throwError if the caller is expected to handle the items separately.
      // For consistency with how errors are handled in other services, let's make it throw.
      return throwError(() => new Error('Authentication token missing for batch operation.'));
    }

    console.log(`[UtilitiesService][${serviceCallId}] executeBatchOperations: Starting for ${initialOperations.length} operations. URL: ${batchUrl}, MaxOps: ${maxOperationsPerBatch}`);

    return new Observable<P[]>(subscriber => {
      const operationsQueue: BatchOperation<P, B>[] = [...initialOperations];
      let activeTimersForRetries = 0;
      let iterationGuard = 0;
      const MAX_ITERATIONS = (initialOperations.length * (retryConfig.maxRetries + 1)) + initialOperations.length + 10;

      const processQueue = async () => {
        iterationGuard++;
        if (iterationGuard > MAX_ITERATIONS) {
          console.error(`[UtilitiesService][${serviceCallId}] executeBatchOperations: MAX ITERATIONS REACHED (${iterationGuard}). Aborting.`);
          operationsQueue.forEach(op => {
            if (!op.processedItem.processingError) {
              op.processedItem.processingError = {message: "Max batch processing iterations reached.", stage: "Batch Orchestration Loop (UtilitiesService)"};
            }
          });
          subscriber.next(initialOperations.map(op => op.processedItem));
          subscriber.complete();
          return;
        }

        if (operationsQueue.length === 0) {
          if (activeTimersForRetries === 0) {
            console.log(`[UtilitiesService][${serviceCallId}] executeBatchOperations: Queue empty, no active retry timers. All operations finalized. Iterations: ${iterationGuard - 1}`);
            subscriber.next(initialOperations.map(op => op.processedItem));
            subscriber.complete();
          } else {
            setTimeout(processQueue, 200);
          }
          return;
        }

        const operationsForCurrentBatch = operationsQueue.splice(0, maxOperationsPerBatch);

        try {
          const result: SingleBatchRequestResult<P, B> = await firstValueFrom(
            this.sendSingleBatchRequest<P, B>(authToken, operationsForCurrentBatch, retryConfig, batchUrl, responseParser)
          );

          if (result.operationsToRetryNext.length > 0) {
            activeTimersForRetries += result.operationsToRetryNext.length;
            for (const opToRetry of result.operationsToRetryNext) {
              const delay = retryConfig.initialDelayMs * Math.pow(retryConfig.backoffFactor, opToRetry.retriesAttempted - 1) + (Math.random() * 1000);
              const titleForLog = opToRetry.processedItem.title ? `"${opToRetry.processedItem.title}"` : "(No Title)";
              console.log(`[UtilitiesService][${serviceCallId}] Op ID ${opToRetry.id} ${titleForLog}: Scheduling retry ${opToRetry.retriesAttempted}/${retryConfig.maxRetries} in ${delay.toFixed(0)}ms.`);
              setTimeout(() => {
                operationsQueue.push(opToRetry);
                activeTimersForRetries--;
              }, delay);
            }
          }
        } catch (batchExecutionError) {
          console.error(`[UtilitiesService][${serviceCallId}] executeBatchOperations: Critical error from sendSingleBatchRequest observable for a batch chunk.`, batchExecutionError);
          operationsForCurrentBatch.forEach(op => {
            if (!op.processedItem.processingError) {
              op.processedItem.processingError = {
                message: `Batch chunk failed unrecoverably: ${batchExecutionError instanceof Error ? batchExecutionError.message : String(batchExecutionError)}`,
                stage: 'Batch Chunk Execution Pipeline (UtilitiesService)',
              };
            }
          });
        } finally {
          setTimeout(processQueue, 0);
        }
      };
      processQueue();
    });
  }

  /**
   * Sends a single HTTP batch request.
   * This method still accepts authToken as it's called by executeBatchOperations which now manages the token.
   */
  public sendSingleBatchRequest<P extends BaseProcessedItem, B>(
    authToken: string,
    operationsInBatch: BatchOperation<P, B>[],
    masterRetryConfig: Required<RetryConfig>,
    batchUrl: string,
    responseParser: BatchResponseParser<P>
  ): Observable<SingleBatchRequestResult<P, B>> {
    const serviceCallId = `utils-sendBatch-${Date.now()}`;
    const boundary = `batch_${serviceCallId}_${Math.random().toString(36).substring(2)}`;
    let multipartRequestBody = '';

    for (const operation of operationsInBatch) {
      const jsonBody = JSON.stringify(operation.body);
      multipartRequestBody += `--${boundary}\r\n`;
      multipartRequestBody += `Content-Type: application/http\r\n`;
      multipartRequestBody += `Content-ID: ${operation.id}\r\n`;
      multipartRequestBody += `\r\n`;
      multipartRequestBody += `${operation.method} ${operation.path} HTTP/1.1\r\n`;
      multipartRequestBody += `Content-Type: application/json; charset=UTF-8\r\n`;
      multipartRequestBody += `Content-Length: ${new TextEncoder().encode(jsonBody).length}\r\n`;
      multipartRequestBody += `\r\n`;
      multipartRequestBody += `${jsonBody}\r\n`;
    }
    multipartRequestBody += `--${boundary}--\r\n`;

    // Headers are created using the provided authToken
    const headers = new HttpHeaders({
      'Authorization': `Bearer ${authToken}`,
      'Accept': 'application/json',
      'Content-Type': `multipart/mixed; boundary=${boundary}`
    });

    const firstOpTitle = operationsInBatch[0]?.processedItem.title ? `"${operationsInBatch[0]?.processedItem.title}"` : "(No Title)";
    const operationDescription = `[${serviceCallId}] Batch POST of ${operationsInBatch.length} items to ${batchUrl}. First op ID: ${operationsInBatch[0]?.id} ${firstOpTitle}`;

    const request$ = this.http.post(
      batchUrl,
      multipartRequestBody,
      {headers, observe: 'response', responseType: 'text'}
    ).pipe(
      map((httpResponse: HttpResponse<string>) => {
        const responseContentType = httpResponse.headers.get('Content-Type');
        const responseBody = httpResponse.body || '';

        if (!responseContentType || !responseContentType.startsWith('multipart/mixed')) {
          console.error(`[UtilitiesService]${operationDescription}: Batch response error: Not multipart/mixed. CT: ${responseContentType}`);
          operationsInBatch.forEach(op => {
            op.processedItem.processingError = {
              message: "Batch response format error", stage: "Batch Response Parsing (UtilitiesService)",
              details: {statusCode: httpResponse.status, ct: responseContentType}
            };
          });
          return {processedOperationsThisAttempt: operationsInBatch, operationsToRetryNext: []};
        }

        this.parseMultipartResponse<P, B>(responseBody, responseContentType, operationsInBatch, responseParser);

        const result: SingleBatchRequestResult<P, B> = {
          processedOperationsThisAttempt: operationsInBatch,
          operationsToRetryNext: []
        };

        for (const op of operationsInBatch) {
          const errorInfo = op.processedItem.processingError;
          if (errorInfo?.details && typeof errorInfo.details === 'object' && 'statusCode' in errorInfo.details) {
            const subOpStatusCode = (errorInfo.details as any).statusCode as number;
            const isRetryableCode = masterRetryConfig.retryableStatusCodes.includes(subOpStatusCode);
            const underMaxRetries = op.retriesAttempted < masterRetryConfig.maxRetries;
            if (isRetryableCode && underMaxRetries) {
              op.retriesAttempted++;
              result.operationsToRetryNext.push(op);
            }
          }
        }
        return result;
      }),
      catchError(err => {
        console.error(`[UtilitiesService]${operationDescription}: HTTP POST for batch failed critically.`, err);
        operationsInBatch.forEach(op => {
          if (!op.processedItem.processingError) {
            op.processedItem.processingError = {
              message: `Batch POST failed: ${this.formatHttpError(err)}`,
              stage: 'Batch HTTP Request Critical Failure (UtilitiesService)',
              details: err instanceof HttpErrorResponse ? {statusCode: err.status, errorBody: err.error || err.message} : {errorBody: String(err)}
            };
          }
        });
        return of({processedOperationsThisAttempt: operationsInBatch, operationsToRetryNext: []});
      })
    );
    return this.retryRequest(request$, masterRetryConfig, operationDescription);
  }

  private parseMultipartResponse<P extends BaseProcessedItem, B>(
    rawResponseBody: string,
    contentTypeHeader: string,
    batchedOperations: BatchOperation<P, B>[],
    responseParser: BatchResponseParser<P>,
  ): void {
    const boundaryMatch = contentTypeHeader.match(/boundary=([^;]+)/);
    if (!boundaryMatch) {
      console.error('[Utils PARSE] No boundary in Content-Type header:', contentTypeHeader);
      batchedOperations.forEach(op => {
        if (!op.processedItem.processingError) op.processedItem.processingError = {
          message: "Batch parse error: no boundary in Content-Type", stage: "Batch Response Parsing (UtilitiesService)", details: {contentTypeHeader}
        };
      });
      return;
    }
    const fullBoundary = `--${boundaryMatch[1].trim()}`;
    const parts = rawResponseBody.split(new RegExp(`(?:\\r\\n)?${fullBoundary}(?:\\r\\n)?`));

    for (let i = 0; i < parts.length; i++) {
      const rawPartContent = parts[i];
      if (rawPartContent.trim() === '' || rawPartContent.trim() === '--') {
        continue;
      }

      const multipartPartHeaderEndIndex = rawPartContent.indexOf('\r\n\r\n');
      if (multipartPartHeaderEndIndex === -1) {
        console.warn(`[Utils PARSE WARN] Part ${i} has no multipart header/body separator (\\r\\n\\r\\n). Content preview:`, rawPartContent.substring(0, 300));
        const potentialOpIdMatch = rawPartContent.match(/Content-ID:\s*response-([^\s\r\n]+)/i);
        if (potentialOpIdMatch && potentialOpIdMatch[1]) {
          const op = batchedOperations.find(o => o.id === potentialOpIdMatch[1]);
          if (op && !op.processedItem.processingError) {
            op.processedItem.processingError = {message: "Malformed part: No header/body separator", stage: "Batch Response Parsing (UtilitiesService)", details: {opId: op.id, partIndex: i}};
          }
        }
        continue;
      }

      const multipartPartHeadersStr = rawPartContent.substring(0, multipartPartHeaderEndIndex).trim();
      const encapsulatedHttpResponseStr = rawPartContent.substring(multipartPartHeaderEndIndex + 4).trimStart();

      const contentIdHeaderRegex = /Content-ID:\s*response-([^\s\r\n<>]+)/i;
      const contentIdHeaderMatch = multipartPartHeadersStr.match(contentIdHeaderRegex);

      if (!contentIdHeaderMatch || contentIdHeaderMatch.length < 2) {
        console.warn(`[Utils PARSE WARN] Part ${i} could not extract 'response-OPERATION_ID' from Content-ID. Headers: "${multipartPartHeadersStr}"`);
        continue;
      }
      const originalOperationId = contentIdHeaderMatch[1].trim();
      const operation = batchedOperations.find(op => op.id === originalOperationId);

      if (!operation) {
        console.warn(`[Utils PARSE WARN] Op ID extracted as "${originalOperationId}" (from response part Content-ID) not found in current batch operations (IDs: ${batchedOperations.map(o => o.id).join(', ')}).`);
        console.warn(`[Utils PARSE DEBUG] Full headers for this failing part ${i}: "${multipartPartHeadersStr}"`);
        continue;
      }

      const itemTitleForLog = operation.processedItem.title ? `"${operation.processedItem.title.substring(0, 30)}..."` : "(No Title)";
      const itemLogPrefix = `[Utils BatchOp ID ${operation.id} ${itemTitleForLog}]:`;

      const encapsulatedHeaderBodySeparatorIndex = encapsulatedHttpResponseStr.indexOf('\r\n\r\n');
      let encapsulatedHeadersSection: string;
      let encapsulatedBodySection: string;

      if (encapsulatedHeaderBodySeparatorIndex === -1) {
        encapsulatedHeadersSection = encapsulatedHttpResponseStr;
        encapsulatedBodySection = "";
      } else {
        encapsulatedHeadersSection = encapsulatedHttpResponseStr.substring(0, encapsulatedHeaderBodySeparatorIndex);
        encapsulatedBodySection = encapsulatedHttpResponseStr.substring(encapsulatedHeaderBodySeparatorIndex + 4).trim();
      }

      const httpStatusLineMatch = encapsulatedHeadersSection.match(/^HTTP\/\d\.\d\s+(\d+)\s+([\s\S]*?)(?:\r\n|$)/im);

      if (!httpStatusLineMatch || httpStatusLineMatch.length < 2) {
        operation.processedItem.processingError = {
          message: "Malformed encapsulated sub-response: no HTTP status line found.",
          stage: "Batch Response Parsing (Encapsulated Status Line - UtilitiesService)",
          details: {opId: operation.id, headersPreview: encapsulatedHeadersSection.substring(0, 100)}
        };
        console.warn(`${itemLogPrefix} Error: ${operation.processedItem.processingError.message}`);
        continue;
      }

      const statusCode = parseInt(httpStatusLineMatch[1], 10);
      const statusText = (httpStatusLineMatch[2] || "").trim();
      let responseJson: any = {};

      try {
        if (encapsulatedBodySection) {
          responseJson = JSON.parse(encapsulatedBodySection);
        }
        responseParser(operation.processedItem, responseJson, statusCode, statusText, operation.id);

        if (statusCode >= 200 && statusCode < 300) {
          if (!operation.processedItem.processingError) {
            console.log(`${itemLogPrefix} Processed by parser as Success (Status ${statusCode}).`);
          } else {
            console.warn(`${itemLogPrefix} Processed by parser (Status ${statusCode}), but an error was set:`, operation.processedItem.processingError);
          }
        } else {
          if (!operation.processedItem.processingError) {
            const apiErrorMessage = responseJson.error?.message || statusText || 'Unknown error in sub-response.';
            operation.processedItem.processingError = {
              message: `Batch item failed: ${apiErrorMessage} (Status: ${statusCode})`,
              stage: 'Batch Item Error (UtilitiesService)',
              details: {statusCode: statusCode, errorBody: responseJson.error || responseJson, opId: operation.id}
            };
            console.warn(`${itemLogPrefix} Failed (Status ${statusCode}). API Error: "${apiErrorMessage}". Set generic error by Utils.`);
          } else {
            console.warn(`${itemLogPrefix} Failed (Status ${statusCode}). Error already set by parser:`, operation.processedItem.processingError);
          }
        }
      } catch (e: any) {
        console.error(`${itemLogPrefix} Exception during sub-response processing (JSON parse or parser callback) (Status ${statusCode}):`, e);
        operation.processedItem.processingError = {
          message: `Exception processing sub-response: ${e.message || 'Unknown exception'}`,
          stage: "Batch Sub-Response Processing Exception (UtilitiesService)",
          details: {statusCode: statusCode, statusText: statusText, responseBodyPreview: encapsulatedBodySection.substring(0, 100), opId: operation.id, exception: String(e)}
        };
      }
    }
  }

  public async dataUriToBlob(dataInput: string, expectedMimeType: string): Promise<Blob | null> {
    try {
      let base64Data: string;
      let mimeType: string = expectedMimeType || 'application/octet-stream';

      if (dataInput && dataInput.startsWith('data:')) {
        const splitDataURI = dataInput.split(',');
        if (splitDataURI.length < 2) {
          console.error("Invalid data URI format: missing comma separator.");
          return null;
        }
        const metaPart = splitDataURI[0];
        base64Data = splitDataURI[1];
        mimeType = metaPart.split(':')[1]?.split(';')[0] || mimeType;
      } else if (dataInput) {
        base64Data = dataInput;
      } else {
        console.error("Input data is empty. Cannot convert to Blob.");
        return null;
      }

      const fetchResponse = await fetch(`data:${mimeType};base64,${base64Data}`);
      if (!fetchResponse.ok) {
        throw new Error(`Failed to fetch base64 data as blob: ${fetchResponse.status} ${fetchResponse.statusText}`);
      }
      return await fetchResponse.blob();
    } catch (error) {
      console.error("Error converting data string to Blob:", error instanceof Error ? error.message : error);
      return null;
    }
  }

  public formatHttpError(error: HttpErrorResponse | Error | unknown): string {
    if (error instanceof HttpErrorResponse) {
      let errorMessage = `HTTP ${error.status} ${error.statusText || 'Error'}`;
      const googleError = error.error?.error;
      if (googleError && typeof googleError === 'object' && 'message' in googleError && googleError.message) {
        errorMessage += `: ${googleError.message}`;
        if ('details' in googleError && Array.isArray(googleError.details) && googleError.details.length > 0) {
          errorMessage += ` Details: ${JSON.stringify(googleError.details)}`;
        } else if ('status' in googleError && googleError.status) {
          errorMessage += ` Status: ${googleError.status}`;
        }
      } else if (typeof error.error === 'string') {
        errorMessage += `: ${error.error}`;
      } else if (error.message) {
        errorMessage += `: ${error.message}`;
      } else if (error.error) {
        errorMessage += `: ${JSON.stringify(error.error)}`;
      }
      return errorMessage;
    } else if (error instanceof Error) {
      return `Error: ${error.message}`;
    } else {
      return `Unknown error occurred: ${JSON.stringify(error)}`;
    }
  }

  public arrayBufferToBase64(buffer: ArrayBuffer): string {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
  }

  private readonly mimeMap: {[key: string]: string} = {txt: 'text/plain', html: 'text/html', htm: 'text/html', css: 'text/css', js: 'application/javascript', xml: 'text/xml', csv: 'text/csv', md: 'text/markdown', jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', ico: 'image/vnd.microsoft.icon', mp3: 'audio/mpeg', wav: 'audio/wav', ogg: 'audio/ogg', aac: 'audio/aac', mp4: 'video/mp4', webm: 'video/webm', avi: 'video/x-msvideo', mov: 'video/quicktime', pdf: 'application/pdf', doc: 'application/msword', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xls: 'application/vnd.ms-excel', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', ppt: 'application/vnd.ms-powerpoint', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', odt: 'application/vnd.oasis.opendocument.text', ods: 'application/vnd.oasis.opendocument.spreadsheet', odp: 'application/vnd.oasis.opendocument.presentation', zip: 'application/zip', rar: 'application/vnd.rar', '7z': 'application/x-7z-compressed', tar: 'application/x-tar', gz: 'application/gzip', json: 'application/json', rtf: 'application/rtf', };
  public getMimeTypeFromExtension(filename: string): string {const extension = filename.split('.').pop()?.toLowerCase(); return extension ? (this.mimeMap[extension] || 'application/octet-stream') : 'application/octet-stream';}


  public async generateHash(input: string): Promise<string> {
    try {
      const encoder = new TextEncoder();
      const data = encoder.encode(input);
      const hashBuffer = await crypto.subtle.digest('SHA-256', data);
      const hashArray = Array.from(new Uint8Array(hashBuffer));
      const hashBase64 = btoa(String.fromCharCode(...hashArray));
      return hashBase64;
    } catch (error) {
      console.error('Error generating SHA-256 hash:', error);
      throw new Error('Failed to generate file identifier hash.');
    }
  }

  public tryDecodeURIComponent(uriComponent: string): string {
    if (!uriComponent) return '';
    let decoded = uriComponent;
    try {
      for (let i = 0; i < 5; i++) {
        const previouslyDecoded = decoded;
        decoded = decodeURIComponent(decoded.replace(/\+/g, ' '));
        if (decoded === previouslyDecoded) break;
      }
    } catch (e) {
      try {
        return decode(uriComponent);
      } catch (decodeError) {
        return uriComponent;
      }
    }
    try {
      return decode(decoded);
    } catch (finalDecodeError) {
      return decoded;
    }
  }

  public getDirectory(pathStr: string | null | undefined): string {
    if (!pathStr) return "";
    const normalized = pathStr.trim().replace(/\\/g, '/');
    if (normalized.endsWith('/')) {
      return normalized.replace(/\/$/, '');
    }
    const lastSlash = normalized.lastIndexOf('/');
    if (lastSlash === -1) {
      return "";
    }
    return normalized.substring(0, lastSlash);
  }

  public resolveRelativePath(
    basePathInput: string | null | undefined,
    relativePathInput: string | null | undefined
  ): string | null {
    if (typeof relativePathInput !== 'string') return null;

    const decodedRelativePath = this.tryDecodeURIComponent(relativePathInput);
    const normRelativePath = decodedRelativePath.trim().replace(/\\/g, '/');
    const actualBaseDir = this.getDirectory(basePathInput);

    if (!normRelativePath) return actualBaseDir || null;

    if (normRelativePath.startsWith('/')) {
      const pathSegments = normRelativePath.substring(1).split('/');
      const resolvedParts: string[] = [];
      for (const part of pathSegments) {
        if (part === '..') {
          if (resolvedParts.length > 0) resolvedParts.pop();
        } else if (part !== '.' && part !== '') {
          resolvedParts.push(part);
        }
      }
      return resolvedParts.join('/');
    }

    const baseParts = actualBaseDir ? actualBaseDir.split('/').filter(p => p) : [];
    const relativeParts = normRelativePath.split('/');
    let combinedParts = [...baseParts];

    for (const part of relativeParts) {
      if (part === '..') {
        if (combinedParts.length > 0) {
          combinedParts.pop();
        }
      } else if (part !== '.' && part !== '') {
        combinedParts.push(part);
      }
    }
    return combinedParts.join('/');
  }

}
