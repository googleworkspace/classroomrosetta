/**
 * Copyright 2025 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

import {Injectable} from '@angular/core';
import {HttpErrorResponse} from '@angular/common/http';
import {Observable, throwError, timer} from 'rxjs';
import {retry} from 'rxjs/operators';
import {decode} from 'html-entities';

// Configuration interface for the retry logic (remains the same)
export interface RetryConfig {
  maxRetries?: number;
  initialDelayMs?: number;
  backoffFactor?: number;
  retryableStatusCodes?: number[];
}

@Injectable({
  providedIn: 'root'
})
export class UtilitiesService {

  // --- API Endpoints ---
  public readonly DRIVE_API_UPLOAD_ENDPOINT = 'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,appProperties,webViewLink,parents';
  public readonly DRIVE_API_FILES_ENDPOINT = 'https://www.googleapis.com/drive/v3/files';
  public readonly FORMS_API_CREATE_ENDPOINT = 'https://forms.googleapis.com/v1/forms';
  public readonly FORMS_API_BATCHUPDATE_BASE_ENDPOINT = 'https://forms.googleapis.com/v1/forms/'; // Add {formId}:batchUpdate
  public readonly FORMS_API_GET_BASE_ENDPOINT = 'https://forms.googleapis.com/v1/forms/'; // Add {formId}

  constructor() { }

  /**
   * Wraps an Observable (typically an HTTP request) with retry logic using the modern `retry` operator.
   * Implements exponential backoff for specific HTTP status codes.
   *
   * @param request$ The Observable representing the request to retry.
   * @param config Optional configuration for retry behavior.
   * @param operationName Optional name of the operation for logging purposes.
   * @returns Observable<T> The original Observable augmented with retry logic.
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
   * Converts a data URI string or raw base64 string into a Blob object.
   */
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
      console.error("Data start:", typeof dataInput === 'string' ? dataInput.substring(0, 100) + '...' : '[Not a string]');
      console.error("Expected MIME Type:", expectedMimeType);
      return null;
    }
  }

  /**
   * Helper to format HttpErrorResponse or other errors for better logging.
   */
  public formatHttpError(error: HttpErrorResponse | Error | unknown): string {
    if (error instanceof HttpErrorResponse) {
      let errorMessage = `HTTP ${error.status} ${error.statusText || 'Error'}`;
      const googleError = error.error?.error;
      if (googleError && typeof googleError === 'object') {
        errorMessage += `: ${googleError.message || JSON.stringify(googleError)}`;
        if (googleError.details && Array.isArray(googleError.details) && googleError.details.length > 0) {
          errorMessage += ` Details: ${JSON.stringify(googleError.details)}`;
        } else if (googleError.status) {
          errorMessage += ` Status: ${googleError.status}`;
        }
      } else if (typeof error.error === 'string') {
        errorMessage += `: ${error.error}`;
      } else if (error.message) {
        errorMessage += `: ${error.message}`;
      } else {
        errorMessage += `: ${JSON.stringify(error.error)}`;
      }
      return errorMessage;
    } else if (error instanceof Error) {
      return `Error: ${error.message}`;
    } else {
      return `Unknown error occurred: ${JSON.stringify(error)}`;
    }
  }

  /**
    * Helper function to convert ArrayBuffer to Base64 string.
    */
  public arrayBufferToBase64(buffer: ArrayBuffer): string {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    return window.btoa(binary);
  }

  // --- MIME Type Handling ---
  private readonly mimeMap: {[key: string]: string} = {txt: 'text/plain', html: 'text/html', htm: 'text/html', css: 'text/css', js: 'application/javascript', xml: 'text/xml', csv: 'text/csv', md: 'text/markdown', jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', ico: 'image/vnd.microsoft.icon', mp3: 'audio/mpeg', wav: 'audio/wav', ogg: 'audio/ogg', aac: 'audio/aac', mp4: 'video/mp4', webm: 'video/webm', avi: 'video/x-msvideo', mov: 'video/quicktime', pdf: 'application/pdf', doc: 'application/msword', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xls: 'application/vnd.ms-excel', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', ppt: 'application/vnd.ms-powerpoint', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', odt: 'application/vnd.oasis.opendocument.text', ods: 'application/vnd.oasis.opendocument.spreadsheet', odp: 'application/vnd.oasis.opendocument.presentation', zip: 'application/zip', rar: 'application/vnd.rar', '7z': 'application/x-7z-compressed', tar: 'application/x-tar', gz: 'application/gzip', json: 'application/json', rtf: 'application/rtf', };
  public getMimeTypeFromExtension(filename: string): string {const extension = filename.split('.').pop()?.toLowerCase(); return extension ? (this.mimeMap[extension] || 'application/octet-stream') : 'application/octet-stream';}


  /**
   * Generates a SHA-256 hash of the input string and returns it as a Base64 string.
   */
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

  /**
   * Attempts to decode a URI component multiple times for robustness.
   * Also uses html-entities library as a final fallback for broader character support.
   * @param uriComponent The URI component string to decode.
   * @returns The decoded string, or the original if decoding fails.
   */
  public tryDecodeURIComponent(uriComponent: string): string {
    if (!uriComponent) return '';
    let decoded = uriComponent;
    try {
      for (let i = 0; i < 5; i++) { // Try decoding multiple times
        const previouslyDecoded = decoded;
        decoded = decodeURIComponent(decoded.replace(/\+/g, ' '));
        if (decoded === previouslyDecoded) break;
      }
    } catch (e) {
      // If decodeURIComponent throws (e.g. malformed URI), fall back to html-entities decode on the original.
      // This is often more forgiving for strings that aren't strictly URI encoded.
      try {
        return decode(uriComponent);
      } catch (decodeError) {
        console.warn(`UtilitiesService: Failed to decode URI component with decodeURIComponent and html-entities decode: "${uriComponent}". Error: ${e}, DecodeError: ${decodeError}. Returning original.`);
        return uriComponent; // Return original if all decoding fails
      }
    }
    // Final html-entities decode on the (potentially) URI-decoded string.
    try {
      return decode(decoded);
    } catch (finalDecodeError) {
      console.warn(`UtilitiesService: Failed final html-entities decode for component: "${decoded}". Error: ${finalDecodeError}. Returning as is.`);
      return decoded;
    }
  }

  /**
   * Extracts the directory path from a given file or directory path string.
   * @param pathStr The input path string.
   * @returns The directory path, without a trailing slash. Returns "" for root.
   */
  public getDirectory(pathStr: string | null | undefined): string {
    if (!pathStr) return "";
    const normalized = pathStr.trim().replace(/\\/g, '/'); // Normalize to forward slashes
    if (normalized.endsWith('/')) {
      // If it ends with a slash, it's already a directory path. Remove trailing slash.
      return normalized.replace(/\/$/, '');
    }
    // If it doesn't end with a slash, it could be a file or a directory name without a slash.
    // Find the last slash to get the parent directory.
    const lastSlash = normalized.lastIndexOf('/');
    if (lastSlash === -1) {
      // No slash found, means it's in the root. Parent is root ("").
      return "";
    }
    // "foo/bar/file.txt" -> "foo/bar"
    return normalized.substring(0, lastSlash);
  }

  /**
   * Resolves a relative path against a base path (which can be a file or directory).
   * @param basePathInput The base path, which can be a full file path or a directory path.
   * @param relativePathInput The relative path (can be URI encoded or not).
   * @returns The resolved absolute path from the package root, or null.
   */
  public resolveRelativePath(
    basePathInput: string | null | undefined,
    relativePathInput: string | null | undefined
  ): string | null {
    if (typeof relativePathInput !== 'string') return null;

    // First, URI decode the relative path input
    const decodedRelativePath = this.tryDecodeURIComponent(relativePathInput);

    const normRelativePath = decodedRelativePath.trim().replace(/\\/g, '/');
    const actualBaseDir = this.getDirectory(basePathInput); // Get directory from base (e.g. "a/b/file.html" -> "a/b")

    if (!normRelativePath) return actualBaseDir || null;

    // If relative path starts with '/', it's an absolute path from the package root.
    // The actualBaseDir is ignored in this case.
    if (normRelativePath.startsWith('/')) {
      // Normalize an absolute path by removing leading slash for joining, then re-evaluate.
      // This handles cases like "/foo.txt" or "/dir/foo.txt"
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

    // For relative paths not starting with '/':
    const baseParts = actualBaseDir ? actualBaseDir.split('/').filter(p => p) : []; // Filter out empty strings from split
    const relativeParts = normRelativePath.split('/');
    let combinedParts = [...baseParts];

    for (const part of relativeParts) {
      if (part === '..') {
        if (combinedParts.length > 0) {
          combinedParts.pop();
        }
        // If trying to go above root (e.g. "../../file.txt" from root), combinedParts will be empty.
        // Depending on desired behavior, this could be an error or resolve to root.
        // For IMSCC, it usually means the path is invalid if it tries to go above package root.
        // Current logic will result in path relative to root if it goes above.
      } else if (part !== '.' && part !== '') {
        combinedParts.push(part);
      }
    }
    return combinedParts.join('/');
  }
}
