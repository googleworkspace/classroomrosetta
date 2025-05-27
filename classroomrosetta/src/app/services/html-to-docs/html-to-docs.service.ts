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

import {Injectable, inject} from '@angular/core';
import {HttpClient, HttpHeaders, HttpParams, HttpErrorResponse} from '@angular/common/http';
import {Observable, throwError, of, from} from 'rxjs';
import {switchMap, catchError, tap} from 'rxjs/operators';
import {DriveFile} from '../../interfaces/classroom-interface';
import {UtilitiesService, RetryConfig} from '../utilities/utilities.service';
import {AuthService} from '../auth/auth.service';

@Injectable({
  providedIn: 'root'
})
export class HtmlToDocsService {

  // Inject dependencies
  private http = inject(HttpClient);
  private utils = inject(UtilitiesService);
  private auth = inject(AuthService);

  // Custom property key to store the source identifier HASH
  private readonly APP_PROPERTY_KEY = 'imsccIdentifier';

  constructor() { }

  /**
   * Creates Drive API headers, fetching token internally.
   * @param isUpload Indicates if headers are for a multipart upload (no explicit Content-Type).
   * @returns HttpHeaders object or null if token is missing.
   */
  private createDriveApiHeaders(isUpload: boolean = false): HttpHeaders | null {
    const accessToken = this.auth.getGoogleAccessToken();
    if (!accessToken) {
      console.error('[HtmlToDocsService] Cannot create Drive API headers: Access token is missing.');
      return null;
    }
    let headersConfig: {[header: string]: string | string[]} = {
      'Authorization': `Bearer ${accessToken}`
    };
    // For multipart uploads (FormData), HttpClient sets Content-Type.
    // For JSON metadata POST (like creating an empty folder, not used here directly for doc creation from HTML),
    // 'Content-Type': 'application/json' would be needed.
    // For GET, 'Accept: application/json' is good practice but often default.
    // This service primarily does GET for search and POST FormData for upload.
    if (!isUpload) { // For GET search or if we were doing a JSON POST
      headersConfig['Accept'] = 'application/json'; // Good practice for GET
      // If creating doc via metadata-only POST (not the case here), would add:
      headersConfig['Content-Type'] = 'application/json';
    }

    return new HttpHeaders(headersConfig);
  }


  /**
   * Finds an existing Google Doc based on a HASH of an identifier in appProperties,
   * or creates a new one by uploading an HTML string to Google Drive.
   * Token is fetched internally.
   *
   * @param htmlContent The HTML string to convert IF creating a new doc.
   * @param documentTitle The desired title for the new Google Doc IF creating one.
   * @param itemId The unique identifier (e.g., IMSCC item ID) to search for or associate with the file. THIS WILL BE HASHED.
   * @param parentFolderId Optional. The ID of the Google Drive folder to place the *newly created* document into.
   * @returns Observable<DriveFile> Emitting the Google Drive File resource object (either found or newly created).
   */
  createDocFromHtml(
    htmlContent: string,
    documentTitle: string,
    itemId: string,
    parentFolderId?: string
  ): Observable<DriveFile> {
    // --- Input Validation ---
    if (!itemId) return throwError(() => new Error('[HtmlToDocsService] Item ID (itemId) is required (will be hashed).'));
    if (!documentTitle) return throwError(() => new Error('[HtmlToDocsService] Document title cannot be empty (required for creation).'));

    // Define retry configuration
    const retryConfig: RetryConfig = {maxRetries: 3, initialDelayMs: 2000};

    // --- Step 1: Generate Hash Identifier from itemId ---
    console.log(`[HtmlToDocsService] Generating hash for itemId: ${itemId}`);
    return from(this.utils.generateHash(itemId)).pipe(
      catchError(hashError => {
        console.error(`[HtmlToDocsService] Error generating hash for itemId "${itemId}":`, hashError);
        const errorMessage = hashError instanceof Error ? hashError.message : String(hashError);
        return throwError(() => new Error(`[HtmlToDocsService] Failed to generate identifier hash for ${itemId}. ${errorMessage}`));
      }),
      switchMap(hashedItemId => {
        // --- Step 2: Search for Existing File by HASHED appProperties (with Retry) ---
        const searchHeaders = this.createDriveApiHeaders();
        if (!searchHeaders) {
          return throwError(() => new Error(`[HtmlToDocsService] Authentication token missing for searching doc by hash: ${hashedItemId}`));
        }

        console.log(`[HtmlToDocsService] Searching for existing Drive file with ${this.APP_PROPERTY_KEY}=${hashedItemId} (hash of ${itemId})...`);
        const searchQuery = `appProperties has { key='${this.APP_PROPERTY_KEY}' and value='${this.utils.escapeQueryParam(hashedItemId)}' } and mimeType='application/vnd.google-apps.document' and trashed = false`;
        const searchParams = new HttpParams()
          .set('q', searchQuery)
          .set('fields', 'files(id, name, mimeType, appProperties, webViewLink, parents, trashed)')
          .set('spaces', 'drive');

        const searchRequest$ = this.http.get<{files: DriveFile[]}>(
          this.utils.DRIVE_API_FILES_ENDPOINT, {headers: searchHeaders, params: searchParams}
        );

        return this.utils.retryRequest(
          searchRequest$,
          retryConfig,
          `Search Doc by Hash ${hashedItemId}`
        ).pipe(
          switchMap(response => {
            const suitableFiles = response.files ? response.files.filter(f => f.trashed !== true) : [];

            if (suitableFiles.length > 0) {
              const foundFile = suitableFiles[0];
              console.log(`[HtmlToDocsService] Found existing non-trashed Google Doc for hashed item ${hashedItemId}: ID=${foundFile.id}, Name=${foundFile.name}.`);
              return of(foundFile);
            } else {
              if (response.files && response.files.length > 0) {
                console.warn(`[HtmlToDocsService] Files were found for hashed item ${hashedItemId}, but all were trashed or unsuitable. Proceeding to create a new document.`);
              } else {
                console.log(`[HtmlToDocsService] No existing Google Doc found for hashed item ${hashedItemId}. Creating new Google Doc via Drive import for "${documentTitle}"...`);
              }

              if (!htmlContent && htmlContent !== "") { // Allow empty string for an empty doc, but not null/undefined
                return throwError(() => new Error('[HtmlToDocsService] HTML content is required to create a new document if one is not found.'));
              }

              // Fetch token for creation
              const createHeaders = this.createDriveApiHeaders(true); // true for upload (FormData)
              if (!createHeaders) {
                return throwError(() => new Error(`[HtmlToDocsService] Authentication token missing for creating doc: ${documentTitle}`));
              }

              const metadata = {
                name: documentTitle,
                mimeType: 'application/vnd.google-apps.document', // Important: This tells Drive to convert HTML to Google Doc
                appProperties: {[this.APP_PROPERTY_KEY]: hashedItemId},
                parents: parentFolderId ? [parentFolderId] : undefined
              };
              const metadataBlob = new Blob([JSON.stringify(metadata)], {type: 'application/json'});
              const htmlBlob = new Blob([htmlContent], {type: 'text/html'});
              const formData = new FormData();
              formData.append('metadata', metadataBlob);
              formData.append('file', htmlBlob, `${documentTitle}.html`); // Provide a filename for the HTML part

              const uploadUrlWithFields = `${this.utils.DRIVE_API_UPLOAD_ENDPOINT}${this.utils.DRIVE_API_UPLOAD_ENDPOINT.includes('?') ? '&' : '?'}fields=${encodeURIComponent('id,name,mimeType,appProperties,webViewLink,parents')}`;

              const createRequest$ = this.http.post<DriveFile>(
                uploadUrlWithFields, // Use the upload endpoint
                formData,
                {headers: createHeaders} // HttpClient handles Content-Type for FormData
              );

              return this.utils.retryRequest(
                createRequest$,
                retryConfig,
                `Create Doc "${documentTitle}"`
              ).pipe(
                tap(newlyCreatedFile => {
                  if (!newlyCreatedFile?.id) throw new Error('[HtmlToDocsService] Failed to create/convert document in Google Drive (unexpected response).');
                  console.log(`[HtmlToDocsService] Successfully created Google Doc via Drive import. ID: ${newlyCreatedFile.id}, Name: ${newlyCreatedFile.name}`);
                  console.log(`   Associated ${this.APP_PROPERTY_KEY}: ${newlyCreatedFile.appProperties?.[this.APP_PROPERTY_KEY]} (Hash of: "${itemId}")`);
                  if (parentFolderId && !newlyCreatedFile.parents?.includes(parentFolderId)) {
                    console.warn(`[HtmlToDocsService] Document ${newlyCreatedFile.id} created, but parent folder might not be set as expected immediately. Expected: ${parentFolderId}, Actual: ${newlyCreatedFile.parents}`);
                  }
                }),
                catchError(err => {
                  const formattedError = this.utils.formatHttpError(err);
                  console.error(`[HtmlToDocsService] Error creating Google Doc from HTML for hashed item ${hashedItemId} (final after retries):`, formattedError);
                  return throwError(() => new Error(`[HtmlToDocsService] Failed to create Google Doc for ${documentTitle}. ${formattedError}`));
                })
              );
            }
          }),
          catchError(err => {
            const formattedError = this.utils.formatHttpError(err);
            console.error(`[HtmlToDocsService] Error searching for Google Doc for hashed item ${hashedItemId} (final after retries):`, formattedError);
            return throwError(() => new Error(`[HtmlToDocsService] Drive search failed for ${itemId}. ${formattedError}`));
          })
        );
      })
    );
  }
}
