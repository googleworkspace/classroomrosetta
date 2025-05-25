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
import {HttpClient, HttpHeaders, HttpParams} from '@angular/common/http';
import {Observable, throwError, of, from} from 'rxjs';
import {switchMap, catchError, tap} from 'rxjs/operators';
import {DriveFile} from '../../interfaces/classroom-interface';
import {UtilitiesService, RetryConfig} from '../utilities/utilities.service';

@Injectable({
  providedIn: 'root'
})
export class HtmlToDocsService {

  // Inject dependencies
  private http = inject(HttpClient);
  private utils = inject(UtilitiesService);

  // Custom property key to store the source identifier HASH
  private readonly APP_PROPERTY_KEY = 'imsccIdentifier';

  constructor() { }

  /**
   * Finds an existing Google Doc based on a HASH of an identifier in appProperties,
   * or creates a new one by uploading an HTML string to Google Drive.
   * Includes retry logic for API calls.
   *
   * @param htmlContent The HTML string to convert IF creating a new doc.
   * @param documentTitle The desired title for the new Google Doc IF creating one.
   * @param accessToken A valid Google OAuth 2.0 access token with Drive scope.
   * @param itemId The unique identifier (e.g., IMSCC item ID) to search for or associate with the file. THIS WILL BE HASHED.
   * @param parentFolderId Optional. The ID of the Google Drive folder to place the *newly created* document into.
   * @returns Observable<DriveFile> Emitting the Google Drive File resource object (either found or newly created).
   */
  createDocFromHtml(
    htmlContent: string,
    documentTitle: string,
    accessToken: string,
    itemId: string,
    parentFolderId?: string
  ): Observable<DriveFile> {
    // --- Input Validation ---
    if (!itemId) return throwError(() => new Error('Item ID (itemId) is required (will be hashed).'));
    if (!documentTitle) return throwError(() => new Error('Document title cannot be empty (required for creation).'));
    if (!accessToken) return throwError(() => new Error('Access token is required.'));

    // --- Prepare Headers ---
    const headers = new HttpHeaders({
      'Authorization': `Bearer ${accessToken}`
    });

    // Define retry configuration
    const retryConfig: RetryConfig = {maxRetries: 3, initialDelayMs: 2000};

    // --- Step 1: Generate Hash Identifier from itemId ---
    console.log(`[HtmlToDocsService] Generating hash for itemId: ${itemId}`);
    return from(this.utils.generateHash(itemId)).pipe(
      catchError(hashError => {
        console.error(`[HtmlToDocsService] Error generating hash for itemId "${itemId}":`, hashError);
        const errorMessage = hashError instanceof Error ? hashError.message : String(hashError);
        return throwError(() => new Error(`Failed to generate identifier hash for ${itemId}. ${errorMessage}`));
      }),
      switchMap(hashedItemId => {
      // --- Step 2: Search for Existing File by HASHED appProperties (with Retry) ---
        // The API query already includes 'trashed = false'.
        console.log(`[HtmlToDocsService] Searching for existing Drive file with ${this.APP_PROPERTY_KEY}=${hashedItemId} (hash of ${itemId})...`);
        const searchQuery = `appProperties has { key='${this.APP_PROPERTY_KEY}' and value='${hashedItemId}' } and mimeType='application/vnd.google-apps.document' and trashed = false`;
        const searchParams = new HttpParams()
          .set('q', searchQuery)
          .set('fields', 'files(id, name, mimeType, appProperties, webViewLink, parents, trashed)') // Ensure 'trashed' is in fields
          .set('spaces', 'drive');

        const searchRequest$ = this.http.get<{files: DriveFile[]}>(
          this.utils.DRIVE_API_FILES_ENDPOINT, {headers, params: searchParams}
        );

        return this.utils.retryRequest(
          searchRequest$,
          retryConfig,
          `Search Doc by Hash ${hashedItemId}`
        ).pipe(
          switchMap(response => {
            // --- Step 3: Check Search Results and explicitly filter for non-trashed files ---
            const suitableFiles = response.files ? response.files.filter(f => f.trashed !== true) : [];

            if (suitableFiles.length > 0) {
              // --- File Found & Not Trashed ---
              const foundFile = suitableFiles[0];
              console.log(`[HtmlToDocsService] Found existing non-trashed Google Doc for hashed item ${hashedItemId}: ID=${foundFile.id}, Name=${foundFile.name}.`);
              return of(foundFile);
            } else {
              // --- File Not Found or All Found Files Were Trashed: Proceed with Creation (with Retry) ---
              if (response.files && response.files.length > 0) { // Files were found by API, but our client-side filter removed them
                console.warn(`[HtmlToDocsService] Files were found for hashed item ${hashedItemId}, but all were trashed or unsuitable. Proceeding to create a new document.`);
              } else {
                console.log(`[HtmlToDocsService] No existing Google Doc found for hashed item ${hashedItemId}. Creating new Google Doc via Drive import for "${documentTitle}"...`);
              }

              if (!htmlContent) {
                return throwError(() => new Error('HTML content is required to create a new document.'));
              }

              const metadata = {
                name: documentTitle,
                mimeType: 'application/vnd.google-apps.document',
                appProperties: {[this.APP_PROPERTY_KEY]: hashedItemId},
                parents: parentFolderId ? [parentFolderId] : undefined
              };
              const metadataBlob = new Blob([JSON.stringify(metadata)], {type: 'application/json'});
              const htmlBlob = new Blob([htmlContent], {type: 'text/html'});
              const formData = new FormData();
              formData.append('metadata', metadataBlob);
              formData.append('file', htmlBlob);

              const createRequest$ = this.http.post<DriveFile>(
                this.utils.DRIVE_API_UPLOAD_ENDPOINT, formData, {headers}
              );

              return this.utils.retryRequest(
                createRequest$,
                retryConfig,
                `Create Doc "${documentTitle}"`
              ).pipe(
                tap(newlyCreatedFile => {
                  if (!newlyCreatedFile?.id) throw new Error('Failed to create/convert document in Google Drive (unexpected response).');
                  console.log(`[HtmlToDocsService] Successfully created Google Doc via Drive import. ID: ${newlyCreatedFile.id}, Name: ${newlyCreatedFile.name}`);
                  console.log(`   Associated ${this.APP_PROPERTY_KEY}: ${newlyCreatedFile.appProperties?.[this.APP_PROPERTY_KEY]} (Hash of: "${itemId}")`);
                  if (parentFolderId && !newlyCreatedFile.parents?.includes(parentFolderId)) {
                    console.warn(`[HtmlToDocsService] Document ${newlyCreatedFile.id} created, but might not be in the target folder ${parentFolderId} yet.`);
                  }
                }),
                catchError(err => {
                  const formattedError = this.utils.formatHttpError(err);
                  console.error(`[HtmlToDocsService] Error creating Google Doc from HTML for hashed item ${hashedItemId} (final after retries):`, formattedError);
                  return throwError(() => new Error(`Failed to create Google Doc for ${documentTitle}. ${formattedError}`));
                })
              );
            }
          }),
          catchError(err => {
            const formattedError = this.utils.formatHttpError(err);
            console.error(`[HtmlToDocsService] Error searching for Google Doc for hashed item ${hashedItemId} (final after retries):`, formattedError);
            return throwError(() => new Error(`Drive search failed for ${itemId}. ${formattedError}`));
          })
        );
      })
    );
  }
}
