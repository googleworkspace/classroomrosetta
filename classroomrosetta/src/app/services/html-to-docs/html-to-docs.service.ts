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

import { Injectable, inject } from '@angular/core';
import { HttpClient, HttpHeaders, HttpErrorResponse, HttpParams } from '@angular/common/http';
import { Observable, throwError, of, from } from 'rxjs';
import {switchMap, catchError, tap, map} from 'rxjs/operators';
import { DriveFile } from '../../interfaces/classroom-interface'; // Adjust path
import { UtilitiesService, RetryConfig } from '../utilities/utilities.service'; // Adjust path and import RetryConfig

@Injectable({
  providedIn: 'root'
})
export class HtmlToDocsService {

  // Inject dependencies
  private http = inject(HttpClient);
  private utils = inject(UtilitiesService); // Assumes UtilitiesService provides generateHash, formatHttpError, API endpoints, and retryRequest

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
      itemId: string, // This will be hashed
      parentFolderId?: string
    // createAsPageless parameter removed
  ): Observable<DriveFile> {
      // --- Input Validation ---
      if (!itemId) return throwError(() => new Error('Item ID (itemId) is required (will be hashed).'));
    // htmlContent check moved to later, after search response, to allow searching without it.
      if (!documentTitle) return throwError(() => new Error('Document title cannot be empty (required for creation).'));
      if (!accessToken) return throwError(() => new Error('Access token is required.'));

      // --- Prepare Headers ---
      const headers = new HttpHeaders({
        'Authorization': `Bearer ${accessToken}`
      });

      // Define retry configuration (can be customized)
      const retryConfig: RetryConfig = { maxRetries: 3, initialDelayMs: 2000 };

      // --- Step 1: Generate Hash Identifier from itemId ---
      console.log(`Generating hash for itemId: ${itemId}`);
      return from(this.utils.generateHash(itemId)).pipe(
        catchError(hashError => {
              console.error(`Error generating hash for itemId "${itemId}":`, hashError);
              const errorMessage = hashError instanceof Error ? hashError.message : String(hashError);
              return throwError(() => new Error(`Failed to generate identifier hash for ${itemId}. ${errorMessage}`));
          }),
          switchMap(hashedItemId => {
              // --- Step 2: Search for Existing File by HASHED appProperties (with Retry) ---
              console.log(`Searching for existing Drive file with ${this.APP_PROPERTY_KEY}=${hashedItemId} (hash of ${itemId})...`);
              const searchQuery = `appProperties has { key='${this.APP_PROPERTY_KEY}' and value='${hashedItemId}' } and mimeType='application/vnd.google-apps.document' and trashed = false`;
              const searchParams = new HttpParams()
                .set('q', searchQuery)
                .set('fields', 'files(id, name, mimeType, appProperties, webViewLink, parents)')
                .set('spaces', 'drive');

              const searchRequest$ = this.http.get<{ files: DriveFile[] }>(
                  this.utils.DRIVE_API_FILES_ENDPOINT, { headers, params: searchParams }
              );

              return this.utils.retryRequest(
                searchRequest$,
                retryConfig,
                `Search Doc by Hash ${hashedItemId}`
              ).pipe(
                switchMap(response => {
                  // --- Step 3: Check Search Results ---
                  if (response.files && response.files.length > 0) {
                    // --- File Found ---
                    const foundFile = response.files[0];
                    console.log(`Found existing non-trashed Google Doc for hashed item ${hashedItemId}: ID=${foundFile.id}, Name=${foundFile.name}.`);
                    // Logic for setting pageless mode on existing file removed.
                    return of(foundFile);
                  } else {
                    // --- File Not Found: Proceed with Creation (with Retry) ---
                    console.log(`No existing Google Doc found for hashed item ${hashedItemId}. Creating new Google Doc via Drive import for "${documentTitle}"...`);

                    if (!htmlContent) { // Check htmlContent only if we are certain we need to create a new doc
                      return throwError(() => new Error('HTML content is required to create a new document.'));
                    }

                    const metadata = {
                      name: documentTitle,
                      mimeType: 'application/vnd.google-apps.document',
                      appProperties: { [this.APP_PROPERTY_KEY]: hashedItemId },
                      parents: parentFolderId ? [parentFolderId] : undefined
                    };
                    const metadataBlob = new Blob([JSON.stringify(metadata)], { type: 'application/json' });
                    const htmlBlob = new Blob([htmlContent], { type: 'text/html' });
                    const formData = new FormData();
                    formData.append('metadata', metadataBlob);
                    formData.append('file', htmlBlob);

                    const createRequest$ = this.http.post<DriveFile>(
                        this.utils.DRIVE_API_UPLOAD_ENDPOINT, formData, { headers }
                    );

                    return this.utils.retryRequest(
                      createRequest$,
                      retryConfig,
                      `Create Doc "${documentTitle}"`
                    ).pipe(
                      tap(newlyCreatedFile => {
                        if (!newlyCreatedFile?.id) throw new Error('Failed to create/convert document in Google Drive (unexpected response).');
                        console.log(`Successfully created Google Doc via Drive import. ID: ${newlyCreatedFile.id}, Name: ${newlyCreatedFile.name}`);
                        console.log(`   Associated ${this.APP_PROPERTY_KEY}: ${newlyCreatedFile.appProperties?.[this.APP_PROPERTY_KEY]} (Hash of: "${itemId}")`);
                        if (parentFolderId && !newlyCreatedFile.parents?.includes(parentFolderId)) {
                            console.warn(`Document ${newlyCreatedFile.id} created, but might not be in the target folder ${parentFolderId} yet.`);
                        }
                      }),
                      // Logic for setting pageless mode on newly created file removed.
                      // The newlyCreatedFile is returned directly.
                      catchError(err => {
                          const formattedError = this.utils.formatHttpError(err);
                        console.error(`Error creating Google Doc from HTML for hashed item ${hashedItemId} (final after retries):`, formattedError);
                          return throwError(() => new Error(`Failed to create Google Doc for ${documentTitle}. ${formattedError}`));
                      })
                    );
                  }
                }),
                catchError(err => {
                  const formattedError = this.utils.formatHttpError(err);
                  console.error(`Error searching for Google Doc for hashed item ${hashedItemId} (final after retries):`, formattedError);
                  return throwError(() => new Error(`Drive search failed for ${itemId}. ${formattedError}`));
                })
            );
          })
      );
  }
}
