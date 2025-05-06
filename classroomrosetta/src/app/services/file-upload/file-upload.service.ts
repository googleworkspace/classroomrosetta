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
import { Observable, throwError, of, forkJoin, from } from 'rxjs';
import { switchMap, catchError, map, tap } from 'rxjs/operators';
import { ImsccFile, DriveFile } from '../../interfaces/classroom-interface'; // Adjust path
// Import RetryConfig from UtilitiesService
import { UtilitiesService, RetryConfig } from '../utilities/utilities.service'; // Adjust path

@Injectable({
  providedIn: 'root'
})
export class FileUploadService {

  // Inject dependencies
  private http = inject(HttpClient);
  private utils = inject(UtilitiesService);

  // Custom property key to store the source identifier HASH
  private readonly APP_PROPERTY_KEY = 'imsccIdentifier';
  // Define the fields to be requested from the Drive API for file metadata
  private readonly DRIVE_FILE_FIELDS_TO_REQUEST = 'id,name,mimeType,appProperties,webViewLink,parents,thumbnailLink';


  constructor() { }

  /**
   * Uploads an array of local files (ImsccFile) to a Google Drive folder using HASHED identifiers.
   * Includes retry logic for Drive API calls.
   * Ensures specified fields (including thumbnailLink) are requested in API responses.
   *
   * @param localFilesToUpload Array of objects containing the ImsccFile and the target file name.
   * @param accessToken A valid Google OAuth 2.0 access token with Drive scope.
   * @param parentFolderId The ID of the Google Drive folder where files should be uploaded/searched.
   * @returns Observable<DriveFile[]> Emitting an array of the Google Drive File resource objects (found or uploaded).
   */
  uploadLocalFiles(
    localFilesToUpload: Array<{file: ImsccFile; targetFileName: string}>,
    accessToken: string,
    parentFolderId: string
  ): Observable<DriveFile[]> {
      if (!localFilesToUpload || localFilesToUpload.length === 0) {
        console.log("[FileUploadService] No local files provided for upload.");
        return of([]);
      }
    if (!accessToken) return throwError(() => new Error('[FileUploadService] Access token is required for file upload.'));
    if (!parentFolderId) return throwError(() => new Error('[FileUploadService] Parent folder ID is required for file upload.'));

      const headers = new HttpHeaders({
        'Authorization': `Bearer ${accessToken}`
      });
      const retryConfig: RetryConfig = { maxRetries: 3, initialDelayMs: 2000 };

    const findOrUploadObservables: Observable<DriveFile>[] = localFilesToUpload.map(fileToUpload => {
        const originalFileIdentifier = fileToUpload.file.name;
        if (!originalFileIdentifier) {
          console.error(`[FileUploadService] Skipping file "${fileToUpload.targetFileName}": Missing original file name (identifier).`);
          return throwError(() => new Error(`[FileUploadService] Missing identifier for file: ${fileToUpload.targetFileName}`));
        }

        return from(this.utils.generateHash(originalFileIdentifier)).pipe(
          catchError(hashError => {
            console.error(`[FileUploadService] Error generating hash for "${originalFileIdentifier}":`, hashError);
              const errorMessage = hashError instanceof Error ? hashError.message : String(hashError);
            return throwError(() => new Error(`[FileUploadService] Failed to generate identifier for ${fileToUpload.targetFileName}. ${errorMessage}`));
          }),
          switchMap(hashedIdentifier => {
            console.log(`[FileUploadService] Searching for existing file with hash "${hashedIdentifier}" (from "${originalFileIdentifier}") in folder ${parentFolderId}...`);
            const searchQuery = `'${parentFolderId}' in parents and appProperties has { key='${this.APP_PROPERTY_KEY}' and value='${hashedIdentifier}' } and trashed = false`;
            const searchParams = new HttpParams()
              .set('q', searchQuery)
              .set('fields', `files(${this.DRIVE_FILE_FIELDS_TO_REQUEST})`)
              .set('spaces', 'drive');

            const searchRequest$ = this.http.get<{ files: DriveFile[] }>(
                this.utils.DRIVE_API_FILES_ENDPOINT, { headers, params: searchParams }
            );

            return this.utils.retryRequest(
              searchRequest$,
              retryConfig,
              `Search File by Hash ${hashedIdentifier}`
            ).pipe(
              switchMap(response => {
                if (response.files && response.files.length > 0) {
                  const foundFile = response.files[0];
                  console.log(`[FileUploadService] Found existing file for hash "${hashedIdentifier}": ID=${foundFile.id}, Name=${foundFile.name}. Thumbnail: ${foundFile.thumbnailLink ? 'Available' : 'Not Available'}. Skipping upload.`);
                  return of(foundFile);
                } else {
                  console.log(`[FileUploadService] No existing file found for hash "${hashedIdentifier}". Uploading "${fileToUpload.targetFileName}"...`);

                  let blobPreparation$: Observable<Blob | null>; // Changed to allow null for error case propagation

                  if (typeof fileToUpload.file.data === 'string' && fileToUpload.file.data.startsWith('data:')) {
                    // It's a data URI (likely a base64 image from FileUploadComponent)
                    console.log(`[FileUploadService] Preparing blob from data URI for: ${fileToUpload.targetFileName}`);
                    blobPreparation$ = from(this.utils.dataUriToBlob(fileToUpload.file.data, fileToUpload.file.mimeType));
                  } else if (fileToUpload.file.data instanceof ArrayBuffer) {
                    // It's an ArrayBuffer (other binary file)
                    console.log(`[FileUploadService] Preparing blob directly from ArrayBuffer for: ${fileToUpload.targetFileName}`);
                    try {
                      blobPreparation$ = of(new Blob([fileToUpload.file.data], {type: fileToUpload.file.mimeType}));
                    } catch (e: any) {
                      console.error(`[FileUploadService] Error creating Blob from ArrayBuffer for ${fileToUpload.targetFileName}:`, e);
                      blobPreparation$ = throwError(() => new Error(`Failed to create Blob from ArrayBuffer for ${fileToUpload.targetFileName}: ${e.message}`));
                    }
                  } else if (typeof fileToUpload.file.data === 'string') {
                    // It's a plain string (e.g., text, XML, HTML)
                    console.log(`[FileUploadService] Preparing blob from plain string data for: ${fileToUpload.targetFileName}`);
                    try {
                      blobPreparation$ = of(new Blob([fileToUpload.file.data], {type: fileToUpload.file.mimeType}));
                    } catch (e: any) {
                      console.error(`[FileUploadService] Error creating Blob from string for ${fileToUpload.targetFileName}:`, e);
                      blobPreparation$ = throwError(() => new Error(`Failed to create Blob from string for ${fileToUpload.targetFileName}: ${e.message}`));
                    }
                  } else {
                    console.error(`[FileUploadService] Unexpected data type for file ${fileToUpload.targetFileName}. Cannot prepare blob.`);
                    blobPreparation$ = throwError(() => new Error(`Unexpected data type for file: ${fileToUpload.targetFileName}`));
                  }

                  return blobPreparation$.pipe(
                    switchMap(fileBlob => {
                      if (!fileBlob) {
                        console.error(`[FileUploadService] Skipping upload for "${fileToUpload.targetFileName}": Blob preparation resulted in null or error was not caught properly.`);
                        return throwError(() => new Error(`Blob preparation failed for file: ${fileToUpload.targetFileName}`));
                      }
                      console.log(`[FileUploadService] Uploading blob for "${fileToUpload.targetFileName}" (Type: ${fileBlob.type}, Size: ${fileBlob.size} bytes) to folder ${parentFolderId}...`);
                      const metadata = {
                        name: fileToUpload.targetFileName,
                        parents: [parentFolderId],
                        appProperties: { [this.APP_PROPERTY_KEY]: hashedIdentifier }
                      };
                      const metadataBlob = new Blob([JSON.stringify(metadata)], { type: 'application/json' });
                      const formData = new FormData();
                      formData.append('metadata', metadataBlob);
                      formData.append('file', fileBlob, fileToUpload.targetFileName);

                      // The DRIVE_API_UPLOAD_ENDPOINT should already include ?uploadType=multipart
                      // And we add 'fields' to get specific response data including thumbnailLink.
                      const uploadUrlWithFields = `${this.utils.DRIVE_API_UPLOAD_ENDPOINT}${this.utils.DRIVE_API_UPLOAD_ENDPOINT.includes('?') ? '&' : '?'}fields=${encodeURIComponent(this.DRIVE_FILE_FIELDS_TO_REQUEST)}`;

                      console.log(`[FileUploadService] Upload URL with fields: ${uploadUrlWithFields}`);

                      const uploadRequest$ = this.http.post<DriveFile>(
                        uploadUrlWithFields,
                        formData,
                        {headers}
                      );

                      return this.utils.retryRequest(
                        uploadRequest$,
                        retryConfig,
                        `Upload File "${fileToUpload.targetFileName}"`
                      ).pipe(
                        // This inner switchMap is for the explicit GET request after upload
                        switchMap(initialUploadedFileResponse => {
                          console.log(`[FileUploadService] Initial response received for upload of "${fileToUpload.targetFileName}":`, JSON.stringify(initialUploadedFileResponse, null, 2));
                          if (!initialUploadedFileResponse || !initialUploadedFileResponse.id) {
                            console.error(`[FileUploadService] Failed to upload file "${fileToUpload.targetFileName}" to Google Drive. Initial response missing ID:`, initialUploadedFileResponse);
                            throw new Error(`[FileUploadService] Failed to upload file "${fileToUpload.targetFileName}" to Google Drive (initial response missing ID).`);
                          }
                          console.log(`[FileUploadService] File created with ID: ${initialUploadedFileResponse.id}. Now fetching full metadata including thumbnailLink...`);

                          const getFileMetadataUrl = `${this.utils.DRIVE_API_FILES_ENDPOINT}/${initialUploadedFileResponse.id}`;
                          const getFileParams = new HttpParams().set('fields', this.DRIVE_FILE_FIELDS_TO_REQUEST);
                          const getFileRequest$ = this.http.get<DriveFile>(getFileMetadataUrl, {headers, params: getFileParams});

                          return this.utils.retryRequest(
                            getFileRequest$,
                            retryConfig,
                            `Get Metadata for File ID ${initialUploadedFileResponse.id}`
                          ).pipe(
                            tap(fullFileMetadata => {
                              console.log(`[FileUploadService] Successfully fetched full metadata for file: ID=${fullFileMetadata.id}, Name=${fullFileMetadata.name}. WebViewLink: ${fullFileMetadata.webViewLink}, Thumbnail: ${fullFileMetadata.thumbnailLink ? 'Available' : 'Not Available'}`);
                              if (fullFileMetadata.appProperties) {
                                console.log(`   Associated ${this.APP_PROPERTY_KEY}: ${fullFileMetadata.appProperties[this.APP_PROPERTY_KEY]} (Hash of: "${originalFileIdentifier}")`);
                              } else {
                                console.warn(`   No appProperties found on fetched metadata for file: ${fullFileMetadata.id}`);
                              }
                            })
                          );
                        }),
                        catchError(err => {
                          console.error(`[FileUploadService] Raw error object during upload/metadata fetch of "${fileToUpload.targetFileName}":`, err);
                          const formattedError = this.utils.formatHttpError(err);
                          console.error(`[FileUploadService] Error during upload or metadata fetch for "${fileToUpload.targetFileName}" (final after retries - formatted):`, formattedError);
                          let detail = formattedError;
                          if (err instanceof HttpErrorResponse) {
                            detail = `Status: ${err.status}, StatusText: ${err.statusText}, Message: ${err.message}, Body: ${JSON.stringify(err.error)}`;
                          } else if (err instanceof Error) {
                            detail = err.message;
                          }
                          return throwError(() => new Error(`[FileUploadService] Upload or metadata fetch failed for ${fileToUpload.targetFileName}. Details: ${detail}`));
                        })
                      );
                    }),
                    catchError(err => {
                        const errorMessage = err instanceof Error ? err.message : String(err);
                      console.error(`[FileUploadService] Failed to prepare blob for file "${fileToUpload.targetFileName}":`, errorMessage);
                      return throwError(() => new Error(`[FileUploadService] Blob preparation failed for ${fileToUpload.targetFileName}. ${errorMessage}`));
                    })
                  );
                }
              }),
              catchError(err => {
                  const formattedError = this.utils.formatHttpError(err);
                console.error(`[FileUploadService] Error searching for existing file with hash "${hashedIdentifier}" (final after retries):`, formattedError, err);
                return throwError(() => new Error(`[FileUploadService] Search failed for ${fileToUpload.targetFileName}. ${formattedError}`));
              })
            );
          })
        );
      });

      return forkJoin(findOrUploadObservables).pipe(
        map(results => {
          console.log(`[FileUploadService] Find-or-upload process completed for ${results.length} files.`, results);
            return results;
        }),
        catchError(err => {
          const formattedError = err instanceof Error ? err.message : this.utils.formatHttpError(err);
          console.error('[FileUploadService] Error during bulk file find-or-upload process:', formattedError, err);
          return throwError(() => new Error(`[FileUploadService] Bulk file processing failed. ${formattedError}`));
        })
      );
  }
}
