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
import {HttpClient, HttpHeaders, HttpErrorResponse, HttpParams} from '@angular/common/http';
import {Observable, throwError, of, forkJoin, from} from 'rxjs';
import {switchMap, catchError, map, tap} from 'rxjs/operators';
import {ImsccFile, DriveFile} from '../../interfaces/classroom-interface';
import {UtilitiesService, RetryConfig} from '../utilities/utilities.service';
import {AuthService} from '../auth/auth.service';

@Injectable({
  providedIn: 'root'
})
export class FileUploadService {

  // Inject dependencies
  private http = inject(HttpClient);
  private utils = inject(UtilitiesService);
  private auth = inject(AuthService); // Inject AuthService

  // Custom property key to store the source identifier HASH
  private readonly APP_PROPERTY_KEY = 'imsccIdentifier';
  // Define the fields to be requested from the Drive API for file metadata
  private readonly DRIVE_FILE_FIELDS_TO_REQUEST = 'id,name,mimeType,appProperties,webViewLink,parents,thumbnailLink';

  constructor() { }

  /**
   * Creates Drive API headers, fetching token internally.
   * @returns HttpHeaders object or null if token is missing.
   */
  private createDriveApiHeaders(): HttpHeaders | null {
    const accessToken = this.auth.getGoogleAccessToken();
    if (!accessToken) {
      console.error('[FileUploadService] Cannot create Drive API headers: Access token is missing.');
      return null;
    }

    return new HttpHeaders({
      'Authorization': `Bearer ${accessToken}`
    });
  }

  /**
   * Uploads an array of local files (ImsccFile) to a Google Drive folder using HASHED identifiers.
   * Token is fetched internally just before API calls.
   *
   * @param localFilesToUpload Array of objects containing the ImsccFile and the target file name.
   * @param parentFolderId The ID of the Google Drive folder where files should be uploaded/searched.
   * @returns Observable<DriveFile[]> Emitting an array of the Google Drive File resource objects (found or uploaded).
   */
  uploadLocalFiles(
    localFilesToUpload: Array<{file: ImsccFile; targetFileName: string}>,
    parentFolderId: string
  ): Observable<DriveFile[]> {
    if (!localFilesToUpload || localFilesToUpload.length === 0) {
      console.log("[FileUploadService] No local files provided for upload.");
      return of([]);
    }

    if (!parentFolderId) return throwError(() => new Error('[FileUploadService] Parent folder ID is required for file upload.'));

    const retryConfig: RetryConfig = {maxRetries: 3, initialDelayMs: 2000};

    const findOrUploadObservables: Observable<DriveFile>[] = localFilesToUpload.map(fileToUpload => {
      const originalFileIdentifier = fileToUpload.file.name;
      if (!originalFileIdentifier) {
        console.error(`[FileUploadService] Skipping file "${fileToUpload.targetFileName}": Missing original file name (identifier).`);
        return throwError(() => new Error(`[FileUploadService] Missing identifier for file: ${fileToUpload.targetFileName}`));
      }

      let fileDataSize: number;
      if (fileToUpload.file.data instanceof ArrayBuffer) {
        fileDataSize = fileToUpload.file.data.byteLength;
      } else if (typeof fileToUpload.file.data === 'string') {
        fileDataSize = fileToUpload.file.data.length;
      } else {
        console.warn(`[FileUploadService] Could not determine data size for file "${fileToUpload.targetFileName}" (original identifier: "${originalFileIdentifier}"). Using size 0 for hash generation.`);
        fileDataSize = 0;
      }
      const compositeIdentifierString = `${originalFileIdentifier}|size:${fileDataSize}`;

      return from(this.utils.generateHash(compositeIdentifierString)).pipe(
        catchError(hashError => {
          console.error(`[FileUploadService] Error generating hash for "${originalFileIdentifier}":`, hashError);
          const errorMessage = hashError instanceof Error ? hashError.message : String(hashError);
          return throwError(() => new Error(`[FileUploadService] Failed to generate identifier for ${fileToUpload.targetFileName}. ${errorMessage}`));
        }),
        switchMap(hashedIdentifier => {
          const searchHeaders = this.createDriveApiHeaders();
          if (!searchHeaders) {
            return throwError(() => new Error(`[FileUploadService] Authentication token missing for searching file: ${fileToUpload.targetFileName}`));
          }

          const searchQuery = `'${parentFolderId}' in parents and appProperties has { key='${this.APP_PROPERTY_KEY}' and value='${this.utils.escapeQueryParam(hashedIdentifier)}' } and trashed = false`;
          const searchParams = new HttpParams()
            .set('q', searchQuery)
            .set('fields', `files(${this.DRIVE_FILE_FIELDS_TO_REQUEST})`)
            .set('spaces', 'drive');

          const searchRequest$ = this.http.get<{files: DriveFile[]}>(
            this.utils.DRIVE_API_FILES_ENDPOINT, {headers: searchHeaders, params: searchParams}
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
                let blobPreparation$: Observable<Blob | null>;
                // ... (blobPreparation$ logic remains the same) ...
                if (typeof fileToUpload.file.data === 'string' && fileToUpload.file.data.startsWith('data:')) {
                  blobPreparation$ = from(this.utils.dataUriToBlob(fileToUpload.file.data, fileToUpload.file.mimeType));
                } else if (fileToUpload.file.data instanceof ArrayBuffer) {
                  try {
                    blobPreparation$ = of(new Blob([fileToUpload.file.data], {type: fileToUpload.file.mimeType}));
                  } catch (e: any) {
                    console.error(`[FileUploadService] Error creating Blob from ArrayBuffer for ${fileToUpload.targetFileName}:`, e);
                    blobPreparation$ = throwError(() => new Error(`Failed to create Blob from ArrayBuffer for ${fileToUpload.targetFileName}: ${e.message}`));
                  }
                } else if (typeof fileToUpload.file.data === 'string') {
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
                      console.error(`[FileUploadService] Skipping upload for "${fileToUpload.targetFileName}": Blob preparation resulted in null.`);
                      return throwError(() => new Error(`Blob preparation failed for file: ${fileToUpload.targetFileName}`));
                    }

                    const uploadHeaders = this.createDriveApiHeaders();
                    if (!uploadHeaders) {
                      return throwError(() => new Error(`[FileUploadService] Authentication token missing for uploading file: ${fileToUpload.targetFileName}`));
                    }

                    const metadata = {
                      name: fileToUpload.targetFileName,
                      parents: [parentFolderId],
                      appProperties: {[this.APP_PROPERTY_KEY]: hashedIdentifier}
                    };
                    const metadataBlob = new Blob([JSON.stringify(metadata)], {type: 'application/json'});
                    const formData = new FormData();
                    formData.append('metadata', metadataBlob);
                    formData.append('file', fileBlob, fileToUpload.targetFileName);

                    const uploadUrlWithFields = `${this.utils.DRIVE_API_UPLOAD_ENDPOINT}${this.utils.DRIVE_API_UPLOAD_ENDPOINT.includes('?') ? '&' : '?'}fields=${encodeURIComponent(this.DRIVE_FILE_FIELDS_TO_REQUEST)}`;

                    const uploadRequest$ = this.http.post<DriveFile>(
                      uploadUrlWithFields,
                      formData,
                      {headers: uploadHeaders}
                    );

                    return this.utils.retryRequest(
                      uploadRequest$,
                      retryConfig,
                      `Upload File "${fileToUpload.targetFileName}"`
                    ).pipe(
                      tap(uploadedFileResponse => {
                        console.log(`[FileUploadService] File uploaded. Response: ID=${uploadedFileResponse.id}, Name=${uploadedFileResponse.name}. WebViewLink: ${uploadedFileResponse.webViewLink}, Thumbnail: ${uploadedFileResponse.thumbnailLink ? 'Available' : 'Not Available'}`);
                        if (uploadedFileResponse.appProperties?.[this.APP_PROPERTY_KEY] === hashedIdentifier) {
                          console.log(`   App property "${this.APP_PROPERTY_KEY}" successfully set and verified in upload response.`);
                        } else {
                          console.warn(`   App property "${this.APP_PROPERTY_KEY}" mismatch or missing in upload response for ${uploadedFileResponse.name}. Expected: ${hashedIdentifier}, Got: ${uploadedFileResponse.appProperties?.[this.APP_PROPERTY_KEY]}`);
                        }
                        if (!uploadedFileResponse.thumbnailLink) {
                          console.warn(`   ThumbnailLink not available in upload response for ${uploadedFileResponse.name} (ID: ${uploadedFileResponse.id}). It might generate shortly.`);
                        }
                      }),
                      catchError(err => {
                        console.error(`[FileUploadService] Raw error object during upload of "${fileToUpload.targetFileName}":`, err);
                        const formattedError = this.utils.formatHttpError(err);
                        console.error(`[FileUploadService] Error during upload for "${fileToUpload.targetFileName}" (final after retries - formatted):`, formattedError);
                        let detail = formattedError;
                        if (err instanceof HttpErrorResponse) {
                          detail = `Status: ${err.status}, StatusText: ${err.statusText}, Message: ${err.message}, Body: ${JSON.stringify(err.error)}`;
                        } else if (err instanceof Error) {
                          detail = err.message;
                        }
                        return throwError(() => new Error(`[FileUploadService] Upload failed for ${fileToUpload.targetFileName}. Details: ${detail}`));
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
        console.log(`[FileUploadService] Find-or-upload process completed for ${results.length} files.`);
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
