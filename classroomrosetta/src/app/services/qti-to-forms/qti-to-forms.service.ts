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
import {Observable, throwError, of, from, forkJoin} from 'rxjs';
import {switchMap, catchError, map, tap, mergeMap} from 'rxjs/operators';
import {ImsccFile, DriveFile, Material} from '../../interfaces/classroom-interface';
import {
  GoogleForm, FormInfo, FormItem, QuestionItem, Question, Option,
  BatchUpdateFormRequest, FormRequest, BatchUpdateFormResponse, Image as FormsImage,
} from '../../interfaces/forms-interface';
import {UtilitiesService, RetryConfig} from '../utilities/utilities.service';
import {FileUploadService} from '../file-upload/file-upload.service';
import {AuthService} from '../auth/auth.service';
import {decode} from 'html-entities';
import {environment} from '../../../environments/environment';

// Helper interface for parsed choice data
interface ParsedChoice {
  identifier: string;
  value: string;
}

// Helper interface for parsed grading info
interface ParsedGradingInfo {
  points: number;
  correctAnswerValues: string[];
}

// Intermediate structure for parsed QTI items before forming final FormRequests
interface IntermediateFormItemDefinition {
  type: 'question' | 'image_standalone';
  title?: string;
  description?: string;
  question?: Question;
  imageFileToUpload?: ImsccFile;
  originalImgSrc?: string;
  imageAltText?: string;
  videoItem?: {video: {youtubeUri: string, altText?: string}};
  pageBreakItem?: {};
  sectionHeaderItem?: {};
  gridItem?: {
    rows: string[];
    columns: string[];
    type: 'GRID' | 'CHECKBOX_GRID';
  };
}

// Interface for the expected Apps Script response
interface AppsScriptFormUpdateResponse {
  success: boolean;
  message: string;
  createdItems?: number;
  errors?: string[];
}

// Interface for the Apps Script API scripts.run request payload
interface AppsScriptRunRequest {
  function: string;
  parameters?: any[];
  devMode?: boolean;
  sessionState?: string;
}

// Interface for the Apps Script API scripts.run response
interface AppsScriptRunResponse {
  done?: boolean;
  response?: {
    result?: any;
  };
  error?: {
    code?: number;
    message?: string;
    details?: any[];
  };
}


@Injectable({
  providedIn: 'root'
})
export class QtiToFormsService {

  private http = inject(HttpClient);
  private utils = inject(UtilitiesService);
  private fileUploadService = inject(FileUploadService);
  private auth = inject(AuthService);
  private readonly APP_PROPERTY_KEY = 'imsccIdentifier';

  private readonly APPS_SCRIPT_RUN_ENDPOINT = environment.formsItemsApi;

  constructor() {
    if (!this.APPS_SCRIPT_RUN_ENDPOINT) {
      console.error("FATAL: Apps Script Form API URL (scripts.run endpoint) is not configured in the environment file. Form item creation will fail.");
    }
  }

  /**
   * Creates API headers, fetching token internally.
   * @param contentType Optional content type, defaults to 'application/json'.
   * @returns HttpHeaders object or null if token is missing.
   */
  private createApiHeaders(contentType: string = 'application/json'): HttpHeaders | null {
    const accessToken = this.auth.getGoogleAccessToken();
    if (!accessToken) {
      console.error('[QTI Service] Cannot create API headers: Access token is missing.');
      return null;
    }
    let headersConfig: {[header: string]: string | string[]} = {
      'Authorization': `Bearer ${accessToken}`
    };
    if (contentType) {
      headersConfig['Content-Type'] = contentType;
    }
    return new HttpHeaders(headersConfig);
  }


  /**
   * Helper function to check if a string is likely a filename based on common image extensions.
   */
  private isLikelyFilename(text: string | null | undefined): boolean {
    if (!text || typeof text !== 'string') return false;
    return /\.(jpeg|jpg|gif|png|svg|bmp|webp|tif|tiff)$/i.test(text.trim());
  }

  /**
   * Creates a Google Form from a QTI file.
   * Fetches OAuth token internally.
   *
   * @param qtiFile The QTI file content.
   * @param allPackageFiles All files from the IMSCC package for resource resolution.
   * @param formTitle The desired title for the new Google Form.
   * @param itemId The unique identifier (e.g., IMSCC item ID) for the form.
   * @param parentFolderId The ID of the Google Drive folder to place the form.
   * @returns Observable<Material | null> Emitting the Material object for the created/found form, or null on failure.
   */
  createFormFromQti(
    qtiFile: ImsccFile,
    allPackageFiles: ImsccFile[],
    formTitle: string,
    itemId: string,
    parentFolderId: string
  ): Observable<Material | null> {
    if (!qtiFile?.data || typeof qtiFile.data !== 'string') return throwError(() => new Error('[QTI Service] QTI file data is missing or not a string.'));
    if (!allPackageFiles) return throwError(() => new Error('[QTI Service] Package files array is required to resolve resources.'));
    if (!itemId) return throwError(() => new Error('[QTI Service] Item ID (itemId) is required.'));
    if (!formTitle) return throwError(() => new Error('[QTI Service] Form title cannot be empty.'));
    if (!parentFolderId) return throwError(() => new Error('[QTI Service] Parent folder ID is required.'));
    if (!this.APPS_SCRIPT_RUN_ENDPOINT) return throwError(() => new Error('[QTI Service] Apps Script Form API URL (scripts.run endpoint) is not configured.'));

    console.log(`[QTI Service] Starting QTI to Form conversion for item: ${itemId}, title: "${formTitle}"`);

    const retryConfig: RetryConfig = {maxRetries: 3, initialDelayMs: 2000};

    return from(this.utils.generateHash(itemId)).pipe(
      catchError(hashError => {
        console.error(`[QTI Service] Error generating hash for itemId "${itemId}":`, hashError);
        return throwError(() => new Error(`[QTI Service] Failed to generate identifier hash. ${hashError.message || hashError}`));
      }),
      switchMap(hashedItemId => {
        const driveApiHeaders = this.createApiHeaders();
        if (!driveApiHeaders) {
          return throwError(() => new Error('[QTI Service] Authentication token missing for Drive search.'));
        }
        console.log(`[QTI Service] Searching for existing Google Form with ${this.APP_PROPERTY_KEY}=${hashedItemId} in folder ${parentFolderId}...`);
        const searchQuery = `'${parentFolderId}' in parents and appProperties has { key='${this.APP_PROPERTY_KEY}' and value='${this.utils.escapeQueryParam(hashedItemId)}' } and mimeType='application/vnd.google-apps.form' and trashed = false`;
        const searchParams = new HttpParams().set('q', searchQuery).set('fields', `files(id, name, webViewLink, appProperties, parents)`);
        const driveSearchRequest$ = this.http.get<{files: DriveFile[]}>(this.utils.DRIVE_API_FILES_ENDPOINT, {headers: driveApiHeaders, params: searchParams});

        return this.utils.retryRequest(driveSearchRequest$, retryConfig, `Search Form by Hash ${hashedItemId}`).pipe(
          switchMap(searchResponse => {
            if (searchResponse.files && searchResponse.files.length > 0) {
              const foundFormDriveFile = searchResponse.files[0];
              console.log(`[QTI Service] Found existing Google Form: ID=${foundFormDriveFile.id}, Name=${foundFormDriveFile.name}.`);
              return of({form: {formUrl: foundFormDriveFile.webViewLink || `https://docs.google.com/forms/d/${foundFormDriveFile.id}/viewform`, title: foundFormDriveFile.name}} as Material);
            } else {
              console.log(`[QTI Service] No existing Form found. Creating new Form and converting QTI...`);
              let intermediateItems: IntermediateFormItemDefinition[] = [];
              try {
                const parser = new DOMParser();
                const qtiDoc = parser.parseFromString(qtiFile.data as string, "application/xml");
                const parseErrorNode = qtiDoc.querySelector('parsererror');
                if (parseErrorNode) throw new Error(`Failed to parse QTI XML: ${parseErrorNode.textContent}`);
                intermediateItems = this._parseQtiToIntermediateItems(qtiDoc, qtiFile.name, allPackageFiles);
                if (intermediateItems.length === 0) console.warn(`[QTI Service] QTI file "${qtiFile.name}" parsed, but no questions/images generated.`);
                else console.log(`[QTI Service] Parsed QTI and identified ${intermediateItems.length} potential form items.`);
              } catch (parseErr) {
                const errorMessage = parseErr instanceof Error ? parseErr.message : String(parseErr);
                return throwError(() => new Error(`[QTI Service] QTI parsing failed: ${errorMessage}`));
              }

              return this._uploadImagesAndBuildFormRequests(intermediateItems, parentFolderId).pipe(
                switchMap(qtiRequests => {
                  if (qtiRequests.length === 0 && intermediateItems.length > 0) {
                    console.warn(`[QTI Service] No form item requests generated for "${formTitle}" after processing images, though intermediate items existed.`);
                  } else if (qtiRequests.length === 0 && intermediateItems.length === 0) {
                    console.log(`[QTI Service] No content (questions/images) found in QTI to add to form "${formTitle}". Creating an empty form.`);
                  }

                  const formsApiHeadersForCreate = this.createApiHeaders('application/json');
                  if (!formsApiHeadersForCreate) {
                    return throwError(() => new Error('[QTI Service] Authentication token missing for Form creation.'));
                  }

                  const formBody: {info: FormInfo} = {info: {title: formTitle, documentTitle: formTitle}};
                  const formCreateRequest$ = this.http.post<GoogleForm>(this.utils.FORMS_API_CREATE_ENDPOINT, formBody, {headers: formsApiHeadersForCreate});

                  return this.utils.retryRequest(formCreateRequest$, retryConfig, `Create Form "${formTitle}"`).pipe(
                    tap(createdForm => console.log(`[QTI Service] Initial Form created. ID: ${createdForm.formId}`)),
                    mergeMap(createdForm => {
                      if (!createdForm?.formId) throw new Error('Form creation failed or did not return ID.');
                      const formId = createdForm.formId;

                      const formsApiHeadersForBatch = this.createApiHeaders('application/json');
                      if (!formsApiHeadersForBatch) {
                        return throwError(() => new Error('[QTI Service] Authentication token missing for making form a quiz.'));
                      }
                      const makeQuizUrl = `${this.utils.FORMS_API_BATCHUPDATE_BASE_ENDPOINT}${formId}:batchUpdate`;
                      const makeQuizRequest: FormRequest = {updateSettings: {settings: {quizSettings: {isQuiz: true}}, updateMask: 'quizSettings.isQuiz'}};
                      const makeQuizBody: BatchUpdateFormRequest = {requests: [makeQuizRequest]};
                      const makeQuizHttpRequest$ = this.http.post<BatchUpdateFormResponse>(makeQuizUrl, makeQuizBody, {headers: formsApiHeadersForBatch});
                      return this.utils.retryRequest(makeQuizHttpRequest$, retryConfig, `Make Form ${formId} a Quiz`).pipe(map(() => createdForm), catchError(quizErr => throwError(() => new Error(`Failed to set form as quiz. ${this.utils.formatHttpError(quizErr)}`))));
                    }),
                    mergeMap(createdForm => {
                      if (!createdForm?.formId) throw new Error('Form ID missing after quiz setup.');
                      if (qtiRequests.length === 0) {
                        console.log(`[QTI Service] No items to add to form ${createdForm.formId}.`);
                        return of(createdForm);
                      }

                      const appsScriptApiHeadersForRun = this.createApiHeaders('application/json');
                      if (!appsScriptApiHeadersForRun) {
                        return throwError(() => new Error('[QTI Service] Authentication token missing for Apps Script execution.'));
                      }

                      const appsScriptRunPayload: AppsScriptRunRequest = {
                        function: "createFormItemsInGoogleForm", // Name of the Apps Script function
                        parameters: [createdForm.formId, qtiRequests], // formId, then array of FormRequest objects
                        devMode: false // Set to true for testing with head deployment of Apps Script
                      };

                      console.log('[QTI Service] Calling Apps Script API to add items. Endpoint:', this.APPS_SCRIPT_RUN_ENDPOINT);

                      const addItemsViaAppsScript$ = this.http.post<AppsScriptRunResponse>(
                        this.APPS_SCRIPT_RUN_ENDPOINT,
                        appsScriptRunPayload,
                        {headers: appsScriptApiHeadersForRun}
                      );

                      return this.utils.retryRequest(addItemsViaAppsScript$, retryConfig, `Add Items to Form ${createdForm.formId} via Apps Script`).pipe(
                        map(appsScriptRunResponse => {
                          if (appsScriptRunResponse.error) {
                            console.error(`[QTI Service] Apps Script execution error for form ${createdForm.formId}:`, appsScriptRunResponse.error);
                            throw new Error(`Apps Script execution failed: ${appsScriptRunResponse.error.message || 'Unknown Apps Script error'}`);
                          }
                          if (appsScriptRunResponse.response && appsScriptRunResponse.response.result) {
                            const result = appsScriptRunResponse.response.result as AppsScriptFormUpdateResponse;
                            console.log(`[QTI Service] Apps Script function result for form ${createdForm.formId}: Success=${result.success}, Msg=${result.message}, Created=${result.createdItems}`);
                            if (!result.success) {
                              console.warn(`[QTI Service] Apps Script function reported failure for form ${createdForm.formId}: ${result.message}`, result.errors);
                            }
                          } else {
                            console.warn(`[QTI Service] Apps Script response for form ${createdForm.formId} did not contain expected result structure.`);
                          }
                          return createdForm;
                        }),
                        catchError(appsScriptErr => {
                          console.error(`[QTI Service] HTTP Error calling Apps Script API for form ${createdForm.formId}: ${this.utils.formatHttpError(appsScriptErr)}`);
                          return throwError(() => new Error(`[QTI Service] Failed to execute Apps Script to add items: ${this.utils.formatHttpError(appsScriptErr)}`));
                        })
                      );
                    }),
                    mergeMap(finalForm => {
                      if (!finalForm?.formId) return of(null);
                      const formId = finalForm.formId;

                      const driveApiHeadersForUpdate = this.createApiHeaders('application/json');
                      if (!driveApiHeadersForUpdate) {
                        console.warn(`[QTI Service] Token missing for Drive properties update for Form ${formId}. Skipping update.`);
                        return of({form: {formUrl: finalForm.responderUri || `https://docs.google.com/forms/d/${formId}/viewform`, title: finalForm.info?.title || formTitle}} as Material);
                      }

                      const driveUpdateUrl = `${this.utils.DRIVE_API_FILES_ENDPOINT}/${formId}`;
                      const driveUpdateBody = {appProperties: {[this.APP_PROPERTY_KEY]: hashedItemId}};
                      const driveUpdateParams = new HttpParams().set('addParents', parentFolderId).set('removeParents', 'root');
                      const driveUpdateRequest$ = this.http.patch<DriveFile>(driveUpdateUrl, driveUpdateBody, {headers: driveApiHeadersForUpdate, params: driveUpdateParams});

                      return this.utils.retryRequest(driveUpdateRequest$, retryConfig, `Update Drive Props for Form ${formId}`).pipe(
                        map(driveFile => ({form: {formUrl: finalForm.responderUri || `https://docs.google.com/forms/d/${formId}/viewform`, title: driveFile.name || finalForm.info?.title || formTitle}} as Material)),
                        catchError(driveErr => {
                          console.error(`[QTI Service] Error updating Drive properties for Form ${formId}: ${this.utils.formatHttpError(driveErr)}`);
                          return of({form: {formUrl: finalForm.responderUri || `https://docs.google.com/forms/d/${formId}/viewform`, title: finalForm.info?.title || formTitle}} as Material);
                        })
                      );
                    }),
                    catchError(err => {
                      console.error(`[QTI Service] Error in Form creation/update pipeline for ${hashedItemId}: ${err.message || err}`);
                      return of(null);
                    })
                  );
                })
              );
            }
          }),
          catchError(err => {
            console.error(`[QTI Service] Error searching for existing Form for ${hashedItemId}: ${this.utils.formatHttpError(err)}`);
            return of(null);
          })
        );
      })
    );
  }

  /**
   * Uploads images found in intermediate items and then builds FormRequest objects.
   * Handles retrying thumbnail fetches for uploaded images.
   */
  private _uploadImagesAndBuildFormRequests(
    intermediateItems: IntermediateFormItemDefinition[],
    parentFolderIdForImages: string,
  ): Observable<FormRequest[]> {
    if (intermediateItems.length === 0) {
      return of([]);
    }

    const itemsThatNeedImageUpload = intermediateItems.filter(item => !!item.imageFileToUpload);
    if (itemsThatNeedImageUpload.length === 0) {
      // If no images to upload, directly build requests (still async due to _buildFormRequestsFromIntermediate)
      return this._buildFormRequestsFromIntermediate(intermediateItems, new Map()).pipe(
        map(formRequests => { // Ensure indices are set
          return formRequests.map((req, idx) => {
            if (req.createItem && req.createItem.location) {
              req.createItem.location.index = idx;
            }
            return req;
          });
        })
      );
    }

    const filesToBatchUpload: Array<{file: ImsccFile; targetFileName: string}> = itemsThatNeedImageUpload.map(itemDef => ({
      file: itemDef.imageFileToUpload!,
      targetFileName: itemDef.imageFileToUpload!.name
    }));

    console.log(`[QTI Service] Attempting to upload ${filesToBatchUpload.length} images to Drive folder ${parentFolderIdForImages}.`);

    return this.fileUploadService.uploadLocalFiles(filesToBatchUpload, parentFolderIdForImages).pipe(
      switchMap(uploadedDriveFiles => { // MODIFIED: map to switchMap
        console.log(`[QTI Service] Successfully processed ${uploadedDriveFiles.length} of ${filesToBatchUpload.length} image uploads (some may have been found existing).`);
        const driveFileMap = new Map<string, DriveFile>();

        uploadedDriveFiles.forEach(driveFile => {
          if (driveFile && driveFile.name) {
            const originalImsccFile = filesToBatchUpload.find(f => f.targetFileName === driveFile.name)?.file;
            if (originalImsccFile) {
              driveFileMap.set(originalImsccFile.name, driveFile);
            } else {
              console.warn(`[QTI Service] Could not map uploaded/found Drive file "${driveFile.name}" back to an original ImsccFile name.`);
            }
          }
        });

        return this._buildFormRequestsFromIntermediate(intermediateItems, driveFileMap);
      }),
      map(formRequests => {
        return formRequests.map((req, idx) => {
          if (req.createItem && req.createItem.location) {
            req.createItem.location.index = idx;
          }
          return req;
        });
      }),
      catchError(uploadOrBuildError => {
        console.error('[QTI Service] Image upload or form request building process failed overall:', uploadOrBuildError);
        console.warn('[QTI Service] Attempting to build form requests without any uploaded/resolved images due to previous error.');
        return this._buildFormRequestsFromIntermediate(intermediateItems, new Map()).pipe(
          map(requestsWithoutImages => {
            return requestsWithoutImages.map((req, idx) => {
              if (req.createItem && req.createItem.location) {
                req.createItem.location.index = idx;
              }
              return req;
            });
          }),
          catchError(fallbackBuildError => {
            console.error('[QTI Service] Fallback attempt to build requests without images also failed:', fallbackBuildError);
            return of([]);
          })
        );
      })
    );
  }

  /**
   * Builds an array of FormRequest objects from intermediate definitions,
   * asynchronously handling thumbnail fetching for images if needed.
   */
  private _buildFormRequestsFromIntermediate(
    intermediateItems: IntermediateFormItemDefinition[],
    driveFileMap: Map<string, DriveFile>
  ): Observable<FormRequest[]> { // MODIFIED: Returns Observable
    if (intermediateItems.length === 0) {
      return of([]);
    }

    const itemObservables: Observable<FormRequest | null>[] = intermediateItems.map(itemDef => {
      if (itemDef.imageFileToUpload) {
        const originalImsccImageFileName = itemDef.imageFileToUpload.name;
        const initialDriveFile = driveFileMap.get(originalImsccImageFileName) || null;

        // Check if thumbnail is missing AND we have a file ID to attempt a re-fetch
        if (initialDriveFile && initialDriveFile.id && !initialDriveFile.thumbnailLink) {
          console.log(`[QTI Service] Thumbnail missing for "${initialDriveFile.name}" (ID: ${initialDriveFile.id}). Will attempt to re-fetch metadata for thumbnail.`);
          const driveApiHeaders = this.createApiHeaders();
          if (!driveApiHeaders) {
            console.warn(`[QTI Service] Cannot re-fetch Drive file metadata for ${initialDriveFile.name}: Auth token missing. Proceeding with current data.`);
            return of(this._constructSingleFormRequest(itemDef, initialDriveFile));
          }

          const fetchUrl = `${this.utils.DRIVE_API_FILES_ENDPOINT}/${initialDriveFile.id}`;
          // Request specifically the fields needed, including thumbnailLink
          const fetchParams = new HttpParams().set('fields', 'id,name,thumbnailLink,webViewLink');

          const getFileMetadata$ = this.http.get<DriveFile>(fetchUrl, {headers: driveApiHeaders, params: fetchParams});

          const thumbnailRetryConfig: RetryConfig = {
            maxRetries: 3,
            initialDelayMs: 3000, // Start with 3 seconds
            backoffFactor: 2,
          };
          const requestNameForLogging = `Fetch Drive File Metadata for thumbnail of '${initialDriveFile.name}' (ID: ${initialDriveFile.id})`;

          return this.utils.retryRequest(getFileMetadata$, thumbnailRetryConfig, requestNameForLogging).pipe(
            map(updatedDriveFileWithThumbnail => {
              // updatedDriveFileWithThumbnail contains the latest metadata.
              // thumbnailLink may or may not be present even after retries.
              if (updatedDriveFileWithThumbnail.thumbnailLink) {
                console.log(`[QTI Service] Thumbnail successfully fetched for "${updatedDriveFileWithThumbnail.name}" after retry.`);
              } else {
                console.warn(`[QTI Service] Thumbnail still not available for "${updatedDriveFileWithThumbnail.name}" after retries. Will use fallback link.`);
              }
              return this._constructSingleFormRequest(itemDef, updatedDriveFileWithThumbnail);
            }),
            catchError(err => {
              console.error(`[QTI Service] Error re-fetching metadata for ${initialDriveFile.name} (ID: ${initialDriveFile.id}) after retries: ${this.utils.formatHttpError(err)}. Proceeding with original data or fallback link.`);
              // Fallback to using the initialDriveFile (before retry attempt) if retry fails.
              return of(this._constructSingleFormRequest(itemDef, initialDriveFile));
            })
          );
        } else {
          // Thumbnail already exists on initialDriveFile, or no file ID to fetch,
          // or it's a data URI, or no DriveFile mapping found.
          // Proceed with the initialDriveFile data (or lack thereof).
          return of(this._constructSingleFormRequest(itemDef, initialDriveFile));
        }
      } else {
        // Not an image item that requires Drive upload/processing, construct request directly.
        return of(this._constructSingleFormRequest(itemDef, null));
      }
    });

    return forkJoin(itemObservables).pipe(
      map(results => {
        const validRequests = results.filter(req => req !== null) as FormRequest[];
        console.log(`[QTI Service] Built ${validRequests.length} FormRequest objects after async processing of items.`);
        return validRequests; // Indices will be set by the caller
      }),
      catchError(err => {
        console.error('[QTI Service] Error in forkJoin while building form requests from intermediate items:', err);
        return of([]); // Return empty array on error during forkJoin
      })
    );
  }

  /**
   * Helper method to construct a single FormRequest object from an IntermediateFormItemDefinition
   * and the relevant DriveFile (which might be post-thumbnail-retry).
   * @param itemDef The intermediate item definition.
   * @param driveFileForImageProcessing The DriveFile to use for image URI generation (could be null).
   * @returns A FormRequest object or null if the item cannot be converted.
   */
  private _constructSingleFormRequest(
    itemDef: IntermediateFormItemDefinition,
    driveFileForImageProcessing: DriveFile | null
  ): FormRequest | null {
    let formItem: FormItem | null = null;
    let imageForForm: FormsImage | undefined = undefined;

    if (itemDef.imageFileToUpload) {
      const currentDriveFile = driveFileForImageProcessing;

      let driveImageUri: string | null = null;
      let altTextForImageObject = "Image";

      if (itemDef.imageAltText && !this.isLikelyFilename(itemDef.imageAltText)) {
        altTextForImageObject = itemDef.imageAltText;
      }

      if (currentDriveFile) {
        if (currentDriveFile.thumbnailLink) {
          // Remove sizing parameters like =s220 to get the original image if possible
          driveImageUri = currentDriveFile.thumbnailLink.replace(/=s\d+$/, '');
        } else if (currentDriveFile.id) {
          // Fallback to direct Drive link if thumbnail is still missing
          driveImageUri = `https://drive.google.com/uc?id=${currentDriveFile.id}`;
          console.log(`[QTI Service] Constructing image for "${itemDef.title || itemDef.imageFileToUpload.name}": ThumbnailLink missing for Drive file "${currentDriveFile.name}" (ID: ${currentDriveFile.id}) even after potential retry. Using constructed 'uc?id=' link: ${driveImageUri}. Ensure file is publicly viewable or accessible by Forms.`);
        } else {
          console.warn(`[QTI Service] DriveFile for "${itemDef.imageFileToUpload.name}" (original src: ${itemDef.originalImgSrc}) is missing both thumbnailLink and ID. Cannot generate image URI.`);
        }
      } else if (typeof itemDef.imageFileToUpload.data === 'string' && itemDef.imageFileToUpload.data.startsWith('data:image')) {
        // If it's a base64 data URI (e.g., from unzipping or already processed)
        driveImageUri = itemDef.imageFileToUpload.data;
      } else {
        console.warn(`[QTI Service] No DriveFile object found/available for image: "${itemDef.imageFileToUpload.name}" (original src: ${itemDef.originalImgSrc}), and data is not a base64 URI. Image will be skipped.`);
      }

      if (driveImageUri) {
        imageForForm = {sourceUri: driveImageUri, altText: altTextForImageObject};
      } else {
        console.warn(`[QTI Service] Image "${itemDef.imageFileToUpload.name}" for item "${itemDef.title}" could not be processed (no URI generated). Item will be created without this image.`);
      }
    }

    // Build the FormItem based on itemDef.type
    if (itemDef.type === 'question' && itemDef.question) {
      const questionItem: QuestionItem = {question: itemDef.question};
      if (imageForForm) {
        questionItem.image = imageForForm;
      }
      formItem = {title: itemDef.title, description: itemDef.description, questionItem: questionItem};
    } else if (itemDef.type === 'image_standalone') {
      let standaloneImageTitle = itemDef.title;
      if (!standaloneImageTitle && itemDef.imageAltText && !this.isLikelyFilename(itemDef.imageAltText)) {
        standaloneImageTitle = itemDef.imageAltText;
      } else if (!standaloneImageTitle || this.isLikelyFilename(standaloneImageTitle)) {
        standaloneImageTitle = "Image";
      }

      if (imageForForm) {
        formItem = {
          title: standaloneImageTitle,
          description: (itemDef.description && !this.isLikelyFilename(itemDef.description)) ? itemDef.description : undefined,
          imageItem: {image: imageForForm}
        };
      } else {
        // If imageForForm could not be created (e.g. no URI), we skip creating the standalone image item.
        console.warn(`[QTI Service] Standalone image (original src: ${itemDef.originalImgSrc || 'unknown'}) could not be processed because image URI is missing. Skipping this image item.`);
      }
    } else if (itemDef.videoItem) {
      formItem = {title: itemDef.title, description: itemDef.description, videoItem: itemDef.videoItem};
    } else if (itemDef.pageBreakItem) {
      formItem = {title: itemDef.title, description: itemDef.description, pageBreakItem: itemDef.pageBreakItem};
    } else if (itemDef.sectionHeaderItem) {
      formItem = {title: itemDef.title, description: itemDef.description, sectionHeaderItem: itemDef.sectionHeaderItem};
    }

    if (formItem) {
      // The index is initially set to 0; it will be correctly reassigned by the caller
      // after all items are processed and their final order is known.
      return {createItem: {item: formItem, location: {index: 0}}};
    }
    return null;
  }


  private _parseQtiToIntermediateItems(
    qtiDoc: XMLDocument,
    qtiFilePath: string,
    allPackageFiles: ImsccFile[]
  ): IntermediateFormItemDefinition[] {
    let allFoundItems: IntermediateFormItemDefinition[] = [];
    const itemElements = Array.from(qtiDoc.querySelectorAll('assessment > section > item, assessment > item, item'));
    console.log(`[QTI Service] Found ${itemElements.length} <item> elements in "${qtiFilePath}".`);

    if (itemElements.length === 0) {
      const itemBodies = Array.from(qtiDoc.querySelectorAll('itemBody'));
      if (itemBodies.length > 0) {
        console.warn(`[QTI Service] No <item> elements found. Processing ${itemBodies.length} <itemBody> elements directly (less reliable).`);
        itemBodies.forEach((bodyEl, idx) => {
          const htmlMattext = bodyEl.querySelector('mattext[texttype="text/html"]');
          if (htmlMattext && htmlMattext.textContent) {
            allFoundItems.push(...this.parseHtmlContentWithinQti(htmlMattext, qtiFilePath, allPackageFiles, `direct_itembody_${idx}_html`));
          } else {
            allFoundItems.push(...this.parseStandardQtiItems(bodyEl, qtiFilePath, allPackageFiles, `direct_itembody_${idx}_standard`));
          }
        });
      } else {
        console.warn(`[QTI Service] No <item> or <itemBody> elements found in QTI file "${qtiFilePath}". Cannot parse questions.`);
      }
      return allFoundItems;
    }

    itemElements.forEach((itemElement, index) => {
      const itemIdent = itemElement.getAttribute('ident') || itemElement.getAttribute('identifier');
      const itemLabel = itemElement.getAttribute('label');
      // console.log(`[QTI Service] Processing <item> #${index + 1} (ident: ${itemIdent || 'N/A'}, label: ${itemLabel || 'N/A'}).`);

      const itemIdentifierForParsing = itemIdent || itemLabel || `item_${index}`;

      const standardItems = this.parseStandardQtiItems(itemElement, qtiFilePath, allPackageFiles, itemIdentifierForParsing);

      if (standardItems.length > 0) {
        allFoundItems.push(...standardItems);
      } else {
        const mattextElement = itemElement.querySelector('presentation > material > mattext[texttype="text/html"], itemBody > div > mattext[texttype="text/html"], itemBody > mattext[texttype="text/html"]');
        if (mattextElement && mattextElement.textContent) {
          const itemsFromHtml = this.parseHtmlContentWithinQti(mattextElement, qtiFilePath, allPackageFiles, itemIdentifierForParsing + '_html');
          if (itemsFromHtml.length > 0) {
            allFoundItems.push(...itemsFromHtml);
          }
        }
      }
    });

    return allFoundItems;
  }

  // Sub-parser for HTML content embedded within QTI (e.g., in <mattext>)
  private parseHtmlContentWithinQti(
    htmlSourceElement: Element,
    qtiFilePath: string,
    allPackageFiles: ImsccFile[],
    baseIdentifier: string
  ): IntermediateFormItemDefinition[] {
    const intermediateItems: IntermediateFormItemDefinition[] = [];
    if (!htmlSourceElement.textContent) return intermediateItems;

    const htmlContent = decode(htmlSourceElement.textContent.trim());
    const parser = new DOMParser();
    const htmlDoc = parser.parseFromString(htmlContent, 'text/html');
    const bodyChildren = Array.from(htmlDoc.body.children);
    let questionCounter = 0; // Used for titling if no other text is found

    bodyChildren.forEach((element: Element) => { // elIdx removed as not used

      if (element.tagName.toLowerCase() === 'img') {
        const imgSrc = element.getAttribute('src');
        const rawAltText = element.getAttribute('alt');
        let imageFileToUpload: ImsccFile | undefined;
        if (imgSrc) {
          const imagePath = this._resolveImagePath(imgSrc, qtiFilePath);
          if (imagePath) imageFileToUpload = allPackageFiles.find(f => f.name.toLowerCase() === imagePath.toLowerCase());
        }
        intermediateItems.push({
          type: 'image_standalone',
          title: (rawAltText && !this.isLikelyFilename(rawAltText)) ? rawAltText : "Image",
          imageFileToUpload: imageFileToUpload,
          originalImgSrc: imgSrc || undefined,
          imageAltText: rawAltText || (imageFileToUpload ? imageFileToUpload.name : '')
        });
      } else if (element.tagName.toLowerCase() === 'p' || element.tagName.toLowerCase() === 'div' || element.tagName.toLowerCase() === 'span' || element.tagName.toLowerCase() === 'table' || element.tagName.toLowerCase() === 'ul' || element.tagName.toLowerCase() === 'ol') {
        const imagesInElement = Array.from(element.querySelectorAll('img'));
        // Get text content from the current element, excluding its own image alt texts for this specific purpose.
        const textContentFromParent = this.getTextContent(element, true).trim();

        if (imagesInElement.length > 0) {
          imagesInElement.forEach((imgElement) => { // imgIdx removed as not used
            questionCounter++;
            const imgSrc = imgElement.getAttribute('src');
            const rawAltText = imgElement.getAttribute('alt'); // This is the alt text specific to this image
            let imageFileToUpload: ImsccFile | undefined;
            if (imgSrc) {
              const imagePath = this._resolveImagePath(imgSrc, qtiFilePath);
              if (imagePath) imageFileToUpload = allPackageFiles.find(f => f.name.toLowerCase() === imagePath.toLowerCase());
            }

            // Prioritize text from the parent element containing the image as the question title.
            // If no such text, use the image's alt text (if descriptive).
            // Fallback to a generic title if needed.
            let titleForIntermediateItem = textContentFromParent ||
              (rawAltText && !this.isLikelyFilename(rawAltText) ? rawAltText : `Image Question ${questionCounter}`);
            if (this.isLikelyFilename(titleForIntermediateItem)) titleForIntermediateItem = `Image Question ${questionCounter}`;

            let descriptionForIntermediateItem = "";
            // If there was parent text AND descriptive alt text, use alt text as description.
            if (textContentFromParent && rawAltText && !this.isLikelyFilename(rawAltText) && rawAltText !== titleForIntermediateItem) {
              descriptionForIntermediateItem = rawAltText;
            } else if (!textContentFromParent && rawAltText && !this.isLikelyFilename(rawAltText) && rawAltText === titleForIntermediateItem) {
              // If title came from alt text, no separate description from it.
              descriptionForIntermediateItem = "";
            }


            intermediateItems.push({
              type: 'question', // Treat images with surrounding text as questions by default
              title: titleForIntermediateItem,
              description: descriptionForIntermediateItem || undefined, // Ensure undefined if empty
              question: {textQuestion: {paragraph: false}}, // Default to short answer for image-based questions
              imageFileToUpload: imageFileToUpload,
              originalImgSrc: imgSrc || undefined,
              imageAltText: rawAltText || (imageFileToUpload ? imageFileToUpload.name : '') // Store image's own alt text
            });
          });
        } else if (textContentFromParent) { // If no images in this element, but there is text
          questionCounter++;
          let qText = textContentFromParent;
          // Avoid double-numbering if text already starts like "1. Question..."
          if (!textContentFromParent.match(/^\s*\d+[\.\)]\s+/)) {
            qText = `${questionCounter}. ${textContentFromParent}`;
          }
          intermediateItems.push({
            type: 'question',
            title: qText,
            description: undefined,
            question: {textQuestion: {paragraph: true}}, // Assume paragraph for text-only items from HTML
            imageAltText: undefined // No image directly associated with this text item for alt text purposes
          });
        }
      }
      // Other HTML elements could be parsed here if needed (e.g. for textItem, etc.)
    });
    return intermediateItems;
  }

  private parseStandardQtiItems(
    itemElement: Element,
    sourceFileName: string,
    allPackageFiles: ImsccFile[],
    itemIdentifierOverride?: string
  ): IntermediateFormItemDefinition[] {
    const intermediateItems: IntermediateFormItemDefinition[] = [];
    const itemIdentifier = itemIdentifierOverride || itemElement.getAttribute('ident') || itemElement.getAttribute('identifier') || `qti_item_${Date.now()}`;
    let itemTitleAttr = itemElement.getAttribute('title') || '';

    // const itemBody = itemElement.querySelector('itemBody'); // Not directly used, sub-elements are queried
    // const presentation = itemElement.querySelector('presentation'); // Not directly used

    let questionText = '';
    let itemDescription: string | undefined = undefined;
    let imageFileToUpload: ImsccFile | undefined;
    let originalImgSrc: string | undefined;
    let imageAltText: string | undefined;

    const presentationFlowMaterial = itemElement.querySelector('presentation > flow > material');
    if (presentationFlowMaterial) {
      const mainMatTextElement = presentationFlowMaterial.querySelector('mattext');
      if (mainMatTextElement) {
        questionText = this.getTextContent(mainMatTextElement).trim();
        const matHtmlContent = mainMatTextElement.innerHTML;
        if (matHtmlContent) {
          const parser = new DOMParser();
          const tempDoc = parser.parseFromString(decode(matHtmlContent), "text/html");
          const imgTag = tempDoc.querySelector('img');
          if (imgTag) {
            originalImgSrc = imgTag.getAttribute('src') || undefined;
            const rawAlt = imgTag.getAttribute('alt');
            if (rawAlt) imageAltText = rawAlt;

            if (originalImgSrc) {
              const imagePath = this._resolveImagePath(originalImgSrc, sourceFileName);
              if (imagePath) {
                imageFileToUpload = allPackageFiles.find(f => f.name.toLowerCase() === imagePath.toLowerCase());
                if (!imageFileToUpload) console.warn(`   [QTI Standard - Item ${itemIdentifier}] Image (from main material: ${originalImgSrc}, resolved: ${imagePath}) not found in package files.`);
              } else {
                console.warn(`   [QTI Standard - Item ${itemIdentifier}] Could not resolve image path for src: ${originalImgSrc}`);
              }
            }
          }
        }
      }
      if (!questionText && imageFileToUpload && imageAltText && !this.isLikelyFilename(imageAltText)) {
        questionText = imageAltText;
      }
    }

    if (!questionText) {
      const promptElement = itemElement.querySelector(':scope > itemBody > prompt, :scope > presentation > prompt, :scope > prompt');
      if (promptElement) {
        questionText = this.getTextContent(promptElement).trim();
      }
    }

    if (!questionText && itemTitleAttr && !this.isLikelyFilename(itemTitleAttr)) {
      questionText = itemTitleAttr;
    }

    if (!questionText && imageFileToUpload) {
      questionText = (imageAltText && !this.isLikelyFilename(imageAltText)) ? imageAltText : "Image-based question";
    }

    const cleanQuestionText = questionText.replace(/[\n\r\t]+/g, ' ').replace(/\s\s+/g, ' ').trim();
    if (!cleanQuestionText && !imageFileToUpload) {
      // console.warn(`   Skipping standard QTI item ${itemIdentifier}: No question text or image could be extracted after all checks.`);
      return intermediateItems;
    }

    const descriptionMetaElement = itemElement.querySelector('itemmetadata > qtimetadata > qti_metadatafield[fieldlabel="qmd_description"] > fieldentry, itemmetadata > qtimetadata > qmd_description');
    if (descriptionMetaElement) {
      itemDescription = this.getTextContent(descriptionMetaElement).trim();
    } else {
      const rubricBlock = itemElement.querySelector(':scope > itemBody > rubricBlock, :scope > presentation > rubricBlock, :scope > rubricBlock');
      if (rubricBlock) itemDescription = this.getTextContent(rubricBlock).trim();
    }

    const choiceInteraction = itemElement.querySelector('choiceInteraction');
    const textEntryInteraction = itemElement.querySelector('textEntryInteraction');
    const extendedTextInteraction = itemElement.querySelector('extendedTextInteraction');

    const responseLid = itemElement.querySelector('response_lid[rtype="MultipleChoice"], response_lid[rtype="TrueFalse"], response_lid[rtype="Selection"]');
    const renderChoice = responseLid?.querySelector('render_choice');

    const responseStr = itemElement.querySelector('response_str[rtype="String"]');
    const renderFib = responseStr?.querySelector('render_fib');

    let question: Question | undefined;
    const allParsedChoices: ParsedChoice[] = [];
    let questionTypeForGrading: 'CHOICE' | 'TEXT' | 'UNKNOWN' = 'UNKNOWN';

    if (choiceInteraction || renderChoice) {
      questionTypeForGrading = 'CHOICE';
      const interactionElement = choiceInteraction || renderChoice!;

      let maxChoicesAttr: string | null = null;
      if (choiceInteraction) maxChoicesAttr = choiceInteraction.getAttribute('maxChoices');
      else if (responseLid) maxChoicesAttr = responseLid.getAttribute('cardinality') === 'Multiple' ? "0" : "1";

      const isCheckbox = maxChoicesAttr !== '1';
      const choiceTypeValue: 'RADIO' | 'CHECKBOX' = isCheckbox ? 'CHECKBOX' : 'RADIO';
      const options: Option[] = [];

      const choices = Array.from(interactionElement.querySelectorAll('simpleChoice, response_label'));

      choices.forEach(choiceEl => {
        let choiceTextVal = '';
        let choiceId = choiceEl.getAttribute('identifier') || choiceEl.getAttribute('ident');

        if (choiceEl.tagName.toLowerCase() === 'simplechoice') {
          choiceTextVal = this.getTextContent(choiceEl).trim();
        } else if (choiceEl.tagName.toLowerCase() === 'response_label') {
          const mattextChoice = choiceEl.querySelector('material > mattext, mattext');
          if (mattextChoice) choiceTextVal = this.getTextContent(mattextChoice).trim();
        }

        if (choiceTextVal && choiceId) {
          if (!options.find(o => o.value === choiceTextVal)) {
            options.push({value: choiceTextVal});
            allParsedChoices.push({identifier: choiceId, value: choiceTextVal});
          }
        }
      });
      if (options.length > 0) {
        question = {choiceQuestion: {type: choiceTypeValue, options: options, shuffle: renderChoice?.getAttribute('shuffle') === 'yes'}};
      }
    } else if (textEntryInteraction || extendedTextInteraction || (responseStr && renderFib)) {
      questionTypeForGrading = 'TEXT';
      const isParagraph = !!extendedTextInteraction || (!!renderFib && (renderFib.getAttribute('rows') || '1') !== '1');
      question = {textQuestion: {paragraph: isParagraph}};
    }

    if (!question && cleanQuestionText) { // If no interaction type matched but we have text
      // console.log(`   [QTI Standard - Item ${itemIdentifier}] No specific interaction found for "${cleanQuestionText.substring(0, 50)}...", defaulting to a paragraph text question.`);
      question = {textQuestion: {paragraph: true}};
      questionTypeForGrading = 'TEXT'; // Treat as text for grading purposes.
    }

    if (question) {
      const gradingInfo = this.parseResponseProcessing(itemElement, questionTypeForGrading, allParsedChoices);
      if (gradingInfo && gradingInfo.correctAnswerValues.length > 0 && !(question.textQuestion?.paragraph && questionTypeForGrading === 'TEXT')) { // Only grade non-paragraph text for now
        question.required = true; // Typically, graded questions are required
        question.grading = {pointValue: gradingInfo.points, correctAnswers: {answers: gradingInfo.correctAnswerValues.map(val => ({value: val}))}};
      }

      intermediateItems.push({
        type: 'question',
        title: cleanQuestionText,
        description: itemDescription || undefined,
        question: question,
        imageFileToUpload: imageFileToUpload,
        originalImgSrc: originalImgSrc,
        imageAltText: imageAltText
      });
    } else if (imageFileToUpload) { // If only an image was found, create a standalone image item
      intermediateItems.push({
        type: 'image_standalone',
        title: cleanQuestionText || (imageAltText && !this.isLikelyFilename(imageAltText) ? imageAltText : "Image"), // Use text or alt text as title
        description: itemDescription || undefined,
        imageFileToUpload: imageFileToUpload,
        originalImgSrc: originalImgSrc,
        imageAltText: imageAltText
      });
    } else {
      // console.warn(`   Skipping standard QTI item ${itemIdentifier}: No question text, identifiable interaction, or standalone image found after all parsing attempts.`);
    }
    return intermediateItems;
  }


  private parseResponseProcessing(item: Element, questionType: 'CHOICE' | 'TEXT' | 'UNKNOWN', choicesMap?: ParsedChoice[]): ParsedGradingInfo | null {
    if (questionType === 'UNKNOWN') return null;
    const respProcessing = item.querySelector('resprocessing, responseProcessing');
    if (!respProcessing) return null;

    let points = 0;
    let pointsExplicitlySet = false;
    const correctAnswerValues: string[] = [];

    const weightMetaField = Array.from(item.querySelectorAll('itemmetadata > qtimetadata > qti_metadatafield'))
      .find(field => field.querySelector('fieldlabel')?.textContent?.trim().toLowerCase() === 'qmd_weighting');
    if (weightMetaField) {
      const weightEntry = weightMetaField.querySelector('fieldentry');
      if (weightEntry?.textContent) {
        const parsedPoints = parseFloat(weightEntry.textContent.trim());
        if (!isNaN(parsedPoints) && parsedPoints >= 0) {
          points = Math.round(parsedPoints);
          pointsExplicitlySet = true;
        }
      }
    } else {
      const scoreOutcome = Array.from(item.querySelectorAll('outcomeDeclaration[identifier="SCORE"], outcomeDeclaration[identifier="MAXSCORE"]'))
        .find(od => od.querySelector('defaultValue > value'));
      if (scoreOutcome) {
        const scoreValue = scoreOutcome.querySelector('defaultValue > value')?.textContent;
        if (scoreValue) {
          const parsedPoints = parseFloat(scoreValue.trim());
          if (!isNaN(parsedPoints) && parsedPoints >= 0) {
            points = Math.round(parsedPoints);
            pointsExplicitlySet = true;
          }
        }
      }
    }

    const respconditions = Array.from(respProcessing.querySelectorAll('respcondition'));
    respconditions.forEach(condition => {
      const setvar = condition.querySelector('setvar[varname="SCORE"], setvar[varname="score"], setvar'); // D2L uses "score", QTI spec "SCORE"
      const scoreFromSetvarText = setvar?.textContent; // This is the value being set
      const action = setvar?.getAttribute('action'); // Should be 'Set'

      if (action === 'Set' && scoreFromSetvarText) {
        const scoreFromSetvar = parseFloat(scoreFromSetvarText.trim());
        if (!isNaN(scoreFromSetvar)) {
          if (scoreFromSetvar === 100.0) {
            if (!pointsExplicitlySet || points === 0) {
              points = 1;
              pointsExplicitlySet = true;
            }
          } else if (scoreFromSetvar > 0 && scoreFromSetvar !== 100.0) {
            if (!pointsExplicitlySet || scoreFromSetvar > points) {
              points = Math.round(scoreFromSetvar);
              pointsExplicitlySet = true;
            }
          }
        }

        const conditionVar = condition.querySelector('conditionvar');
        const varequal = conditionVar?.querySelector('varequal') || null; // Could be 'varequal' or 'varequal_multiple' etc.
        const varAnd = conditionVar?.querySelector('and');
        const varSubset = conditionVar?.querySelector('varsubset');

        const processCorrectValue = (valueProviderElement: Element | null) => {
          if (!valueProviderElement) return;
          const correctIdentifierOrValue = this.getTextContent(valueProviderElement)?.trim();
          if (correctIdentifierOrValue) {
            if (!isNaN(scoreFromSetvar) && scoreFromSetvar > 0) { // Only add if it contributes to a positive score
              if (questionType === 'CHOICE' && choicesMap) {
                const matchingChoice = choicesMap.find(c => c.identifier === correctIdentifierOrValue);
                if (matchingChoice && !correctAnswerValues.includes(matchingChoice.value)) {
                  correctAnswerValues.push(matchingChoice.value);
                }
              } else if (questionType === 'TEXT') { // For text questions, the value itself is the answer
                if (!correctAnswerValues.includes(correctIdentifierOrValue)) {
                  correctAnswerValues.push(correctIdentifierOrValue);
                }
              }
            }
          }
        };

        processCorrectValue(varequal);
        if (varAnd) varAnd.querySelectorAll('varequal').forEach(ve => processCorrectValue(ve));
        if (varSubset) { // varsubset contains space-separated identifiers
          const subsetIdentifiers = varSubset.textContent?.trim().split(/\s+/);
          subsetIdentifiers?.forEach(id => {
            if (!isNaN(scoreFromSetvar) && scoreFromSetvar > 0) {
              if (questionType === 'CHOICE' && choicesMap) {
                const matchingChoice = choicesMap.find(c => c.identifier === id);
                if (matchingChoice && !correctAnswerValues.includes(matchingChoice.value)) {
                  correctAnswerValues.push(matchingChoice.value);
                }
              }
            }
          });
        }
      }
    });

    // Fallback: Check <responseDeclaration> if no answers found via <resprocessing>
    if (correctAnswerValues.length === 0) {
      const responseDeclarations = Array.from(item.querySelectorAll('responseDeclaration'));
      // Try to find the most relevant response declaration
      const primaryResponseDecl = responseDeclarations.find(rd => rd.getAttribute('identifier')?.toUpperCase() === 'RESPONSE') ||
        responseDeclarations.find(rd => rd.querySelector('correctResponse')) || // Any with a correctResponse tag
        responseDeclarations[0]; // Fallback to the first one

      if (primaryResponseDecl) {
        // Check mapping for default score if points not set
        const mapping = primaryResponseDecl.querySelector('mapping');
        if (mapping) {
          const defaultValue = mapping.getAttribute('defaultValue'); // This is often the score for correct
          if (defaultValue) {
            const mappedPoints = parseFloat(defaultValue);
            if (!isNaN(mappedPoints) && mappedPoints >= 0) {
              if (!pointsExplicitlySet || mappedPoints > points) { // If mapping provides higher points or not set
                points = Math.round(mappedPoints);
                pointsExplicitlySet = true;
              }
            }
          }
        }

        const correctResponse = primaryResponseDecl.querySelector('correctResponse');
        if (correctResponse) {
          correctResponse.querySelectorAll('value').forEach(valueEl => {
            const correctIdentifierOrValue = this.getTextContent(valueEl)?.trim();
            if (correctIdentifierOrValue) {
              if (questionType === 'CHOICE' && choicesMap) {
                const matchingChoiceById = choicesMap.find(c => c.identifier === correctIdentifierOrValue);
                if (matchingChoiceById && !correctAnswerValues.includes(matchingChoiceById.value)) {
                  correctAnswerValues.push(matchingChoiceById.value);
                } else { // If not matched by ID, try to match by the textual value itself (less common for QTI)
                  const matchingChoiceByValue = choicesMap.find(c => this.getTextContent(valueEl)?.trim().toLowerCase() === c.value.toLowerCase());
                  if (matchingChoiceByValue && !correctAnswerValues.includes(matchingChoiceByValue.value)) {
                    correctAnswerValues.push(matchingChoiceByValue.value);
                  }
                }
              } else if (questionType === 'TEXT') {
                if (!correctAnswerValues.includes(correctIdentifierOrValue)) {
                  correctAnswerValues.push(correctIdentifierOrValue);
                }
              }
            }
          });
        }
      }
    }

    if (correctAnswerValues.length > 0 && points === 0 && !pointsExplicitlySet) {
      points = 1; // Default to 1 point if answers found but no points explicitly set or derived
    }


    if (correctAnswerValues.length > 0) {
      return {points: points, correctAnswerValues: correctAnswerValues};
    }
    return null;
  }

  private getTextContent(element: Element | null, excludeImageAltText = false): string {
    if (!element) return '';
    // Clone the element to avoid modifying the original DOM during text extraction
    let clonedElement = element.cloneNode(true) as Element;

    // If requested, remove all <img> tags from the clone before text extraction
    // to prevent their 'alt' text from being included.
    if (excludeImageAltText) {
      clonedElement.querySelectorAll('img').forEach(img => img.remove());
    }

    let textAccumulator: string[] = [];

    function extractText(node: Node) {
      if (node.nodeType === Node.TEXT_NODE) {
        textAccumulator.push(node.textContent || '');
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as Element;
        const tagName = el.tagName.toLowerCase();
        // Handle block-level elements or <br> by adding newlines appropriately
        if (tagName === 'br') {
          textAccumulator.push('\n');
        } else if (tagName === 'p' && textAccumulator.length > 0 && !textAccumulator[textAccumulator.length - 1].endsWith('\n\n')) {
          // Ensure a blank line before a new paragraph if not already there
          if (textAccumulator.length > 0 && !textAccumulator[textAccumulator.length - 1].endsWith('\n')) {
            textAccumulator.push('\n'); // Add one newline if previous didn't end with one
          }
          textAccumulator.push('\n'); // Add second newline for paragraph break
        }

        // Recursively process child nodes
        // For <img> tags, if excludeImageAltText is true, they've already been removed.
        // If false, their alt text (if rendered as text by browser) would be picked up if not handled specially.
        // However, standard .textContent doesn't include alt text. This custom walk aims for visual text.
        for (let i = 0; i < el.childNodes.length; i++) {
          extractText(el.childNodes[i]);
        }

        // Add a space after certain block-like elements if text doesn't already end with whitespace/newline
        // This helps separate words that might be joined if tags are immediately consecutive.
        if (['div', 'li', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'td', 'th'].includes(tagName)) {
          if (textAccumulator.length > 0 && !textAccumulator[textAccumulator.length - 1].match(/(\s|\n)$/)) {
            textAccumulator.push(' ');
          }
        }
      }
    }

    extractText(clonedElement);
    let rawText = textAccumulator.join('');

    // Decode HTML entities
    let decodedText = rawText;
    try {decodedText = decode(decodedText);} // Using 'html-entities' library
    catch (libError) {/* console.warn("[getTextContent] html-entities decoding failed.", libError); */}

    // Normalize multiple newlines/whitespace
    decodedText = decodedText.replace(/\n\s*\n/g, '\n\n'); // Collapse multiple blank lines into one
    let plainText = decodedText.replace(/<[^>]*>/g, ' '); // Strip any remaining HTML tags (should be minimal after DOM walk)
    plainText = plainText.replace(/[\s\u00A0]+/g, ' ').trim(); // Normalize spaces, non-breaking spaces, and trim

    // Basic manual entity replacement for common cases if decode didn't catch them or for safety.
    // Note: 'decode' should handle these, but this is a fallback.
    plainText = plainText.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

    return plainText.trim();
  }

  private _resolveImagePath(imgSrc: string, qtiFilePath: string): string | null {
    if (!imgSrc) return null;
    let relativePath = imgSrc;
    // Default base for resolution is the directory of the QTI file itself
    let baseForResolution = this.utils.getDirectory(qtiFilePath);

    if (imgSrc.startsWith('$IMS-CC-FILEBASE$')) {
      // Path is relative to the root of the IMSCC package
      relativePath = imgSrc.substring('$IMS-CC-FILEBASE$'.length);
      if (relativePath.startsWith('/')) { // Remove leading slash if present
        relativePath = relativePath.substring(1);
      }
      baseForResolution = ""; // Resolve from package root
    } else if (imgSrc.startsWith('/')) {
      // Absolute path from the perspective of the "website" hosting the QTI, treat as relative to package root in IMSCC context
      relativePath = imgSrc.substring(1);
      baseForResolution = ""; // Resolve from package root
    }
    // If imgSrc is already a relative path (e.g., "images/pic.jpg"), baseForResolution (QTI file's dir) will be used.

    const decodedRelativePath = this.utils.tryDecodeURIComponent(relativePath);
    return this.utils.resolveRelativePath(baseForResolution, decodedRelativePath);
  }
}
