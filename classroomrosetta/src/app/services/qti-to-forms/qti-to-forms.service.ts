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
import {switchMap, catchError, map, tap, mergeMap} from 'rxjs/operators';
import {ImsccFile, DriveFile, Material} from '../../interfaces/classroom-interface'; // Assuming path is correct
import {
  GoogleForm, FormInfo, FormItem, QuestionItem, Question, Option,
  BatchUpdateFormRequest, FormRequest, BatchUpdateFormResponse, Image as FormsImage,
} from '../../interfaces/forms-interface'; // Assuming path is correct
import {UtilitiesService, RetryConfig} from '../utilities/utilities.service'; // Assuming path is correct
import {FileUploadService} from '../file-upload/file-upload.service'; // Assuming path is correct
import {AuthService} from '../auth/auth.service'; // Import AuthService
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
  private auth = inject(AuthService); // Inject AuthService
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
      console.error('[QtiToFormsService] Cannot create API headers: Access token is missing.');
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
    itemId: string, // No longer takes accessToken
    parentFolderId: string
  ): Observable<Material | null> {
    if (!qtiFile?.data || typeof qtiFile.data !== 'string') return throwError(() => new Error('[QTI Service] QTI file data is missing or not a string.'));
    if (!allPackageFiles) return throwError(() => new Error('[QTI Service] Package files array is required to resolve resources.'));
    if (!itemId) return throwError(() => new Error('[QTI Service] Item ID (itemId) is required.'));
    if (!formTitle) return throwError(() => new Error('[QTI Service] Form title cannot be empty.'));
    if (!parentFolderId) return throwError(() => new Error('[QTI Service] Parent folder ID is required.'));
    if (!this.APPS_SCRIPT_RUN_ENDPOINT) return throwError(() => new Error('[QTI Service] Apps Script Form API URL (scripts.run endpoint) is not configured.'));

    console.log(`[QTI Service] Starting QTI to Form conversion for item: ${itemId}, title: "${formTitle}"`);

    // Headers will be created just-in-time by helper for each API call.
    const retryConfig: RetryConfig = {maxRetries: 3, initialDelayMs: 2000};

    return from(this.utils.generateHash(itemId)).pipe(
      catchError(hashError => {
        console.error(`[QTI Service] Error generating hash for itemId "${itemId}":`, hashError);
        return throwError(() => new Error(`[QTI Service] Failed to generate identifier hash. ${hashError.message || hashError}`));
      }),
      switchMap(hashedItemId => {
        const driveApiHeaders = this.createApiHeaders(); // For Drive search
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

              // _uploadImagesAndBuildFormRequests will handle its own token for FileUploadService
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
                        function: "createFormItemsInGoogleForm",
                        parameters: [createdForm.formId, qtiRequests],
                        devMode: false // Set to true for testing with latest Apps Script code if needed
                      };

                      console.log('[QTI Service] Calling Apps Script API to add items. Endpoint:', this.APPS_SCRIPT_RUN_ENDPOINT);
                      // console.debug('[QTI Service] Apps Script Payload:', JSON.stringify(appsScriptRunPayload, null, 2)); // Verbose

                      const addItemsViaAppsScript$ = this.http.post<AppsScriptRunResponse>(
                        this.APPS_SCRIPT_RUN_ENDPOINT,
                        appsScriptRunPayload,
                        {headers: appsScriptApiHeadersForRun}
                      );

                      return this.utils.retryRequest(addItemsViaAppsScript$, retryConfig, `Add Items to Form ${createdForm.formId} via Apps Script`).pipe(
                        map(appsScriptRunResponse => {
                          // console.log(`[QTI Service] Apps Script API raw response for form ${createdForm.formId}:`, JSON.stringify(appsScriptRunResponse, null, 2)); // Verbose
                          if (appsScriptRunResponse.error) {
                            console.error(`[QTI Service] Apps Script execution error for form ${createdForm.formId}:`, appsScriptRunResponse.error);
                            throw new Error(`Apps Script execution failed: ${appsScriptRunResponse.error.message || 'Unknown Apps Script error'}`);
                          }
                          if (appsScriptRunResponse.response && appsScriptRunResponse.response.result) {
                            const result = appsScriptRunResponse.response.result as AppsScriptFormUpdateResponse;
                            console.log(`[QTI Service] Apps Script function result for form ${createdForm.formId}: Success=${result.success}, Msg=${result.message}, Created=${result.createdItems}`);
                            if (!result.success) {
                              console.warn(`[QTI Service] Apps Script function reported failure for form ${createdForm.formId}: ${result.message}`, result.errors);
                              // Potentially throw an error here if partial success is not acceptable
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
                      if (!finalForm?.formId) return of(null); // Should have been caught earlier
                      const formId = finalForm.formId;

                      const driveApiHeadersForUpdate = this.createApiHeaders('application/json');
                      if (!driveApiHeadersForUpdate) {
                        console.warn(`[QTI Service] Token missing for Drive properties update for Form ${formId}. Skipping update.`);
                        // Return the form material without Drive properties updated if token fails here.
                        return of({form: {formUrl: finalForm.responderUri || `https://docs.google.com/forms/d/${formId}/viewform`, title: finalForm.info?.title || formTitle}} as Material);
                      }

                      const driveUpdateUrl = `${this.utils.DRIVE_API_FILES_ENDPOINT}/${formId}`;
                      const driveUpdateBody = {appProperties: {[this.APP_PROPERTY_KEY]: hashedItemId}};
                      // Moving from 'root' to parentFolderId. If it's already in parentFolderId, this might be redundant but generally safe.
                      const driveUpdateParams = new HttpParams().set('addParents', parentFolderId).set('removeParents', 'root');
                      const driveUpdateRequest$ = this.http.patch<DriveFile>(driveUpdateUrl, driveUpdateBody, {headers: driveApiHeadersForUpdate, params: driveUpdateParams});

                      return this.utils.retryRequest(driveUpdateRequest$, retryConfig, `Update Drive Props for Form ${formId}`).pipe(
                        map(driveFile => ({form: {formUrl: finalForm.responderUri || `https://docs.google.com/forms/d/${formId}/viewform`, title: driveFile.name || finalForm.info?.title || formTitle}} as Material)),
                        catchError(driveErr => {
                          console.error(`[QTI Service] Error updating Drive properties for Form ${formId}: ${this.utils.formatHttpError(driveErr)}`);
                          // Return the form material even if Drive properties update fails, as the form itself exists.
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
            return of(null); // If search fails, treat as form not found and proceed to create (or return null if that's desired behavior for search failure)
          })
        );
      })
    );
  }

  private _uploadImagesAndBuildFormRequests(
    intermediateItems: IntermediateFormItemDefinition[],
    // accessToken: string, // Removed accessToken parameter
    parentFolderIdForImages: string,
  ): Observable<FormRequest[]> {
    if (intermediateItems.length === 0) {
      return of([]);
    }

    const itemsThatNeedImageUpload = intermediateItems.filter(item => !!item.imageFileToUpload);
    if (itemsThatNeedImageUpload.length === 0) {
      // No images to upload, directly build requests
      return of(this._buildFormRequestsFromIntermediate(intermediateItems, new Map()));
    }

    const filesToBatchUpload: Array<{file: ImsccFile; targetFileName: string}> = itemsThatNeedImageUpload.map(itemDef => ({
      file: itemDef.imageFileToUpload!,
      targetFileName: itemDef.imageFileToUpload!.name // Use original name as target, FileUploadService handles if it exists
    }));

    console.log(`[QTI Service] Attempting to upload ${filesToBatchUpload.length} images to Drive folder ${parentFolderIdForImages}.`);

    // FileUploadService.uploadLocalFiles now fetches its own token
    return this.fileUploadService.uploadLocalFiles(filesToBatchUpload, parentFolderIdForImages).pipe(
      map(uploadedDriveFiles => {
        console.log(`[QTI Service] Successfully processed ${uploadedDriveFiles.length} of ${filesToBatchUpload.length} image uploads (some may have been found existing).`);
        const driveFileMap = new Map<string, DriveFile>();

        uploadedDriveFiles.forEach(driveFile => {
          if (driveFile && driveFile.name) {
            // Attempt to map back using targetFileName which should be unique if FileUploadService renames.
            // If FileUploadService returns the *original* name in DriveFile.name, this is simpler.
            // Assuming FileUploadService returns DriveFile.name as the name it has in Drive (which was targetFileName).
            const originalImsccFile = filesToBatchUpload.find(f => f.targetFileName === driveFile.name)?.file;
            if (originalImsccFile) {
              driveFileMap.set(originalImsccFile.name, driveFile); // Map by original IMSCC file name
            } else {
              console.warn(`[QTI Service] Could not map uploaded/found Drive file "${driveFile.name}" back to an original ImsccFile name.`);
            }
          }
        });
        return this._buildFormRequestsFromIntermediate(intermediateItems, driveFileMap);
      }),
      catchError(uploadError => {
        console.error('[QTI Service] Batch image upload/find process failed:', uploadError);
        console.warn('[QTI Service] Proceeding to build form requests without any uploaded/found image URIs.');
        // Still try to build requests, images will be missing
        return of(this._buildFormRequestsFromIntermediate(intermediateItems, new Map()));
      })
    );
  }

  // _buildFormRequestsFromIntermediate, _parseQtiToIntermediateItems, parseStandardQtiItems,
  // parseResponseProcessing, getTextContent, _resolveImagePath remain the same internally
  // as they don't make direct HTTP calls needing tokens.

  private _buildFormRequestsFromIntermediate(
    intermediateItems: IntermediateFormItemDefinition[],
    driveFileMap: Map<string, DriveFile> // Maps original ImsccFile.name to DriveFile
  ): FormRequest[] {
    const formRequests: FormRequest[] = [];
    intermediateItems.forEach(itemDef => {
      let formItem: FormItem | null = null;
      let imageForForm: FormsImage | undefined = undefined;

      if (itemDef.imageFileToUpload) {
        // Use the original image file name (from IMSCC) to look up in the map
        const driveFile = driveFileMap.get(itemDef.imageFileToUpload.name);
        let driveImageUri: string | null = null;
        let altTextForImageObject = "Image"; // Default alt text

        // Determine alt text for the FormsImage object
        if (itemDef.imageAltText && !this.isLikelyFilename(itemDef.imageAltText)) {
          altTextForImageObject = itemDef.imageAltText;
        } else if (itemDef.imageAltText && this.isLikelyFilename(itemDef.imageAltText)) {
          // console.log(`[QTI Service] imageAltText "${itemDef.imageAltText}" for image file "${itemDef.imageFileToUpload.name}" appears to be a filename. Using generic alt text for Forms API image object.`);
        }


        if (driveFile) {
          // console.log(`[QTI Service] Processing DriveFile for image "${itemDef.imageFileToUpload.name}": ID=${driveFile.id}, Name=${driveFile.name}`);
          // Prefer direct link if available, otherwise construct one.
          // Note: Google Forms might be picky about direct image URLs. Publicly accessible URLs are best.
          // A direct download link like 'uc?id=...' might work if the image is public or accessible to the Form.
          // Thumbnail links are often smaller and might be better for performance if quality is acceptable.
          if (driveFile.thumbnailLink) { // Check if thumbnailLink is a valid, usable URL
             driveImageUri = driveFile.thumbnailLink.replace(/=s\d+$/, ''); // Remove size parameter for potentially larger image
             // console.log(`[QTI Service] Using thumbnailLink (cleaned) for ${driveFile.name}: ${driveImageUri}`);
          } else if (driveFile.id) { // Fallback to constructing a download link
            // This link forces a download, might not be ideal for direct embedding in Forms if it doesn't resolve to an image preview.
            // A webViewLink might be better if it leads to a viewable image, but often it's an HTML page.
            // The direct /uc?id= link is often used for direct image access if permissions allow.
            driveImageUri = `https://drive.google.com/uc?id=${driveFile.id}`; // Simpler UC link
            console.log(`[QTI Service] ThumbnailLink missing for ${driveFile.name}. Using constructed 'uc?id=' link: ${driveImageUri}. Ensure file is publicly viewable or accessible by Forms.`);
          } else {
            console.warn(`[QTI Service] DriveFile for "${itemDef.imageFileToUpload.name}" is missing both thumbnailLink and ID. Cannot generate image URI.`);
          }
        } else if (typeof itemDef.imageFileToUpload.data === 'string' && itemDef.imageFileToUpload.data.startsWith('data:image')) {
          // If it's a base64 data URI (e.g., from unzipping or already processed)
          driveImageUri = itemDef.imageFileToUpload.data;
          // console.log(`[QTI Service] Using existing base64 data URI for ${itemDef.imageFileToUpload.name} as no DriveFile mapping found or Drive upload failed.`);
        } else {
          console.warn(`[QTI Service] No DriveFile object found in map for image: "${itemDef.imageFileToUpload.name}", and data is not a base64 URI.`);
        }

        if (driveImageUri) {
          imageForForm = {sourceUri: driveImageUri, altText: altTextForImageObject};
          // console.log(`[QTI Service] Prepared FormsImage with sourceUri and altText for item titled: "${itemDef.title || 'untitled item'}"`);
        } else {
          console.warn(`[QTI Service] Image "${itemDef.imageFileToUpload.name}" for item "${itemDef.title}" could not be processed for a usable URI. Item will be created without this image.`);
        }
      }

      // Build the FormItem based on itemDef.type
      if (itemDef.type === 'question' && itemDef.question) {
        const questionItem: QuestionItem = {question: itemDef.question};
        if (imageForForm) {
          questionItem.image = imageForForm; // Add image to the question item
        }
        formItem = {title: itemDef.title, description: itemDef.description, questionItem: questionItem};
      } else if (itemDef.type === 'image_standalone') {
        let standaloneImageTitle = itemDef.title;
        // Use alt text as title if title is missing or looks like a filename
        if (!standaloneImageTitle && itemDef.imageAltText && !this.isLikelyFilename(itemDef.imageAltText)) {
          standaloneImageTitle = itemDef.imageAltText;
        } else if (!standaloneImageTitle || this.isLikelyFilename(standaloneImageTitle)) {
          standaloneImageTitle = "Image"; // Default title for standalone image
        }

        if (imageForForm) { // Only create item if image was successfully processed
          formItem = {
            title: standaloneImageTitle, // Use the determined title
            description: (itemDef.description && !this.isLikelyFilename(itemDef.description)) ? itemDef.description : undefined,
            imageItem: {image: imageForForm}
          };
        } else {
          console.warn(`[QTI Service] Standalone image (original src: ${itemDef.originalImgSrc}) could not be processed. Skipping this image item.`);
        }
      } else if (itemDef.videoItem) {
        formItem = {title: itemDef.title, description: itemDef.description, videoItem: itemDef.videoItem};
      } else if (itemDef.pageBreakItem) {
        formItem = {title: itemDef.title, description: itemDef.description, pageBreakItem: itemDef.pageBreakItem};
      } else if (itemDef.sectionHeaderItem) {
        formItem = {title: itemDef.title, description: itemDef.description, sectionHeaderItem: itemDef.sectionHeaderItem};
      }
      // Add other item types (gridItem etc.) if needed

      if (formItem) {
        formRequests.push({createItem: {item: formItem, location: {index: formRequests.length}}});
      }
    });
    console.log(`[QTI Service] Built ${formRequests.length} final FormRequest objects.`);
    return formRequests;
  }


  private _parseQtiToIntermediateItems(
    qtiDoc: XMLDocument,
    qtiFilePath: string,
    allPackageFiles: ImsccFile[]
  ): IntermediateFormItemDefinition[] {
    const intermediateItems: IntermediateFormItemDefinition[] = [];
    // console.log(`[QTI HTML Parser] Parsing QTI from "${qtiFilePath}"...`);
    const itemElement = qtiDoc.querySelector('item, assessmentItem'); // Top-level item
    if (!itemElement) {
        // console.warn(`[QTI Service] No <item> or <assessmentItem> found in "${qtiFilePath}". Trying to parse individual itemBody elements if any.`);
        // Fallback for QTI files that might just be a list of itemBody elements without a main <item>
        const itemBodies = Array.from(qtiDoc.querySelectorAll('itemBody'));
        if (itemBodies.length > 0) {
            itemBodies.forEach((bodyEl, idx) => {
                const itemsFromFragment = this.parseHtmlContentWithinQti(bodyEl, qtiFilePath, allPackageFiles, `fragment_${idx}`);
                intermediateItems.push(...itemsFromFragment);
            });
            return intermediateItems;
        }
        return intermediateItems; // Still empty if no item/assessmentItem and no itemBody
    }


    // Try to find HTML content first (Blackboard style)
    const mattextElement = itemElement.querySelector('presentation > material > mattext[texttype="text/html"], itemBody > div > mattext[texttype="text/html"], itemBody > mattext[texttype="text/html"]');
    if (mattextElement && mattextElement.textContent) {
      const itemsFromHtml = this.parseHtmlContentWithinQti(mattextElement, qtiFilePath, allPackageFiles, itemElement.getAttribute('ident') || itemElement.getAttribute('identifier') || 'html_qti_item');
      intermediateItems.push(...itemsFromHtml);
      // If HTML parsing yields items, we might assume it's the primary content.
      // However, check if there are also standard QTI interactions outside this HTML block.
      // This is a complex case; for now, prioritize HTML if it yields content.
      if (intermediateItems.length > 0) {
        return intermediateItems;
      }
    }

    // If no items from HTML or no HTML content, parse as standard QTI
    const standardItems = this.parseStandardQtiItems(itemElement, qtiFilePath, allPackageFiles);
    intermediateItems.push(...standardItems);

    return intermediateItems;
  }

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
    let questionCounter = 0; // Counter for questions derived from this HTML block

    bodyChildren.forEach((element: Element, elIdx: number) => {
        const currentIdentifier = `${baseIdentifier}_html_${elIdx}`;
        if (element.tagName.toLowerCase() === 'img') { // Top-level image
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
            const textContentFromParent = this.getTextContent(element, true).trim(); // Exclude alt text of descendant images

            if (imagesInElement.length > 0) {
                imagesInElement.forEach((imgElement, imgIdx) => {
                    questionCounter++;
                    const imgSrc = imgElement.getAttribute('src');
                    const rawAltText = imgElement.getAttribute('alt');
                    let imageFileToUpload: ImsccFile | undefined;
                    if (imgSrc) {
                        const imagePath = this._resolveImagePath(imgSrc, qtiFilePath);
                        if (imagePath) imageFileToUpload = allPackageFiles.find(f => f.name.toLowerCase() === imagePath.toLowerCase());
                    }

                    let titleForIntermediateItem = textContentFromParent || // Use surrounding text if available
                        (rawAltText && !this.isLikelyFilename(rawAltText) ? rawAltText : `Image Question ${questionCounter}`);
                    if (this.isLikelyFilename(titleForIntermediateItem)) titleForIntermediateItem = `Image Question ${questionCounter}`;

                    let descriptionForIntermediateItem = "";
                     if (rawAltText && !this.isLikelyFilename(rawAltText) && rawAltText !== titleForIntermediateItem) {
                        descriptionForIntermediateItem = rawAltText;
                    }

                    intermediateItems.push({
                        type: 'question', // Assume image implies a question if not standalone
                        title: titleForIntermediateItem,
                        description: descriptionForIntermediateItem,
                        question: {textQuestion: {paragraph: false}}, // Default to short answer for image questions from HTML
                        imageFileToUpload: imageFileToUpload,
                        originalImgSrc: imgSrc || undefined,
                        imageAltText: rawAltText || (imageFileToUpload ? imageFileToUpload.name : '')
                    });
                });
            } else if (textContentFromParent) { // Text-only paragraph/div
                questionCounter++;
                let qText = textContentFromParent;
                if (!textContentFromParent.match(/^\s*\d+[\.\)]\s*/)) { // Add numbering if not already present
                    qText = `${questionCounter}. ${textContentFromParent}`;
                }
                intermediateItems.push({
                    type: 'question',
                    title: qText,
                    description: undefined,
                    question: {textQuestion: {paragraph: true}}, // Assume paragraph for text blocks from HTML
                    imageAltText: undefined
                });
            }
        }
    });
    return intermediateItems;
  }


  private parseStandardQtiItems(itemElement: Element, sourceFileName: string, allPackageFiles: ImsccFile[]): IntermediateFormItemDefinition[] {
    const intermediateItems: IntermediateFormItemDefinition[] = [];
    // console.log(`[Standard QTI Parser] Parsing standard QTI item from "${sourceFileName}"...`);

    const itemIdentifier = itemElement.getAttribute('ident') || itemElement.getAttribute('identifier') || `qti_item_${Date.now()}`;
    let itemTitle = itemElement.getAttribute('title') || '';

    const itemBody = itemElement.querySelector('itemBody');
    const presentation = itemElement.querySelector('presentation'); // QTI 1.2

    // Extract question text (prompt) and description
    let questionText = '';
    let itemDescription: string | undefined = undefined;
    let imageFileToUpload: ImsccFile | undefined;
    let originalImgSrc: string | undefined;
    let imageAltText: string | undefined;

    const promptElement = itemBody?.querySelector('prompt') || presentation?.querySelector('prompt');
    if (promptElement) {
        questionText = this.getTextContent(promptElement).trim();
    }

    // Look for material/mattext as the primary question content if prompt is empty or not specific enough
    const materialContainer = itemBody || presentation; // Check both
    if (materialContainer) {
        const matElements = Array.from(materialContainer.querySelectorAll('material > mattext[texttype="text/html"], material > mattext, mattext[texttype="text/html"], mattext'));
        let tempHtmlForImageExtraction = "";
        matElements.forEach(mat => {
            const content = mat.textContent || "";
            tempHtmlForImageExtraction += content;
            if (!questionText && !this.isLikelyFilename(this.getTextContent(mat).trim())) { // Prioritize non-filename text for question
                questionText = this.getTextContent(mat).trim();
            }
        });

        // Image extraction from HTML content within material/mattext
        if (tempHtmlForImageExtraction) {
            const parser = new DOMParser();
            const tempDoc = parser.parseFromString(decode(tempHtmlForImageExtraction), "text/html");
            const imgTag = tempDoc.querySelector('img');
            if (imgTag) {
                originalImgSrc = imgTag.getAttribute('src') || undefined;
                imageAltText = imgTag.getAttribute('alt') || undefined;
                if (originalImgSrc) {
                    const imagePath = this._resolveImagePath(originalImgSrc, sourceFileName);
                    if (imagePath) {
                        imageFileToUpload = allPackageFiles.find(f => f.name.toLowerCase() === imagePath.toLowerCase());
                        if (!imageFileToUpload) console.warn(`   [QTI Standard] Image file not found for src: ${originalImgSrc} (resolved: ${imagePath})`);
                    }
                }
                // If questionText is still empty and alt text is descriptive, use it.
                if (!questionText && imageAltText && !this.isLikelyFilename(imageAltText)) {
                    questionText = imageAltText;
                }
            }
        }
    }

    if (!questionText && itemTitle) questionText = itemTitle; // Fallback to item title if no other text found
    if (!questionText && !itemTitle && imageFileToUpload) { // If only an image, use a default title
        questionText = imageAltText && !this.isLikelyFilename(imageAltText) ? imageAltText : "Image-based question";
    }

    if (!questionText) {
      console.warn(`   Skipping standard QTI item ${itemIdentifier}: No question text or title found.`);
      return intermediateItems; // Skip this item if no text at all
    }
    const cleanQuestionText = questionText.replace(/[\n\r\t]+/g, ' ').replace(/\s\s+/g, ' ').trim();

    // Extract description from metadata or specific elements
    const descriptionMetaElement = itemElement.querySelector('itemmetadata > qtimetadata > qti_metadatafield[fieldlabel="qmd_description"] > fieldentry');
    if (descriptionMetaElement) {
        itemDescription = this.getTextContent(descriptionMetaElement).trim();
    } else {
        const rubricBlock = itemBody?.querySelector('rubricBlock'); // Common place for instructions/description
        if (rubricBlock) itemDescription = this.getTextContent(rubricBlock).trim();
    }


    // Determine question type and parse choices/grading
    const choiceInteraction = itemElement.querySelector('itemBody choiceInteraction, choiceInteraction'); // QTI 2.1+ and QTI 1.2
    const textEntryInteraction = itemElement.querySelector('itemBody textEntryInteraction, textEntryInteraction');
    const extendedTextInteraction = itemElement.querySelector('itemBody extendedTextInteraction, extendedTextInteraction');

    // QTI 1.2 specific response elements
    const responseLid = presentation?.querySelector('response_lid'); // QTI 1.2
    const renderChoice = responseLid?.querySelector('render_choice'); // QTI 1.2 for choices
    const responseStr = presentation?.querySelector('response_str'); // QTI 1.2 for text entry

    let question: Question | undefined;
    const allParsedChoices: ParsedChoice[] = [];
    let questionTypeForGrading: 'CHOICE' | 'TEXT' | 'UNKNOWN' = 'UNKNOWN';

    if (choiceInteraction || renderChoice) {
      questionTypeForGrading = 'CHOICE';
      const interactionElement = choiceInteraction || renderChoice!;
      const maxChoices = choiceInteraction?.getAttribute('maxChoices');
      const isCheckbox = maxChoices !== '0' && maxChoices !== '1' || renderChoice?.getAttribute('shuffle') === 'yes' || responseLid?.getAttribute('cardinality') === 'Multiple';
      const choiceTypeValue: 'RADIO' | 'CHECKBOX' = isCheckbox ? 'CHECKBOX' : 'RADIO';
      const options: Option[] = [];
      const choices = Array.from(interactionElement.querySelectorAll('simpleChoice, response_label > render_choice > response_label')); // QTI 2.1 and 1.2

      choices.forEach(choice => {
        let choiceTextVal = '';
        let choiceId = choice.getAttribute('identifier') || (choice.tagName.toLowerCase() === 'response_label' ? choice.getAttribute('ident') : null);

        // For simpleChoice, text is direct child or in flow. For QTI 1.2 response_label, it's often in a following material/mattext.
        let textEl = choice;
        if (choice.tagName.toLowerCase() === 'response_label') {
            textEl = choice.querySelector('material > mattext, mattext') || choice;
        }

        choiceTextVal = this.getTextContent(textEl).trim();

        if (choiceTextVal && choiceId) {
          if (!options.find(o => o.value === choiceTextVal)) { // Avoid duplicate option values
            options.push({value: choiceTextVal});
            allParsedChoices.push({identifier: choiceId, value: choiceTextVal});
          }
        }
      });
      if (options.length > 0) {
        question = {choiceQuestion: {type: choiceTypeValue, options: options, shuffle: renderChoice?.getAttribute('shuffle') === 'yes'}};
      }
    } else if (textEntryInteraction || extendedTextInteraction || responseStr) {
      questionTypeForGrading = 'TEXT';
      const isParagraph = !!extendedTextInteraction || (!!responseStr && presentation?.querySelector('response_str > render_fib[rows]') !== null);
      question = {textQuestion: {paragraph: isParagraph}};
    }

    if (question) {
      const gradingInfo = this.parseResponseProcessing(itemElement, questionTypeForGrading, allParsedChoices);
      if (gradingInfo && gradingInfo.correctAnswerValues.length > 0 && !(question.textQuestion?.paragraph)) {
        question.required = true; // Typically make gradable questions required
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
    } else if (imageFileToUpload) { // If it's just an image with no clear question structure from QTI
        intermediateItems.push({
            type: 'image_standalone',
            title: cleanQuestionText || (imageAltText && !this.isLikelyFilename(imageAltText) ? imageAltText : "Image"),
            description: itemDescription || undefined,
            imageFileToUpload: imageFileToUpload,
            originalImgSrc: originalImgSrc,
            imageAltText: imageAltText
        });
    } else {
      console.warn(`   Skipping standard QTI item ${itemIdentifier}: Could not determine valid question structure and no standalone image found.`);
    }
    return intermediateItems;
  }


  private parseResponseProcessing(item: Element, questionType: 'CHOICE' | 'TEXT' | 'UNKNOWN', choicesMap?: ParsedChoice[]): ParsedGradingInfo | null {
    if (questionType === 'UNKNOWN') return null;
    const respProcessing = item.querySelector('resprocessing, responseProcessing'); // QTI 1.2 and 2.1
    if (!respProcessing) return null;

    let points = 1; // Default points
    const correctAnswerValues: string[] = [];

    // Attempt to find points from qmd_weighting (common in QTI 1.2)
    const weightMetaField = Array.from(item.querySelectorAll('itemmetadata > qtimetadata > qti_metadatafield'))
        .find(field => field.querySelector('fieldlabel')?.textContent?.trim().toLowerCase() === 'qmd_weighting');
    if (weightMetaField) {
        const weightEntry = weightMetaField.querySelector('fieldentry');
        if (weightEntry?.textContent) {
            points = Math.round(parseFloat(weightEntry.textContent.trim())) || 1;
        }
    }

    // QTI 1.2 style: <respcondition> with <setvar> for score
    const respconditions = Array.from(respProcessing.querySelectorAll('respcondition'));
    respconditions.forEach(condition => {
      const setvar = condition.querySelector('setvar');
      const scoreFromSetvar = parseFloat(setvar?.textContent || '0');
      const action = setvar?.getAttribute('action');

      if (action === 'Set' && scoreFromSetvar > 0) {
        // If points haven't been set by qmd_weighting or it's different, update.
        // This gives precedence to explicit scoring in respcondition if it's higher than default.
        if (points === 1 && scoreFromSetvar !== 1) points = Math.round(scoreFromSetvar);
        else if (scoreFromSetvar > points) points = Math.round(scoreFromSetvar);


        const conditionVar = condition.querySelector('conditionvar');
        const varequal = conditionVar?.querySelector('varequal'); // Correct answer identifier or value
        if (varequal) {
          const correctIdentifierOrValue = varequal.textContent?.trim();
          if (correctIdentifierOrValue) {
            if (questionType === 'CHOICE' && choicesMap) {
              const matchingChoice = choicesMap.find(c => c.identifier === correctIdentifierOrValue);
              if (matchingChoice && !correctAnswerValues.includes(matchingChoice.value)) {
                correctAnswerValues.push(matchingChoice.value);
              }
            } else if (questionType === 'TEXT') { // Direct value for text questions
              const cleanCorrectValue = this.getTextContent(varequal)?.trim(); // Get text content properly
              if (cleanCorrectValue && !correctAnswerValues.includes(cleanCorrectValue)) {
                correctAnswerValues.push(cleanCorrectValue);
              }
            }
          }
        }
      }
    });

    // QTI 2.1 style: <responseDeclaration> with <correctResponse>
    // This might also be a fallback if respconditions didn't yield answers.
    if (correctAnswerValues.length === 0) {
        const responseDeclarations = Array.from(item.querySelectorAll('responseDeclaration'));
        // Try to find the primary response declaration, often named "RESPONSE" or the one with mapping/correctResponse
        const primaryResponseDecl = responseDeclarations.find(rd => rd.getAttribute('identifier')?.toUpperCase() === 'RESPONSE') ||
                                   responseDeclarations.find(rd => rd.querySelector('correctResponse') || rd.querySelector('mapping')) ||
                                   responseDeclarations[0];

        if (primaryResponseDecl) {
            const mapping = primaryResponseDecl.querySelector('mapping');
            if (mapping?.hasAttribute('defaultValue') && parseFloat(mapping.getAttribute('defaultValue') || '0') > 0) {
                // If points from qmd_weighting was 1 (default), override with mapping defaultValue if it's a score
                if (points === 1) points = Math.round(parseFloat(mapping.getAttribute('defaultValue')!)) || points;
            }

            const correctResponse = primaryResponseDecl.querySelector('correctResponse');
            if (correctResponse) {
                correctResponse.querySelectorAll('value').forEach(valueEl => {
                const correctIdentifierOrValue = this.getTextContent(valueEl)?.trim(); // Use getTextContent
                if (correctIdentifierOrValue) {
                    if (questionType === 'CHOICE' && choicesMap) {
                    // For QTI 2.1, value in correctResponse is usually the choice identifier
                    const matchingChoice = choicesMap.find(c => c.identifier === correctIdentifierOrValue);
                    if (matchingChoice && !correctAnswerValues.includes(matchingChoice.value)) {
                        correctAnswerValues.push(matchingChoice.value);
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


    if (correctAnswerValues.length > 0) {
      return {points: points > 0 ? points : 1, correctAnswerValues: correctAnswerValues}; // Ensure points is at least 1 if answers exist
    }
    return null;
  }

  private getTextContent(element: Element | null, excludeImageAltText = false): string {
    if (!element) return '';
    let clonedElement = element.cloneNode(true) as Element;
    if (excludeImageAltText) {
      clonedElement.querySelectorAll('img').forEach(img => img.remove());
    }
    // Try to get text content, preferring text nodes directly to avoid innerHTML issues with entities
    let text = '';
    if (clonedElement.childNodes.length > 0) {
        for (let i = 0; i < clonedElement.childNodes.length; i++) {
            const node = clonedElement.childNodes[i];
            if (node.nodeType === Node.TEXT_NODE) {
                text += node.textContent;
            } else if (node.nodeType === Node.ELEMENT_NODE && (node as Element).tagName.toLowerCase() !== 'img') { // Avoid alt from img if not already removed
                text += (node as Element).textContent; // Recursive textContent for child elements
            }
        }
    } else {
        text = clonedElement.textContent || '';
    }

    let decodedText = text;
    // Minimal decoding for common entities if DOM-based decoding isn't robust enough or fails
    try {decodedText = decode(decodedText); } // from html-entities
    catch (libError) {/* console.warn("[getTextContent] html-entities decoding failed.", libError); */}

    // Basic cleanup
    let plainText = decodedText.replace(/<[^>]*>/g, ' '); // Remove any remaining tags (should be minimal if textContent was effective)
    plainText = plainText.replace(/\s+/g, ' ').trim(); // Normalize whitespace
    // Final entity pass for common ones that might remain if decode didn't catch them or they were re-introduced
    plainText = plainText.replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
    return plainText.replace(/\s+/g, ' ').trim(); // One last normalize
  }

  private _resolveImagePath(imgSrc: string, qtiFilePath: string): string | null {
    if (!imgSrc) return null;
    let relativePath = imgSrc;
    let baseForResolution = this.utils.getDirectory(qtiFilePath); // e.g., "folder1/subfolder" if qtiFilePath is "folder1/subfolder/file.xml"

    // Handle $IMS-CC-FILEBASE$ placeholder which means path is relative to package root
    if (imgSrc.startsWith('$IMS-CC-FILEBASE$')) {
      relativePath = imgSrc.substring('$IMS-CC-FILEBASE$'.length);
      if (relativePath.startsWith('/')) { // If $IMS-CC-FILEBASE$/resource.gif
          relativePath = relativePath.substring(1);
      }
      baseForResolution = ""; // Path is from root
    } else if (imgSrc.startsWith('/')) {
        // Absolute path from package root (e.g. /images/pic.png)
        relativePath = imgSrc.substring(1);
        baseForResolution = "";
    }
    // If relativePath is still something like "../images/pic.png", baseForResolution will be used.
    // If relativePath is "images/pic.png", baseForResolution will be used.

    const decodedRelativePath = this.utils.tryDecodeURIComponent(relativePath);
    return this.utils.resolveRelativePath(baseForResolution, decodedRelativePath);
  }
}
