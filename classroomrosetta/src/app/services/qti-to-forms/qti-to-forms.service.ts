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

import {Injectable, inject} from '@angular/core';
import {HttpClient, HttpHeaders, HttpErrorResponse, HttpParams} from '@angular/common/http';
import {Observable, throwError, of, from, forkJoin, EMPTY} from 'rxjs';
import {switchMap, catchError, map, tap, mergeMap} from 'rxjs/operators';
import {ImsccFile, DriveFile, Material} from '../../interfaces/classroom-interface'; // Adjust path accordingly
import {
  GoogleForm, FormInfo, FormItem, QuestionItem, Question, ChoiceQuestion, Option, TextQuestion, Grading, CorrectAnswers, CorrectAnswer,
  BatchUpdateFormRequest, FormRequest, CreateItemRequest, Location, BatchUpdateFormResponse, UpdateSettingsRequest, ImageItem, Image as FormsImage,
  ScaleQuestion, DateQuestion, TimeQuestion, RowQuestion, VideoItem, PageBreakItem, SectionHeaderItem // Ensure all are imported
} from '../../interfaces/forms-interface'; // Adjust path accordingly
import {UtilitiesService, RetryConfig} from '../utilities/utilities.service'; // Adjust path accordingly
import {FileUploadService} from '../file-upload/file-upload.service';
import {decode} from 'html-entities';
import {AuthService} from '../auth/auth.service';
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
  title?: string; // For questions: actual question text (used for ImageItem title by Apps Script). For standalone items: their title.
  description?: string; // For questions: actual question description/help text (NOT filename). For standalone items: their description.
  question?: Question;
  imageFileToUpload?: ImsccFile;
  originalImgSrc?: string;
  imageAltText?: string; // Specifically the alt text from the <img> tag, for the ImageItem's own altText/helpText in Apps Script
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
  private authService = inject(AuthService);
  private readonly APP_PROPERTY_KEY = 'imsccIdentifier';

  private readonly APPS_SCRIPT_RUN_ENDPOINT = environment.formsItemsApi;

  constructor() {
    if (!this.APPS_SCRIPT_RUN_ENDPOINT) {
      console.error("FATAL: Apps Script Form API URL (scripts.run endpoint) is not configured in the environment file. Form item creation will fail.");
    }
  }

  /**
   * Helper function to check if a string is likely a filename based on common image extensions.
   */
  private isLikelyFilename(text: string | null | undefined): boolean {
    if (!text || typeof text !== 'string') return false;
    return /\.(jpeg|jpg|gif|png|svg|bmp|webp|tif|tiff)$/i.test(text.trim());
  }

  createFormFromQti(
    qtiFile: ImsccFile,
    allPackageFiles: ImsccFile[],
    formTitle: string,
    accessToken: string,
    itemId: string,
    parentFolderId: string
  ): Observable<Material | null> {
    if (!qtiFile?.data || typeof qtiFile.data !== 'string') return throwError(() => new Error('QTI file data is missing or not a string.'));
    if (!allPackageFiles) return throwError(() => new Error('Package files array is required to resolve resources.'));
    if (!itemId) return throwError(() => new Error('Item ID (itemId) is required.'));
    if (!formTitle) return throwError(() => new Error('Form title cannot be empty.'));
    if (!accessToken) return throwError(() => new Error('Access token is required.'));
    if (!parentFolderId) return throwError(() => new Error('Parent folder ID is required.'));
    if (!this.APPS_SCRIPT_RUN_ENDPOINT) return throwError(() => new Error('Apps Script Form API URL (scripts.run endpoint) is not configured.'));

    console.log(`[QTI Service] Starting QTI to Form conversion for item: ${itemId}, title: "${formTitle}"`);

    const formsApiHeaders = new HttpHeaders({'Authorization': `Bearer ${accessToken}`, 'Content-Type': 'application/json'});
    const driveApiHeaders = new HttpHeaders({'Authorization': `Bearer ${accessToken}`});
    const appsScriptApiHeaders = new HttpHeaders({
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    });
    const retryConfig: RetryConfig = { maxRetries: 3, initialDelayMs: 2000 };

    return from(this.utils.generateHash(itemId)).pipe(
      catchError(hashError => {
        console.error(`[QTI Service] Error generating hash for itemId "${itemId}":`, hashError);
        return throwError(() => new Error(`Failed to generate identifier hash. ${hashError.message || hashError}`));
      }),
      switchMap(hashedItemId => {
        console.log(`[QTI Service] Searching for existing Google Form with ${this.APP_PROPERTY_KEY}=${hashedItemId} in folder ${parentFolderId}...`);
        const searchQuery = `'${parentFolderId}' in parents and appProperties has { key='${this.APP_PROPERTY_KEY}' and value='${hashedItemId}' } and mimeType='application/vnd.google-apps.form' and trashed = false`;
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
                return throwError(() => new Error(`QTI parsing failed: ${errorMessage}`));
              }

              return this._uploadImagesAndBuildFormRequests(intermediateItems, accessToken, parentFolderId, retryConfig, itemId).pipe(
                switchMap(qtiRequests => {
                  if (qtiRequests.length === 0 && intermediateItems.length > 0) {
                    console.warn(`[QTI Service] No form item requests generated for "${formTitle}" after processing images, though intermediate items existed. Form might be empty or only have title/description.`);
                  } else if (qtiRequests.length === 0 && intermediateItems.length === 0) {
                    console.log(`[QTI Service] No content (questions/images) found in QTI to add to form "${formTitle}". Creating an empty form.`);
                  }

                  const formBody: {info: FormInfo} = {info: {title: formTitle, documentTitle: formTitle}};
                  const formCreateRequest$ = this.http.post<GoogleForm>(this.utils.FORMS_API_CREATE_ENDPOINT, formBody, {headers: formsApiHeaders});

                  return this.utils.retryRequest(formCreateRequest$, retryConfig, `Create Form "${formTitle}"`).pipe(
                    tap(createdForm => console.log(`[QTI Service] Initial Form created. ID: ${createdForm.formId}`)),
                    mergeMap(createdForm => { // Make Quiz
                      if (!createdForm?.formId) throw new Error('Form creation failed or did not return ID.');
                      const formId = createdForm.formId;
                      const makeQuizUrl = `${this.utils.FORMS_API_BATCHUPDATE_BASE_ENDPOINT}${formId}:batchUpdate`;
                      const makeQuizRequest: FormRequest = {updateSettings: {settings: {quizSettings: {isQuiz: true}}, updateMask: 'quizSettings.isQuiz'}};
                      const makeQuizBody: BatchUpdateFormRequest = {requests: [makeQuizRequest]};
                      const makeQuizHttpRequest$ = this.http.post<BatchUpdateFormResponse>(makeQuizUrl, makeQuizBody, {headers: formsApiHeaders});
                      return this.utils.retryRequest(makeQuizHttpRequest$, retryConfig, `Make Form ${formId} a Quiz`).pipe(map(() => createdForm), catchError(quizErr => throwError(() => new Error(`Failed to set form as quiz. ${this.utils.formatHttpError(quizErr)}`))));
                    }),
                    mergeMap(createdForm => { // Add Items via Apps Script
                      if (!createdForm?.formId) throw new Error('Form ID missing after quiz setup.');
                      if (qtiRequests.length === 0) {
                        console.log(`[QTI Service] No items to add to form ${createdForm.formId}.`);
                        return of(createdForm);
                      }

                      const appsScriptRunPayload: AppsScriptRunRequest = {
                        function: "createFormItemsInGoogleForm",
                        parameters: [createdForm.formId, qtiRequests],
                        devMode: false
                      };

                      console.log('[QTI Service] Calling Apps Script API to add items. Endpoint:', this.APPS_SCRIPT_RUN_ENDPOINT, 'Payload:', JSON.stringify(appsScriptRunPayload, null, 2));

                      const addItemsViaAppsScript$ = this.http.post<AppsScriptRunResponse>(
                        this.APPS_SCRIPT_RUN_ENDPOINT,
                        appsScriptRunPayload,
                        {headers: appsScriptApiHeaders}
                      );

                      return this.utils.retryRequest(addItemsViaAppsScript$, retryConfig, `Add Items to Form ${createdForm.formId} via Apps Script`).pipe(
                        map(appsScriptRunResponse => {
                          console.log(`[QTI Service] Apps Script API raw response for form ${createdForm.formId}:`, JSON.stringify(appsScriptRunResponse, null, 2));
                          if (appsScriptRunResponse.error) {
                            console.error(`[QTI Service] Apps Script execution error for form ${createdForm.formId}:`, appsScriptRunResponse.error);
                            throw new Error(`Apps Script execution failed: ${appsScriptRunResponse.error.message || 'Unknown Apps Script error'}`);
                          }
                          if (appsScriptRunResponse.response && appsScriptRunResponse.response.result) {
                            const result = appsScriptRunResponse.response.result as AppsScriptFormUpdateResponse;
                            console.log(`[QTI Service] Apps Script function result for form ${createdForm.formId}:`, result);
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
                          return throwError(() => new Error(`Failed to execute Apps Script to add items: ${this.utils.formatHttpError(appsScriptErr)}`));
                        })
                      );
                    }),
                    mergeMap(finalForm => { // Update Drive Properties
                      if (!finalForm?.formId) return of(null);
                      const formId = finalForm.formId;
                      const driveUpdateUrl = `${this.utils.DRIVE_API_FILES_ENDPOINT}/${formId}`;
                      const driveUpdateBody = {appProperties: {[this.APP_PROPERTY_KEY]: hashedItemId}};
                      const driveUpdateParams = new HttpParams().set('addParents', parentFolderId).set('removeParents', 'root');
                      const driveUpdateRequest$ = this.http.patch<DriveFile>(driveUpdateUrl, driveUpdateBody, {headers: driveApiHeaders, params: driveUpdateParams});
                      return this.utils.retryRequest(driveUpdateRequest$, retryConfig, `Update Drive Props for Form ${formId}`).pipe(
                        map(driveFile => ({form: {formUrl: finalForm.responderUri || `https://docs.google.com/forms/d/${formId}/viewform`, title: driveFile.name || finalForm.info?.title || formTitle}} as Material)),
                        catchError(driveErr => {
                          console.error(`[QTI Service] Error updating Drive properties for Form ${formId}: ${this.utils.formatHttpError(driveErr)}`);
                          return of(null);
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

  private _uploadImagesAndBuildFormRequests(
    intermediateItems: IntermediateFormItemDefinition[],
    accessToken: string,
    parentFolderIdForImages: string,
    retryConfig: RetryConfig,
    qtiItemIdForHashing?: string
  ): Observable<FormRequest[]> {
    if (intermediateItems.length === 0) {
      return of([]);
    }

    const itemsThatNeedImageUpload = intermediateItems.filter(item => !!item.imageFileToUpload);
    if (itemsThatNeedImageUpload.length === 0) {
      return of(this._buildFormRequestsFromIntermediate(intermediateItems, new Map()));
    }

    const filesToBatchUpload: Array<{file: ImsccFile; targetFileName: string}> = itemsThatNeedImageUpload.map(itemDef => ({
      file: itemDef.imageFileToUpload!,
      targetFileName: itemDef.imageFileToUpload!.name
    }));

    console.log(`[QTI Service] Attempting to upload ${filesToBatchUpload.length} images to Drive folder ${parentFolderIdForImages}.`);

    return this.fileUploadService.uploadLocalFiles(filesToBatchUpload, accessToken, parentFolderIdForImages).pipe(
      map(uploadedDriveFiles => {
        console.log(`[QTI Service] Successfully processed ${uploadedDriveFiles.length} of ${filesToBatchUpload.length} image uploads (some may have been found existing).`);
        const driveFileMap = new Map<string, DriveFile>();

        uploadedDriveFiles.forEach(driveFile => {
          if (driveFile && driveFile.name) {
            const originalImsccFileName = filesToBatchUpload.find(f => f.targetFileName === driveFile.name)?.file.name;
            if (originalImsccFileName) {
              driveFileMap.set(originalImsccFileName, driveFile);
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
        return of(this._buildFormRequestsFromIntermediate(intermediateItems, new Map()));
      })
    );
  }

  private _buildFormRequestsFromIntermediate(
    intermediateItems: IntermediateFormItemDefinition[],
    driveFileMap: Map<string, DriveFile>
  ): FormRequest[] {
    const formRequests: FormRequest[] = [];
    intermediateItems.forEach(itemDef => {
      let formItem: FormItem | null = null;
      let imageForForm: FormsImage | undefined = undefined;

      if (itemDef.imageFileToUpload) {
        const driveFile = driveFileMap.get(itemDef.imageFileToUpload.name);
        let driveImageUri: string | null = null;

          let altTextForImageObject = "Image";
          if (itemDef.imageAltText && !this.isLikelyFilename(itemDef.imageAltText)) {
            altTextForImageObject = itemDef.imageAltText;
          } else if (itemDef.imageAltText && this.isLikelyFilename(itemDef.imageAltText)) {
            console.log(`[QTI Service] imageAltText "${itemDef.imageAltText}" for image file "${itemDef.imageFileToUpload.name}" appears to be a filename. Using generic alt text for Forms API image object.`);
          }

          if (driveFile) {
            console.log(`[QTI Service] Processing DriveFile for image "${itemDef.imageFileToUpload.name}":`, JSON.stringify(driveFile, null, 2));
            console.log(`   Available links - ID: ${driveFile.id}, Name: ${driveFile.name}, ThumbnailLink: ${driveFile.thumbnailLink}, WebViewLink: ${driveFile.webViewLink}`);

              if (driveFile.thumbnailLink) {
                driveImageUri = driveFile.thumbnailLink.replace(/=s\d+$/, '');
                console.log(`[QTI Service] Using thumbnailLink (cleaned) for ${driveFile.name}: ${driveImageUri}`);
              } else if (driveFile.id) {
                driveImageUri = `https://drive.google.com/uc?id=${driveFile.id}&export=download`;
                console.log(`[QTI Service] ThumbnailLink missing for ${driveFile.name}. Using constructed 'uc?id=...&export=download' link as fallback: ${driveImageUri}.`);
              } else {
                console.warn(`[QTI Service] DriveFile for "${itemDef.imageFileToUpload.name}" is missing both thumbnailLink and ID. Cannot generate image URI.`);
              }
            } else if (typeof itemDef.imageFileToUpload.data === 'string' && itemDef.imageFileToUpload.data.startsWith('data:image')) {
              driveImageUri = itemDef.imageFileToUpload.data;
              console.log(`[QTI Service] Using existing base64 data URI for ${itemDef.imageFileToUpload.name} as no DriveFile mapping found or Drive upload failed.`);
            } else {
              console.warn(`[QTI Service] No DriveFile object found in map for image: "${itemDef.imageFileToUpload.name}", and data is not a base64 URI.`);
            }

          if (driveImageUri) {
            imageForForm = {sourceUri: driveImageUri, altText: altTextForImageObject};
            console.log(`[QTI Service] Prepared FormsImage with sourceUri: ${driveImageUri} and altText: "${altTextForImageObject}" for item titled: "${itemDef.title || 'untitled item'}"`);
          } else {
            console.warn(`[QTI Service] Image "${itemDef.imageFileToUpload.name}" for item "${itemDef.title}" could not be processed for a usable URI. Item will be created without this image.`);
          }
        }

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
              console.warn(`[QTI Service] Standalone image (original src: ${itemDef.originalImgSrc}) could not be processed. Skipping this image item.`);
            }
      } else if (itemDef.videoItem) {
        formItem = {title: itemDef.title, description: itemDef.description, videoItem: itemDef.videoItem};
      } else if (itemDef.pageBreakItem) {
        formItem = {title: itemDef.title, description: itemDef.description, pageBreakItem: itemDef.pageBreakItem};
      } else if (itemDef.sectionHeaderItem) {
        formItem = {title: itemDef.title, description: itemDef.description, sectionHeaderItem: itemDef.sectionHeaderItem};
      }

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
    console.log(`[QTI HTML Parser] Parsing QTI from "${qtiFilePath}"...`);
    const itemElement = qtiDoc.querySelector('item, assessmentItem');
    if (!itemElement) return intermediateItems;

    const mattextElement = itemElement.querySelector('presentation > material > mattext[texttype="text/html"], itemBody > div > mattext[texttype="text/html"], itemBody > mattext[texttype="text/html"]');
    if (mattextElement && mattextElement.textContent) {
      const htmlContent = decode(mattextElement.textContent.trim());
      const parser = new DOMParser();
      const htmlDoc = parser.parseFromString(htmlContent, 'text/html');
      const bodyChildren = Array.from(htmlDoc.body.children);
      let questionCounter = 0;

      bodyChildren.forEach((element: Element) => {
        if (element.tagName.toLowerCase() === 'p' || element.tagName.toLowerCase() === 'span' || element.tagName.toLowerCase() === 'div') {
          const imagesInElement = Array.from(element.querySelectorAll('img'));
          const textContentFromParent = this.getTextContent(element, true).trim();

          if (imagesInElement.length > 0) {
            imagesInElement.forEach(imgElement => {
              questionCounter++;
              const imgSrc = imgElement.getAttribute('src');
              const rawAltText = imgElement.getAttribute('alt');
              let imageFileToUpload: ImsccFile | undefined = undefined;

                      if (imgSrc) {
                        const imagePath = this._resolveImagePath(imgSrc, qtiFilePath);
                        if (imagePath) {
                          imageFileToUpload = allPackageFiles.find(f => f.name.toLowerCase() === imagePath.toLowerCase());
                          if (!imageFileToUpload) {
                            console.warn(`   [QTI HTML] Image file not found in package for src: ${imgSrc} (resolved to: ${imagePath})`);
                          } else if (typeof imageFileToUpload.data !== 'string' || !imageFileToUpload.data.startsWith('data:image')) {
                            console.warn(`   [QTI HTML] Image file ${imageFileToUpload.name} is not a base64 data URI. Ensure it's pre-processed if direct embedding is intended.`);
                          }
                        } else {
                            console.warn(`   [QTI HTML] Could not resolve image path for src: ${imgSrc}`);
                          }
                        }

                      let titleForIntermediateItem = textContentFromParent ||
                        (rawAltText && !this.isLikelyFilename(rawAltText) ? rawAltText : `Image ${questionCounter}`);
                      if (this.isLikelyFilename(titleForIntermediateItem)) {
                        titleForIntermediateItem = `Image related to question ${questionCounter}`;
                      }

                      let descriptionForIntermediateItem = "";
                      if (rawAltText && !this.isLikelyFilename(rawAltText)) {
                        if (rawAltText !== titleForIntermediateItem) {
                          descriptionForIntermediateItem = rawAltText;
                        }
                      } else if (rawAltText && this.isLikelyFilename(rawAltText)) {
                        console.log(`[QTI HTML Parser] Alt text "${rawAltText}" is a filename. Question description (help text) will be empty.`);
                      }

                      intermediateItems.push({
                        type: 'question',
                        title: titleForIntermediateItem,
                        description: descriptionForIntermediateItem,
                        question: {textQuestion: {paragraph: false}},
                        imageFileToUpload: imageFileToUpload,
                        originalImgSrc: imgSrc || undefined,
                        imageAltText: rawAltText || (imageFileToUpload ? imageFileToUpload.name : '')
                      });
                    });
            } else if (textContentFromParent) {
              questionCounter++;
              let qText = textContentFromParent; // Corrected: was textContentFromElement
              if (!textContentFromParent.match(/^\d+\.\s*/)) { // Corrected: was textContentFromElement
                qText = `${questionCounter}. ${textContentFromParent}`; // Corrected: was textContentFromElement
              }
              intermediateItems.push({
                type: 'question',
                title: qText,
                description: undefined,
                question: {textQuestion: {paragraph: false}},
                imageAltText: undefined
              });
            }
          }
        });
      const overallGradingInfo = this.parseResponseProcessing(itemElement, 'TEXT');
      if (overallGradingInfo) {
        console.log(`   [QTI HTML] Overall QTI item has grading info (Points: ${overallGradingInfo.points}), not applied to individual HTML-parsed questions as their structure is inferred.`);
      }
      return intermediateItems;
    } else {
      return this.parseStandardQtiItems(qtiDoc, qtiFilePath);
    }
  }

  private parseStandardQtiItems(qtiDoc: XMLDocument, sourceFileName: string): IntermediateFormItemDefinition[] {
    const intermediateItems: IntermediateFormItemDefinition[] = [];
    console.log(`[Standard QTI Parser] Parsing QTI from "${sourceFileName}"...`);
    const items = Array.from(qtiDoc.querySelectorAll('item, assessmentItem'));
    items.forEach((itemElement, index) => {
      const itemIdentifier = itemElement.getAttribute('ident') || itemElement.getAttribute('identifier') || `qti_item_${index}`;
      let itemTitle = itemElement.getAttribute('title') || '';
      const itemBody = itemElement.querySelector('itemBody');
      const presentation = itemElement.querySelector('presentation');
      const questionTextContainer = presentation?.querySelector('material') || itemBody;
      let questionTextSourceElement = questionTextContainer?.querySelector('prompt') ||
        (presentation && questionTextContainer?.querySelector('mattext')) ||
        questionTextContainer?.querySelector('p, div') ||
        questionTextContainer;
      let questionText = questionTextSourceElement ? this.getTextContent(questionTextSourceElement).trim() : '';
      if (!questionText && itemTitle) questionText = itemTitle;
      if (!questionText && !itemTitle) {
        console.warn(`   Skipping standard QTI item ${itemIdentifier}: No question text or title.`);
        return;
      }
      const cleanQuestionText = questionText.replace(/[\n\r]+/g, ' ').replace(/[ ]+/g, ' ').trim();

      const choiceInteraction = itemElement.querySelector('choiceInteraction');
      const textEntryInteraction = itemElement.querySelector('textEntryInteraction');
      const extendedTextInteraction = itemElement.querySelector('extendedTextInteraction');
      const responseLid = presentation?.querySelector('response_lid');
      const renderChoice = responseLid?.querySelector('render_choice');
      const responseStr = presentation?.querySelector('response_str');
      let questionType: 'CHOICE' | 'TEXT' | 'UNKNOWN' = 'UNKNOWN';
      if (choiceInteraction || renderChoice) questionType = 'CHOICE';
      else if (textEntryInteraction || extendedTextInteraction || responseStr) questionType = 'TEXT';

      let question: Question | undefined;
      const allParsedChoices: ParsedChoice[] = [];
      if (questionType === 'CHOICE') {
        const interactionElement = choiceInteraction || renderChoice!;
        const maxChoices = choiceInteraction?.getAttribute('maxChoices');
        const cardinality = responseLid?.getAttribute('cardinality');
        const isCheckbox = maxChoices !== '1' || cardinality === 'multiple' || renderChoice?.getAttribute('shuffle')?.toLowerCase() === 'yes';
        const choiceTypeValue: 'RADIO' | 'CHECKBOX' = isCheckbox ? 'CHECKBOX' : 'RADIO';
        const options: Option[] = [];
        const choices = Array.from(interactionElement.querySelectorAll('simpleChoice, response_label'));
        choices.forEach(choice => {
          let choiceTextVal = '';
          let choiceId = choice.getAttribute('identifier') || (choice.tagName.toLowerCase() === 'response_label' ? choice.getAttribute('ident') : null);
          let textEl = choice.tagName.toLowerCase() === 'simplechoice' ? choice : (choice.closest('flow_label')?.querySelector('material') || choice.nextElementSibling)?.querySelector('mattext, p, div') || choice;
          if (textEl) choiceTextVal = this.getTextContent(textEl).trim();
          if (choiceTextVal && choiceId) {
            options.push({value: choiceTextVal});
            allParsedChoices.push({identifier: choiceId, value: choiceTextVal});
          }
        });
        if (options.length > 0) question = {choiceQuestion: {type: choiceTypeValue, options: options}};
      } else if (questionType === 'TEXT') {
        const isParagraph = !!extendedTextInteraction || (!!responseStr && responseLid?.querySelector('render_fib') !== null);
        question = {textQuestion: {paragraph: isParagraph}};
      }


      if (question) {
        const gradingInfo = this.parseResponseProcessing(itemElement, questionType, allParsedChoices);
        if (gradingInfo && gradingInfo.correctAnswerValues.length > 0 && !(question.textQuestion?.paragraph)) {
          question.required = true;
          question.grading = {pointValue: gradingInfo.points, correctAnswers: {answers: gradingInfo.correctAnswerValues.map(val => ({value: val}))}};
        }

        const descriptionElement = itemElement.querySelector('itemmetadata > qtimetadata > qti_metadatafield[fieldlabel="qmd_description"] > fieldentry') ||
          (itemBody ? itemBody.querySelector('rubricBlock, *[role="tooltip"], .accessibility_description') : null);
        const itemDescription = this.getTextContent(descriptionElement).trim();


        intermediateItems.push({
          type: 'question',
          title: cleanQuestionText,
          description: itemDescription || undefined,
          question: question
        });
      } else {
        console.warn(`   Skipping standard QTI item ${itemIdentifier}: Could not determine valid question structure.`);
      }
    });
    return intermediateItems;
  }

  private parseResponseProcessing(item: Element, questionType: 'CHOICE' | 'TEXT' | 'UNKNOWN', choicesMap?: ParsedChoice[]): ParsedGradingInfo | null {
    if (questionType === 'UNKNOWN') return null;
    const respProcessing = item.querySelector('resprocessing, responseProcessing');
    if (!respProcessing) return null;
    let points = 1;
    const correctAnswerValues: string[] = [];
    const weightMetaLabel = item.querySelector('itemmetadata > qtimetadata > qti_metadatafield > fieldlabel');
    if (weightMetaLabel && (weightMetaLabel.textContent || '').trim().toLowerCase() === 'qmd_weighting') {
      const weightEntry = weightMetaLabel.nextElementSibling;
      if (weightEntry && weightEntry.textContent) {
        points = Math.round(parseFloat(weightEntry.textContent.trim())) || 1;
      }
    }
    const respconditions = Array.from(respProcessing.querySelectorAll('respcondition'));
    if (respconditions.length > 0) {
      respconditions.forEach(condition => {
        const setvar = condition.querySelector('setvar');
        const score = parseFloat(setvar?.textContent || '0');
        const action = setvar?.getAttribute('action');
        if (action === 'Set' && score > 0) {
          if (points === 1 && score !== 1) points = Math.round(score);
          const conditionVar = condition.querySelector('conditionvar');
          const varequal = conditionVar?.querySelector('varequal');
          if (varequal) {
            const correctIdentifierOrValue = varequal.textContent?.trim();
            if (correctIdentifierOrValue) {
              if (questionType === 'CHOICE' && choicesMap) {
                const matchingChoice = choicesMap.find(c => c.identifier === correctIdentifierOrValue);
                if (matchingChoice && !correctAnswerValues.includes(matchingChoice.value)) {
                  correctAnswerValues.push(matchingChoice.value);
                }
              } else if (questionType === 'TEXT') {
                const cleanCorrectValue = this.getTextContent(varequal)?.trim();
                if (cleanCorrectValue && !correctAnswerValues.includes(cleanCorrectValue)) {
                  correctAnswerValues.push(cleanCorrectValue);
                }
              }
            }
          }
        }
      });
    } else {
      const responseDeclaration = item.querySelector(`responseDeclaration[identifier="RESPONSE"]`);
      const mapping = responseDeclaration?.querySelector('mapping');
      if (mapping?.hasAttribute('defaultValue') && parseFloat(mapping.getAttribute('defaultValue') || '0') > 0) {
        points = Math.round(parseFloat(mapping.getAttribute('defaultValue')!)) || points;
      }
      const correctResponse = responseDeclaration?.querySelector('correctResponse');
      if (correctResponse) {
        correctResponse.querySelectorAll('value').forEach(valueEl => {
          const correctIdentifierOrValue = valueEl.textContent?.trim();
          if (correctIdentifierOrValue) {
            if (questionType === 'CHOICE' && choicesMap) {
              const matchingChoice = choicesMap.find(c => c.identifier === correctIdentifierOrValue);
              if (matchingChoice && !correctAnswerValues.includes(matchingChoice.value)) {
                correctAnswerValues.push(matchingChoice.value);
              }
            } else if (questionType === 'TEXT') {
              const cleanCorrectValue = this.getTextContent(valueEl)?.trim();
              if (cleanCorrectValue && !correctAnswerValues.includes(cleanCorrectValue)) {
                correctAnswerValues.push(cleanCorrectValue);
              }
            }
          }
        });
      }
    }
    if (correctAnswerValues.length > 0) {
      return {points: points, correctAnswerValues: correctAnswerValues};
    }
    return null;
  }

  private getTextContent(element: Element | null, excludeImageAltText = false): string {
    if (!element) return '';
    let clonedElement = element.cloneNode(true) as Element;
    if (excludeImageAltText) {
      clonedElement.querySelectorAll('img').forEach(img => img.remove());
    }
    let rawHtml = clonedElement.innerHTML;
    let decodedText = rawHtml;
    if (typeof document !== 'undefined' && document.createElement) {
      try {
        const tempDecoder = document.createElement('div');
        tempDecoder.innerHTML = rawHtml;
        decodedText = tempDecoder.textContent || tempDecoder.innerText || '';
      } catch (e) {console.warn("[getTextContent] DOM parsing failed.", e);}
    }
    try {decodedText = decode(decodedText);}
    catch (libError) {console.warn("[getTextContent] html-entities decoding failed.", libError);}
    let plainText = decodedText.replace(/<annotation[^>]*>.*?<\/annotation>/gis, ' ');
    plainText = plainText.replace(/<[^>]*>/g, ' ');
    plainText = plainText.replace(/\s+/g, ' ').trim();
    plainText = plainText.replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&').replace(/\s+/g, ' ').trim();
    return plainText;
  }

  private _resolveImagePath(imgSrc: string, qtiFilePath: string): string | null {
    if (!imgSrc) return null;
    let relativePath = imgSrc;
    let baseForResolution = this.utils.getDirectory(qtiFilePath);
    if (imgSrc.startsWith('$IMS-CC-FILEBASE$')) {
      relativePath = imgSrc.substring('$IMS-CC-FILEBASE$'.length);
      baseForResolution = "";
    }
    if (relativePath.startsWith('./')) {
      relativePath = relativePath.substring(2);
    }
    const decodedRelativePath = this.utils.tryDecodeURIComponent(relativePath);
    return this.utils.resolveRelativePath(baseForResolution, decodedRelativePath);
  }
}
