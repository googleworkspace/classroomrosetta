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

import { inject, Injectable } from '@angular/core';
import {HttpClient, HttpHeaders, HttpParams, HttpErrorResponse} from '@angular/common/http';
import { Observable, throwError, of, EMPTY, forkJoin, timer, mergeMap } from 'rxjs';
import {catchError, map, expand, reduce, switchMap, shareReplay, tap} from 'rxjs/operators';
import {
    Classroom,
    ClassroomListResponse,
    Topic,
    ListTopicsResponse,
    CourseWork,
    ProcessedCourseWork, // Interface from classroom-interface.ts
    Material,
  ImsccFile,
  CourseWorkMaterial // Represents the API resource for CourseWorkMaterial
} from '../../interfaces/classroom-interface'; // Adjust path as needed
import {UtilitiesService, RetryConfig} from '../utilities/utilities.service'; // Adjust path as needed

@Injectable({
  providedIn: 'root'
})
export class ClassroomService {

  // Base URLs for the Google Classroom API v1
  private coursesApiUrl = 'https://classroom.googleapis.com/v1/courses';
  private topicsApiUrl = (courseId: string) => `${this.coursesApiUrl}/${courseId}/topics`;
  private courseWorkApiUrl = (courseId: string) => `${this.coursesApiUrl}/${courseId}/courseWork`;
  private courseWorkMaterialsApiUrl = (courseId: string) => `${this.coursesApiUrl}/${courseId}/courseWorkMaterials`;


  private http = inject(HttpClient);
  private utils = inject(UtilitiesService);
  private pageSize = '50'; // Consider making this configurable or smaller if not all classrooms are needed at once
  private materialLimit = 20; // Classroom API limit for materials per item

  private defaultRetryConfig: RetryConfig = {maxRetries: 3, initialDelayMs: 2000};

  constructor() { }

  // --- Get Classrooms ---
  getActiveClassrooms(authToken: string): Observable<Classroom[]> {
    if (!authToken) {
      console.error('ClassroomService: Auth token is missing for getActiveClassrooms.');
      return of([]); // Return empty array or throw error based on desired handling
    }
    const context = 'getActiveClassrooms';
    console.log(`[ClassroomService] ${context}: Fetching active classrooms.`);
    return this.fetchClassroomPage(authToken, undefined, context).pipe(
      expand(response => {
        if (response.nextPageToken) {
          console.log(`[ClassroomService] ${context}: Fetching next page of classrooms (Token: ${response.nextPageToken}).`);
          return this.fetchClassroomPage(authToken, response.nextPageToken, `${context} (paginated)`);
        }
        return EMPTY;
      }),
      map(response => response.courses || []), // Ensure courses is not undefined
      reduce((acc, courses) => acc.concat(courses), [] as Classroom[]),
      tap(allCourses => console.log(`[ClassroomService] ${context}: Successfully fetched ${allCourses.length} active classrooms.`)),
      catchError(err => this.handleError(err, `${context} (Accumulation)`))
    );
  }

  private fetchClassroomPage(authToken: string, pageToken?: string, context: string = 'fetchClassroomPage'): Observable<ClassroomListResponse> {
    const headers = this.createAuthHeaders(authToken);
    let params = new HttpParams()
      .set('courseStates', 'ACTIVE')
      .set('pageSize', this.pageSize);
    if (pageToken) {
      params = params.set('pageToken', pageToken);
    }
    const operationDescription = `${context} (Page Token: ${pageToken ?? 'initial'})`;
    console.log(`[ClassroomService] ${operationDescription}: Making API call.`);
    const request$ = this.http.get<ClassroomListResponse>(this.coursesApiUrl, {headers, params});
    return this.utils.retryRequest(request$, this.defaultRetryConfig, operationDescription).pipe(
      catchError(err => this.handleError(err, operationDescription))
    );
  }

  /**
   * Assigns multiple pieces of content (ProcessedCourseWork) to multiple classrooms.
   * Returns an array of ProcessedCourseWork, each potentially updated with an error or API response data.
   */
  assignContentToClassrooms(
    authToken: string,
    classroomIds: string[],
    assignments: ProcessedCourseWork[]
  ): Observable<ProcessedCourseWork[]> {
    console.log(`[ClassroomService] assignContentToClassrooms: Starting assignment process for ${assignments.length} items to ${classroomIds.length} classrooms.`);
    if (!authToken) {
      console.error('[ClassroomService] assignContentToClassrooms: Auth token is required.');
      return throwError(() => new Error('Authentication token is required.'));
    }
    if (!classroomIds?.length) {
      console.warn('[ClassroomService] assignContentToClassrooms: No classroom IDs provided.');
      return of(assignments.map(a => ({...a, processingError: {message: "No classroom selected.", stage: "Pre-flight Check"}})));
    }
    if (!assignments?.length) {
      console.warn('[ClassroomService] assignContentToClassrooms: No assignments provided.');
      return of([]);
    }

    const topicRequestCache = new Map<string, Observable<string | undefined>>();
    const allOperationsObservables: Observable<ProcessedCourseWork>[] = [];

    for (const courseId of classroomIds) {
      console.log(`[ClassroomService] Processing assignments for Classroom ID: ${courseId}`);
      for (const originalAssignment of assignments) {
        const assignmentForProcessing: ProcessedCourseWork = {...originalAssignment}; // Clone to avoid mutation issues if retrying
        const itemLogPrefix = `[ClassroomService] Item "${assignmentForProcessing.title}" (ID: ${assignmentForProcessing.associatedWithDeveloper?.id || 'N/A'}) for Course ID ${courseId}:`;

        if (!assignmentForProcessing.title || !assignmentForProcessing.workType) {
          console.warn(`${itemLogPrefix} Skipping assignment due to missing title or workType.`);
             assignmentForProcessing.processingError = {
                 message: 'Skipped: Missing title or workType.',
               stage: 'Pre-flight Check (ClassroomService)'
             };
             allOperationsObservables.push(of(assignmentForProcessing));
             continue;
        }
        console.log(`${itemLogPrefix} Preparing for Classroom API. Topic: "${assignmentForProcessing.associatedWithDeveloper?.topic || 'None'}"`);

        const topicName = assignmentForProcessing.associatedWithDeveloper?.topic;
        const uniqueMaterials = this.deduplicateMaterials(assignmentForProcessing.materials || []);
        console.log(`${itemLogPrefix} Deduplicated materials count: ${uniqueMaterials.length}`);


        const cacheKey = `${courseId}:${topicName?.trim().toLowerCase() || 'undefined_topic_key'}`; // Ensure a valid key even if topicName is undefined
        let topicId$: Observable<string | undefined> | undefined = topicRequestCache.get(cacheKey);

        if (!topicId$) {
          console.log(`${itemLogPrefix} Topic ID for "${topicName || 'None'}" not in cache. Fetching or creating.`);
            topicId$ = this.getOrCreateTopicId(authToken, courseId, topicName).pipe(
              tap(resolvedTopicId => console.log(`${itemLogPrefix} Resolved Topic ID for "${topicName || 'None'}": ${resolvedTopicId || 'None'}`)),
              shareReplay(1), // Cache the result of this observable
                catchError(topicError => {
                  console.error(`${itemLogPrefix} CRITICAL error resolving topic "${topicName}". Error: ${topicError.message}`);
                  // Propagate a specific error structure
                  return throwError(() => ({
                    message: `Failed to resolve/create topic "${topicName || 'None'}": ${topicError.message || topicError}`,
                    stage: 'Topic Management',
                    details: topicError.details || topicError.toString()
                  }));
                })
            );
            topicRequestCache.set(cacheKey, topicId$);
        } else {
          console.log(`${itemLogPrefix} Using cached Topic ID for "${topicName || 'None'}".`);
        }

        const createClassroomItemObservables = (
            currentAssignmentData: ProcessedCourseWork,
            materialsForThisPart: Material[]
        ): Observable<ProcessedCourseWork> => {
          console.log(`${itemLogPrefix} (Part: "${currentAssignmentData.title}"): Attempting to get Topic ID.`);
            return topicId$.pipe(
              switchMap(topicIdValue => {
                console.log(`${itemLogPrefix} (Part: "${currentAssignmentData.title}"): Topic ID resolved to "${topicIdValue}". Proceeding to create Classroom item.`);
                    const assignmentDataForApiCall: ProcessedCourseWork = {
                      ...currentAssignmentData, // This includes the (potentially part-specific) title
                      materials: materialsForThisPart,
                      topicId: topicIdValue,
                      // Ensure descriptionForClassroom is used for the API description
                      description: currentAssignmentData.descriptionForClassroom
                    };
                // Remove redundant/internal fields not for API
                const {descriptionForDisplay, localFilesToUpload, qtiFile, htmlContent, webLinkUrl, richtext, associatedWithDeveloper, processingError, classroomCourseWorkId, classroomLink, ...apiPayload} = assignmentDataForApiCall;


                return this.createClassroomItem(authToken, courseId, apiPayload as ProcessedCourseWork).pipe(
                    map(apiResponse => {
                      const resultItem: ProcessedCourseWork = {
                        ...currentAssignmentData, // Use currentAssignmentData which has the correct title for this part
                        materials: materialsForThisPart, // Reflect materials used for THIS part
                        topicId: topicIdValue,
                        classroomCourseWorkId: apiResponse.id,
                        classroomLink: apiResponse.alternateLink,
                        state: apiResponse.state || 'DRAFT', // Default to DRAFT if not present
                        processingError: undefined // Clear any previous error
                            };
                      console.log(`${itemLogPrefix} (Part: "${resultItem.title}"): Successfully created in Classroom. API Response ID: ${apiResponse.id}, Link: ${apiResponse.alternateLink}`);
                            return resultItem;
                        }),
                        catchError(classroomCreationError => {
                          console.error(`${itemLogPrefix} (Part: "${currentAssignmentData.title}"): Failed to create in Classroom. Error: ${classroomCreationError.message || classroomCreationError}`);
                          const errorItem: ProcessedCourseWork = {...currentAssignmentData, materials: materialsForThisPart, topicId: topicIdValue};
                            errorItem.processingError = {
                              message: `Failed to create item "${currentAssignmentData.title}" in Classroom: ${classroomCreationError.message || classroomCreationError}`,
                              stage: currentAssignmentData.workType === 'MATERIAL' ? 'Classroom Material Creation' : 'Classroom CourseWork Creation',
                              details: classroomCreationError.details || classroomCreationError.toString()
                            };
                          return of(errorItem); // Return the item with error info
                        })
                    );
                }),
                catchError(topicResolutionError => {
                  // This catch is for errors from the topicId$ observable itself
                  console.error(`${itemLogPrefix} (Part: "${currentAssignmentData.title}"): Failed to resolve topic. Error: ${topicResolutionError.message || topicResolutionError}`);
                  const errorItem: ProcessedCourseWork = {...currentAssignmentData, materials: materialsForThisPart}; // currentAssignmentData for the part
                    errorItem.processingError = {
                        message: `Failed to resolve/create topic "${topicName || 'None'}": ${topicResolutionError.message || topicResolutionError}`,
                      stage: 'Topic Management',
                      details: topicResolutionError.details || topicResolutionError.toString()
                    };
                  return of(errorItem); // Return the item with error info
                })
            );
        };

        if (uniqueMaterials.length <= this.materialLimit) {
          console.log(`${itemLogPrefix} Material count (${uniqueMaterials.length}) is within limit. Creating single Classroom item.`);
          allOperationsObservables.push(createClassroomItemObservables(assignmentForProcessing, uniqueMaterials));
        } else {
          console.log(`${itemLogPrefix} Material count (${uniqueMaterials.length}) exceeds limit of ${this.materialLimit}. Splitting into parts.`);
          const numParts = Math.ceil(uniqueMaterials.length / this.materialLimit);
          const materialChunks = this.chunkArray(uniqueMaterials, this.materialLimit);

          for (let i = 0; i < numParts; i++) {
            const partIndex = i + 1;
            const partTitle = `${assignmentForProcessing.title} (Part ${partIndex} of ${numParts})`;
            console.log(`${itemLogPrefix} Preparing Part ${partIndex}/${numParts}: "${partTitle}"`);
            const partAssignmentData: ProcessedCourseWork = {
              ...assignmentForProcessing, // Base data
              title: partTitle, // Part-specific title
              descriptionForClassroom: `Part ${partIndex} of ${numParts}:\n\n${assignmentForProcessing.descriptionForClassroom || ''}`, // Part-specific description
              // materials will be set by createClassroomItemObservables
            };
            allOperationsObservables.push(createClassroomItemObservables(partAssignmentData, materialChunks[i]));
          }
        }
      }
    }

    if (allOperationsObservables.length === 0) {
      console.warn('[ClassroomService] assignContentToClassrooms: No valid operations to perform after pre-flight checks.');
      // Ensure original assignments are returned with errors if they existed
      return of(assignments.map(a => ({...a, processingError: a.processingError || {message: "No valid operation to perform.", stage: "Pre-flight Check (ClassroomService)"}})));
    }

    console.log(`[ClassroomService] assignContentToClassrooms: Starting forkJoin for ${allOperationsObservables.length} Classroom item creation operations.`);
    return forkJoin(allOperationsObservables).pipe(
      tap(results => {
        const successes = results.filter(r => !r.processingError).length;
        const failures = results.length - successes;
        console.log(`[ClassroomService] assignContentToClassrooms: Batch assignment process finished. Total operations: ${results.length}, Successes: ${successes}, Failures: ${failures}`);
        results.forEach((result, idx) => {
            if(result.processingError) {
              console.warn(`[ClassroomService] Failed item [${idx}]: "${result.title}", Error: ${result.processingError.message}, Stage: ${result.processingError.stage}, Details: ${JSON.stringify(result.processingError.details)}`);
            } else {
              console.log(`[ClassroomService] Successful item [${idx}]: "${result.title}", Classroom ID: ${result.classroomCourseWorkId}`);
            }
        });
      }),
      catchError(err => { // This catchError is for forkJoin itself, though individual errors are handled above
        console.error('[ClassroomService] assignContentToClassrooms: Unexpected error in forkJoin. This should ideally be caught by individual operations.', err);
        return throwError(() => new Error(`Critical error during batch assignment processing: ${err.message || err}`));
      })
    );
  }


  // --- Topic Management Methods ---
  private getOrCreateTopicId(authToken: string, courseId: string, topicName?: string): Observable<string | undefined> {
    const normalizedTopicName = topicName?.trim(); // Handle potentially undefined topicName
    if (!normalizedTopicName) { // Check if it's empty or undefined after trimming
      console.log(`[ClassroomService] getOrCreateTopicId: No topic name provided for course ${courseId}. Item will not have a topic.`);
      return of(undefined);
    }

    const lowerCaseTopicName = normalizedTopicName.toLowerCase();
    const context = `getOrCreateTopicId (Course: ${courseId}, Topic: "${normalizedTopicName}")`;
    console.log(`[ClassroomService] ${context}: Starting.`);

    return this.listAllTopics(authToken, courseId).pipe(
      map(allTopics => {
        const found = allTopics.find(topic => topic.name?.toLowerCase() === lowerCaseTopicName);
        console.log(`[ClassroomService] ${context}: Searched ${allTopics.length} topics. Found existing: ${found ? found.topicId : 'No'}`);
        return found;
      }),
      switchMap(existingTopic => {
        if (existingTopic?.topicId) {
          console.log(`[ClassroomService] ${context}: Using existing topic ID: ${existingTopic.topicId}`);
          return of(existingTopic.topicId);
        } else {
          console.log(`[ClassroomService] ${context}: Topic not found. Creating new topic.`);
          return this.createTopic(authToken, courseId, normalizedTopicName).pipe(
            map(newTopic => {
              console.log(`[ClassroomService] ${context}: New topic created with ID: ${newTopic.topicId}`);
              return newTopic.topicId;
            })
          );
        }
      }),
      catchError(err => {
        console.error(`[ClassroomService] ${context}: Error during operation.`, err);
        return this.handleError(err, context); // Let handleError format and rethrow
      })
    );
  }

  private listAllTopics(authToken: string, courseId: string): Observable<Topic[]> {
     const context = `listAllTopics (Course: ${courseId})`;
    console.log(`[ClassroomService] ${context}: Fetching all topics.`);
     return this.fetchTopicPage(authToken, courseId, undefined, context).pipe(
       expand(response => {
         if (response.nextPageToken) {
           console.log(`[ClassroomService] ${context}: Fetching next page of topics (Token: ${response.nextPageToken}).`);
           return this.fetchTopicPage(authToken, courseId, response.nextPageToken, `${context} (paginated)`);
         }
         return EMPTY;
       }),
        map(response => response.topic || []),
        reduce((acc, topics) => acc.concat(topics), [] as Topic[]),
       tap(allTopics => console.log(`[ClassroomService] ${context}: Successfully fetched ${allTopics.length} topics.`)),
        catchError(err => this.handleError(err, `${context} (Accumulation)`))
     );
  }

  private fetchTopicPage(authToken: string, courseId: string, pageToken?: string, context: string = 'fetchTopicPage'): Observable<ListTopicsResponse> {
    const headers = this.createAuthHeaders(authToken);
    let params = new HttpParams().set('pageSize', this.pageSize);
    if (pageToken) {
      params = params.set('pageToken', pageToken);
    }
    const url = this.topicsApiUrl(courseId);
    const operationDescription = `${context} (Course: ${courseId}, Page Token: ${pageToken ?? 'initial'})`;
    console.log(`[ClassroomService] ${operationDescription}: Making API call to ${url}.`);
    const request$ = this.http.get<ListTopicsResponse>(url, {headers, params});
    return this.utils.retryRequest(request$, this.defaultRetryConfig, operationDescription).pipe(
        catchError(err => this.handleError(err, operationDescription))
    );
  }

  private createTopic(authToken: string, courseId: string, topicName: string): Observable<Topic> {
    const headers = this.createAuthHeaders(authToken);
    const url = this.topicsApiUrl(courseId);
    const body = { name: topicName };
    const context = `createTopic (Course: ${courseId}, Topic: "${topicName}")`;
    console.log(`[ClassroomService] ${context}: Making API call to ${url} with body:`, body);
    const request$ = this.http.post<Topic>(url, body, {headers});
    return this.utils.retryRequest(request$, this.defaultRetryConfig, context).pipe(
       map(topic => {
         console.log(`[ClassroomService] ${context}: Topic created successfully. ID: ${topic.topicId}, Name: "${topic.name}"`);
            return topic;
       }),
       catchError(err => this.handleError(err, context))
    );
  }

  // --- Unified CourseWork / CourseWorkMaterial Creation Method ---
  private createClassroomItem(
    authToken: string,
    courseId: string,
    itemData: ProcessedCourseWork // Expects ProcessedCourseWork with only API relevant fields
  ): Observable<CourseWork | CourseWorkMaterial> {
    const headers = this.createAuthHeaders(authToken);
    const itemLogPrefix = `[ClassroomService] createClassroomItem (Course: ${courseId}, Title: "${itemData.title}", Type: ${itemData.workType})`;

    console.log(`${itemLogPrefix}: Preparing to create item.`);

    if (!itemData.title || !itemData.workType) {
      const errorMsg = `${itemLogPrefix}: Missing required field(s) (title, workType). Cannot create item.`;
      console.error(errorMsg, 'Item Data:', itemData);
      return throwError(() => new Error(errorMsg));
    }

    if (itemData.workType === 'MATERIAL') {
      const materialBody: CourseWorkMaterial = {
        title: itemData.title,
        description: itemData.description, // Ensure this is descriptionForClassroom
        materials: itemData.materials,
        state: itemData.state || 'PUBLISHED', // Default to PUBLISHED if not specified
        topicId: itemData.topicId,
        scheduledTime: itemData.scheduledTime
        // assigneeMode is not set for materials as it's always ALL_STUDENTS
      };
      const url = this.courseWorkMaterialsApiUrl(courseId);
      console.log(`${itemLogPrefix}: Submitting as CourseWorkMaterial to ${url}. Body:`, JSON.stringify(materialBody));
      const request$ = this.http.post<CourseWorkMaterial>(url, materialBody, {headers});
      return this.utils.retryRequest(request$, this.defaultRetryConfig, `${itemLogPrefix} as MATERIAL`).pipe(
        map(createdMaterial => {
          console.log(`${itemLogPrefix}: CourseWorkMaterial created successfully. ID: ${createdMaterial.id}`);
          return createdMaterial;
        }),
        catchError(err => {
          console.error(`${itemLogPrefix}: Error creating CourseWorkMaterial.`, err);
          return this.handleError(err, `${itemLogPrefix} as MATERIAL`);
        })
      );
    } else {
      // Ensure workType is one of the valid CourseWork types
      const validWorkTypes: Array<CourseWork['workType']> = ['ASSIGNMENT', 'SHORT_ANSWER_QUESTION', 'MULTIPLE_CHOICE_QUESTION'];
      if (!validWorkTypes.includes(itemData.workType as any)) {
        const errorMsg = `${itemLogPrefix}: Invalid workType "${itemData.workType}" for CourseWork.`;
        console.error(errorMsg, 'Item Data:', itemData);
        return throwError(() => new Error(errorMsg));
      }

      const courseWorkBody: CourseWork = {
        title: itemData.title,
        description: itemData.description, // Ensure this is descriptionForClassroom
        materials: itemData.materials,
        workType: itemData.workType as 'ASSIGNMENT' | 'SHORT_ANSWER_QUESTION' | 'MULTIPLE_CHOICE_QUESTION',
        state: itemData.state || 'PUBLISHED', // Default to PUBLISHED
        topicId: itemData.topicId,
        maxPoints: itemData.maxPoints,
        assignment: itemData.workType === 'ASSIGNMENT' ? itemData.assignment : undefined,
        multipleChoiceQuestion: itemData.workType === 'MULTIPLE_CHOICE_QUESTION' ? itemData.multipleChoiceQuestion : undefined,
        dueDate: itemData.dueDate,
        dueTime: itemData.dueTime,
        scheduledTime: itemData.scheduledTime,
        submissionModificationMode: itemData.submissionModificationMode
        // assigneeMode defaults to ALL_STUDENTS if not specified by API
      };
      const url = this.courseWorkApiUrl(courseId);
      console.log(`${itemLogPrefix}: Submitting as CourseWork to ${url}. Body:`, JSON.stringify(courseWorkBody));
      const request$ = this.http.post<CourseWork>(url, courseWorkBody, {headers});
      return this.utils.retryRequest(request$, this.defaultRetryConfig, itemLogPrefix).pipe(
        map(createdWork => {
          console.log(`${itemLogPrefix}: CourseWork created successfully. ID: ${createdWork.id}`);
          return createdWork;
        }),
        catchError(err => {
          console.error(`${itemLogPrefix}: Error creating CourseWork.`, err);
          return this.handleError(err, itemLogPrefix);
        })
      );
    }
  }

  // --- Utility and Error Handling ---
  private createAuthHeaders(authToken: string): HttpHeaders {
     return new HttpHeaders({
        'Authorization': `Bearer ${authToken}`,
        'Accept': 'application/json',
       'Content-Type': 'application/json' // Ensure Content-Type for POST/PATCH
     });
  }

  private chunkArray<T>(array: T[], size: number): T[][] {
    if (!array) return [];
    const result: T[][] = [];
    for (let i = 0; i < array.length; i += size) {
        result.push(array.slice(i, i + size));
    }
    return result;
  }

  private getMaterialKey(material: Material): string | null {
      if (material.driveFile?.driveFile?.id) return `drive-${material.driveFile.driveFile.id}`;
      if (material.youtubeVideo?.id) return `youtube-${material.youtubeVideo.id}`;
      if (material.link?.url) return `link-${material.link.url}`;
      if (material.form?.formUrl) return `form-${material.form.formUrl}`;
    return null; // Should not happen if material is valid
  }

  private deduplicateMaterials(materials: Material[]): Material[] {
      if (!materials || materials.length === 0) return [];
      const seenKeys = new Set<string>();
      const uniqueMaterials: Material[] = [];
      for (const material of materials) {
          const key = this.getMaterialKey(material);
          if (key !== null && !seenKeys.has(key)) {
              seenKeys.add(key);
              uniqueMaterials.push(material);
          } else if (key === null) { // If key is null (e.g. unexpected material structure), include it but log
            console.warn('[ClassroomService] deduplicateMaterials: Encountered material with no identifiable key, including as is.', material);
              uniqueMaterials.push(material);
          } else {
            console.log(`[ClassroomService] deduplicateMaterials: Duplicate material skipped (Key: ${key})`);
          }
      }
      return uniqueMaterials;
  }

  private handleError(error: HttpErrorResponse | Error, context: string = 'Unknown Operation'): Observable<never> {
    let userMessage = `Failed during ${context}; please try again later or check console for details.`;
    let detailedMessage = `Context: ${context} - An unknown error occurred!`;
    let statusCode: number | undefined = undefined;
    let errorDetailsForPropagation: any = error; // Default to the original error object

    if (error instanceof HttpErrorResponse) {
        statusCode = error.status;
        detailedMessage = `Context: ${context} - Server error: Code ${error.status}, Message: ${error.message || 'No message body'}`;
      userMessage = `The server returned an error (Code: ${error.status}) while processing ${context}. Please check details or try again.`;
      errorDetailsForPropagation = error.error || error.message; // Prefer error.error if available
        try {
          const errorBody = JSON.stringify(error.error); // Attempt to stringify for logging
            detailedMessage += `, Body: ${errorBody}`;
          const googleApiError = error.error?.error?.message; // Standard Google API error structure
            if (googleApiError) {
               userMessage = `Google API Error in ${context}: ${googleApiError} (Code: ${error.status})`;
               detailedMessage += ` | Google API Specific: ${googleApiError}`;
            }
        } catch (e) { /* Ignore JSON stringify errors if error.error is not an object */}
    } else if (error instanceof Error) {
       detailedMessage = `Context: ${context} - Client/Network error: ${error.message}`;
      userMessage = `A network or client-side error occurred in ${context}. Please check your connection or the console. Details: ${error.message}`;
      errorDetailsForPropagation = error.message;
    }

    console.error(`[ClassroomService] handleError: ${detailedMessage}`, 'Full Error Object:', error);

    // Create a new error object that includes the user-friendly message and structured details
    const finalError = new Error(userMessage);
    (finalError as any).status = statusCode; // Attach status if available
    (finalError as any).details = errorDetailsForPropagation; // Attach more detailed error info
    (finalError as any).stage = context; // Add context as stage

    return throwError(() => finalError);
  }
}
