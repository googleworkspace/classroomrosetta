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

import {inject, Injectable} from '@angular/core';
import {HttpClient, HttpHeaders, HttpParams, HttpErrorResponse} from '@angular/common/http';
import {Observable, throwError, of, EMPTY, forkJoin} from 'rxjs';
import {catchError, map, expand, reduce, switchMap, shareReplay, tap} from 'rxjs/operators';
import {
  Classroom,
  ClassroomListResponse,
  Topic,
  ListTopicsResponse,
  CourseWork,
  ProcessedCourseWork,
  Material,
  CourseWorkMaterial
} from '../../interfaces/classroom-interface';
import {UtilitiesService, RetryConfig} from '../utilities/utilities.service';

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
  /**
   * Fetches all active classrooms for the authenticated user.
   * Handles pagination automatically.
   * @param authToken The OAuth 2.0 token for authentication.
   * @returns An Observable array of active Classroom objects.
   */
  getActiveClassrooms(authToken: string): Observable<Classroom[]> {
    if (!authToken) {
      console.error('ClassroomService: Auth token is missing for getActiveClassrooms.');
      return of([]); // Return empty array or throw error based on desired handling
    }
    const context = 'getActiveClassrooms';
    console.log(`[ClassroomService] ${context}: Fetching active classrooms.`);
    // Fetch the first page and then expand to get subsequent pages if nextPageToken exists
    return this.fetchClassroomPage(authToken, undefined, context).pipe(
      expand(response => {
        // If there's a nextPageToken, fetch the next page
        if (response.nextPageToken) {
          console.log(`[ClassroomService] ${context}: Fetching next page of classrooms (Token: ${response.nextPageToken}).`);
          return this.fetchClassroomPage(authToken, response.nextPageToken, `${context} (paginated)`);
        }
        // If no nextPageToken, complete the expand operator
        return EMPTY;
      }),
      // Map each response to its 'courses' array (or an empty array if undefined)
      map(response => response.courses || []),
      // Reduce the stream of course arrays into a single array of all courses
      reduce((acc, courses) => acc.concat(courses), [] as Classroom[]),
      tap(allCourses => console.log(`[ClassroomService] ${context}: Successfully fetched ${allCourses.length} active classrooms.`)),
      catchError(err => this.handleError(err, `${context} (Accumulation)`))
    );
  }

  /**
   * Fetches a single page of classrooms.
   * @param authToken The OAuth 2.0 token.
   * @param pageToken Optional token for fetching a specific page.
   * @param context A string describing the context of the call for logging.
   * @returns An Observable of ClassroomListResponse.
   */
  private fetchClassroomPage(authToken: string, pageToken?: string, context: string = 'fetchClassroomPage'): Observable<ClassroomListResponse> {
    const headers = this.createAuthHeaders(authToken);
    let params = new HttpParams()
      .set('courseStates', 'ACTIVE') // Filter for active courses
      .set('pageSize', this.pageSize);
    if (pageToken) {
      params = params.set('pageToken', pageToken);
    }
    const operationDescription = `${context} (Page Token: ${pageToken ?? 'initial'})`;
    console.log(`[ClassroomService] ${operationDescription}: Making API call.`);
    const request$ = this.http.get<ClassroomListResponse>(this.coursesApiUrl, {headers, params});
    // Retry the request using the utility service
    return this.utils.retryRequest(request$, this.defaultRetryConfig, operationDescription).pipe(
      catchError(err => this.handleError(err, operationDescription))
    );
  }

  /**
   * Assigns multiple pieces of content (ProcessedCourseWork) to multiple classrooms.
   * Handles topic creation/retrieval and material splitting if necessary.
   * @param authToken The OAuth 2.0 token.
   * @param classroomIds Array of classroom IDs to assign content to.
   * @param assignments Array of ProcessedCourseWork objects to assign.
   * @returns An Observable array of ProcessedCourseWork, each potentially updated with an error or API response data.
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
      // Return assignments with an error message if no classrooms are selected
      return of(assignments.map(a => ({...a, processingError: {message: "No classroom selected.", stage: "Pre-flight Check"}})));
    }
    if (!assignments?.length) {
      console.warn('[ClassroomService] assignContentToClassrooms: No assignments provided.');
      return of([]);
    }

    // Cache for topic ID requests to avoid redundant API calls for the same topic in the same course
    const topicRequestCache = new Map<string, Observable<string | undefined>>();
    const allOperationsObservables: Observable<ProcessedCourseWork>[] = [];

    // Iterate over each classroom ID
    for (const courseId of classroomIds) {
      console.log(`[ClassroomService] Processing assignments for Classroom ID: ${courseId}`);
      // Iterate over each assignment to be processed
      for (const originalAssignment of assignments) {
        // Clone the assignment to avoid mutating the original object, especially important if retries occur
        const assignmentForProcessing: ProcessedCourseWork = {...originalAssignment};
        const itemLogPrefix = `[ClassroomService] Item "${assignmentForProcessing.title}" (ID: ${assignmentForProcessing.associatedWithDeveloper?.id || 'N/A'}) for Course ID ${courseId}:`;

        // Basic validation for the assignment
        if (!assignmentForProcessing.title || !assignmentForProcessing.workType) {
          console.warn(`${itemLogPrefix} Skipping assignment due to missing title or workType.`);
          assignmentForProcessing.processingError = {
            message: 'Skipped: Missing title or workType.',
            stage: 'Pre-flight Check (ClassroomService)'
          };
          allOperationsObservables.push(of(assignmentForProcessing)); // Add to operations to maintain result structure
          continue; // Skip to the next assignment
        }
        console.log(`${itemLogPrefix} Preparing for Classroom API. Topic: "${assignmentForProcessing.associatedWithDeveloper?.topic || 'None'}"`);

        const topicName = assignmentForProcessing.associatedWithDeveloper?.topic;
        // Deduplicate materials to prevent API errors or redundant attachments
        const uniqueMaterials = this.deduplicateMaterials(assignmentForProcessing.materials || []);
        console.log(`${itemLogPrefix} Deduplicated materials count: ${uniqueMaterials.length}`);


        // Create a unique cache key for topic requests per course and topic name
        const cacheKey = `${courseId}:${topicName?.trim().toLowerCase() || 'undefined_topic_key'}`;
        let topicId$: Observable<string | undefined> | undefined = topicRequestCache.get(cacheKey);

        // If topic ID observable is not in cache, create it
        if (!topicId$) {
          console.log(`${itemLogPrefix} Topic ID for "${topicName || 'None'}" not in cache. Fetching or creating.`);
          topicId$ = this.getOrCreateTopicId(authToken, courseId, topicName).pipe(
            tap(resolvedTopicId => console.log(`${itemLogPrefix} Resolved Topic ID for "${topicName || 'None'}": ${resolvedTopicId || 'None'}`)),
            shareReplay(1), // Cache the result of this observable to share among subscribers
            catchError(topicError => {
              console.error(`${itemLogPrefix} CRITICAL error resolving topic "${topicName}". Error: ${topicError.message}`);
              // Propagate a structured error
              return throwError(() => ({
                message: `Failed to resolve/create topic "${topicName || 'None'}": ${topicError.message || topicError}`,
                stage: 'Topic Management',
                details: topicError.details || topicError.toString()
              }));
            })
          );
          topicRequestCache.set(cacheKey, topicId$); // Store the observable in the cache
        } else {
          console.log(`${itemLogPrefix} Using cached Topic ID for "${topicName || 'None'}".`);
        }

        // Function to create an observable for a single classroom item (CourseWork or Material)
        // This is used for both single items and parts of a split item
        const createClassroomItemObservables = (
          currentAssignmentData: ProcessedCourseWork, // Data for the current item/part
          materialsForThisPart: Material[] // Materials specific to this item/part
        ): Observable<ProcessedCourseWork> => {
          console.log(`${itemLogPrefix} (Part: "${currentAssignmentData.title}"): Attempting to get Topic ID.`);
          return topicId$.pipe( // Use the cached/created topicId$ observable
            switchMap(topicIdValue => { // Once topic ID is resolved
              console.log(`${itemLogPrefix} (Part: "${currentAssignmentData.title}"): Topic ID resolved to "${topicIdValue}". Proceeding to create Classroom item.`);
              // Prepare data for the API call
              const assignmentDataForApiCall: ProcessedCourseWork = {
                ...currentAssignmentData,
                materials: materialsForThisPart,
                topicId: topicIdValue,
                description: currentAssignmentData.descriptionForClassroom // Use the classroom-specific description
              };
              // Remove fields not intended for the Classroom API payload
              const {descriptionForDisplay, localFilesToUpload, qtiFile, htmlContent, webLinkUrl, richtext, associatedWithDeveloper, processingError, classroomCourseWorkId, classroomLink, ...apiPayload} = assignmentDataForApiCall;

              // Call the method to create the actual item in Classroom
              return this.createClassroomItem(authToken, courseId, apiPayload as ProcessedCourseWork).pipe(
                map(apiResponse => {
                  // Map successful API response to a ProcessedCourseWork object
                  const resultItem: ProcessedCourseWork = {
                    ...currentAssignmentData,
                    materials: materialsForThisPart,
                    topicId: topicIdValue,
                    classroomCourseWorkId: apiResponse.id,
                    classroomLink: apiResponse.alternateLink,
                    state: apiResponse.state || 'DRAFT',
                    processingError: undefined // Clear any previous error
                  };
                  console.log(`${itemLogPrefix} (Part: "${resultItem.title}"): Successfully created in Classroom. API Response ID: ${apiResponse.id}, Link: ${apiResponse.alternateLink}`);
                  return resultItem;
                }),
                catchError(classroomCreationError => {
                  // Handle errors during Classroom item creation
                  console.error(`${itemLogPrefix} (Part: "${currentAssignmentData.title}"): Failed to create in Classroom. Error: ${classroomCreationError.message || classroomCreationError}`);
                  const errorItem: ProcessedCourseWork = {...currentAssignmentData, materials: materialsForThisPart, topicId: topicIdValue};
                  errorItem.processingError = {
                    message: `Failed to create item "${currentAssignmentData.title}" in Classroom: ${classroomCreationError.message || classroomCreationError}`,
                    stage: currentAssignmentData.workType === 'MATERIAL' ? 'Classroom Material Creation' : 'Classroom CourseWork Creation',
                    details: classroomCreationError.details || classroomCreationError.toString()
                  };
                  return of(errorItem); // Return the item with error information
                })
              );
            }),
            catchError(topicResolutionError => {
              // Handle errors from the topicId$ observable itself (e.g., failure to create/fetch topic)
              console.error(`${itemLogPrefix} (Part: "${currentAssignmentData.title}"): Failed to resolve topic. Error: ${topicResolutionError.message || topicResolutionError}`);
              const errorItem: ProcessedCourseWork = {...currentAssignmentData, materials: materialsForThisPart};
              errorItem.processingError = {
                message: `Failed to resolve/create topic "${topicName || 'None'}": ${topicResolutionError.message || topicResolutionError}`,
                stage: 'Topic Management',
                details: topicResolutionError.details || topicResolutionError.toString()
              };
              return of(errorItem); // Return the item with error information
            })
          );
        };

        // Check if materials exceed the API limit
        if (uniqueMaterials.length <= this.materialLimit) {
          console.log(`${itemLogPrefix} Material count (${uniqueMaterials.length}) is within limit. Creating single Classroom item.`);
          // If within limit, create a single Classroom item
          allOperationsObservables.push(createClassroomItemObservables(assignmentForProcessing, uniqueMaterials));
        } else {
          // If materials exceed limit, split the assignment into parts
          console.log(`${itemLogPrefix} Material count (${uniqueMaterials.length}) exceeds limit of ${this.materialLimit}. Splitting into parts.`);
          const numParts = Math.ceil(uniqueMaterials.length / this.materialLimit);
          const materialChunks = this.chunkArray(uniqueMaterials, this.materialLimit); // Split materials into chunks

          // Create an observable for each part
          for (let i = 0; i < numParts; i++) {
            const partIndex = i + 1;
            const partTitle = `${assignmentForProcessing.title} (Part ${partIndex} of ${numParts})`;
            console.log(`${itemLogPrefix} Preparing Part ${partIndex}/${numParts}: "${partTitle}"`);
            // Create data for this specific part
            const partAssignmentData: ProcessedCourseWork = {
              ...assignmentForProcessing, // Base data from the original assignment
              title: partTitle, // Part-specific title
              descriptionForClassroom: `Part ${partIndex} of ${numParts}:\n\n${assignmentForProcessing.descriptionForClassroom || ''}`, // Part-specific description
              // materials will be set by createClassroomItemObservables for this chunk
            };
            allOperationsObservables.push(createClassroomItemObservables(partAssignmentData, materialChunks[i]));
          }
        }
      } // End loop for assignments
    } // End loop for classroomIds

    // If no operations were prepared (e.g., all assignments failed pre-flight checks)
    if (allOperationsObservables.length === 0) {
      console.warn('[ClassroomService] assignContentToClassrooms: No valid operations to perform after pre-flight checks.');
      // Return original assignments, ensuring any existing processingErrors are preserved or a new one is added
      return of(assignments.map(a => ({...a, processingError: a.processingError || {message: "No valid operation to perform.", stage: "Pre-flight Check (ClassroomService)"}})));
    }

    // Use forkJoin to execute all prepared operations in parallel
    console.log(`[ClassroomService] assignContentToClassrooms: Starting forkJoin for ${allOperationsObservables.length} Classroom item creation operations.`);
    return forkJoin(allOperationsObservables).pipe(
      tap(results => {
        // Log summary of batch assignment
        const successes = results.filter(r => !r.processingError).length;
        const failures = results.length - successes;
        console.log(`[ClassroomService] assignContentToClassrooms: Batch assignment process finished. Total operations: ${results.length}, Successes: ${successes}, Failures: ${failures}`);
        // Log details for each item
        results.forEach((result, idx) => {
          if (result.processingError) {
            console.warn(`[ClassroomService] Failed item [${idx}]: "${result.title}", Error: ${result.processingError.message}, Stage: ${result.processingError.stage}, Details: ${JSON.stringify(result.processingError.details)}`);
          } else {
            console.log(`[ClassroomService] Successful item [${idx}]: "${result.title}", Classroom ID: ${result.classroomCourseWorkId}`);
          }
        });
      }),
      catchError(err => { // Catch errors from forkJoin itself (should be rare as individual errors are handled)
        console.error('[ClassroomService] assignContentToClassrooms: Unexpected error in forkJoin. This should ideally be caught by individual operations.', err);
        return throwError(() => new Error(`Critical error during batch assignment processing: ${err.message || err}`));
      })
    );
  }


  // --- Topic Management Methods ---
  /**
   * Gets the ID of an existing topic by name, or creates it if it doesn't exist.
   * @param authToken The OAuth 2.0 token.
   * @param courseId The ID of the course.
   * @param topicName The name of the topic. If undefined or empty, returns Observable<undefined>.
   * @returns An Observable of the topic ID string, or undefined if no topicName was provided.
   */
  private getOrCreateTopicId(authToken: string, courseId: string, topicName?: string): Observable<string | undefined> {
    const normalizedTopicName = topicName?.trim(); // Trim whitespace, handles undefined
    if (!normalizedTopicName) { // If topicName is empty, null, or undefined after trim
      console.log(`[ClassroomService] getOrCreateTopicId: No topic name provided for course ${courseId}. Item will not have a topic.`);
      return of(undefined); // Return an observable of undefined, indicating no topic
    }

    const lowerCaseTopicName = normalizedTopicName.toLowerCase(); // For case-insensitive comparison
    const context = `getOrCreateTopicId (Course: ${courseId}, Topic: "${normalizedTopicName}")`;
    console.log(`[ClassroomService] ${context}: Starting.`);

    // First, list all topics for the course
    return this.listAllTopics(authToken, courseId).pipe(
      map(allTopics => {
        // Find a topic with a matching name (case-insensitive)
        const found = allTopics.find(topic => topic.name?.toLowerCase() === lowerCaseTopicName);
        console.log(`[ClassroomService] ${context}: Searched ${allTopics.length} topics. Found existing: ${found ? found.topicId : 'No'}`);
        return found; // Return the found topic object or undefined
      }),
      switchMap(existingTopic => {
        // If an existing topic is found, return its ID
        if (existingTopic?.topicId) {
          console.log(`[ClassroomService] ${context}: Using existing topic ID: ${existingTopic.topicId}`);
          return of(existingTopic.topicId);
        } else {
          // If no existing topic is found, create a new one
          console.log(`[ClassroomService] ${context}: Topic not found. Creating new topic.`);
          return this.createTopic(authToken, courseId, normalizedTopicName).pipe(
            map(newTopic => {
              console.log(`[ClassroomService] ${context}: New topic created with ID: ${newTopic.topicId}`);
              return newTopic.topicId; // Return the ID of the newly created topic
            })
          );
        }
      }),
      catchError(err => {
        console.error(`[ClassroomService] ${context}: Error during operation.`, err);
        return this.handleError(err, context); // Use the centralized error handler
      })
    );
  }

  /**
   * Lists all topics for a given course, handling pagination.
   * @param authToken The OAuth 2.0 token.
   * @param courseId The ID of the course.
   * @returns An Observable array of Topic objects.
   */
  private listAllTopics(authToken: string, courseId: string): Observable<Topic[]> {
    const context = `listAllTopics (Course: ${courseId})`;
    console.log(`[ClassroomService] ${context}: Fetching all topics.`);
    // Fetch the first page of topics
    return this.fetchTopicPage(authToken, courseId, undefined, context).pipe(
      expand(response => {
        // If nextPageToken exists, fetch the next page
        if (response.nextPageToken) {
          console.log(`[ClassroomService] ${context}: Fetching next page of topics (Token: ${response.nextPageToken}).`);
          return this.fetchTopicPage(authToken, courseId, response.nextPageToken, `${context} (paginated)`);
        }
        return EMPTY; // Complete the expand operator
      }),
      map(response => response.topic || []), // Extract topics from each response
      reduce((acc, topics) => acc.concat(topics), [] as Topic[]), // Accumulate all topics into a single array
      tap(allTopics => console.log(`[ClassroomService] ${context}: Successfully fetched ${allTopics.length} topics.`)),
      catchError(err => this.handleError(err, `${context} (Accumulation)`))
    );
  }

  /**
   * Fetches a single page of topics for a course.
   * @param authToken The OAuth 2.0 token.
   * @param courseId The ID of the course.
   * @param pageToken Optional token for fetching a specific page.
   * @param context A string describing the context of the call for logging.
   * @returns An Observable of ListTopicsResponse.
   */
  private fetchTopicPage(authToken: string, courseId: string, pageToken?: string, context: string = 'fetchTopicPage'): Observable<ListTopicsResponse> {
    const headers = this.createAuthHeaders(authToken);
    let params = new HttpParams().set('pageSize', this.pageSize); // Use configured page size
    if (pageToken) {
      params = params.set('pageToken', pageToken);
    }
    const url = this.topicsApiUrl(courseId);
    const operationDescription = `${context} (Course: ${courseId}, Page Token: ${pageToken ?? 'initial'})`;
    console.log(`[ClassroomService] ${operationDescription}: Making API call to ${url}.`);
    const request$ = this.http.get<ListTopicsResponse>(url, {headers, params});
    // Retry the request using the utility service
    return this.utils.retryRequest(request$, this.defaultRetryConfig, operationDescription).pipe(
      catchError(err => this.handleError(err, operationDescription))
    );
  }

  /**
   * Creates a new topic in a course.
   * @param authToken The OAuth 2.0 token.
   * @param courseId The ID of the course.
   * @param topicName The name for the new topic.
   * @returns An Observable of the created Topic object.
   */
  private createTopic(authToken: string, courseId: string, topicName: string): Observable<Topic> {
    const headers = this.createAuthHeaders(authToken);
    const url = this.topicsApiUrl(courseId);
    const body = {name: topicName}; // API request body
    const context = `createTopic (Course: ${courseId}, Topic: "${topicName}")`;
    console.log(`[ClassroomService] ${context}: Making API call to ${url} with body:`, body);
    const request$ = this.http.post<Topic>(url, body, {headers});
    // Retry the request using the utility service
    return this.utils.retryRequest(request$, this.defaultRetryConfig, context).pipe(
      map(topic => {
        console.log(`[ClassroomService] ${context}: Topic created successfully. ID: ${topic.topicId}, Name: "${topic.name}"`);
        return topic;
      }),
      catchError(err => this.handleError(err, context))
    );
  }

  // --- Unified CourseWork / CourseWorkMaterial Creation Method ---
  /**
   * Creates either a CourseWork item (assignment, question) or a CourseWorkMaterial item in Classroom.
   * The type is determined by `itemData.workType`.
   * @param authToken The OAuth 2.0 token.
   * @param courseId The ID of the course.
   * @param itemData Data for the item to be created. Should be ProcessedCourseWork with only API-relevant fields.
   * @returns An Observable of the created CourseWork or CourseWorkMaterial object.
   */
  private createClassroomItem(
    authToken: string,
    courseId: string,
    itemData: ProcessedCourseWork // Expects ProcessedCourseWork with only API relevant fields
  ): Observable<CourseWork | CourseWorkMaterial> {
    const headers = this.createAuthHeaders(authToken);
    const itemLogPrefix = `[ClassroomService] createClassroomItem (Course: ${courseId}, Title: "${itemData.title}", Type: ${itemData.workType})`;

    console.log(`${itemLogPrefix}: Preparing to create item.`);

    // Validate required fields
    if (!itemData.title || !itemData.workType) {
      const errorMsg = `${itemLogPrefix}: Missing required field(s) (title, workType). Cannot create item.`;
      console.error(errorMsg, 'Item Data:', itemData);
      return throwError(() => new Error(errorMsg));
    }

    // Handle CourseWorkMaterial creation
    if (itemData.workType === 'MATERIAL') {
      const materialBody: CourseWorkMaterial = {
        title: itemData.title,
        description: itemData.description, // This should be descriptionForClassroom
        materials: itemData.materials,
        state: itemData.state || 'DRAFT', // Default to DRAFT if not specified
        topicId: itemData.topicId,
        scheduledTime: itemData.scheduledTime
        // assigneeMode is not applicable for materials (always ALL_STUDENTS by API default)
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
    } else { // Handle CourseWork creation (ASSIGNMENT, SHORT_ANSWER_QUESTION, MULTIPLE_CHOICE_QUESTION)
      // Validate workType for CourseWork
      const validWorkTypes: Array<CourseWork['workType']> = ['ASSIGNMENT', 'SHORT_ANSWER_QUESTION', 'MULTIPLE_CHOICE_QUESTION'];
      if (!validWorkTypes.includes(itemData.workType as any)) {
        const errorMsg = `${itemLogPrefix}: Invalid workType "${itemData.workType}" for CourseWork.`;
        console.error(errorMsg, 'Item Data:', itemData);
        return throwError(() => new Error(errorMsg));
      }

      const courseWorkBody: CourseWork = {
        title: itemData.title,
        description: itemData.description, // This should be descriptionForClassroom
        materials: itemData.materials,
        workType: itemData.workType as 'ASSIGNMENT' | 'SHORT_ANSWER_QUESTION' | 'MULTIPLE_CHOICE_QUESTION',
        state: itemData.state || 'DRAFT', // Default to DRAFT
        topicId: itemData.topicId,
        maxPoints: itemData.maxPoints,
        // Conditionally add assignment-specific or question-specific fields
        assignment: itemData.workType === 'ASSIGNMENT' ? itemData.assignment : undefined,
        multipleChoiceQuestion: itemData.workType === 'MULTIPLE_CHOICE_QUESTION' ? itemData.multipleChoiceQuestion : undefined,
        dueDate: itemData.dueDate,
        dueTime: itemData.dueTime,
        scheduledTime: itemData.scheduledTime,
        submissionModificationMode: itemData.submissionModificationMode
        // assigneeMode defaults to ALL_STUDENTS if not specified, as per Classroom API behavior
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
  /**
   * Creates standard HTTP headers for API requests, including Authorization.
   * @param authToken The OAuth 2.0 token.
   * @returns HttpHeaders object.
   */
  private createAuthHeaders(authToken: string): HttpHeaders {
    return new HttpHeaders({
      'Authorization': `Bearer ${authToken}`,
      'Accept': 'application/json',
      'Content-Type': 'application/json' // Important for POST/PATCH requests
    });
  }

  /**
   * Splits an array into chunks of a specified size.
   * @param array The array to chunk.
   * @param size The size of each chunk.
   * @returns A new array containing the chunks.
   */
  private chunkArray<T>(array: T[], size: number): T[][] {
    if (!array) return [];
    const result: T[][] = [];
    for (let i = 0; i < array.length; i += size) {
      result.push(array.slice(i, i + size));
    }
    return result;
  }

  /**
   * Generates a unique key for a material based on its type and ID/URL.
   * Used for deduplication.
   * @param material The Material object.
   * @returns A string key or null if no identifiable key can be generated.
   */
  private getMaterialKey(material: Material): string | null {
    if (material.driveFile?.driveFile?.id) return `drive-${material.driveFile.driveFile.id}`;
    if (material.youtubeVideo?.id) return `youtube-${material.youtubeVideo.id}`;
    if (material.link?.url) return `link-${material.link.url}`;
    if (material.form?.formUrl) return `form-${material.form.formUrl}`;
    // If other material types are supported, add their key generation here
    return null; // Should not happen if material is a recognized type and valid
  }

  /**
   * Deduplicates an array of Material objects based on their unique keys.
   * @param materials Array of Material objects.
   * @returns A new array with unique Material objects.
   */
  private deduplicateMaterials(materials: Material[]): Material[] {
    if (!materials || materials.length === 0) return [];
    const seenKeys = new Set<string>();
    const uniqueMaterials: Material[] = [];
    for (const material of materials) {
      const key = this.getMaterialKey(material);
      if (key !== null && !seenKeys.has(key)) { // If key is valid and not seen before
        seenKeys.add(key);
        uniqueMaterials.push(material);
      } else if (key === null) { // If material has no identifiable key (e.g., malformed or new type)
        console.warn('[ClassroomService] deduplicateMaterials: Encountered material with no identifiable key, including as is.', material);
        uniqueMaterials.push(material); // Include it but log a warning
      } else { // If key has been seen (duplicate)
        console.log(`[ClassroomService] deduplicateMaterials: Duplicate material skipped (Key: ${key})`);
      }
    }
    return uniqueMaterials;
  }

  /**
   * Centralized error handler for API requests.
   * Logs the error and throws a new error with a user-friendly message and details.
   * @param error The error object (HttpErrorResponse or Error).
   * @param context A string describing the operation context where the error occurred.
   * @returns An Observable that throws the processed error.
   */
  private handleError(error: HttpErrorResponse | Error, context: string = 'Unknown Operation'): Observable<never> {
    let userMessage = `Failed during ${context}; please try again later or check console for details.`;
    let detailedMessage = `Context: ${context} - An unknown error occurred!`;
    let statusCode: number | undefined = undefined;
    let errorDetailsForPropagation: any = error; // Default to the original error object for propagation

    if (error instanceof HttpErrorResponse) {
      statusCode = error.status;
      detailedMessage = `Context: ${context} - Server error: Code ${error.status}, Message: ${error.message || 'No message body'}`;
      userMessage = `The server returned an error (Code: ${error.status}) while processing ${context}. Please check details or try again.`;
      errorDetailsForPropagation = error.error || error.message; // Prefer error.error (parsed body) if available
      try {
        const errorBody = JSON.stringify(error.error); // Attempt to stringify the error body for logging
        detailedMessage += `, Body: ${errorBody}`;
        // Check for standard Google API error structure in the response body
        const googleApiError = error.error?.error?.message;
        if (googleApiError) {
          userMessage = `Google API Error in ${context}: ${googleApiError} (Code: ${error.status})`;
          detailedMessage += ` | Google API Specific: ${googleApiError}`;
        }
      } catch (e) { /* Ignore JSON stringify errors if error.error is not a simple object */}
    } else if (error instanceof Error) { // Client-side or network errors
      detailedMessage = `Context: ${context} - Client/Network error: ${error.message}`;
      userMessage = `A network or client-side error occurred in ${context}. Please check your connection or the console. Details: ${error.message}`;
      errorDetailsForPropagation = error.message;
    }

    console.error(`[ClassroomService] handleError: ${detailedMessage}`, 'Full Error Object:', error);

    // Create a new error object to be thrown, including user-friendly message and structured details
    const finalError = new Error(userMessage);
    (finalError as any).status = statusCode; // Attach HTTP status code if available
    (finalError as any).details = errorDetailsForPropagation; // Attach more detailed error info (e.g., server response body)
    (finalError as any).stage = context; // Add context as a 'stage' property for easier debugging

    return throwError(() => finalError); // Return an observable that emits the error
  }
}
