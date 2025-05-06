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
  private pageSize = '50';
  private materialLimit = 20;

  private defaultRetryConfig: RetryConfig = {maxRetries: 3, initialDelayMs: 2000};

  constructor() { }

  // --- Get Classrooms ---
  getActiveClassrooms(authToken: string): Observable<Classroom[]> {
    if (!authToken) {
      console.error('ClassroomService: Auth token is missing for getActiveClassrooms.');
      return of([]);
    }
    const context = 'getActiveClassrooms';
    return this.fetchClassroomPage(authToken, undefined, context).pipe(
      expand(response => response.nextPageToken
          ? this.fetchClassroomPage(authToken, response.nextPageToken, `${context} (paginated)`)
          : EMPTY
      ),
      map(response => response.courses || []),
      reduce((acc, courses) => acc.concat(courses), [] as Classroom[]),
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
    if (!authToken) return throwError(() => new Error('Authentication token is required.'));
    if (!classroomIds?.length) {
      console.warn('ClassroomService: No classroom IDs provided for assignment.');
      return of(assignments.map(a => ({...a, processingError: {message: "No classroom selected.", stage: "Pre-flight Check"}})));
    }
    if (!assignments?.length) {
      console.warn('ClassroomService: No assignments provided for processing.');
      return of([]);
    }

    const topicRequestCache = new Map<string, Observable<string | undefined>>();
    const allOperationsObservables: Observable<ProcessedCourseWork>[] = [];

    for (const courseId of classroomIds) {
      for (const originalAssignment of assignments) {
        const assignmentForProcessing: ProcessedCourseWork = { ...originalAssignment };

        if (!assignmentForProcessing.title || !assignmentForProcessing.workType) {
             console.warn(`Skipping assignment due to missing title or workType:`, assignmentForProcessing.title);
             assignmentForProcessing.processingError = {
                 message: 'Skipped: Missing title or workType.',
                 stage: 'Pre-flight Check'
             };
             allOperationsObservables.push(of(assignmentForProcessing));
             continue;
        }

        const topicName = assignmentForProcessing.associatedWithDeveloper?.topic;
        const uniqueMaterials = this.deduplicateMaterials(assignmentForProcessing.materials || []);

        const cacheKey = `${courseId}:${topicName?.trim().toLowerCase() || 'undefined'}`;
        let topicId$: Observable<string | undefined> | undefined = topicRequestCache.get(cacheKey);

        if (!topicId$) {
            topicId$ = this.getOrCreateTopicId(authToken, courseId, topicName).pipe(
                shareReplay(1),
                catchError(topicError => {
                    console.error(`Critical error resolving topic "${topicName}" for course ${courseId}. Error: ${topicError.message}`);
                    return throwError(() => topicError);
                })
            );
            topicRequestCache.set(cacheKey, topicId$);
        }

        const createClassroomItemObservables = ( // Renamed from createCourseWorkObservablesForAssignment
            currentAssignmentData: ProcessedCourseWork,
            materialsForThisPart: Material[]
        ): Observable<ProcessedCourseWork> => {
            return topicId$.pipe(
              switchMap(topicIdValue => {
                    const assignmentDataForApiCall: ProcessedCourseWork = {
                      ...currentAssignmentData,
                      materials: materialsForThisPart,
                      topicId: topicIdValue,
                      description: currentAssignmentData.descriptionForClassroom // Ensure API description uses classroom version
                    };

                  // No longer destructuring descriptionForDisplay out. Pass the full object.
                  // const { descriptionForDisplay, ...apiPayload } = assignmentDataForApiCall; // REMOVED THIS LINE

                  return this.createClassroomItem(authToken, courseId, assignmentDataForApiCall).pipe( // Pass full assignmentDataForApiCall
                    map(apiResponse => {
                      const resultItem: ProcessedCourseWork = {
                        ...currentAssignmentData, // Start with data for this part
                        materials: materialsForThisPart,
                        topicId: topicIdValue,
                        classroomCourseWorkId: apiResponse.id,
                        classroomLink: apiResponse.alternateLink,
                        state: apiResponse.state || 'DRAFT',
                        processingError: undefined
                            };
                          console.log(`Successfully created "${resultItem.title}" in Classroom. API Response ID: ${apiResponse.id}, Link: ${apiResponse.alternateLink}`);
                            return resultItem;
                        }),
                        catchError(classroomCreationError => {
                          const errorItem: ProcessedCourseWork = {...currentAssignmentData, materials: materialsForThisPart, topicId: topicIdValue};
                            errorItem.processingError = {
                                message: `Failed to create in Classroom: ${classroomCreationError.message || classroomCreationError}`,
                              stage: currentAssignmentData.workType === 'MATERIAL' ? 'Classroom Material Creation' : 'Classroom CourseWork Creation',
                              details: classroomCreationError.details || classroomCreationError.toString()
                            };
                            return of(errorItem);
                        })
                    );
                }),
                catchError(topicResolutionError => {
                    const errorItem: ProcessedCourseWork = { ...currentAssignmentData, materials: materialsForThisPart };
                    errorItem.processingError = {
                        message: `Failed to resolve/create topic "${topicName || 'None'}": ${topicResolutionError.message || topicResolutionError}`,
                      stage: 'Topic Management',
                      details: topicResolutionError.details || topicResolutionError.toString()
                    };
                    return of(errorItem);
                })
            );
        };

        if (uniqueMaterials.length <= this.materialLimit) {
          allOperationsObservables.push(createClassroomItemObservables(assignmentForProcessing, uniqueMaterials));
        } else {
          console.log(`Assignment "${assignmentForProcessing.title}" exceeds material limit (${uniqueMaterials.length} > ${this.materialLimit}). Splitting into parts for course ${courseId}.`);
          const numParts = Math.ceil(uniqueMaterials.length / this.materialLimit);
          const materialChunks = this.chunkArray(uniqueMaterials, this.materialLimit);

          for (let i = 0; i < numParts; i++) {
            const partIndex = i + 1;
            const partAssignmentData: ProcessedCourseWork = {
              ...assignmentForProcessing,
                title: `${assignmentForProcessing.title} (Part ${partIndex} of ${numParts})`,
              // descriptionForDisplay is inherited from assignmentForProcessing
                descriptionForClassroom: `Part ${partIndex} of ${numParts}:\n\n${assignmentForProcessing.descriptionForClassroom || ''}`,
              // materials will be set by createClassroomItemObservables
            };
            allOperationsObservables.push(createClassroomItemObservables(partAssignmentData, materialChunks[i]));
          }
        }
      }
    }

    if (allOperationsObservables.length === 0) {
        console.warn('ClassroomService: No valid operations to perform.');
      return of(assignments.map(a => ({...a, processingError: a.processingError || {message: "No valid operation to perform.", stage: "Pre-flight Check"}})));
    }

    return forkJoin(allOperationsObservables).pipe(
      tap(results => {
        const successes = results.filter(r => !r.processingError).length;
        const failures = results.length - successes;
        console.log(`ClassroomService: Batch assignment process finished. Successes: ${successes}, Failures: ${failures}`);
        results.forEach(result => {
            if(result.processingError) {
                console.warn(`Failed item: "${result.title}", Error: ${result.processingError.message}, Stage: ${result.processingError.stage}`);
            }
        });
      })
    );
  }


  // --- Topic Management Methods ---
  private getOrCreateTopicId(authToken: string, courseId: string, topicName?: string): Observable<string | undefined> {
    if (!topicName || topicName.trim() === '') {
      return of(undefined);
    }
    const normalizedTopicName = topicName.trim();
    const lowerCaseTopicName = normalizedTopicName.toLowerCase();
    const context = `getOrCreateTopicId (Course: ${courseId}, Topic: ${normalizedTopicName})`;

    return this.listAllTopics(authToken, courseId).pipe(
      map(allTopics => allTopics.find(topic => topic.name?.toLowerCase() === lowerCaseTopicName)),
      switchMap(existingTopic => {
        if (existingTopic?.topicId) {
          console.log(`Found existing topic "${normalizedTopicName}" with ID: ${existingTopic.topicId} in course ${courseId}`);
          return of(existingTopic.topicId);
        } else {
          console.log(`Topic "${normalizedTopicName}" not found in course ${courseId}. Creating new topic.`);
          return this.createTopic(authToken, courseId, normalizedTopicName).pipe(
            map(newTopic => newTopic.topicId)
          );
        }
      }),
      catchError(err => this.handleError(err, context))
    );
  }

  private listAllTopics(authToken: string, courseId: string): Observable<Topic[]> {
     const context = `listAllTopics (Course: ${courseId})`;
     return this.fetchTopicPage(authToken, courseId, undefined, context).pipe(
        expand(response => response.nextPageToken
            ? this.fetchTopicPage(authToken, courseId, response.nextPageToken, `${context} (paginated)`)
            : EMPTY
        ),
        map(response => response.topic || []),
        reduce((acc, topics) => acc.concat(topics), [] as Topic[]),
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
    const request$ = this.http.get<ListTopicsResponse>(url, {headers, params});
    return this.utils.retryRequest(request$, this.defaultRetryConfig, operationDescription).pipe(
        catchError(err => this.handleError(err, operationDescription))
    );
  }

  private createTopic(authToken: string, courseId: string, topicName: string): Observable<Topic> {
    const headers = this.createAuthHeaders(authToken);
    const url = this.topicsApiUrl(courseId);
    const body = { name: topicName };
    const context = `createTopic (Course: ${courseId}, Topic: ${topicName})`;
    const request$ = this.http.post<Topic>(url, body, {headers});
    return this.utils.retryRequest(request$, this.defaultRetryConfig, context).pipe(
       map(topic => {
            console.log(`Topic "${topic.name}" created successfully with ID: ${topic.topicId} in course ${courseId}`);
            return topic;
       }),
       catchError(err => this.handleError(err, context))
    );
  }

  // --- Unified CourseWork / CourseWorkMaterial Creation Method ---
  private createClassroomItem(
    authToken: string,
    courseId: string,
    itemData: ProcessedCourseWork // Expects full ProcessedCourseWork
  ): Observable<CourseWork | CourseWorkMaterial> {
    const headers = this.createAuthHeaders(authToken);
    const context = `createClassroomItem (Course: ${courseId}, Title: ${itemData.title})`;

    if (!itemData.title || !itemData.workType) {
      const errorMsg = `Cannot create Classroom item within ${context}: Missing required field(s) (title, workType).`;
      console.error(errorMsg, itemData);
      return throwError(() => new Error(errorMsg));
    }

    if (itemData.workType === 'MATERIAL') {
      const materialBody: CourseWorkMaterial = {
        title: itemData.title,
        description: itemData.descriptionForClassroom,
        materials: itemData.materials,
        state: itemData.state || 'PUBLISHED',
        topicId: itemData.topicId,
        scheduledTime: itemData.scheduledTime
      };
      const url = this.courseWorkMaterialsApiUrl(courseId);
      console.log('ClassroomService: Submitting CourseWorkMaterial Body:', materialBody);
      const request$ = this.http.post<CourseWorkMaterial>(url, materialBody, {headers});
      return this.utils.retryRequest(request$, this.defaultRetryConfig, `${context} as MATERIAL`).pipe(
        map(createdMaterial => {
          console.log(`CourseWorkMaterial "${createdMaterial.title}" created successfully with ID: ${createdMaterial.id} in course ${courseId}`);
          return createdMaterial;
        }),
        catchError(err => this.handleError(err, `${context} as MATERIAL`))
      );
    } else {
      const courseWorkBody: CourseWork = {
        title: itemData.title,
        description: itemData.descriptionForClassroom,
        materials: itemData.materials,
        workType: itemData.workType as 'ASSIGNMENT' | 'SHORT_ANSWER_QUESTION' | 'MULTIPLE_CHOICE_QUESTION',
        state: itemData.state || 'PUBLISHED',
        topicId: itemData.topicId,
        maxPoints: itemData.maxPoints,
        assignment: itemData.workType === 'ASSIGNMENT' ? itemData.assignment : undefined,
        multipleChoiceQuestion: itemData.workType === 'MULTIPLE_CHOICE_QUESTION' ? itemData.multipleChoiceQuestion : undefined,
        dueDate: itemData.dueDate,
        dueTime: itemData.dueTime,
        scheduledTime: itemData.scheduledTime,
        submissionModificationMode: itemData.submissionModificationMode
      };
      const url = this.courseWorkApiUrl(courseId);
      console.log('ClassroomService: Submitting CourseWork Body:', courseWorkBody);
      const request$ = this.http.post<CourseWork>(url, courseWorkBody, {headers});
      return this.utils.retryRequest(request$, this.defaultRetryConfig, context).pipe(
        map(createdWork => {
          console.log(`CourseWork "${createdWork.title}" created successfully with ID: ${createdWork.id} in course ${courseId}`);
          return createdWork;
        }),
        catchError(err => this.handleError(err, context))
      );
    }
  }

  // --- Utility and Error Handling ---
  private createAuthHeaders(authToken: string): HttpHeaders {
     return new HttpHeaders({
        'Authorization': `Bearer ${authToken}`,
        'Accept': 'application/json',
        'Content-Type': 'application/json'
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
      return null;
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
          } else if (key === null) {
              uniqueMaterials.push(material);
          }
      }
      return uniqueMaterials;
  }

  private handleError(error: HttpErrorResponse | Error, context: string = 'Unknown Operation'): Observable<never> {
    let userMessage = `Failed during ${context}; please try again later or check console for details.`;
    let detailedMessage = `Context: ${context} - An unknown error occurred!`;
    let statusCode: number | undefined = undefined;

    if (error instanceof HttpErrorResponse) {
        statusCode = error.status;
        detailedMessage = `Context: ${context} - Server error: Code ${error.status}, Message: ${error.message || 'No message body'}`;
        userMessage = `The server returned an error (Code: ${error.status}). Please check details or try again.`;
        try {
            const errorBody = JSON.stringify(error.error);
            detailedMessage += `, Body: ${errorBody}`;
            const googleApiError = error.error?.error?.message;
            if (googleApiError) {
               userMessage = `Google API Error in ${context}: ${googleApiError} (Code: ${error.status})`;
               detailedMessage += ` | Google API Specific: ${googleApiError}`;
            }
        } catch (e) { /* Ignore JSON stringify errors */ }
    } else if (error instanceof Error) {
       detailedMessage = `Context: ${context} - Client/Network error: ${error.message}`;
       userMessage = `A network or client-side error occurred in ${context}. Please check your connection or the console.`;
    }

    console.error('ClassroomService Error:', detailedMessage, 'Full Error Object:', error);
    const finalError = new Error(userMessage);
    if (statusCode !== undefined) {
        (finalError as any).status = statusCode;
    }
    (finalError as any).details = detailedMessage;
    return throwError(() => finalError);
  }
}
