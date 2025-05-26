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
  CourseWorkMaterial,
} from '../../interfaces/classroom-interface';

import {
  UtilitiesService,
  RetryConfig,
  BatchOperation,
  BatchResponseParser
} from '../utilities/utilities.service'; // Assuming path is correct
import {AuthService} from '../auth/auth.service'; // Import AuthService

@Injectable({
  providedIn: 'root'
})
export class ClassroomService {

  // Base API URLs for Google Classroom
  private coursesApiUrl = 'https://classroom.googleapis.com/v1/courses';
  private topicsApiUrl = (courseId: string) => `${this.coursesApiUrl}/${courseId}/topics`;

  // Injected services
  private http = inject(HttpClient);
  private utils = inject(UtilitiesService);
  private auth = inject(AuthService); // Inject AuthService

  // Default pagination and material limits
  private pageSize = '50';
  private materialLimit = 20;

  // Default configuration for retrying failed HTTP requests
  private defaultRetryConfig: Required<RetryConfig> = {
    maxRetries: 5,
    initialDelayMs: 2000,
    backoffFactor: 2,
    retryableStatusCodes: [429, 500, 503, 504]
  };

  constructor() { }

  /**
   * Fetches all active classrooms for the authenticated user.
   * Token is fetched internally.
   */
  getActiveClassrooms(): Observable<Classroom[]> {
    const context = 'getActiveClassrooms';
    const authToken = this.auth.getGoogleAccessToken();
    if (!authToken) {
      console.error(`[ClassroomService] ${context}: Auth token is missing.`);
      return throwError(() => new Error('Authentication token is required for getActiveClassrooms.'));
    }

    return this.fetchClassroomPage(undefined, context, authToken).pipe( // Pass token to first call
      expand(response => {
        if (response.nextPageToken) {
          return this.fetchClassroomPage(response.nextPageToken, `${context} (paginated)`, authToken);
        }
        return EMPTY;
      }),
      map(response => response.courses || []),
      reduce((acc, courses) => acc.concat(courses), [] as Classroom[]),
      tap(allCourses => console.log(`[ClassroomService] ${context}: Successfully fetched ${allCourses.length} active classrooms.`)),
      catchError(err => this.handleError(err, `${context} (Accumulation)`))
    );
  }

  /**
   * Fetches a single page of classrooms.
   * Expects authToken to be passed, or fetches it if not provided (for internal consistency).
   */
  private fetchClassroomPage(pageToken?: string, context: string = 'fetchClassroomPage', authToken?: string): Observable<ClassroomListResponse> {
    const tokenToUse = authToken || this.auth.getGoogleAccessToken();
    if (!tokenToUse) {
      const errorMsg = `[ClassroomService] ${context}: Auth token is missing.`;
      console.error(errorMsg);
      return throwError(() => new Error(errorMsg));
    }
    const headers = this.createAuthHeaders(tokenToUse);
    let params = new HttpParams().set('courseStates', 'ACTIVE').set('pageSize', this.pageSize);
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
   * Parser function specific to Google Classroom API responses for batch operations.
   */
  private classroomBatchResponseParser: BatchResponseParser<ProcessedCourseWork> = (
    processedItem, responseJson, statusCode, statusText, operationId
  ): void => {
    const itemTitleForLog = processedItem.title ? `"${processedItem.title.substring(0, 20)}..."` : "(No Title)";
    const itemLogPrefix = `[CS BatchOp ID ${operationId} ${itemTitleForLog} PARSER]:`;

    if (statusCode >= 200 && statusCode < 300) {
      processedItem.classroomCourseWorkId = responseJson.id;
      processedItem.classroomLink = responseJson.alternateLink;
      processedItem.state = responseJson.state || 'DRAFT';
      processedItem.processingError = undefined;
      console.log(`${itemLogPrefix} Parsed Success (Status ${statusCode}). ID: ${responseJson.id}, Link: ${responseJson.alternateLink}`);
    } else {
      const apiErrorMessage = responseJson.error?.message || statusText || 'Unknown error in sub-response.';
      console.warn(`${itemLogPrefix} Parsed Failure (Status ${statusCode}). API Error: "${apiErrorMessage}"`);
      processedItem.processingError = {
        message: `Batch item failed: ${apiErrorMessage} (Status: ${statusCode})`,
        stage: processedItem.workType === 'MATERIAL' ? 'Batch Material Creation Error (CS)' : 'Batch CourseWork Creation Error (CS)',
        details: {statusCode: statusCode, errorBody: responseJson.error || responseJson, opId: operationId}
      };
    }
  };

  /**
   * Assigns content to multiple classrooms. Token is fetched internally.
   */
  assignContentToClassrooms(
    classroomIds: string[],
    assignments: ProcessedCourseWork[]
  ): Observable<ProcessedCourseWork[]> {
    const serviceCallId = `cs-assign-${Date.now()}`;
    console.log(`[ClassroomService][${serviceCallId}] assignContentToClassrooms: Starting for ${assignments.length} items to ${classroomIds.length} classrooms.`);

    const authToken = this.auth.getGoogleAccessToken(); // Fetch token for this entire operation
    if (!authToken) {
      const errorMsg = `[ClassroomService][${serviceCallId}] Auth token missing. Cannot assign content.`;
      console.error(errorMsg);
      // Mark all assignments as failed due to auth token issue
      const assignmentsWithError = assignments.map(a => ({
        ...a,
        processingError: {message: "Authentication token missing for batch operation.", stage: "Pre-flight Auth (CS)"}
      }));
      return of(assignmentsWithError); // Return assignments marked with error
    }

    if (!classroomIds?.length) {
      return of(assignments.map(a => ({...a, processingError: {message: "No classroom selected.", stage: "Pre-flight Check (CS)"}})));
    }
    if (!assignments?.length) {
      return of([]);
    }

    const topicRequestCache = new Map<string, Observable<string | undefined>>();
    const preparationObservables: Observable<void>[] = [];
    const allItemsToTrackForBatchResult: ProcessedCourseWork[] = [];
    const batchOperationsToExecute: BatchOperation<ProcessedCourseWork, CourseWork | CourseWorkMaterial>[] = [];
    let operationIdCounter = 0;

    for (const courseId of classroomIds) {
      for (const originalAssignment of assignments) {
        const baseItemForProcessing: ProcessedCourseWork = {
          ...originalAssignment,
          title: originalAssignment.title || 'Untitled Assignment',
          materials: originalAssignment.materials ? [...originalAssignment.materials] : [],
          processingError: undefined
        };
        const itemLogPrefix = `[CS PREP][${serviceCallId}] Item "${baseItemForProcessing.title}" (DevID: ${baseItemForProcessing.associatedWithDeveloper?.id}) for Course ID ${courseId}:`;

        if (!baseItemForProcessing.title || !baseItemForProcessing.workType) {
          baseItemForProcessing.processingError = {message: 'Skipped: Missing title or workType.', stage: 'Pre-flight Check (CS)'};
          if (!allItemsToTrackForBatchResult.find(item => item === baseItemForProcessing)) {
            allItemsToTrackForBatchResult.push(baseItemForProcessing);
          }
          console.warn(`${itemLogPrefix} Missing title or workType.`);
          continue;
        }

        const prepObservableForItem = new Observable<void>(observer => {
          const topicName = baseItemForProcessing.associatedWithDeveloper?.topic;
          const cacheKey = `${courseId}:${topicName?.trim().toLowerCase() || 'undefined_topic_key'}`;
          let topicId$ = topicRequestCache.get(cacheKey);

          if (!topicId$) {
            // getOrCreateTopicId will fetch its own token
            topicId$ = this.getOrCreateTopicId(courseId, topicName).pipe(
              tap(resolvedTopicId => console.log(`${itemLogPrefix} Topic "${topicName || 'None'}" for course ${courseId} resolved to ID: ${resolvedTopicId || 'None'}`)),
              shareReplay(1),
              catchError(topicError => {
                console.error(`${itemLogPrefix} CRITICAL error resolving topic "${topicName}".`, topicError);
                baseItemForProcessing.processingError = {
                  message: `Failed to resolve/create topic "${topicName || 'None'}": ${topicError.message || String(topicError)}`,
                  stage: 'Topic Management (CS)',
                  details: topicError.details || String(topicError)
                };
                if (!allItemsToTrackForBatchResult.find(item => item === baseItemForProcessing)) {
                  allItemsToTrackForBatchResult.push(baseItemForProcessing);
                }
                return of(undefined);
              })
            );
            topicRequestCache.set(cacheKey, topicId$);
          }

          topicId$.subscribe(topicIdValue => {
            if (baseItemForProcessing.processingError) {
              if (!allItemsToTrackForBatchResult.find(item => item === baseItemForProcessing)) {
                allItemsToTrackForBatchResult.push(baseItemForProcessing);
              }
              observer.next(); observer.complete(); return;
            }

            const uniqueMaterials = this.deduplicateMaterials(baseItemForProcessing.materials || []);
            let materialChunks = uniqueMaterials.length > this.materialLimit ?
              this.chunkArray(uniqueMaterials, this.materialLimit) :
              [uniqueMaterials];
            if (materialChunks.length === 0 && uniqueMaterials.length === 0) {
              materialChunks.push([]);
            }

            materialChunks.forEach((materialChunkForPart, index) => {
              const numParts = materialChunks.length || 1;
              const partSuffix = numParts > 1 ? ` (Part ${index + 1} of ${numParts})` : '';
              const effectiveItemTitle = `${baseItemForProcessing.title}${partSuffix}`;
              const effectiveItemForPart: ProcessedCourseWork = {
                ...baseItemForProcessing,
                title: effectiveItemTitle,
                materials: materialChunkForPart,
                classroomCourseWorkId: undefined,
                classroomLink: undefined,
                state: baseItemForProcessing.state || 'DRAFT',
                processingError: undefined,
              };
              if (numParts > 1) {
                effectiveItemForPart.descriptionForClassroom = `Part ${index + 1} of ${numParts}:\n\n${baseItemForProcessing.descriptionForClassroom || ''}`;
              }
              allItemsToTrackForBatchResult.push(effectiveItemForPart);
              let path: string;
              let operationBody: CourseWork | CourseWorkMaterial;
              if (effectiveItemForPart.workType === 'MATERIAL') {
                path = `/v1/courses/${courseId}/courseWorkMaterials`;
                operationBody = { /* ... as before ... */
                  title: effectiveItemTitle,
                  description: effectiveItemForPart.descriptionForClassroom,
                  materials: materialChunkForPart,
                  state: effectiveItemForPart.state,
                  topicId: topicIdValue,
                  scheduledTime: effectiveItemForPart.scheduledTime,
                };
              } else {
                path = `/v1/courses/${courseId}/courseWork`;
                operationBody = { /* ... as before ... */
                  title: effectiveItemTitle,
                  description: effectiveItemForPart.descriptionForClassroom,
                  materials: materialChunkForPart,
                  workType: effectiveItemForPart.workType as CourseWork['workType'],
                  state: effectiveItemForPart.state,
                  topicId: topicIdValue,
                  maxPoints: effectiveItemForPart.maxPoints,
                  assignment: effectiveItemForPart.workType === 'ASSIGNMENT' ? effectiveItemForPart.assignment : undefined,
                  multipleChoiceQuestion: effectiveItemForPart.workType === 'MULTIPLE_CHOICE_QUESTION' ? effectiveItemForPart.multipleChoiceQuestion : undefined,
                  dueDate: effectiveItemForPart.dueDate,
                  dueTime: effectiveItemForPart.dueTime,
                  scheduledTime: effectiveItemForPart.scheduledTime,
                  submissionModificationMode: effectiveItemForPart.submissionModificationMode,
                };
              }
              batchOperationsToExecute.push({
                id: `op-${serviceCallId}-${operationIdCounter++}`,
                method: 'POST',
                path: path,
                body: operationBody,
                processedItem: effectiveItemForPart,
                retriesAttempted: 0
              });
            });
            observer.next();
            observer.complete();
          });
        });
        preparationObservables.push(prepObservableForItem);
      }
    }

    return (preparationObservables.length > 0 ? forkJoin(preparationObservables) : of(null as any)).pipe(
      switchMap(() => {
        console.log(`[ClassroomService][${serviceCallId}] Prep complete. Batch ops: ${batchOperationsToExecute.length}. Tracked items: ${allItemsToTrackForBatchResult.length}.`);
        if (batchOperationsToExecute.length > 0) {
          // Pass the fetched authToken to UtilitiesService
          return this.utils.executeBatchOperations<ProcessedCourseWork, CourseWork | CourseWorkMaterial>(
            batchOperationsToExecute,
            this.defaultRetryConfig,
            this.utils.GOOGLE_CLASSROOM_BATCH_ENDPOINT_URL,
            this.utils.GOOGLE_CLASSROOM_MAX_OPERATIONS_PER_BATCH,
            this.classroomBatchResponseParser
          ).pipe(
            map(() => allItemsToTrackForBatchResult)
          );
        } else {
          return of(allItemsToTrackForBatchResult);
        }
      }),
      tap((finalItems: ProcessedCourseWork[]) => {
        const successes = finalItems.filter(r => !r.processingError).length;
        const failures = finalItems.length - successes;
        console.log(`[ClassroomService][${serviceCallId}] assignContentToClassrooms FINALIZED. Total items: ${finalItems.length}, Successes: ${successes}, Failures: ${failures}.`);
      }),
      catchError((err: any) => {
        console.error(`[ClassroomService][${serviceCallId}] Critical error in assignContentToClassrooms pipeline:`, err);
        allItemsToTrackForBatchResult.forEach(item => {
          if (!item.processingError) {
            item.processingError = {
              message: `Critical batch pipeline error: ${err.message || String(err)}`,
              stage: 'Batch Pipeline Root (CS)'
            };
          }
        });
        return of(allItemsToTrackForBatchResult);
      })
    );
  }

  /**
   * Gets the ID of a topic by name. If it doesn't exist, creates it.
   * Token is fetched internally.
   */
  private getOrCreateTopicId(courseId: string, topicName?: string): Observable<string | undefined> {
    const normalizedTopicName = topicName?.trim();
    if (!normalizedTopicName) return of(undefined);

    const context = `getOrCreateTopicId (Course: ${courseId}, Topic: "${normalizedTopicName}")`;
    // listAllTopics and createTopic will fetch their own tokens

    return this.listAllTopics(courseId).pipe( // No authToken passed
      switchMap(allTopics => {
        const found = allTopics.find(topic => topic.name?.toLowerCase() === normalizedTopicName.toLowerCase());
        if (found?.topicId) return of(found.topicId);
        return this.createTopic(courseId, normalizedTopicName).pipe(map(newTopic => newTopic.topicId)); // No authToken passed
      }),
      catchError(err => {
        return throwError(() => this.handleError(err, context));
      })
    );
  }

  /**
   * Lists all topics in a given course. Token is fetched internally.
   */
  private listAllTopics(courseId: string): Observable<Topic[]> {
    const context = `listAllTopics (Course: ${courseId})`;
    // fetchTopicPage will fetch its own token
    return this.fetchTopicPageInternal(courseId, undefined, context).pipe(
      expand(response => response.nextPageToken ? this.fetchTopicPageInternal(courseId, response.nextPageToken, `${context} (paginated)`) : EMPTY),
      map(response => response.topic || []),
      reduce((acc, topics) => acc.concat(topics), [] as Topic[]),
      catchError(err => this.handleError(err, `${context} (Accumulation)`))
    );
  }

  /**
   * Internal version of fetchTopicPage that gets its own token.
   */
  private fetchTopicPageInternal(courseId: string, pageToken?: string, context: string = 'fetchTopicPageInternal'): Observable<ListTopicsResponse> {
    const authToken = this.auth.getGoogleAccessToken();
    if (!authToken) {
      const errorMsg = `[ClassroomService] ${context}: Auth token is missing.`;
      console.error(errorMsg);
      return throwError(() => new Error(errorMsg));
    }
    const headers = this.createAuthHeaders(authToken);
    let params = new HttpParams().set('pageSize', this.pageSize);
    if (pageToken) params = params.set('pageToken', pageToken);
    const url = this.topicsApiUrl(courseId);
    const operationDescription = `${context} (Course: ${courseId}, Page: ${pageToken ?? 'initial'})`;
    const request$ = this.http.get<ListTopicsResponse>(url, {headers, params});
    return this.utils.retryRequest(request$, this.defaultRetryConfig, operationDescription).pipe(
      catchError(err => this.handleError(err, operationDescription))
    );
  }

  /**
   * Creates a new topic in a course. Token is fetched internally.
   */
  private createTopic(courseId: string, topicName: string): Observable<Topic> {
    const authToken = this.auth.getGoogleAccessToken();
    if (!authToken) {
      const errorMsg = `[ClassroomService] createTopic (Course: ${courseId}, Topic: "${topicName}"): Auth token is missing.`;
      console.error(errorMsg);
      return throwError(() => new Error(errorMsg));
    }
    const headers = this.createAuthHeaders(authToken);
    const url = this.topicsApiUrl(courseId);
    const body = {name: topicName};
    const context = `createTopic (Course: ${courseId}, Topic: "${topicName}")`;
    const request$ = this.http.post<Topic>(url, body, {headers});
    return this.utils.retryRequest(request$, this.defaultRetryConfig, context).pipe(
      catchError(err => this.handleError(err, context))
    );
  }

  /**
   * Creates standard HTTP headers for authenticated API calls.
   * This method still accepts authToken as a parameter.
   */
  private createAuthHeaders(authToken: string): HttpHeaders {
    return new HttpHeaders({
      'Authorization': `Bearer ${authToken}`,
      'Accept': 'application/json',
      'Content-Type': 'application/json'
    });
  }

  /**
   * Splits an array into chunks of a specified size.
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
   * Generates a unique key for a Material object.
   */
  private getMaterialKey(material: Material): string | null {
    if (material.driveFile?.driveFile?.id) return `drive-${material.driveFile.driveFile.id}`;
    if (material.youtubeVideo?.id) return `youtube-${material.youtubeVideo.id}`;
    if (material.link?.url) return `link-${this.utils.tryDecodeURIComponent(material.link.url).toLowerCase()}`;
    if (material.form?.formUrl) return `form-${this.utils.tryDecodeURIComponent(material.form.formUrl).toLowerCase()}`;
    return null;
  }

  /**
   * Deduplicates an array of Material objects.
   */
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

  /**
   * Centralized error handler.
   */
  private handleError(error: HttpErrorResponse | Error, context: string = 'Unknown Operation'): Observable<never> {
    let userMessage = `Failed during ${context}; please try again or check console.`;
    let detailedMessage = `Context: ${context} - Unknown error.`;
    let statusCode: number | undefined;
    let errorDetailsForPropagation: any = error;

    if (error instanceof HttpErrorResponse) {
      statusCode = error.status;
      detailedMessage = `Context: ${context} - ${this.utils.formatHttpError(error)}`;
      userMessage = `Server error (Code: ${error.status}) in ${context}.`;
      errorDetailsForPropagation = error.error || {message: error.message, status: error.status};
      const googleApiError = error.error?.error?.message;
      if (googleApiError) userMessage = `Google API Error in ${context}: ${googleApiError} (${error.status})`;
    } else if (error instanceof Error) {
      detailedMessage = `Context: ${context} - Client error: ${error.message}`;
      userMessage = `Client error in ${context}: ${error.message}`;
      errorDetailsForPropagation = {message: error.message, name: error.name};
    }

    console.error(`[ClassroomService] handleError: ${detailedMessage}`, error);
    const finalError = new Error(userMessage);
    (finalError as any).status = statusCode;
    (finalError as any).details = errorDetailsForPropagation;
    (finalError as any).stage = context;
    return throwError(() => finalError);
  }
}
