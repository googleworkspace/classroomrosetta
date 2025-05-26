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

import { Injectable, inject } from '@angular/core';
import {HttpClient, HttpHeaders, HttpParams, HttpErrorResponse} from '@angular/common/http';
import { Observable, of, throwError, from } from 'rxjs';
import { map, catchError, switchMap, shareReplay, tap } from 'rxjs/operators';
import {UtilitiesService} from '../utilities/utilities.service';
import {AuthService} from '../auth/auth.service'; // Import AuthService

/**
 * Interface for Google Drive API File resource (simplified)
 */
interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
  parents?: string[];
  appProperties?: {[key: string]: string};
  // Add other fields like webViewLink, thumbnailLink if needed by this service directly
  webViewLink?: string;
  thumbnailLink?: string;
}

/**
 * Interface for Drive API file list response (simplified)
 */
interface DriveFileList {
  files: DriveFile[];
}

@Injectable({
  providedIn: 'root'
})
export class DriveFolderService {

  // Inject dependencies
  private http = inject(HttpClient);
  private utils = inject(UtilitiesService);
  private auth = inject(AuthService); // Inject AuthService

  // Use Drive endpoints from UtilitiesService
  private readonly DRIVE_API_URL = this.utils.DRIVE_API_FILES_ENDPOINT;

  // Root folder name constant
  private readonly ROOT_IMPORT_FOLDER_NAME = 'LMS Import';

  // Custom property key to store the source identifier HASH
  private readonly ITEM_ID_HASH_PROPERTY_KEY = 'itemIdHash';

  // Cache for findOrCreateFolder operations
  private findOrCreateFolderCache = new Map<string, Observable<string>>();

  // Cache for the root import folder ID observable specifically
  private rootImportFolderIdObservable$: Observable<string> | null = null;

  constructor() { }

  /**
   * Creates standard HTTP headers for authenticated Drive API calls.
   * Fetches token internally.
   * @param contentType Optional content type, defaults to 'application/json'.
   * @returns HttpHeaders object or null if token is missing.
   */
  private createDriveApiHeaders(contentType: string = 'application/json'): HttpHeaders | null {
    const accessToken = this.auth.getGoogleAccessToken();
    if (!accessToken) {
      console.error('[DriveFolderService] Cannot create Drive API headers: Access token is missing.');
      return null;
    }
    let headers = new HttpHeaders({
      'Authorization': `Bearer ${accessToken}`
    });
    if (contentType) {
      headers = headers.set('Content-Type', contentType);
    }
    return headers;
  }


  /**
   * Finds a folder by its associated ItemId HASH stored in appProperties.
   * Token is fetched internally.
   */
  private findFolderByHashedItemId(hashedItemId: string): Observable<string | null> {
    const headers = this.createDriveApiHeaders();
    if (!headers) {
      return throwError(() => new Error('[DriveFolderService] Authentication token missing for findFolderByHashedItemId.'));
    }

    const query = `appProperties has { key='${this.ITEM_ID_HASH_PROPERTY_KEY}' and value='${this.escapeQueryParam(hashedItemId)}' } and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    const params = new HttpParams()
      .set('q', query)
      .set('fields', 'files(id, name, appProperties)')
      .set('spaces', 'drive');

    // console.log(`Searching for folder by ItemId HASH "${hashedItemId}"`);

    const searchRequest$ = this.http.get<DriveFileList>(this.DRIVE_API_URL, { headers, params });

    return this.utils.retryRequest(
        searchRequest$,
      {maxRetries: 3, initialDelayMs: 1000},
        `Find Folder by Hash ${hashedItemId}`
    ).pipe(
      map(response => {
        if (response.files && response.files.length > 0) {
          const foundFolder = response.files[0];
          console.log(`Found folder "${foundFolder.name}" with ID: ${foundFolder.id} matching ItemId HASH "${hashedItemId}"`);
          return foundFolder.id;
        } else {
          // console.log(`Folder with ItemId HASH "${hashedItemId}" not found.`);
          return null;
        }
      }),
      catchError(error => {
        const formattedError = this.utils.formatHttpError(error);
        console.error(`Error finding folder by ItemId HASH "${hashedItemId}" (final after retries):`, formattedError);
        console.warn(`Search failed for ItemId HASH "${hashedItemId}", proceeding as if not found.`);
        return of(null);
      })
    );
  }

  /**
   * Finds a folder by name within a specific parent folder.
   * Token is fetched internally.
   */
  private findFolderByName(folderName: string, parentFolderId: string): Observable<string | null> {
    const headers = this.createDriveApiHeaders();
    if (!headers) {
      return throwError(() => new Error('[DriveFolderService] Authentication token missing for findFolderByName.'));
    }

    const query = `name='${this.escapeQueryParam(folderName)}' and '${parentFolderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    const params = new HttpParams()
      .set('q', query)
      .set('fields', 'files(id, name)')
      .set('spaces', 'drive');

    // console.log(`Searching for folder "${folderName}" in parent "${parentFolderId}"`);

    const searchRequest$ = this.http.get<DriveFileList>(this.DRIVE_API_URL, { headers, params });

    return this.utils.retryRequest(
        searchRequest$,
        { maxRetries: 3, initialDelayMs: 1000 },
        `Find Folder by Name "${folderName}"`
    ).pipe(
      map(response => {
        if (response.files && response.files.length > 0) {
          console.log(`Found folder "${folderName}" with ID: ${response.files[0].id}`);
          return response.files[0].id;
        } else {
          // console.log(`Folder "${folderName}" not found in parent "${parentFolderId}".`);
          return null;
        }
      }),
      catchError(error => {
        const formattedError = this.utils.formatHttpError(error);
        console.error(`Error finding folder "${folderName}" in parent "${parentFolderId}" (final after retries):`, formattedError);
        console.warn(`Search failed for folder "${folderName}", proceeding as if not found.`);
        return of(null);
      })
    );
  }

  /**
   * Creates a new folder in Google Drive.
   * Token is fetched internally.
   */
  private createFolder(folderName: string, parentFolderId: string, hashedItemId?: string): Observable<string> {
    const headers = this.createDriveApiHeaders('application/json'); // Explicit content type for POST
    if (!headers) {
      return throwError(() => new Error('[DriveFolderService] Authentication token missing for createFolder.'));
    }

    const metadata: Partial<DriveFile> & {mimeType: string, parents: string[]} = {
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentFolderId]
    };
    if (hashedItemId) {
      metadata.appProperties = { [this.ITEM_ID_HASH_PROPERTY_KEY]: hashedItemId };
    }
    const params = new HttpParams().set('fields', 'id, name, appProperties');

    const createRequest$ = this.http.post<DriveFile>(this.DRIVE_API_URL, metadata, { headers, params });

    return this.utils.retryRequest(
        createRequest$,
        { maxRetries: 3, initialDelayMs: 1500 },
        `Create Folder "${folderName}"`
    ).pipe(
      map(response => {
        console.log(`Created folder "${response.name}" with ID: ${response.id}` + (hashedItemId ? ` and ItemId HASH: ${response.appProperties?.[this.ITEM_ID_HASH_PROPERTY_KEY]}` : ''));
        return response.id;
      }),
      catchError(error => {
        const formattedError = this.utils.formatHttpError(error);
        console.error(`Error creating folder "${folderName}" in parent "${parentFolderId}" (final after retries):`, formattedError);
        return throwError(() => new Error(`Failed to create folder "${folderName}". ${formattedError}`));
      })
    );
  }

 /**
 * Finds or creates a folder, using HASHED ItemId if provided.
 * Implements caching for the find/create operation. Token is fetched by underlying methods.
 */
  public findOrCreateFolder(folderName: string, parentFolderId: string, itemId?: string): Observable<string> {
    // Token will be fetched by findFolderByHashedItemId, findFolderByName, or createFolder.
    let cacheKeyPrefix = itemId ? 'itemHash' : 'name';

    let operation$: Observable<string>;

    if (itemId) {
      operation$ = from(this.utils.generateHash(itemId)).pipe(
        tap(hashedItemId => console.log(`Generated hash for itemId "${itemId}" for folder "${folderName}": ${hashedItemId}`)),
        switchMap(hashedItemId => {
          const itemSpecificCacheKey = `${cacheKeyPrefix}_${parentFolderId}_${hashedItemId}`;
          if (this.findOrCreateFolderCache.has(itemSpecificCacheKey)) {
            return this.findOrCreateFolderCache.get(itemSpecificCacheKey)!;
          }
          const newOp$ = this.findFolderByHashedItemId(hashedItemId).pipe( // No accessToken
            switchMap(folderIdFromHash => {
              if (folderIdFromHash) {
                return of(folderIdFromHash);
              } else {
                return this.createFolder(folderName, parentFolderId, hashedItemId); // No accessToken
              }
            }),
            shareReplay(1),
            catchError(err => {
              this.findOrCreateFolderCache.delete(itemSpecificCacheKey);
              console.error(`Error in findOrCreateFolder (hashed path) for key ${itemSpecificCacheKey}, folder "${folderName}":`, err.message || err);
              return throwError(() => err);
            })
          );
          this.findOrCreateFolderCache.set(itemSpecificCacheKey, newOp$);
          return newOp$;
        }),
        catchError(hashError => {
          console.error(`Error generating hash for itemId "${itemId}" (folder "${folderName}"):`, hashError);
          const errorMessage = hashError instanceof Error ? hashError.message : String(hashError);
          return throwError(() => new Error(`Failed to generate identifier hash for ${itemId} (folder "${folderName}"). ${errorMessage}`));
        })
      );
    } else {
      const nameSpecificCacheKey = `${cacheKeyPrefix}_${parentFolderId}_${this.sanitizeFolderName(folderName)}`;
      if (this.findOrCreateFolderCache.has(nameSpecificCacheKey)) {
        return this.findOrCreateFolderCache.get(nameSpecificCacheKey)!;
      }
      operation$ = this.findFolderByName(folderName, parentFolderId).pipe( // No accessToken
        switchMap(folderIdFromName => {
          if (folderIdFromName) {
            return of(folderIdFromName);
          } else {
            return this.createFolder(folderName, parentFolderId); // No accessToken
          }
        }),
        shareReplay(1),
        catchError(err => {
          this.findOrCreateFolderCache.delete(nameSpecificCacheKey);
          console.error(`Error in findOrCreateFolder (name path) for key ${nameSpecificCacheKey}, folder "${folderName}":`, err.message || err);
          return throwError(() => err);
        })
      );
      this.findOrCreateFolderCache.set(nameSpecificCacheKey, operation$);
    }
    return operation$;
}


  /**
   * Gets the ID of the root "LMS Import" folder, creating it if necessary.
   * Token is fetched by underlying findOrCreateFolder.
   */
  public getRootImportFolderId(): Observable<string> {
    if (!this.rootImportFolderIdObservable$) {
      this.rootImportFolderIdObservable$ = this.findOrCreateFolder(
        this.ROOT_IMPORT_FOLDER_NAME,
        'root' // No itemId
      ).pipe(
        tap(id => console.log(`Root Import Folder ID ("${this.ROOT_IMPORT_FOLDER_NAME}") obtained: ${id}`)),
        shareReplay(1)
      );
    }
    return this.rootImportFolderIdObservable$;
  }

  /**
   * Gets the ID of the Course Folder within the root import folder.
   * Token is fetched by underlying methods.
   */
  public getCourseFolderId(courseName: string): Observable<string> {
    const sanitizedCourseName = this.sanitizeFolderName(courseName);
    return this.getRootImportFolderId().pipe( // No accessToken
      switchMap(rootImportFolderId =>
        this.findOrCreateFolder(sanitizedCourseName, rootImportFolderId /* no itemId */) // No accessToken
      ),
      tap(id => console.log(`Course Folder ID ("${sanitizedCourseName}"): ${id}`))
    );
  }

  /**
   * Gets the ID of the Topic Folder within a specific course folder.
   * Token is fetched by underlying findOrCreateFolder.
   */
  public getTopicFolderId(topicName: string, courseFolderId: string): Observable<string> {
    const sanitizedTopicName = this.sanitizeFolderName(topicName);
    return this.findOrCreateFolder(sanitizedTopicName, courseFolderId /* no itemId */).pipe( // No accessToken
        tap(id => console.log(`Topic Folder ID ("${sanitizedTopicName}") in course ${courseFolderId}: ${id}`))
    );
  }

  /**
   * Gets the ID of the Assignment Folder within a specific topic folder.
   * Token is fetched by underlying findOrCreateFolder.
   */
  public getAssignmentFolderId(assignmentName: string, topicFolderId: string, itemId: string): Observable<string> {
    const sanitizedAssignmentName = this.sanitizeFolderName(assignmentName);
    return this.findOrCreateFolder(sanitizedAssignmentName, topicFolderId, itemId).pipe( // No accessToken
        tap(id => console.log(`Assignment Folder ID ("${sanitizedAssignmentName}", itemId: ${itemId}) in topic ${topicFolderId}: ${id}`))
    );
  }

  /**
   * Orchestrates the creation/retrieval of the full folder path for an assignment item.
   * Token is fetched by underlying methods.
   */
  public ensureAssignmentFolderStructure(
    courseName: string,
    topicName: string,
    assignmentName: string,
    itemId: string
  ): Observable<string> {
    if (!courseName || !topicName || !assignmentName) {
      return throwError(() => new Error('Course, Topic, and Assignment names are required for folder structure.'));
    }
    if (!itemId) {
      return throwError(() => new Error('An ItemId is required to ensure the assignment folder structure (for unique identification).'));
    }
    console.log(`Ensuring folder structure for: Course="${courseName}", Topic="${topicName}", Assignment="${assignmentName}" (ItemId: ${itemId})`);

    return this.getCourseFolderId(courseName).pipe( // No accessToken
      switchMap(courseFolderId => this.getTopicFolderId(topicName, courseFolderId)), // No accessToken
      switchMap(topicFolderId => this.getAssignmentFolderId(assignmentName, topicFolderId, itemId)), // No accessToken
      tap(assignmentFolderId => console.log(`Successfully ensured folder structure. Final Assignment Folder ID: ${assignmentFolderId} (for ItemId: ${itemId})`)),
      catchError(err => {
        const baseMessage = `Failed to ensure folder structure for assignment "${assignmentName}" (itemId: ${itemId})`;
        console.error(`${baseMessage}:`, err.message || err);
        const detailedError = err instanceof Error ? err.message : String(err);
        return throwError(() => new Error(`${baseMessage}. Error: ${detailedError}`));
      })
    );
  }

  /**
   * Helper function to sanitize folder names.
   */
  private sanitizeFolderName(name: string): string {
    if (!name) return 'Untitled';
    return name.replace(/[\\/]/g, '_').replace(/:/g, '-').trim() || 'Untitled';
  }

   /**
   * Helper function to escape single quotes and backslashes for Drive API query parameters.
   */
  private escapeQueryParam(value: string): string {
    if (!value) return '';
    return value.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  }
}
