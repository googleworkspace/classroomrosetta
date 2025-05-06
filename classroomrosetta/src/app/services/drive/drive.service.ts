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
import { HttpClient, HttpHeaders, HttpParams, HttpErrorResponse } from '@angular/common/http';
import { Observable, of, throwError, from } from 'rxjs';
import { map, catchError, switchMap, shareReplay, tap } from 'rxjs/operators';
// Import RetryConfig from UtilitiesService
import { UtilitiesService, RetryConfig } from '../utilities/utilities.service'; // Adjust path accordingly

/**
 * Interface for Google Drive API File resource (simplified)
 */
interface DriveFile {
  id: string;
  name: string;
  mimeType: string;
  parents?: string[];
  appProperties?: {[key: string]: string};
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
  private utils = inject(UtilitiesService); // Inject UtilitiesService for hashing, helpers, and retryRequest

  // Use Drive endpoints from UtilitiesService if available, otherwise define here
  private readonly DRIVE_API_URL = this.utils.DRIVE_API_FILES_ENDPOINT || 'https://www.googleapis.com/drive/v3/files';

  // Root folder name constant
  private readonly ROOT_IMPORT_FOLDER_NAME = 'LMS Import';

  // Custom property key to store the source identifier HASH
  // TODO: Standardize this key across all services (e.g., 'imsccIdentifierHash')
  private readonly ITEM_ID_HASH_PROPERTY_KEY = 'itemIdHash'; // Renamed to reflect it stores a hash

  // Cache for the root import folder ID
  private rootImportFolderId$: Observable<string> | null = null;

  constructor() { }

  /**
   * Finds a folder by its associated ItemId HASH stored in appProperties.
   * Includes retry logic.
   *
   * @param hashedItemId The HASHED ItemId to search for.
   * @param accessToken A valid Google OAuth 2.0 access token.
   * @returns Observable emitting the folder ID if found, or null otherwise.
   */
  private findFolderByHashedItemId(hashedItemId: string, accessToken: string): Observable<string | null> {
    const headers = new HttpHeaders({
      'Authorization': `Bearer ${accessToken}`
    });
    const query = `appProperties has { key='${this.ITEM_ID_HASH_PROPERTY_KEY}' and value='${this.escapeQueryParam(hashedItemId)}' } and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    const params = new HttpParams()
      .set('q', query)
      .set('fields', 'files(id, name, appProperties)')
      .set('spaces', 'drive');

    console.log(`Searching for folder by ItemId HASH "${hashedItemId}"`);

    // Define the request
    const searchRequest$ = this.http.get<DriveFileList>(this.DRIVE_API_URL, { headers, params });

    // Wrap with retry logic
    return this.utils.retryRequest(
        searchRequest$,
        { maxRetries: 3, initialDelayMs: 1000 }, // Example config
        `Find Folder by Hash ${hashedItemId}`
    ).pipe(
      map(response => {
        if (response.files && response.files.length > 0) {
          const foundFolder = response.files[0];
          console.log(`Found folder "${foundFolder.name}" with ID: ${foundFolder.id} matching ItemId HASH "${hashedItemId}"`);
          return foundFolder.id;
        } else {
          console.log(`Folder with ItemId HASH "${hashedItemId}" not found.`);
          return null;
        }
      }),
      // Final catchError after retries
      catchError(error => {
        const formattedError = this.utils.formatHttpError(error);
        console.error(`Error finding folder by ItemId HASH "${hashedItemId}" (final after retries):`, formattedError);
        console.warn(`Search failed for ItemId HASH "${hashedItemId}", proceeding as if not found.`);
        return of(null); // Return null to allow creation logic
      })
    );
  }

  /**
   * Finds a folder by name within a specific parent folder.
   * Includes retry logic.
   *
   * @param folderName The name of the folder to find.
   * @param parentFolderId The ID of the parent folder ('root' for My Drive).
   * @param accessToken A valid Google OAuth 2.0 access token.
   * @returns Observable emitting the folder ID if found, or null otherwise.
   */
  private findFolderByName(folderName: string, parentFolderId: string, accessToken: string): Observable<string | null> {
    const headers = new HttpHeaders({
      'Authorization': `Bearer ${accessToken}`
    });
    const query = `name='${this.escapeQueryParam(folderName)}' and '${parentFolderId}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    const params = new HttpParams()
      .set('q', query)
      .set('fields', 'files(id, name)')
      .set('spaces', 'drive');

    console.log(`Searching for folder "${folderName}" in parent "${parentFolderId}"`);

    // Define the request
    const searchRequest$ = this.http.get<DriveFileList>(this.DRIVE_API_URL, { headers, params });

    // Wrap with retry logic
    return this.utils.retryRequest(
        searchRequest$,
        { maxRetries: 3, initialDelayMs: 1000 }, // Example config
        `Find Folder by Name "${folderName}"`
    ).pipe(
      map(response => {
        if (response.files && response.files.length > 0) {
          console.log(`Found folder "${folderName}" with ID: ${response.files[0].id}`);
          return response.files[0].id;
        } else {
          console.log(`Folder "${folderName}" not found in parent "${parentFolderId}".`);
          return null;
        }
      }),
      // Final catchError after retries
      catchError(error => {
        const formattedError = this.utils.formatHttpError(error);
        console.error(`Error finding folder "${folderName}" in parent "${parentFolderId}" (final after retries):`, formattedError);
        console.warn(`Search failed for folder "${folderName}", proceeding as if not found.`);
        return of(null); // Return null to allow creation logic
      })
    );
  }

  /**
   * Creates a new folder in Google Drive, optionally adding an ItemId HASH to appProperties.
   * Includes retry logic.
   *
   * @param folderName The name for the new folder.
   * @param parentFolderId The ID of the parent folder where the new folder should be created.
   * @param accessToken A valid Google OAuth 2.0 access token.
   * @param hashedItemId Optional HASHED ItemId to store in the folder's appProperties.
   * @returns Observable emitting the ID of the newly created folder.
   */
  private createFolder(folderName: string, parentFolderId: string, accessToken: string, hashedItemId?: string): Observable<string> {
    const headers = new HttpHeaders({
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    });
    const metadata: Partial<DriveFile> & {mimeType: string, parents: string[]} = {
      name: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentFolderId]
    };
    if (hashedItemId) {
      metadata.appProperties = { [this.ITEM_ID_HASH_PROPERTY_KEY]: hashedItemId };
      console.log(`Creating folder "${folderName}" in parent "${parentFolderId}" with ItemId HASH "${hashedItemId}"`);
    } else {
      console.log(`Creating folder "${folderName}" in parent "${parentFolderId}" (no ItemId HASH)`);
    }
    const params = new HttpParams().set('fields', 'id, name, appProperties');

    // Define the request
    const createRequest$ = this.http.post<DriveFile>(this.DRIVE_API_URL, metadata, { headers, params });

    // Wrap with retry logic
    return this.utils.retryRequest(
        createRequest$,
        { maxRetries: 3, initialDelayMs: 1500 }, // Example config
        `Create Folder "${folderName}"`
    ).pipe(
      map(response => {
        console.log(`Created folder "${response.name}" with ID: ${response.id}` + (hashedItemId ? ` and ItemId HASH: ${response.appProperties?.[this.ITEM_ID_HASH_PROPERTY_KEY]}` : ''));
        return response.id;
      }),
      // Final catchError after retries
      catchError(error => {
        const formattedError = this.utils.formatHttpError(error);
        console.error(`Error creating folder "${folderName}" in parent "${parentFolderId}" (final after retries):`, formattedError);
        // Re-throw the error after logging, as folder creation failure might be critical
        return throwError(() => new Error(`Failed to create folder "${folderName}". ${formattedError}`));
      })
    );
  }

 /**
 * Finds or creates a folder, using HASHED ItemId if provided.
 * Includes retry logic inherited from called methods.
 *
 * @param folderName The name of the folder.
 * @param parentFolderId The ID of the parent folder.
 * @param accessToken A valid Google OAuth 2.0 access token.
 * @param itemId Optional ORIGINAL ItemId associated with this folder. Will be hashed if provided.
 * @returns Observable emitting the ID of the found or created folder.
 */
private findOrCreateFolder(folderName: string, parentFolderId: string, accessToken: string, itemId?: string): Observable<string> {
    // Note: Retry logic is applied within findFolderByHashedItemId, findFolderByName, and createFolder.
    // No additional retry wrapper needed here, but we need to handle errors from the inner calls.
    if (itemId) {
      console.log(`Attempting to find or create folder "${folderName}" using ItemId "${itemId}" (will be hashed)`);
      return from(this.utils.generateHash(itemId)).pipe(
        catchError(hashError => {
          console.error(`Error generating hash for itemId "${itemId}":`, hashError);
          const errorMessage = hashError instanceof Error ? hashError.message : String(hashError);
          return throwError(() => new Error(`Failed to generate identifier hash for ${itemId}. ${errorMessage}`));
        }),
        switchMap(hashedItemId => {
          console.log(`Generated hash for itemId "${itemId}": ${hashedItemId}`);
          return this.findFolderByHashedItemId(hashedItemId, accessToken).pipe( // Retries handled inside
            switchMap(folderIdFromHash => {
              if (folderIdFromHash) {
                console.log(`Using existing folder found by ItemId HASH "${hashedItemId}". ID: ${folderIdFromHash}`);
                return of(folderIdFromHash);
              } else {
                console.log(`Folder with ItemId HASH "${hashedItemId}" not found. Creating new folder "${folderName}" with this HASH.`);
                return this.createFolder(folderName, parentFolderId, accessToken, hashedItemId); // Retries handled inside
              }
            })
          );
        }),
        // Catch errors from hashing, finding, or creating
        catchError(err => {
            // Log the error from the find/create process for this specific folder
            console.error(`Failed to find or create folder "${folderName}" associated with itemId "${itemId}" (final after retries):`, err.message || err);
            // Propagate the error
            return throwError(() => err);
        })
      );
    } else {
      console.log(`Attempting to find or create folder "${folderName}" by name in parent "${parentFolderId}" (no ItemId provided).`);
      return this.findFolderByName(folderName, parentFolderId, accessToken).pipe( // Retries handled inside
        switchMap(folderIdFromName => {
          if (folderIdFromName) {
            console.log(`Using existing folder found by name "${folderName}". ID: ${folderIdFromName}`);
            return of(folderIdFromName);
          } else {
            console.log(`Folder with name "${folderName}" not found in parent "${parentFolderId}". Creating new folder.`);
            return this.createFolder(folderName, parentFolderId, accessToken); // Retries handled inside
          }
        }),
        // Catch errors from finding or creating by name
        catchError(err => {
            console.error(`Failed to find or create folder "${folderName}" by name (final after retries):`, err.message || err);
            return throwError(() => err);
        })
      );
    }
}


  /**
   * Gets the ID of the root "LMS Import" folder, creating it if necessary.
   * Uses caching to avoid repeated lookups.
   * (No itemId/hashing involved for the root import folder itself)
   *
   * @param accessToken A valid Google OAuth 2.0 access token.
   * @returns Observable emitting the ID of the "LMS Import" folder.
   */
  public getRootImportFolderId(accessToken: string): Observable<string> {
    if (!this.rootImportFolderId$) {
      // findOrCreateFolder internally calls methods with retry logic
      this.rootImportFolderId$ = this.findOrCreateFolder(this.ROOT_IMPORT_FOLDER_NAME, 'root', accessToken /* no itemId */).pipe(
        tap(id => console.log(`Root Import Folder ID (${this.ROOT_IMPORT_FOLDER_NAME}): ${id}`)),
        shareReplay(1) // Cache the result
      );
    }
    return this.rootImportFolderId$;
  }

  /**
   * Gets the ID of the Course Folder within the root import folder.
   * (No itemId/hashing involved for the course folder itself)
   *
   * @param courseName The name for the course folder.
   * @param accessToken A valid Google OAuth 2.0 access token.
   * @returns Observable emitting the ID of the course folder.
   */
  public getCourseFolderId(courseName: string, accessToken: string): Observable<string> {
    const sanitizedCourseName = this.sanitizeFolderName(courseName);
    // findOrCreateFolder internally calls methods with retry logic
    return this.getRootImportFolderId(accessToken).pipe(
      switchMap(rootImportFolderId =>
        this.findOrCreateFolder(sanitizedCourseName, rootImportFolderId, accessToken /* no itemId */)
      )
    );
  }

  /**
   * Gets the ID of the Topic Folder within a specific course folder.
   * (No itemId/hashing involved for the topic folder itself)
   *
   * @param topicName The name for the topic folder.
   * @param courseFolderId The ID of the parent course folder.
   * @param accessToken A valid Google OAuth 2.0 access token.
   * @returns Observable emitting the ID of the topic folder.
   */
  public getTopicFolderId(topicName: string, courseFolderId: string, accessToken: string): Observable<string> {
    const sanitizedTopicName = this.sanitizeFolderName(topicName);
    // findOrCreateFolder internally calls methods with retry logic
    return this.findOrCreateFolder(sanitizedTopicName, courseFolderId, accessToken /* no itemId */);
  }

  /**
   * Gets the ID of the Assignment Folder within a specific topic folder.
   * Uses HASHED ItemId for searching and creation.
   *
   * @param assignmentName The name for the assignment folder.
   * @param topicFolderId The ID of the parent topic folder.
   * @param accessToken A valid Google OAuth 2.0 access token.
   * @param itemId The specific ORIGINAL ItemId associated with this assignment (will be hashed).
   * @returns Observable emitting the ID of the assignment folder.
   */
  public getAssignmentFolderId(assignmentName: string, topicFolderId: string, accessToken: string, itemId: string): Observable<string> {
    const sanitizedAssignmentName = this.sanitizeFolderName(assignmentName);
    // Pass the ORIGINAL itemId here; findOrCreateFolder will handle hashing and retries internally
    return this.findOrCreateFolder(sanitizedAssignmentName, topicFolderId, accessToken, itemId);
  }

  /**
   * Orchestrates the creation/retrieval of the full folder path for an assignment item.
   * Root -> LMS Import -> Course -> Topic -> Assignment (identified by HASHED ItemId)
   * Includes retry logic inherited from called methods.
   *
   * @param courseName Name of the course.
   * @param topicName Name of the topic.
   * @param assignmentName Name of the assignment.
   * @param itemId The ORIGINAL ItemId, used to identify/create the assignment folder via its hash.
   * @param accessToken A valid Google OAuth 2.0 access token.
   * @returns Observable emitting the ID of the final assignment folder.
   */
  public ensureAssignmentFolderStructure(
    courseName: string,
    topicName: string,
    assignmentName: string,
    itemId: string, // Now mandatory and used for the assignment folder (will be hashed internally)
    accessToken: string
  ): Observable<string> {
    // Input validation
    if (!courseName || !topicName || !assignmentName) {
      return throwError(() => new Error('Course, Topic, and Assignment names are required.'));
    }
    if (!itemId) {
        return throwError(() => new Error('An ItemId is required to ensure the assignment folder structure.'));
    }

    // Chain the folder creation/retrieval calls. Retry logic is handled within the get*FolderId methods via findOrCreateFolder.
    return this.getCourseFolderId(courseName, accessToken).pipe(
      switchMap(courseFolderId => this.getTopicFolderId(topicName, courseFolderId, accessToken)),
      switchMap(topicFolderId => this.getAssignmentFolderId(assignmentName, topicFolderId, accessToken, itemId)),
      tap(assignmentFolderId => console.log(`Ensured folder structure complete. Assignment Folder ID: ${assignmentFolderId} (associated with ItemId: ${itemId}, stored as hash)`)),
      // Add a final catchError here to handle failures in the overall chain
      catchError(err => {
          console.error(`Failed to ensure folder structure for assignment "${assignmentName}" (itemId: ${itemId}):`, err.message || err);
          return throwError(() => new Error(`Could not ensure folder structure for assignment "${assignmentName}". ${err.message || err}`));
      })
    );
  }


  /**
   * Helper function to sanitize folder names (basic example).
   * (No changes needed)
   */
  private sanitizeFolderName(name: string): string {
    if (!name) return 'Untitled';
    return name.replace(/[\\/]/g, '-').trim() || 'Untitled';
  }

   /**
   * Helper function to escape single quotes and backslashes for Drive API query parameters.
   * (No changes needed)
   */
  private escapeQueryParam(value: string): string {
    return value.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
   }

} // End of Service
