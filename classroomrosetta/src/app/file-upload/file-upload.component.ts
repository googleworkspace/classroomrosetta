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

import {Component, inject, ViewChild, ChangeDetectorRef} from '@angular/core';
import JSZip from 'jszip';
import {Observable, from, of, throwError, forkJoin} from 'rxjs';
import {map, switchMap, concatMap, catchError, tap, finalize, toArray, shareReplay} from 'rxjs/operators';
import {CommonModule} from '@angular/common';
import {MatButtonModule} from '@angular/material/button';
import {MatInputModule} from '@angular/material/input';
import {MatFormFieldModule} from '@angular/material/form-field';
import {MatIconModule} from '@angular/material/icon';
import {MatProgressSpinnerModule} from '@angular/material/progress-spinner';
import {ClassroomService} from '../services/classroom/classroom.service';
import {ConverterService} from '../services/converter/converter.service';
import {DriveFolderService} from '../services/drive/drive.service';
import {HtmlToDocsService} from '../services/html-to-docs/html-to-docs.service';
import {FileUploadService} from '../services/file-upload/file-upload.service';
import {QtiToFormsService} from '../services/qti-to-forms/qti-to-forms.service';
import {UtilitiesService} from '../services/utilities/utilities.service';
import {AuthService} from '../services/auth/auth.service';
import {
  ProcessedCourseWork,
  SubmissionData,
  Material,
  ImsccFile,
  DriveFile,
} from '../interfaces/classroom-interface';
import {CourseworkDisplayComponent} from '../coursework-display/coursework-display.component';
import {decode} from 'html-entities';

// Local helper function for normalizing spaces
function normalizeSpacesString(str: string): string {
  if (!str) return '';
  return str.replace(/\s+/g, ' ').trim();
}

interface UnzippedFile {
  name: string;
  data: string | ArrayBuffer;
  mimeType: string;
}

export interface StagedProcessingError {
  message: string;
  stage: string;
  details?: any;
}

// Represents the result of processing an item's content (Drive uploads, QTI, HTML conversions)
// before it's sent to ClassroomService.
export interface ProcessingResult {
  itemId: string | undefined; // IMSCC Item ID
  assignmentName: string;     // Original assignment name
  topicName: string;          // Topic name derived from IMSCC structure (used for Drive folder org)
  assignmentFolderId: string; // Drive Folder ID where content was placed
  createdDoc?: DriveFile;     // Details if a Google Doc was created from HTML
  createdForm?: Material;     // Details if a Google Form was created from QTI (as a Material object)
  uploadedFiles?: DriveFile[];// List of other files uploaded to Drive
  error: StagedProcessingError | null; // Error if this stage failed for the item
  finalHtmlDescription?: string;      // The final HTML description (e.g., after local links are converted to Drive links)
  finalPlainTextDescription?: string; // The plain text version of the final HTML, for Classroom API
}

@Component({
  selector: 'app-file-upload',
  standalone: true,
  imports: [
    CommonModule, MatButtonModule, CourseworkDisplayComponent, MatInputModule,
    MatFormFieldModule, MatIconModule, MatProgressSpinnerModule
  ],
  templateUrl: './file-upload.component.html',
  styleUrl: './file-upload.component.scss'
})
export class FileUploadComponent {
  selectedFile: File | null = null;
  unzippedFiles: UnzippedFile[] = [];
  @ViewChild('fileInput') fileInput: any = null;
  assignments: ProcessedCourseWork[] = []; // Holds all ProcessedCourseWork items derived from the IMSCC package
  isProcessing: boolean = false;
  loadingMessage: string = '';
  errorMessage: string | null = null;
  successMessage: string | null = null;

  // Injecting services
  classroom = inject(ClassroomService);
  converter = inject(ConverterService);
  drive = inject(DriveFolderService);
  docs = inject(HtmlToDocsService);
  files = inject(FileUploadService);
  qti = inject(QtiToFormsService);
  auth = inject(AuthService);
  util = inject(UtilitiesService);
  private changeDetectorRef = inject(ChangeDetectorRef);

  onClickFileInputButton(): void {
    this.fileInput.nativeElement.click();
  }

  onChangeFileInput(): void {
    const filesFromInput: FileList | null = this.fileInput.nativeElement.files;
    if (filesFromInput && filesFromInput.length > 0) {
      this.selectedFile = filesFromInput[0];
    } else {
      this.selectedFile = null;
    }
    // Reset state
    this.assignments = [];
    this.unzippedFiles = [];
    this.isProcessing = false;
    this.errorMessage = null;
    this.successMessage = null;
    this.changeDetectorRef.markForCheck();
  }

  async onUpload(): Promise<void> {
    if (this.selectedFile && !this.isProcessing) {
      this.isProcessing = true;
      this.loadingMessage = 'Unzipping file...';
      this.errorMessage = null;
      this.successMessage = null;
      this.assignments = [];
      this.unzippedFiles = [];
      this.changeDetectorRef.markForCheck();
      try {
        await this.unzipAndConvert(this.selectedFile);
      } catch (error: any) {
        console.error('[Orchestrator] Error during file upload/unzip trigger:', error);
        this.errorMessage = `Error during setup: ${error?.message || String(error)}`;
        this.isProcessing = false;
        this.loadingMessage = '';
        this.changeDetectorRef.markForCheck();
      }
    } else if (this.isProcessing) {
      console.warn('[Orchestrator] Processing already in progress.');
    } else {
      console.warn('[Orchestrator] No file selected for upload.');
      this.errorMessage = 'Please select a file first.';
    }
  }

  async unzipAndConvert(file: File): Promise<void> {
    const accumulatedAssignments: ProcessedCourseWork[] = [];
    try {
      const zip = new JSZip();
      const loadedZip = await zip.loadAsync(file);
      this.loadingMessage = 'Reading package contents...';
      this.changeDetectorRef.markForCheck();
      const filePromises: Promise<UnzippedFile | null>[] = [];

      loadedZip.forEach((relativePath, zipEntry) => {
        if (!zipEntry.dir) {
          const promise = (async (): Promise<UnzippedFile | null> => {
            let data: string | ArrayBuffer;
            const mimeType = this.util.getMimeTypeFromExtension(relativePath);
            // Heuristic to determine if content is text-based or binary
            if (mimeType.startsWith('text/') ||
              ['application/xml', 'application/json', 'application/javascript', 'image/svg+xml'].includes(mimeType) ||
              /\.(html|htm|css|js|xml|qti|txt|md)$/i.test(relativePath)) {
              data = await zipEntry.async('string');
            } else if (mimeType.startsWith('image/')) { // Convert images to base64 data URIs
              const base64Data = await zipEntry.async('base64');
              data = `data:${mimeType};base64,${base64Data}`;
            } else { // Treat others as binary
              data = await zipEntry.async('arraybuffer');
            }
            return {name: relativePath, data: data, mimeType: mimeType};
          })();
          filePromises.push(promise);
        }
      });

      const resolvedFiles = await Promise.all(filePromises);
      this.unzippedFiles = resolvedFiles.filter(f => f !== null) as UnzippedFile[];
      console.log(`[Orchestrator] Unzipped files prepared: ${this.unzippedFiles.length} files.`);

      const imsccFilesForConverter: ImsccFile[] = this.unzippedFiles.map(uf => ({
        name: uf.name, data: uf.data, mimeType: uf.mimeType
      }));

      this.loadingMessage = 'Converting course structure...';
      this.changeDetectorRef.markForCheck();

      // ConverterService.convertImscc emits ProcessedCourseWork items
      this.converter.convertImscc(imsccFilesForConverter)
        .pipe(
          finalize(() => {
            console.log('[Orchestrator] IMSCC conversion stream finalized.');
            if (!this.isProcessing && !this.errorMessage) { // Check if an error didn't already stop processing
              this.successMessage = `Conversion complete. Found ${accumulatedAssignments.length} items. Ready for submission.`;
            } else if (!this.errorMessage && this.isProcessing) { // Still processing, but conversion itself is done
              // This might occur if `isProcessing` is managed by a broader scope, but typically finalize means this part is done.
              // If an error happened in `subscribe.error`, `isProcessing` should be false.
            }
            this.isProcessing = false; // Ensure processing is marked false if not already by an error.
            this.loadingMessage = '';
            this.assignments = [...accumulatedAssignments]; // Store all converted assignments
            console.log('[Orchestrator] Final assignments count after conversion:', this.assignments.length);
            this.changeDetectorRef.markForCheck();
          })
        )
        .subscribe({
          next: (assignment) => {
            accumulatedAssignments.push(assignment);
          },
          error: (error) => {
            console.error('[Orchestrator] Error during IMSCC conversion stream:', error);
            this.errorMessage = `Conversion Error: ${error?.message || String(error)}`;
            this.isProcessing = false; // Stop processing on error
            this.loadingMessage = '';
            this.changeDetectorRef.markForCheck();
          }
        });

    } catch (error: any) {
      console.error('[Orchestrator] Error during unzipping or IMSCC conversion setup:', error);
      this.errorMessage = `Unzip/Read Error: ${error?.message || String(error)}`;
      this.assignments = [];
      this.isProcessing = false;
      this.loadingMessage = '';
      this.changeDetectorRef.markForCheck();
    }
  }

  // Main orchestration method for processing selected assignments
  process(selectedContent: SubmissionData): void {
    if (this.isProcessing) {
      console.warn("[Orchestrator] Processing already in progress.");
      this.errorMessage = "Processing is already underway. Please wait.";
      return;
    }
    if (!selectedContent.assignmentIds || selectedContent.assignmentIds.length === 0) {
      this.errorMessage = "Please select at least one assignment to submit.";
      return;
    }
    if (!selectedContent.classroomIds || selectedContent.classroomIds.length === 0) {
      this.errorMessage = "Please select at least one classroom to submit to.";
      return;
    }

    console.log('[Orchestrator] Starting process with selected content:', selectedContent);
    const token = this.auth.getGoogleAccessToken();
    if (!token) {
      this.errorMessage = "Processing aborted: No Google access token. Please log in.";
      console.error(this.errorMessage);
      return;
    }

    // Prepare assignments for the current processing run by filtering selected items
    // and clearing any pre-existing processing errors from previous runs.
    const assignmentsToProcessIds = new Set(selectedContent.assignmentIds);
    const assignmentsReadyForDriveProcessing = this.assignments
      .filter(assignment => assignment.associatedWithDeveloper?.id && assignmentsToProcessIds.has(assignment.associatedWithDeveloper.id))
      .map(assignment => ({...assignment, processingError: undefined})); // Clear previous errors for this run

    if (assignmentsReadyForDriveProcessing.length === 0) {
      this.errorMessage = "No selected assignments found to process.";
      console.warn(`[Orchestrator] ${this.errorMessage}`);
      return;
    }

    // Update the main `this.assignments` list to reflect cleared errors for UI consistency.
    this.assignments = this.assignments.map(a => {
      if (a.associatedWithDeveloper?.id && assignmentsToProcessIds.has(a.associatedWithDeveloper.id)) {
        // Find the version with the cleared error, or fallback to original 'a' if somehow not found (should not happen)
        return assignmentsReadyForDriveProcessing.find(arp => a.associatedWithDeveloper && arp.associatedWithDeveloper?.id === a.associatedWithDeveloper.id) || a;
      }
      return a;
    });

    const courseName = this.converter.coursename || 'Untitled Course';
    console.log(`[Orchestrator] Processing ${assignmentsReadyForDriveProcessing.length} selected assignments for course: "${courseName}"`);
    console.log(`[Orchestrator] Target Classroom IDs: ${selectedContent.classroomIds.join(', ')}`);

    this.isProcessing = true;
    this.loadingMessage = `Processing ${assignmentsReadyForDriveProcessing.length} assignment(s)... (Step 1: Content Preparation)`;
    this.errorMessage = null;
    this.successMessage = null;
    this.changeDetectorRef.markForCheck();

    const allPackageFilesForServices: ImsccFile[] = this.unzippedFiles.map(uf => ({
      name: uf.name, data: uf.data, mimeType: uf.mimeType
    }));

    // Stage 1: Process content (Drive uploads, QTI conversion, HTML-to-Docs, link fixing)
    this.processAssignmentsAndCreateContent(courseName, assignmentsReadyForDriveProcessing, token, allPackageFilesForServices).pipe(
      tap(driveProcessingResults => {
        console.log('[Orchestrator] Content preparation stage (Drive, QTI, HTML) completed.');
        // Update the main assignments list with results from this stage (errors, final descriptions).
        // This is crucial for preparing items for ClassroomService.
        this.assignments = this.assignments.map(assignment => {
          const result = driveProcessingResults.find(r => r.itemId === assignment.associatedWithDeveloper?.id);
          if (result) { // If a result exists for this assignment
            if (result.error) {
              console.warn(`[Orchestrator] Item "${assignment.title}" (ID: ${result.itemId}) failed content prep: ${result.error.message}`);
              return {
                ...assignment,
                processingError: { // Set error from this stage
                  message: result.error.message || 'Content preparation failed.',
                  stage: result.error.stage || 'Content Preparation',
                  details: result.error.details
                }
              };
            }
            // If successful, update with final descriptions needed by ClassroomService
            if (result.finalHtmlDescription !== undefined && result.finalPlainTextDescription !== undefined) {
              return {
                ...assignment,
                descriptionForDisplay: result.finalHtmlDescription,
                descriptionForClassroom: result.finalPlainTextDescription, // Essential for ClassroomService
                richtext: !!result.finalHtmlDescription,
                processingError: undefined // Explicitly clear if this stage was successful
              };
            }
          }
          return assignment; // Return unchanged if no result found for it (should only be for non-processed items)
        });
        this.changeDetectorRef.markForCheck();
        this.loadingMessage = 'Preparing items for Classroom submission...';
      }),
      switchMap(driveProcessingResults => {
        // Filter items that are selected AND have no processingError from the content preparation stage.
        // These `itemsForClassroom` now have their final `descriptionForClassroom`.
        const itemsForClassroom = this.assignments.filter(assignment =>
          assignmentsToProcessIds.has(assignment.associatedWithDeveloper?.id || '') && !assignment.processingError
        );

        console.log(`[Orchestrator] Items filtered for Classroom submission: ${itemsForClassroom.length} items.`);

        if (itemsForClassroom.length === 0) {
          const message = assignmentsReadyForDriveProcessing.length > 0 ?
            "All selected items failed during content preparation." :
            "No items available for Classroom submission.";
          this.errorMessage = message;
          console.warn(`[Orchestrator] ${message}`);
          // Return an observable that emits a "skipped" status
          return of({type: 'skipped' as const, reason: message, data: [...this.assignments]});
        }

        // Add materials (Drive files, forms) to the ProcessedCourseWork items.
        // This is the final step before calling ClassroomService.
        const updatedAssignmentsForClassroom = this.addContentAsMaterials(driveProcessingResults, itemsForClassroom);
        console.log(`[Orchestrator] Assignments fully prepared for ClassroomService: ${updatedAssignmentsForClassroom.length} items.`);
        if (updatedAssignmentsForClassroom.length > 0) {
          console.log('[Orchestrator] Example of first assignment payload for ClassroomService (materials & description):',
            JSON.stringify(updatedAssignmentsForClassroom[0]?.materials, null, 2),
            JSON.stringify(updatedAssignmentsForClassroom[0]?.descriptionForClassroom, null, 2));
        }

        this.loadingMessage = `Submitting ${updatedAssignmentsForClassroom.length} assignment(s) to ${selectedContent.classroomIds.length} classroom(s)... This may take some time.`;
        this.changeDetectorRef.markForCheck();

        // Stage 2: Call the batch-enabled ClassroomService.
        // It will handle batching requests and individual item retries internally.
        return this.classroom.assignContentToClassrooms(token, selectedContent.classroomIds, updatedAssignmentsForClassroom).pipe(
          map(classroomServiceResults => {
            console.log('[Orchestrator] Received final results from ClassroomService.assignContentToClassrooms.');
            // Update the main assignments list with the final status from ClassroomService.
            // Each item in classroomServiceResults is a ProcessedCourseWork with its final Classroom status.
            this.assignments = this.assignments.map(assignment => {
              const finalResult = classroomServiceResults.find(cr => cr.associatedWithDeveloper?.id === assignment.associatedWithDeveloper?.id);
              return finalResult ? finalResult : assignment; // Replace with the version from ClassroomService
            });
            return {type: 'success' as const, response: classroomServiceResults, data: [...this.assignments]};
          }),
          catchError(classroomPipelineError => { // Catch errors from the assignContentToClassrooms observable itself
            const message = classroomPipelineError?.message || 'Unknown error during Classroom submission pipeline.';
            console.error(`[Orchestrator] Error from ClassroomService.assignContentToClassrooms pipeline: ${message}`, classroomPipelineError);
            // Mark all items that were ATTEMPTED for Classroom submission as failed if an overall pipeline error occurs.
            this.assignments = this.assignments.map(a => {
              // Check if this assignment 'a' was part of the batch sent to ClassroomService
              if (updatedAssignmentsForClassroom.some(itemSent => itemSent.associatedWithDeveloper?.id === a.associatedWithDeveloper?.id)) {
                // Only set error if not already set by a more specific error from batch response
                if (!a.processingError) {
                  a.processingError = {
                    message: `Classroom Submission Pipeline Error: ${message}`,
                    stage: 'Classroom Push Pipeline',
                    details: classroomPipelineError.details || classroomPipelineError
                  };
                }
              }
              return a;
            });
            return throwError(() => ({type: 'error' as const, source: 'Classroom Submission Pipeline', error: new Error(message), data: [...this.assignments], details: classroomPipelineError.details}));
          })
        );
      }),
      catchError((pipelineError: any) => { // Catch errors from the content preparation stage (tap/switchMap)
        const source = pipelineError?.source || 'Content Preparation Pipeline';
        const message = pipelineError?.error?.message || pipelineError?.message || 'An unknown error in content preparation.';
        console.error(`[Orchestrator] Error in main processing pipeline (source: ${source}): ${message}`, pipelineError);
        this.errorMessage = `Error during ${source}: ${message}`;
        // Ensure this.assignments reflects any partial updates or errors from pipelineError.data if available
        this.assignments = pipelineError.data && Array.isArray(pipelineError.data) ? [...pipelineError.data] : [...this.assignments];
        return of({type: 'error' as const, source: source, error: new Error(message), data: [...this.assignments]});
      }),
      finalize(() => {
        console.log("[Orchestrator] Main processing pipeline finalized.");
        this.isProcessing = false;
        this.loadingMessage = '';
        this.changeDetectorRef.markForCheck();
      })
    ).subscribe({
      next: (finalResult) => {
        console.log('[Orchestrator] Final result type of processing pipeline:', finalResult.type);
        this.assignments = [...finalResult.data]; // Ensure UI reflects the absolute final state

        if (finalResult.type === 'success') {
          const itemsProcessedByClassroom = finalResult.response; // Array of ProcessedCourseWork from ClassroomService
          const successfulSubmissions = itemsProcessedByClassroom.filter(r => !r.processingError).length;
          const failedSubmissions = itemsProcessedByClassroom.filter(r => r.processingError).length;

          if (failedSubmissions > 0) {
            this.errorMessage = `Submission to Classroom completed with ${failedSubmissions} item(s) failing. Check item details.`;
            if (successfulSubmissions > 0) {
              this.successMessage = `Successfully submitted ${successfulSubmissions} item(s) to classroom(s).`;
            } else { // No successes, only failures from classroom stage
              this.successMessage = null;
            }
          } else if (successfulSubmissions > 0) {
            this.successMessage = `Successfully submitted ${successfulSubmissions} item(s) to classroom(s).`;
            this.errorMessage = null;
          } else { // No items were successfully submitted by ClassroomService (e.g. all filtered before call or all failed)
            this.errorMessage = "No items were successfully submitted to Classroom. Check selection or previous errors.";
            this.successMessage = null;
          }
        } else if (finalResult.type === 'skipped') {
          this.errorMessage = `Processing skipped: ${finalResult.reason}`;
          this.successMessage = null;
        } else if (finalResult.type === 'error') { // Error from content prep stage or Classroom pipeline's catchError
          this.errorMessage = `Processing Error (${finalResult.source || 'Unknown Stage'}): ${finalResult.error.message}`;
          this.successMessage = null;
        }
        this.changeDetectorRef.markForCheck();
      },
      error: (err) => { // Catch unexpected errors from the overall subscription if something slips through
        const errMsg = err?.error?.message || err?.message || err?.toString() || 'An unexpected error occurred.';
        console.error('[Orchestrator] Critical uncaught error in subscription to process pipeline:', errMsg, err);
        this.errorMessage = `An unexpected critical error occurred: ${errMsg}`;
        this.successMessage = null;
        this.isProcessing = false;
        this.loadingMessage = '';
        this.changeDetectorRef.markForCheck();
      }
    });
  }


  // This method handles the pre-Classroom processing: Drive uploads, QTI to Forms, HTML to Docs, link fixing.
  // It returns ProcessingResult[] which includes final descriptions and any errors from this stage.
  processAssignmentsAndCreateContent(
    courseName: string,
    itemsToProcess: ProcessedCourseWork[], // Items selected for processing, with errors cleared for this run
    accessToken: string,
    allPackageImsccFiles: ImsccFile[]
  ): Observable<ProcessingResult[]> {
    console.log(`[Orchestrator] processAssignmentsAndCreateContent: Starting content preparation for ${itemsToProcess.length} items.`);
    if (itemsToProcess.length === 0) {
      return of([]);
    }

    return from(itemsToProcess).pipe(
      // Process items sequentially for Drive operations to manage load and for clearer per-item loading messages.
      concatMap((item, index) => {
        const itemLogPrefix = `ContentPrep Item ${index + 1}/${itemsToProcess.length} ("${item.title?.substring(0, 30)}..."):`;
        this.loadingMessage = itemLogPrefix;
        this.changeDetectorRef.markForCheck();
        console.log(`${itemLogPrefix} Beginning content preparation.`);

        const assignmentName = item.title || 'Untitled Assignment';
        // Use item's specific topic name for Drive folder organization.
        // ClassroomService will use item.associatedWithDeveloper.topic for the actual Classroom topic.
        const driveTopicFolderName = item.associatedWithDeveloper?.topic || 'General';
        let htmlDescriptionForProcessing = item.descriptionForDisplay; // Initial HTML for this item
        const itemId = item.associatedWithDeveloper?.id;

        // Prepare file data for upload service, ensuring ArrayBuffer for binary files.
        const filesToUploadForDriveService = (item.localFilesToUpload || []).map(ftu => {
          const originalUnzippedFile = this.unzippedFiles.find(uzf => uzf.name === ftu.file.name);
          if (originalUnzippedFile && originalUnzippedFile.data instanceof ArrayBuffer) {
            // If original data is ArrayBuffer, use it directly.
            return {...ftu, file: {...ftu.file, data: originalUnzippedFile.data}};
          } else if (typeof ftu.file.data === 'string' && ftu.file.data.startsWith('data:') && ftu.file.mimeType.startsWith('image/')) {
            // If it's a data URI (e.g., from image unzipping), FileUploadService might need to handle it
            // or it should be converted to Blob/ArrayBuffer here if FileUploadService expects that.
            // For now, pass as is; FileUploadService would need to be robust.
            console.warn(`${itemLogPrefix} File ${ftu.file.name} is a data URI. Ensure FileUploadService can handle this.`);
          } else if (typeof ftu.file.data === 'string' && !ftu.file.data.startsWith('data:')) {
            console.error(`${itemLogPrefix} File ${ftu.file.name} has raw string data (not data URI). This is unexpected for binary uploads.`);
          }
          return ftu;
        });

        const qtiFileForService: ImsccFile | undefined = item.qtiFile?.[0];
        const shouldCreateDoc = !qtiFileForService && !!htmlDescriptionForProcessing && item.workType === 'ASSIGNMENT' && !!item.richtext;

        if (!itemId) {
          console.error(`${itemLogPrefix} Critical error: Item is missing associatedDeveloper.id. Skipping content prep.`);
          return of({
            itemId: undefined, assignmentName, topicName: driveTopicFolderName,
            assignmentFolderId: 'ERROR_NO_ID', error: {message: 'Missing item ID for content prep', stage: 'Pre-flight (ContentPrep)'}
          } as ProcessingResult);
        }

        console.log(`${itemLogPrefix} ItemID: ${itemId}, DriveTopic: "${driveTopicFolderName}", FilesToUpload: ${filesToUploadForDriveService.length}, HasQTI: ${!!qtiFileForService}, ShouldCreateDoc: ${shouldCreateDoc}`);

        // Step 1: Ensure Drive folder structure
        return this.drive.ensureAssignmentFolderStructure(courseName, driveTopicFolderName, assignmentName, itemId, accessToken).pipe(
          switchMap(assignmentFolderId => {
            console.log(`${itemLogPrefix} Drive Folder ready: ${assignmentFolderId}`);

            // Step 2: Upload local files, shareReplay to ensure it executes once per item even if subscribed multiple times.
            const uploadedFiles$: Observable<DriveFile[]> = (filesToUploadForDriveService.length > 0 ?
              this.files.uploadLocalFiles(filesToUploadForDriveService, accessToken, assignmentFolderId) :
              of([])
            ).pipe(
              tap(uploaded => console.log(`${itemLogPrefix} ${uploaded.length} local files uploaded.`)),
              catchError(uploadErr => {
                console.error(`${itemLogPrefix} Error uploading files:`, uploadErr);
                return throwError(() => ({message: `File upload failed: ${uploadErr.message}`, stage: 'File Upload', details: uploadErr, itemId}));
              }),
              shareReplay(1)
            );

            // Step 3: Create primary content (QTI Form or HTML Doc) or process HTML for links
            let primaryContentCreation$: Observable<DriveFile | Material | null>; // Material for Form, DriveFile for Doc
            let currentHtmlContent = htmlDescriptionForProcessing; // Work with a mutable copy for this item

            if (qtiFileForService) { // QTI takes precedence
              console.log(`${itemLogPrefix} Processing QTI to Form.`);
              primaryContentCreation$ = this.qti.createFormFromQti(qtiFileForService, allPackageImsccFiles, assignmentName, accessToken, itemId, assignmentFolderId)
                .pipe(
                  tap(form => console.log(`${itemLogPrefix} QTI Form created/found.`)),
                  catchError(qtiErr => throwError(() => ({message: `QTI to Form failed: ${qtiErr.message}`, stage: 'QTI Processing', details: qtiErr, itemId})))
                );
            } else if (shouldCreateDoc) { // HTML to Doc if applicable
              console.log(`${itemLogPrefix} Processing HTML to Doc.`);
              primaryContentCreation$ = uploadedFiles$.pipe( // Ensure uploads complete before link replacement & doc creation
                switchMap(uploadedDriveFiles => {
                  if (currentHtmlContent) {
                    currentHtmlContent = this.replaceLocalLinksInHtml(currentHtmlContent, uploadedDriveFiles, filesToUploadForDriveService, itemLogPrefix);
                  }
                  return this.docs.createDocFromHtml(currentHtmlContent || '', assignmentName, accessToken, itemId, assignmentFolderId);
                }),
                tap(doc => console.log(`${itemLogPrefix} HTML Doc created/found.`)),
                catchError(docErr => throwError(() => ({message: `HTML to Doc failed: ${docErr.message}`, stage: 'HTML to Doc', details: docErr, itemId})))
              );
            } else { // No QTI, no Doc creation; if HTML exists, still process it for link replacement
              console.log(`${itemLogPrefix} No QTI/Doc creation. Processing HTML for links if present.`);
              primaryContentCreation$ = uploadedFiles$.pipe(
                map(uploadedDriveFiles => {
                  if (currentHtmlContent) {
                    currentHtmlContent = this.replaceLocalLinksInHtml(currentHtmlContent, uploadedDriveFiles, filesToUploadForDriveService, itemLogPrefix);
                  }
                  return null; // Signifies no primary Doc/Form was created, but HTML (if any) was processed
                }),
                catchError(linkErr => throwError(() => ({message: `HTML link processing failed: ${linkErr.message}`, stage: 'HTML Link Processing', details: linkErr, itemId})))
              );
            }

            // Step 4: Combine results of uploads and primary content creation
            return forkJoin({
              uploadedFilesResult: uploadedFiles$,
              createdContentResult: primaryContentCreation$
            }).pipe(
              map(({uploadedFilesResult, createdContentResult}) => {
                // Finalize descriptions for this item
                const finalHtmlDesc = currentHtmlContent || ''; // currentHtmlContent has link replacements
                const finalPlainTextDesc = this.generatePlainText(finalHtmlDesc);

                console.log(`${itemLogPrefix} Content preparation successful. Final plain text description length: ${finalPlainTextDesc.length}`);
                return {
                  itemId, assignmentName, topicName: driveTopicFolderName, assignmentFolderId,
                  createdDoc: (createdContentResult && 'mimeType' in createdContentResult) ? createdContentResult as DriveFile : undefined,
                  createdForm: (createdContentResult && 'form' in createdContentResult) ? createdContentResult as Material : undefined,
                  uploadedFiles: uploadedFilesResult,
                  error: null, // Success for this item's content prep
                  finalHtmlDescription: finalHtmlDesc,
                  finalPlainTextDescription: finalPlainTextDesc
                } as ProcessingResult;
              })
            );
          }),
          catchError((errorFromStage: any) => { // Catch errors from ensureAssignmentFolderStructure or any sub-stage
            const stage = errorFromStage?.stage || 'Content Preparation Sub-Stage';
            const message = errorFromStage?.message || 'Unknown error during content prep sub-stage.';
            console.error(`${itemLogPrefix} Error during content prep for item. Stage: ${stage}, Msg: ${message}`, errorFromStage.details);
            return of({
              itemId, assignmentName, topicName: driveTopicFolderName,
              assignmentFolderId: 'ERROR_CONTENT_PREP',
              error: {message, stage, details: errorFromStage.details || errorFromStage},
              finalHtmlDescription: htmlDescriptionForProcessing, // Fallback to original HTML on error
              finalPlainTextDescription: this.generatePlainText(htmlDescriptionForProcessing) // Fallback plain text
            } as ProcessingResult);
          })
        );
      }),
      toArray() // Collect results for all items
    );
  }

  // Helper to replace local links in HTML content with Drive links
  private replaceLocalLinksInHtml(
    htmlContent: string,
    uploadedDriveFiles: DriveFile[],
    filesOriginallyQueuedForUpload: Array<{file: ImsccFile; targetFileName: string}>,
    itemLogPrefix: string
  ): string {
    if (!htmlContent) return '';
    // This check prevents errors if DOMParser is not available in certain test environments,
    // though it should be available in a browser context.
    if (typeof DOMParser === 'undefined') {
      console.warn(`${itemLogPrefix} DOMParser not available. Skipping HTML link replacement.`);
      return htmlContent;
    }
    console.log(`${itemLogPrefix} Attempting to replace local links in HTML (length: ${htmlContent.length}).`);
    try {
      const parser = new DOMParser();
      const htmlDoc = parser.parseFromString(htmlContent, 'text/html');
      const body = htmlDoc.body || htmlDoc.documentElement; // Fallback to documentElement if body is null

      // Process <a> tags with data-imscc-original-path
      body.querySelectorAll('a[data-imscc-original-path]').forEach((anchor: Element) => {
        const htmlAnchor = anchor as HTMLAnchorElement;
        const originalPath = htmlAnchor.getAttribute('data-imscc-original-path');
        const displayTitleAttr = decode(htmlAnchor.getAttribute('data-imscc-display-title') || htmlAnchor.textContent || 'Linked File');

        if (originalPath) {
          let matchedDriveFile: DriveFile | undefined;
          const normalizedOriginalPath = normalizeSpacesString(this.util.tryDecodeURIComponent(originalPath).toLowerCase());

          for (const ftu of filesOriginallyQueuedForUpload) {
            // Match against the original zip path (ftu.file.name)
            const decodedFullZipPath = this.util.tryDecodeURIComponent(ftu.file.name);
            const normalizedFtuName = normalizeSpacesString(decodedFullZipPath.toLowerCase());

            // Try to match if the originalPath is a suffix of the full zip path or vice-versa (for flexibility with relative paths)
            if (normalizedFtuName.endsWith(normalizedOriginalPath) || normalizedOriginalPath.endsWith(normalizedFtuName)) {
              matchedDriveFile = uploadedDriveFiles.find(udf => udf.name === ftu.targetFileName); // Match uploaded file by its target name
              if (matchedDriveFile) break;
            }
          }

          if (matchedDriveFile?.webViewLink) {
            htmlAnchor.href = matchedDriveFile.webViewLink;
            htmlAnchor.target = '_blank';
            htmlAnchor.rel = 'noopener noreferrer'; // Security best practice for _blank links
            htmlAnchor.textContent = displayTitleAttr; // Use the decoded display title
            console.log(`${itemLogPrefix} Updated anchor for "${originalPath}" to Drive link: ${matchedDriveFile.webViewLink} with text "${displayTitleAttr}".`);
          } else {
            console.warn(`${itemLogPrefix} No Drive file found for local link anchor: "${originalPath}". Original text: "${displayTitleAttr}"`);
            htmlAnchor.textContent = `[Broken Link: ${displayTitleAttr}]`;
            htmlAnchor.removeAttribute('href'); // Remove href if broken
          }
          // Clean up helper attributes
          htmlAnchor.removeAttribute('data-imscc-original-path');
          htmlAnchor.removeAttribute('data-imscc-media-type');
          htmlAnchor.removeAttribute('data-imscc-display-title');
        }
      });

      // Process custom <span> placeholders (if any were generated by converter for non-<a> links)
      // Example: <span class="imscc-local-file-placeholder" data-original-path="path/to/file.pdf">File.pdf</span>
      body.querySelectorAll('span.imscc-local-file-placeholder[data-original-path]').forEach((span: Element) => {
        const htmlSpan = span as HTMLSpanElement;
        const originalPath = htmlSpan.getAttribute('data-original-path');
        const displayDefaultText = decode(htmlSpan.textContent || 'Linked Content');

        if (originalPath) {
          let matchedDriveFile: DriveFile | undefined;
          const normalizedOriginalPath = normalizeSpacesString(this.util.tryDecodeURIComponent(originalPath).toLowerCase());
          for (const ftu of filesOriginallyQueuedForUpload) {
            const decodedFullZipPath = this.util.tryDecodeURIComponent(ftu.file.name);
            const normalizedFtuName = normalizeSpacesString(decodedFullZipPath.toLowerCase());
            if (normalizedFtuName.endsWith(normalizedOriginalPath) || normalizedOriginalPath.endsWith(normalizedFtuName)) {
              matchedDriveFile = uploadedDriveFiles.find(udf => udf.name === ftu.targetFileName);
              if (matchedDriveFile) break;
            }
          }

          if (matchedDriveFile?.webViewLink) {
            const newAnchor = htmlDoc.createElement('a');
            newAnchor.href = matchedDriveFile.webViewLink;
            newAnchor.target = '_blank';
            newAnchor.rel = 'noopener noreferrer';
            newAnchor.textContent = displayDefaultText;
            span.parentNode?.replaceChild(newAnchor, span); // Replace span with new anchor
            console.log(`${itemLogPrefix} Replaced SPAN placeholder for "${originalPath}" with Drive link.`);
          } else {
            console.warn(`${itemLogPrefix} No Drive file found for SPAN placeholder: "${originalPath}".`);
            htmlSpan.textContent = `[Broken Content: ${displayDefaultText}]`;
            htmlSpan.style.color = "red"; // Indicate broken
          }
        }
      });
      // Consider other placeholder patterns if your converter service creates them.

      return htmlDoc.body.innerHTML;
    } catch (e) {
      console.error(`${itemLogPrefix} Error during HTML DOM manipulation for link replacement:`, e);
      return htmlContent; // Return original HTML on error
    }
  }

  // Generates plain text from HTML, truncating if necessary.
  private generatePlainText(html: string | undefined | null): string {
    if (!html) return '';
    if (typeof document === 'undefined') { // Guard for non-browser environments (e.g. server-side tests)
      console.warn("`document` is not available. Cannot generate plain text from HTML.");
      return '';
    }
    try {
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = html;
      // Extract text content, then normalize whitespace (replace multiple spaces/newlines with a single space)
      let text = (tempDiv.textContent || tempDiv.innerText || '').replace(/\s\s+/g, ' ').trim();
      const maxLen = 30000; // Google Classroom description length limit
      if (text.length > maxLen) {
        text = text.substring(0, maxLen - 3) + "..."; // Truncate and add ellipsis
        console.warn('[Orchestrator] Plain text description was truncated to fit Classroom API limits.');
      }
      return text;
    } catch (e) {
      console.error("[Orchestrator] Error generating plain text from HTML:", e);
      return ""; // Fallback to empty string on error
    }
  }

  // Helper to generate a unique key for a Material object for deduplication.
  private getMaterialKey(material: Material): string | null {
    if (material.driveFile?.driveFile?.id) return `drive-${material.driveFile.driveFile.id}`;
    if (material.youtubeVideo?.id) return `youtube-${material.youtubeVideo.id}`;
    if (material.link?.url) return `link-${this.util.tryDecodeURIComponent(material.link.url).toLowerCase()}`; // Normalize URL for key
    if (material.form?.formUrl) return `form-${this.util.tryDecodeURIComponent(material.form.formUrl).toLowerCase()}`; // Normalize URL for key
    return null; // Should not happen for valid, known material types
  }

  // Deduplicates materials based on their unique key.
  private deduplicateMaterials(materials: Material[]): Material[] {
    if (!materials || materials.length === 0) return [];
    const seenKeys = new Set<string>();
    const uniqueMaterials: Material[] = [];
    for (const material of materials) {
      const key = this.getMaterialKey(material);
      if (key !== null && !seenKeys.has(key)) {
        seenKeys.add(key);
        uniqueMaterials.push(material);
      } else if (key === null) { // If material type is unknown or key generation fails
        console.warn('[Orchestrator] deduplicateMaterials: Encountered material with no identifiable key. Including as is.', material);
        uniqueMaterials.push(material); // Include it anyway, but log
      } else {
        // This material (based on its key) has already been added.
        console.log(`[Orchestrator] deduplicateMaterials: Duplicate material skipped (Key: ${key})`);
      }
    }
    return uniqueMaterials;
  }

  // This method is called AFTER the Drive/Conversion stage (`processAssignmentsAndCreateContent`).
  // It takes items that are ready for Classroom and populates their `materials` array
  // based on the files uploaded/created during the Drive/Conversion stage.
  // The `itemsReadyForClassroom` should already have their final `descriptionForClassroom` set.
  addContentAsMaterials(
    driveProcessingResults: ProcessingResult[], // Results from the Drive/Conversion stage for ALL selected items
    itemsReadyForClassroom: ProcessedCourseWork[]  // Subset of items: selected, passed Drive/Conversion, and have final descriptions
  ): ProcessedCourseWork[] {
    console.log(`[Orchestrator] addContentAsMaterials: Finalizing materials for ${itemsReadyForClassroom.length} items.`);

    return itemsReadyForClassroom.map(item => {
      // Create a new object for modification to avoid side effects on the input array elements if they are directly from state.
      const updatedItem = {...item, materials: [...(item.materials || [])]}; // Start with any pre-existing materials

      const result = driveProcessingResults.find(r => r.itemId === updatedItem.associatedWithDeveloper?.id);
      const itemLogPrefix = `MaterialFinalization Item "${updatedItem.title}" (ID: ${updatedItem.associatedWithDeveloper?.id}):`;

      if (result && !result.error) { // Only add materials if the corresponding Drive/Conversion stage was successful
        const initialMaterialCount = updatedItem.materials.length;

        // Add created Google Doc as a material
        if (result.createdDoc?.id && result.createdDoc?.name) {
          updatedItem.materials.push({
            driveFile: {driveFile: {id: result.createdDoc.id, title: result.createdDoc.name}, shareMode: 'STUDENT_COPY'} // Default to STUDENT_COPY for primary docs
          });
        }
        // Add created Google Form link as a material
        if (result.createdForm?.form?.formUrl) {
          updatedItem.materials.push({
            link: {url: result.createdForm.form.formUrl, title: result.createdForm.form.title || result.assignmentName}
          });
        }
        // Add other uploaded files as materials
        (result.uploadedFiles || []).forEach(uploadedFile => {
          if (uploadedFile?.id && uploadedFile?.name) {
            // Avoid re-adding the primary createdDoc if it somehow also appears in general uploadedFiles
            if (!(result.createdDoc?.id === uploadedFile.id)) {
              updatedItem.materials.push({
                driveFile: {driveFile: {id: uploadedFile.id, title: uploadedFile.name}, shareMode: 'VIEW'} // Default to VIEW for supplementary files
              });
            }
          }
        });

        // Deduplicate all materials (existing + newly added) to ensure clean list for ClassroomService
        updatedItem.materials = this.deduplicateMaterials(updatedItem.materials);
        console.log(`${itemLogPrefix} Initial materials: ${initialMaterialCount}, Final materials after additions & deduplication: ${updatedItem.materials.length}`);

        // Descriptions (descriptionForDisplay, descriptionForClassroom) are assumed to be ALREADY FINALIZED
        // on the `item` object when it's passed into this function (as part of `itemsReadyForClassroom`).
        // They were set using `finalHtmlDescription` and `finalPlainTextDescription` from `driveProcessingResults`
        // when `this.assignments` was updated in the main `process` method.
        // Therefore, no need to update descriptions here again.
        // The fallback logic for descriptions that was here previously is removed,
        // relying on `processAssignmentsAndCreateContent` to be the source of truth for final descriptions.
        if (!updatedItem.descriptionForClassroom && updatedItem.materials.length > 0) {
          console.warn(`${itemLogPrefix} Item has materials but descriptionForClassroom is empty. Consider if a default description is needed if not set by content prep.`);
        }


      } else if (result && result.error) {
        // This case should ideally be filtered out by the logic that creates `itemsReadyForClassroom`,
        // but this log is a safeguard.
        console.warn(`${itemLogPrefix} Skipping material addition; item had an error in Drive/Conversion stage: ${result.error.message}`);
      } else if (!result) {
        console.warn(`${itemLogPrefix} No corresponding Drive/Conversion result found. Materials cannot be added from this stage.`);
      }
      return updatedItem; // Return the item with its finalized materials list
    });
  }
}
