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
import {Observable, from, of, throwError, forkJoin, EMPTY, Subject} from 'rxjs';
import {map, switchMap, concatMap, catchError, tap, finalize, takeUntil, toArray} from 'rxjs/operators';
import {CommonModule} from '@angular/common';
import { MatButtonModule } from '@angular/material/button';
import { MatInputModule } from '@angular/material/input';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatIconModule } from '@angular/material/icon';
import {MatProgressSpinnerModule} from '@angular/material/progress-spinner';
import { ClassroomService } from '../services/classroom/classroom.service';
import { ConverterService } from '../services/converter/converter.service';
import { DriveFolderService } from '../services/drive/drive.service';
import {HtmlToDocsService} from '../services/html-to-docs/html-to-docs.service';
import {FileUploadService} from '../services/file-upload/file-upload.service';
import {QtiToFormsService} from '../services/qti-to-forms/qti-to-forms.service';
import {UtilitiesService, RetryConfig} from '../services/utilities/utilities.service';
import { AuthService } from '../services/auth/auth.service';
import {
  ProcessedCourseWork,
  ProcessingResult,
  SubmissionData,
  Material,
  ImsccFile,
  DriveFile,
} from '../interfaces/classroom-interface';
import { CourseworkDisplayComponent } from '../coursework-display/coursework-display.component';
import {decode} from 'html-entities'; // For decoding HTML entities in text content

// Helper function to escape special characters for use in RegExp
function escapeRegExp(string: string): string {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
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

@Component({
  selector: 'app-file-upload',
  standalone: true,
  imports: [
    CommonModule,
    MatButtonModule,
    CourseworkDisplayComponent,
    MatInputModule,
    MatFormFieldModule,
    MatIconModule,
    MatProgressSpinnerModule
  ],
  templateUrl: './file-upload.component.html',
  styleUrl: './file-upload.component.scss'
})
export class FileUploadComponent {
  selectedFile: File | null = null;
  unzippedFiles: UnzippedFile[] = [];
  @ViewChild('fileInput') fileInput: any = null;
  assignments: ProcessedCourseWork[] = [];
  isProcessing: boolean = false;
  loadingMessage: string = '';
  errorMessage: string | null = null;
  successMessage: string | null = null;

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
    const files: { [key: string]: File } = this.fileInput.nativeElement.files;
    this.selectedFile = files[0];
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
            if (mimeType.startsWith('text/') || ['application/xml', 'application/json', 'application/javascript', 'image/svg+xml'].includes(mimeType)) {
              data = await zipEntry.async('string');
            } else if (mimeType.startsWith('image/')) {
              const base64Data = await zipEntry.async('base64');
              data = `data:${mimeType};base64,${base64Data}`;
            } else {
              data = await zipEntry.async('arraybuffer');
            }
            return { name: relativePath, data: data, mimeType: mimeType };
          })();
          filePromises.push(promise);
        }
      });

      const resolvedFiles = await Promise.all(filePromises);
      this.unzippedFiles = resolvedFiles.filter(f => f !== null) as UnzippedFile[];
      console.log(`[Orchestrator] Unzipped files prepared: ${this.unzippedFiles.length} files.`);

      const imsccFilesForConverter: ImsccFile[] = this.unzippedFiles.map(uf => {
        let dataForConverter: string;
        if (typeof uf.data === 'string') {
          dataForConverter = uf.data;
        } else if (uf.data instanceof ArrayBuffer) {
          console.warn(`[Orchestrator] File "${uf.name}" (mime: ${uf.mimeType}) is ArrayBuffer. ConverterService expects string. Using placeholder.`);
          dataForConverter = `[Binary File Placeholder: ${uf.name}]`;
        } else {
          console.error(`[Orchestrator] Unexpected data type for file ${uf.name} when preparing for ConverterService.`);
          dataForConverter = `[Error: Unexpected data type for ${uf.name}]`;
        }
        return { name: uf.name, data: dataForConverter, mimeType: uf.mimeType };
      });

      this.loadingMessage = 'Converting course structure...';
      this.changeDetectorRef.markForCheck();

      this.converter.convertImscc(imsccFilesForConverter)
        .pipe(
          finalize(() => {
            console.log('[Orchestrator] IMSCC conversion stream finalize.');
            if (!this.errorMessage) {
                 this.isProcessing = false;
                 this.loadingMessage = '';
                 this.successMessage = `Conversion complete. Found ${accumulatedAssignments.length} items. Ready for submission.`;
            }
            this.assignments = [...accumulatedAssignments];
            console.log('[Orchestrator] Final assignments structure after conversion:', this.assignments);
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
            this.isProcessing = false;
            this.loadingMessage = '';
            this.changeDetectorRef.markForCheck();
          }
        });

    } catch (error: any) {
      console.error('[Orchestrator] Error during the unzipping or IMSCC conversion process:', error);
      this.errorMessage = `Unzip/Read Error: ${error?.message || String(error)}`;
      this.assignments = [];
      this.isProcessing = false;
      this.loadingMessage = '';
      this.changeDetectorRef.markForCheck();
    }
  }

  process(selectedContent: SubmissionData) {
    if (this.isProcessing) {
      console.warn("[Orchestrator] Cannot start processing assignments while another operation is in progress.");
        this.errorMessage = "Processing is already underway. Please wait.";
        return;
    }
    if (selectedContent.assignmentIds.length === 0) {
        this.errorMessage = "Please select at least one assignment to submit.";
        return;
    }
     if (selectedContent.classroomIds.length === 0) {
        this.errorMessage = "Please select at least one classroom to submit to.";
        return;
    }

    console.log('[Orchestrator] Starting process function with selected content:', selectedContent);
    const token = this.auth.getGoogleAccessToken();

    if (!token) {
      this.errorMessage = "[Orchestrator] Processing aborted: No Google access token available. Please log in.";
      console.error(this.errorMessage);
      return;
    }

    let assignmentsToProcess = this.assignments.filter(assignment =>
      assignment.associatedWithDeveloper?.id &&
      selectedContent.assignmentIds.includes(assignment.associatedWithDeveloper.id)
    );

    if (assignmentsToProcess.length === 0) {
      console.warn("[Orchestrator] Processing skipped: No assignments selected or found matching the provided IDs.");
      this.errorMessage = "Could not find selected assignments to process.";
      return;
    }

    const assignmentsToProcessIds = new Set(assignmentsToProcess.map(a => a.associatedWithDeveloper?.id));
    this.assignments = this.assignments.map(a => {
        if (a.associatedWithDeveloper?.id && assignmentsToProcessIds.has(a.associatedWithDeveloper.id)) {
          console.log(`[Orchestrator] Clearing previous error for item: ${a.title} (ID: ${a.associatedWithDeveloper.id})`);
          return {...a, processingError: undefined};
        }
        return a;
    });
    assignmentsToProcess = this.assignments.filter(assignment =>
      assignment.associatedWithDeveloper?.id &&
      selectedContent.assignmentIds.includes(assignment.associatedWithDeveloper.id)
    );


    const courseName = this.converter.coursename || 'Untitled Course';
    console.log(`[Orchestrator] Processing ${assignmentsToProcess.length} selected assignments for course: "${courseName}"`);
    console.log(`[Orchestrator] Target Classroom IDs: ${selectedContent.classroomIds.join(', ')}`);

    this.isProcessing = true;
    this.loadingMessage = `Processing ${assignmentsToProcess.length} assignment(s)... (Step 1: Drive Content & QTI/HTML Conversion)`;
    this.errorMessage = null;
    this.successMessage = null;
    this.changeDetectorRef.markForCheck();

    const allPackageFilesForServices: ImsccFile[] = this.unzippedFiles.map(uf => {
      if (typeof uf.data === 'string') {
            return { name: uf.name, data: uf.data, mimeType: uf.mimeType };
      } else {
          return {name: uf.name, data: `[Binary data placeholder for ${uf.name}]`, mimeType: uf.mimeType};
        }
    });

    this.processAssignmentsAndCreateContent(courseName, assignmentsToProcess, token, allPackageFilesForServices).pipe(
      tap(driveProcessingResults => {
        console.log('[Orchestrator] Drive content/conversion stage completed. Results:', JSON.stringify(driveProcessingResults, null, 2));
        this.assignments = this.assignments.map(assignment => {
          const result = driveProcessingResults.find(r => r.itemId === assignment.associatedWithDeveloper?.id);
          if (result && result.error) {
            console.warn(`[Orchestrator] Item "${assignment.title}" (ID: ${result.itemId}) encountered error during Drive/Conversion stage: ${result.error.message}`);
            const errorDetails = result.error as StagedProcessingError;
            return {
              ...assignment,
              processingError: {
                message: errorDetails.message || 'Drive/Conversion operation failed.',
                stage: errorDetails.stage || 'Drive/Conversion Operation',
                details: errorDetails.details
              }
            };
          }
          return assignment;
        });
        console.log('[Orchestrator] Assignments after updating with Drive/Conversion stage errors:', JSON.stringify(this.assignments.filter(a => assignmentsToProcessIds.has(a.associatedWithDeveloper?.id)), null, 2));
        this.changeDetectorRef.markForCheck();

        const itemErrors = driveProcessingResults.filter(r => !!r.error);
        if (itemErrors.length > 0) {
          console.warn(`[Orchestrator] Drive/Conversion stage completed with ${itemErrors.length} item-level errors.`);
        }
        this.loadingMessage = 'Preparing items for Classroom...';
        this.changeDetectorRef.markForCheck();
      }),
      switchMap(driveProcessingResults => {
        console.log('[Orchestrator] In switchMap after Drive/Conversion processing. Preparing items for Classroom submission.');

        const itemsForClassroom = this.assignments.filter(assignment => {
          const isSelected = assignment.associatedWithDeveloper?.id && selectedContent.assignmentIds.includes(assignment.associatedWithDeveloper.id);
          const hasNoError = !assignment.processingError;
          if (isSelected) console.log(`[Orchestrator] Checking item "${assignment.title}" for Classroom: Selected: ${isSelected}, HasNoError: ${hasNoError}`);
          return isSelected && hasNoError;
        });

        console.log(`[Orchestrator] Items filtered for Classroom (selected and no prior error): ${itemsForClassroom.length} items.`);
        if (itemsForClassroom.length > 0) {
          console.log('[Orchestrator] First item being prepared for Classroom:', JSON.stringify(itemsForClassroom[0], null, 2));
        }


        if (itemsForClassroom.length === 0 && assignmentsToProcess.length > 0) {
          this.errorMessage = "All selected items failed during content preparation (Drive/Conversion stage).";
          console.warn(`[Orchestrator] ${this.errorMessage}`);
            this.changeDetectorRef.markForCheck();
          return throwError(() => ({type: 'error' as const, source: 'Content Preparation Stage', error: new Error(this.errorMessage || 'Unknown content preparation error'), data: [...this.assignments]}));
        }
        if (itemsForClassroom.length === 0) {
          console.warn('[Orchestrator] No items to submit to Classroom (either none were selected for processing or all failed in content prep).');
          return of({type: 'skipped' as const, reason: 'No items to submit to Classroom after content preparation.', data: [...this.assignments]});
        }

        const updatedAssignmentsForClassroom = this.addContentAsMaterials(driveProcessingResults, itemsForClassroom);
        console.log(`[Orchestrator] Assignments after addContentAsMaterials for Classroom: ${updatedAssignmentsForClassroom.length} items.`);
        if (updatedAssignmentsForClassroom.length > 0) {
          console.log('[Orchestrator] First assignment payload to be sent to ClassroomService:', JSON.stringify(updatedAssignmentsForClassroom[0], null, 2));
        }


        if (!selectedContent.classroomIds || selectedContent.classroomIds.length === 0) {
          console.warn('[Orchestrator] No classrooms selected. Skipping Classroom submission.');
          return of({ type: 'skipped' as const, reason: 'No classrooms selected', data: [...this.assignments] });
        }

        this.loadingMessage = `Submitting ${updatedAssignmentsForClassroom.length} assignment(s) to ${selectedContent.classroomIds.length} classroom(s)...`;
        this.changeDetectorRef.markForCheck();

        console.log(`[Orchestrator] >>> Calling ClassroomService.assignContentToClassrooms with ${updatedAssignmentsForClassroom.length} assignments for ${selectedContent.classroomIds.length} classrooms.`);
        return this.classroom.assignContentToClassrooms(token, selectedContent.classroomIds, updatedAssignmentsForClassroom).pipe(
          map(classroomResults => {
            console.log('[Orchestrator] Received results from ClassroomService.assignContentToClassrooms:', classroomResults);
            this.assignments = this.assignments.map(assignment => {
              const classroomResult = classroomResults.find(cr => cr.associatedWithDeveloper?.id === assignment.associatedWithDeveloper?.id);
              if (classroomResult) {
                console.log(`[Orchestrator] Updating assignment "${assignment.title}" with Classroom result.`);
                return classroomResult;
              }
              return assignment;
            });
            return { type: 'success' as const, response: classroomResults, data: [...this.assignments] };
          }),
          catchError(classroomError => {
            const message = classroomError?.message || 'Unknown error during Classroom submission phase.';
            console.error(`[Orchestrator] Error from ClassroomService.assignContentToClassrooms pipeline: ${message}`, classroomError);
            this.assignments = this.assignments.map(a => {
              if (updatedAssignmentsForClassroom.find(i => i.associatedWithDeveloper?.id === a.associatedWithDeveloper?.id)) {
                    if (!a.processingError) {
                      a.processingError = {message: `Classroom Submission Error: ${message}`, stage: 'Classroom Push', details: classroomError.details || classroomError};
                    }
                }
                return a;
            });
            return throwError(() => ({type: 'error' as const, source: 'Classroom Submission', error: new Error(message), data: [...this.assignments], details: classroomError.details}));
          })
        );
      }),
      catchError((pipelineError: any) => {
        const source = pipelineError?.source || 'Content Preparation Pipeline';
        const message = pipelineError?.error?.message || pipelineError?.message || 'An unknown error occurred in the content preparation pipeline.';
        console.error(`[Orchestrator] Error in main processing pipeline (source: ${source}): ${message}`, pipelineError);
        this.errorMessage = `Error during ${source}: ${message}`;
        this.assignments = pipelineError.data && Array.isArray(pipelineError.data) ? [...pipelineError.data] : [...this.assignments];
        return of({ type: 'error' as const, source: source, error: new Error(message), data: [...this.assignments] });
      }),
      finalize(() => {
        console.log("[Orchestrator] Main processing pipeline finalize block.");
          this.isProcessing = false;
          this.loadingMessage = '';
          this.changeDetectorRef.markForCheck();
      })
    ).subscribe({
      next: (finalResult) => {
        console.log('[Orchestrator] Final result of processing pipeline:', finalResult);
        this.assignments = [...finalResult.data];

        if (finalResult.type === 'success') {
          const itemsWithErrors = this.assignments.filter(a => a.processingError && selectedContent.assignmentIds.includes(a.associatedWithDeveloper?.id || '')).length;
          const successfulItems = finalResult.response.filter(r => !r.processingError).length;

          if (itemsWithErrors > 0) {
            this.errorMessage = `Submission completed with ${itemsWithErrors} item(s) failing. Check item details.`;
            if (successfulItems > 0) {
              this.successMessage = `Successfully processed ${successfulItems} item(s).`;
            } else {
              this.successMessage = null;
            }
          } else if (successfulItems > 0) {
            this.successMessage = `Successfully submitted ${successfulItems} item(s) to classroom(s).`;
            this.errorMessage = null;
          } else {
            this.errorMessage = "No items were successfully submitted. Please check the selection or previous errors.";
            this.successMessage = null;
          }

        } else if (finalResult.type === 'skipped') {
          this.errorMessage = `Submission skipped: ${finalResult.reason}`;
          this.successMessage = null;
        } else if (finalResult.type === 'error') {
          this.errorMessage = finalResult.error.message;
           this.successMessage = null;
        }
        this.changeDetectorRef.markForCheck();
      },
      error: (err) => {
        const errMsg = err?.error?.message || err?.message || err?.toString() || 'An unexpected error occurred during the final processing stage.';
        console.error('[Orchestrator] Uncaught error in subscription to process pipeline:', errMsg, err);
        this.errorMessage = `An unexpected error occurred: ${errMsg}`;
        this.successMessage = null;
        this.isProcessing = false;
        this.loadingMessage = '';
        this.changeDetectorRef.markForCheck();
      }
    });
  }


  processAssignmentsAndCreateContent(
    courseName: string,
    itemsToProcess: ProcessedCourseWork[],
    accessToken: string,
    allPackageImsccFiles: ImsccFile[]
  ): Observable<ProcessingResult[]> {

    console.log(`[Orchestrator] processAssignmentsAndCreateContent: Starting Drive processing for ${itemsToProcess.length} items in course "${courseName}"`);

    if (itemsToProcess.length === 0) {
        return of([]);
    }

    return from(itemsToProcess).pipe(
      concatMap((item, index) => {
        const itemLogPrefix = `[Orchestrator] Drive Processing Item ${index + 1}/${itemsToProcess.length} ("${item.title?.substring(0, 30)}..."):`;
        this.loadingMessage = `${itemLogPrefix} (Drive Content & QTI/HTML Conversion)`;
        this.changeDetectorRef.markForCheck();
        console.log(`${itemLogPrefix} Start.`);

        const assignmentName = item.title || 'Untitled Assignment';
        const topicName = item.associatedWithDeveloper?.topic || 'General';
        let currentDescriptionForDisplay = item.descriptionForDisplay; // Use a mutable variable for HTML
        const itemId = item.associatedWithDeveloper?.id;

        const filesToUploadForDrive = item.localFilesToUpload || [];

        const qtiFileArray = item.qtiFile;
        const qtiFileForService: ImsccFile | undefined = (qtiFileArray && qtiFileArray.length > 0) ? qtiFileArray[0] : undefined;

        const shouldCreateDoc = !qtiFileForService && !!currentDescriptionForDisplay && item.workType === 'ASSIGNMENT' && !!item.richtext;

        if (!itemId) {
          console.error(`${itemLogPrefix} Skipping item due to missing associated developer ID.`);
          return of({
            itemId: undefined, assignmentName, topicName,
            assignmentFolderId: 'ERROR_NO_ID', createdDoc: undefined, createdForm: undefined,
            uploadedFiles: undefined,
            error: {message: 'Missing associated developer ID', stage: 'Pre-flight Check (Drive/Conversion)'}
          } as ProcessingResult);
        }

        console.log(`${itemLogPrefix} Item ID: ${itemId}, Topic: "${topicName}", Files to Upload: ${filesToUploadForDrive.length}, Has QTI: ${!!qtiFileForService}, Should Create Doc: ${shouldCreateDoc}`);

        return this.drive.ensureAssignmentFolderStructure(
          courseName, topicName, assignmentName, itemId, accessToken
        ).pipe(
          switchMap(assignmentFolderId => {
            console.log(`${itemLogPrefix} Drive Folder ensured/created. ID: ${assignmentFolderId}`);

            const filesForDriveUploadService = filesToUploadForDrive.map(ftu => {
              const originalUnzippedFile = this.unzippedFiles.find(uzf => uzf.name === ftu.file.name);
              if (originalUnzippedFile && originalUnzippedFile.data instanceof ArrayBuffer) {
                return {file: {...ftu.file, data: originalUnzippedFile.data}, targetFileName: ftu.targetFileName};
              }
              return ftu;
            });


            const uploadFiles$: Observable<DriveFile[]> = filesForDriveUploadService.length > 0
              ? this.files.uploadLocalFiles(filesForDriveUploadService, accessToken, assignmentFolderId).pipe(
                tap(uploadedFiles => console.log(`${itemLogPrefix} Uploaded ${uploadedFiles.length} local file(s).`)),
                catchError(uploadError => {
                  console.error(`${itemLogPrefix} ERROR uploading local files:`, uploadError);
                  return throwError(() => ({
                    message: uploadError.message || 'File upload failed',
                    stage: 'File Upload',
                    details: uploadError.toString(),
                    itemId
                  } as StagedProcessingError & { itemId: string } ));
                })
              )
              : of([]);

            let createContent$: Observable<DriveFile | Material | null>;

            if (qtiFileForService) {
              this.loadingMessage = `${itemLogPrefix} (Creating Form from QTI)`;
              this.changeDetectorRef.markForCheck();
              console.log(`${itemLogPrefix} QTI file found. Calling QtiToFormsService.`);
              createContent$ = this.qti.createFormFromQti(
                qtiFileForService,
                allPackageImsccFiles,
                assignmentName,
                accessToken,
                itemId,
                assignmentFolderId
              ).pipe(
                tap(createdForm => console.log(`${itemLogPrefix} Created/Found Google Form.`)),
                catchError(formError => {
                  console.error(`${itemLogPrefix} ERROR creating Google Form:`, formError);
                  return throwError(() => ({
                    message: formError.message || 'Form creation from QTI failed',
                    stage: 'Form Creation',
                    details: formError.toString(),
                    itemId
                  } as StagedProcessingError & { itemId: string }));
                })
              );
            } else if (shouldCreateDoc) {
              this.loadingMessage = `${itemLogPrefix} (Creating Doc from HTML)`;
              this.changeDetectorRef.markForCheck();
              console.log(`${itemLogPrefix} No QTI, should create Doc. Waiting for file uploads to complete if any.`);
              createContent$ = uploadFiles$.pipe(
                switchMap(uploadedDriveFiles => {
                  console.log(`${itemLogPrefix} File uploads completed for "${item.title}" (or none needed). Uploaded files count: ${uploadedDriveFiles.length}. Proceeding to create Doc from HTML.`);

                  // --- Start Unified Placeholder Replacement Logic ---
                  if (currentDescriptionForDisplay) {
                    console.log(`${itemLogPrefix} HTML content present. Attempting to replace all placeholders in HTML for "${item.title}". Original HTML length: ${currentDescriptionForDisplay.length}`);
                    try {
                      const parser = new DOMParser();
                      const htmlDoc = parser.parseFromString(currentDescriptionForDisplay, 'text/html');
                      const body = htmlDoc.body || htmlDoc.documentElement;
                      const spansToReplace: {span: HTMLSpanElement, newLinkElement: HTMLParagraphElement}[] = [];

                      // 1. Process all SPAN placeholders (general files, images, and now videos)
                      body.querySelectorAll('span').forEach(span => {
                        const spanText = span.textContent || "";
                        let replacementMadeForThisSpan = false;

                        // Try Video Placeholder First (due to its specific class and text structure)
                        if (span.classList.contains('imscc-video-placeholder-text')) {
                          const videoMatch = spanText.match(/\[VIDEO_PLACEHOLDER REF_NAME="([^"]+)" DISPLAY_TITLE="([^"]+)"\]/);
                          if (videoMatch) {
                            const originalVideoSrcDecoded = videoMatch[1]; // Already decoded by ConverterService
                            const displayTitleDecoded = videoMatch[2];     // Already decoded
                            console.log(`${itemLogPrefix} Found video placeholder SPAN: REF_NAME="${originalVideoSrcDecoded}", DISPLAY_TITLE="${displayTitleDecoded}"`);

                            let matchedDriveFile: DriveFile | undefined;
                            for (const ftu of filesForDriveUploadService) {
                              const decodedFtuName = this.util.tryDecodeURIComponent(ftu.file.name);
                              if (decodedFtuName === originalVideoSrcDecoded) {
                                matchedDriveFile = uploadedDriveFiles.find(udf => udf.name === ftu.targetFileName);
                                if (matchedDriveFile) {
                                  console.log(`${itemLogPrefix} Matched video SPAN REF_NAME "${originalVideoSrcDecoded}" to DriveFile "${matchedDriveFile.name}" (target: "${ftu.targetFileName}") with link ${matchedDriveFile.webViewLink}`);
                                }
                                break;
                              }
                            }
                            if (matchedDriveFile?.webViewLink) {
                              const p = htmlDoc.createElement('p');
                              const a = htmlDoc.createElement('a');
                              a.href = matchedDriveFile.webViewLink; a.target = '_blank'; a.rel = 'noopener noreferrer';
                              a.textContent = `Video: ${displayTitleDecoded}`;
                              p.appendChild(a);
                              spansToReplace.push({span, newLinkElement: p});
                              replacementMadeForThisSpan = true;
                            } else {
                              console.warn(`${itemLogPrefix} Could not find uploaded Drive file or webViewLink for video placeholder SPAN REF_NAME="${originalVideoSrcDecoded}". Placeholder SPAN will remain.`);
                            }
                          }
                        }

                        // If not a video placeholder or video replacement failed, try general image/file placeholders
                        if (!replacementMadeForThisSpan) {
                          for (const uploadedDriveFile of uploadedDriveFiles) {
                            const originalFileDetail = filesForDriveUploadService.find(ftu => ftu.targetFileName === uploadedDriveFile.name);
                            if (originalFileDetail && uploadedDriveFile.webViewLink) {
                              const targetFileName = originalFileDetail.targetFileName;

                              const imagePlaceholderMatch = spanText.match(/\[Image:\s*(.+?)\s*-\s*will be attached separately\]/i);
                              if (imagePlaceholderMatch && imagePlaceholderMatch[1].trim() === targetFileName) {
                                const p = htmlDoc.createElement('p');
                                const a = htmlDoc.createElement('a');
                                a.href = uploadedDriveFile.webViewLink; a.target = '_blank'; a.rel = 'noopener noreferrer';
                                a.textContent = `Image: ${decode(targetFileName)}`;
                                p.appendChild(a);
                                spansToReplace.push({span, newLinkElement: p});
                                replacementMadeForThisSpan = true;
                                break;
                              }

                              const filePlaceholderMatch = spanText.match(/\[Attached File:\s*(.+?)\]/i);
                              if (filePlaceholderMatch && filePlaceholderMatch[1].trim() === targetFileName) {
                                const originalLinkText = decode(spanText.substring(0, filePlaceholderMatch.index).trim() || `File: ${targetFileName}`);
                                const p = htmlDoc.createElement('p');
                                const a = htmlDoc.createElement('a');
                                a.href = uploadedDriveFile.webViewLink; a.target = '_blank'; a.rel = 'noopener noreferrer';
                                a.textContent = originalLinkText;
                                p.appendChild(a);
                                spansToReplace.push({span, newLinkElement: p});
                                replacementMadeForThisSpan = true;
                                break;
                              }
                            }
                          }
                        }
                      });

                      // Perform DOM modifications after iterating
                      spansToReplace.forEach(rep => {
                        if (rep.span.parentNode) {
                          rep.span.parentNode.replaceChild(rep.newLinkElement, rep.span);
                          console.log(`${itemLogPrefix} Replaced placeholder SPAN with link: ${rep.newLinkElement.outerHTML}`);
                        }
                      });

                      currentDescriptionForDisplay = body.innerHTML;
                      console.log(`${itemLogPrefix} HTML after all placeholder replacements. New length: ${currentDescriptionForDisplay.length}`);

                    } catch (e) {
                      console.error(`${itemLogPrefix} Error during HTML DOM manipulation for placeholder replacement:`, e);
                    }
                  } else {
                    console.log(`${itemLogPrefix} No HTML content to modify for "${item.title}".`);
                  }
                  // --- End Unified Placeholder Replacement Logic ---

                  return this.docs.createDocFromHtml(
                    currentDescriptionForDisplay || '',  // Ensure string, even if null/undefined
                    assignmentName, accessToken, itemId, assignmentFolderId
                  ).pipe(
                    tap(createdDoc => console.log(`${itemLogPrefix} Created/Found Google Doc.`)),
                    catchError(docError => {
                      console.error(`${itemLogPrefix} ERROR creating Google Doc:`, docError);
                      return throwError(() => ({
                        message: docError.message || 'Document creation from HTML failed',
                        stage: 'Document Creation',
                        details: docError.toString(),
                        itemId
                      } as StagedProcessingError & { itemId: string }));
                    })
                  );
                })
              );
            } else {
              console.log(`${itemLogPrefix} No QTI and not creating Doc. Content creation step is null.`);
              createContent$ = of(null);
            }

            return forkJoin({
              uploadedFilesResult: uploadFiles$,
              createdContentResult: createContent$
            }).pipe(
              map(({ uploadedFilesResult, createdContentResult }) => {
                let createdDoc: DriveFile | undefined = undefined;
                let createdForm: Material | undefined = undefined;
                if (createdContentResult) {
                    if ('mimeType' in createdContentResult && typeof createdContentResult.mimeType === 'string') {
                        createdDoc = createdContentResult as DriveFile;
                    } else if ('form' in createdContentResult && createdContentResult.form) {
                        createdForm = createdContentResult as Material;
                    }
                }
                console.log(`${itemLogPrefix} Drive content creation/finding complete. Doc: ${!!createdDoc}, Form: ${!!createdForm}, Files: ${uploadedFilesResult.length}`);
                return {
                  itemId: itemId, assignmentName, topicName, assignmentFolderId,
                  createdDoc, createdForm,
                  uploadedFiles: uploadedFilesResult,
                  error: null
                } as ProcessingResult;
              })
            );
          }),
          catchError((errorDetails: any) => {
            const stage = errorDetails?.stage || 'Folder Creation (Drive)';
            const errorMessageText = errorDetails?.message || errorDetails?.error?.message || 'Unknown error during item Drive processing';
            const errorDetailsString = errorDetails?.details || errorDetails?.error?.toString() || errorDetails.toString();

            console.error(`${itemLogPrefix} ERROR during Drive processing. Stage: ${stage}, Message: ${errorMessageText}`);
            return of({
              itemId: itemId, assignmentName, topicName,
              assignmentFolderId: 'ERROR_DRIVE_PROCESSING', createdDoc: undefined, createdForm: undefined,
              uploadedFiles: undefined,
              error: {message: errorMessageText, stage: stage, details: errorDetailsString}
            } as ProcessingResult);
          })
        );
      }),
      toArray()
    );
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
        console.warn('[Orchestrator] deduplicateMaterials: Encountered material with no identifiable key, including as is.', material);
        uniqueMaterials.push(material);
      } else {
        // console.log(`[Orchestrator] deduplicateMaterials: Duplicate material skipped (Key: ${key})`);
      }
    }
    return uniqueMaterials;
  }


  addContentAsMaterials(
    processingResults: ProcessingResult[],
    courseWorkItemsToUpdate: ProcessedCourseWork[]
  ): ProcessedCourseWork[] {
    console.log('[Orchestrator] addContentAsMaterials: Associating Drive content as Materials...');
    const courseWorkMap = new Map<string, ProcessedCourseWork>();
    courseWorkItemsToUpdate.forEach(item => {
      if (item.associatedWithDeveloper?.id) {
        courseWorkMap.set(item.associatedWithDeveloper.id, { ...item, materials: [...(item.materials || [])] });
      }
    });

    processingResults.forEach(result => {
      if (result.itemId) {
        const courseWorkItem = courseWorkMap.get(result.itemId);
        if (courseWorkItem) {
          const itemLogPrefix = `[Orchestrator] Material Association for Item "${courseWorkItem.title}" (ID: ${result.itemId}):`;

          if (!courseWorkItem.materials) {
            courseWorkItem.materials = [];
          }

          console.log(`${itemLogPrefix} Updating materials. Current material count: ${courseWorkItem.materials.length}`);
          const initialMaterialCount = courseWorkItem.materials.length;
          const addedMaterialNamesThisPass: string[] = [];
          let primaryContentMaterialThisPass: Material | null = null;

          if (result.createdDoc?.id && result.createdDoc?.name) {
            const docMaterial: Material = { driveFile: { driveFile: { id: result.createdDoc.id, title: result.createdDoc.name }, shareMode: 'STUDENT_COPY' } };
            if (!courseWorkItem.materials.some(m => m.driveFile?.driveFile?.id === result.createdDoc?.id)) {
              courseWorkItem.materials.push(docMaterial);
              addedMaterialNamesThisPass.push(`"${result.createdDoc.name}" (Doc)`);
              if (!primaryContentMaterialThisPass) primaryContentMaterialThisPass = docMaterial;
            }
          }

          if (result.createdForm?.form?.formUrl) {
            const formMaterial: Material = {link: {url: result.createdForm.form.formUrl, title: result.createdForm.form.title || result.assignmentName}};
            if (!courseWorkItem.materials.some(m => m.link?.url === result.createdForm?.form?.formUrl)) {
              courseWorkItem.materials.push(formMaterial);
              addedMaterialNamesThisPass.push(`"${result.createdForm.form.title || result.assignmentName}" (Form)`);
              if (!primaryContentMaterialThisPass) primaryContentMaterialThisPass = formMaterial;
            }
          }

          const filesToAddAsMaterials = result.uploadedFiles || [];
          if (filesToAddAsMaterials.length > 0) {
            filesToAddAsMaterials.forEach(uploadedFile => {
              if (uploadedFile?.id && uploadedFile?.name) {
                if (!(result.createdDoc?.id === uploadedFile.id)) {
                  if (!courseWorkItem.materials!.some(m => m.driveFile?.driveFile?.id === uploadedFile.id)) {
                    const fileMaterial: Material = { driveFile: { driveFile: { id: uploadedFile.id, title: uploadedFile.name }, shareMode: 'VIEW' } };
                    courseWorkItem.materials!.push(fileMaterial);
                    addedMaterialNamesThisPass.push(`"${uploadedFile.name}"`);
                  }
                }
              }
            });
          }

          courseWorkItem.materials = this.deduplicateMaterials(courseWorkItem.materials);


          if (addedMaterialNamesThisPass.length > 0 && !courseWorkItem.descriptionForClassroom?.trim()) {
              let materialDescription: string;
            if (primaryContentMaterialThisPass) {
              if (primaryContentMaterialThisPass.driveFile) {
                materialDescription = `Please review the attached document: ${addedMaterialNamesThisPass[0]}.`;
              } else if (primaryContentMaterialThisPass.link && primaryContentMaterialThisPass.link.url.includes('google.com/forms')) {
                materialDescription = `Please complete the attached form: ${addedMaterialNamesThisPass[0]}.`;
              } else {
                materialDescription = `Please see the attached content: ${addedMaterialNamesThisPass[0]}.`;
                  }
                if (addedMaterialNamesThisPass.length > 1) {
                  const otherFiles = addedMaterialNamesThisPass.slice(1);
                      if (otherFiles.length > 0) {
                         materialDescription += ` Additional file(s): ${otherFiles.join(', ')}.`;
                      }
                  }
              } else {
                if (addedMaterialNamesThisPass.length === 1) materialDescription = `Please see the attached file: ${addedMaterialNamesThisPass[0]}.`;
                else materialDescription = `Please see the attached files (${addedMaterialNamesThisPass.length}): ${addedMaterialNamesThisPass.join(', ')}.`;
              }
              courseWorkItem.descriptionForClassroom = materialDescription;
            console.log(`${itemLogPrefix} Set classroom description based on new materials: "${materialDescription}"`);
          } else if (addedMaterialNamesThisPass.length > 0) {
            console.log(`${itemLogPrefix} ${addedMaterialNamesThisPass.length} new material(s) associated. Kept existing classroom description.`);
          } else if (courseWorkItem.materials.length > initialMaterialCount) {
            console.log(`${itemLogPrefix} Materials were re-associated or already present. No new materials added in this specific pass. Kept existing classroom description.`);
          } else {
            console.log(`${itemLogPrefix} No new materials were added. Keeping original description.`);
          }
          console.log(`${itemLogPrefix} Final material count: ${courseWorkItem.materials.length}`);


        } else {
          console.warn(`[Orchestrator] addContentAsMaterials: Could not find matching CourseWork item in map for ProcessingResult with itemId: ${result.itemId}.`);
        }
      } else if (result.error) {
        console.error(`[Orchestrator] addContentAsMaterials: Skipping material addition for item ID ${result.itemId || 'Unknown'} due to processing error:`, result.error);
      } else if (!result.itemId) {
        console.error(`[Orchestrator] addContentAsMaterials: Skipping material addition for item "${result.assignmentName}" because it had no ID.`);
      }
    });

    console.log('[Orchestrator] addContentAsMaterials: Finished associating materials.');
    return Array.from(courseWorkMap.values());
  }

}
