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
import {FileUploadService as ActualFileUploadService} from '../services/file-upload/file-upload.service';
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

export interface ProcessingResult {
  itemId: string | undefined;
  assignmentName: string;
  topicName: string;
  assignmentFolderId: string;
  createdDoc?: DriveFile; // Doc from main description
  createdForm?: Material; // Form from QTI
  uploadedFiles?: DriveFile[]; // Includes other uploaded files AND docs converted from HTML assets
  error: StagedProcessingError | null;
  finalHtmlDescription?: string;
  finalPlainTextDescription?: string;
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
  assignments: ProcessedCourseWork[] = [];
  isProcessing: boolean = false;
  loadingMessage: string = '';
  errorMessage: string | null = null;
  successMessage: string | null = null;

  classroom = inject(ClassroomService);
  converter = inject(ConverterService);
  drive = inject(DriveFolderService);
  docs = inject(HtmlToDocsService);
  files = inject(ActualFileUploadService);
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
            if (mimeType.startsWith('text/') ||
              ['application/xml', 'application/json', 'application/javascript', 'image/svg+xml'].includes(mimeType) ||
              /\.(html|htm|css|js|xml|qti|txt|md)$/i.test(relativePath)) {
              data = await zipEntry.async('string');
            } else if (mimeType.startsWith('image/')) {
              const base64Data = await zipEntry.async('base64');
              data = `data:${mimeType};base64,${base64Data}`;
            } else {
              data = await zipEntry.async('arraybuffer');
            }
            return {name: relativePath, data: data, mimeType: mimeType};
          })();
          filePromises.push(promise);
        }
      });

      const resolvedFiles = await Promise.all(filePromises);
      this.unzippedFiles = resolvedFiles.filter(f => f !== null) as UnzippedFile[];
      console.log(`[Orchestrator] Unzipped files prepared: ${this.unzippedFiles.length} files.`, this.unzippedFiles);

      const imsccFilesForConverter: ImsccFile[] = this.unzippedFiles.map(uf => ({
        name: uf.name, data: uf.data, mimeType: uf.mimeType
      }));

      this.loadingMessage = 'Converting course structure...';
      this.changeDetectorRef.markForCheck();

      this.converter.convertImscc(imsccFilesForConverter)
        .pipe(
          finalize(() => {
            console.log('[Orchestrator] IMSCC conversion stream finalized.');
            if (!this.isProcessing && !this.errorMessage) {
              this.successMessage = `Conversion complete. Found ${accumulatedAssignments.length} items. Ready for submission.`;
            }
            this.isProcessing = false;
            this.loadingMessage = '';
            this.assignments = [...accumulatedAssignments];
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
            this.isProcessing = false;
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

    const initialTokenCheck = this.auth.getGoogleAccessToken();
    if (!initialTokenCheck) {
      this.errorMessage = "Processing aborted: User not authenticated. Please log in.";
      console.error(this.errorMessage);
      this.isProcessing = false;
      this.loadingMessage = '';
      this.changeDetectorRef.markForCheck();
      return;
    }

    const assignmentsToProcessIds = new Set(selectedContent.assignmentIds);
    const assignmentsReadyForDriveProcessing = this.assignments
      .filter(assignment => assignment.associatedWithDeveloper?.id && assignmentsToProcessIds.has(assignment.associatedWithDeveloper.id))
      .map(assignment => ({...assignment, processingError: undefined}));

    if (assignmentsReadyForDriveProcessing.length === 0) {
      this.errorMessage = "No selected assignments found to process.";
      console.warn(`[Orchestrator] ${this.errorMessage}`);
      return;
    }

    this.assignments = this.assignments.map(a => {
      if (a.associatedWithDeveloper?.id && assignmentsToProcessIds.has(a.associatedWithDeveloper.id)) {
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

    // Stage 1: Process content.
    this.processAssignmentsAndCreateContent(courseName, assignmentsReadyForDriveProcessing, allPackageFilesForServices).pipe(
      tap(driveProcessingResults => {
        console.log('[Orchestrator] Content preparation stage (Drive, QTI, HTML) completed.');
        this.assignments = this.assignments.map(assignment => {
          const result = driveProcessingResults.find(r => r.itemId === assignment.associatedWithDeveloper?.id);
          if (result) {
            if (result.error) {
              console.warn(`[Orchestrator] Item "${assignment.title}" (ID: ${result.itemId}) failed content prep: ${result.error.message}`);
              return {...assignment, processingError: {...result.error}};
            }
            if (result.finalHtmlDescription !== undefined && result.finalPlainTextDescription !== undefined) {
              return {
                ...assignment,
                descriptionForDisplay: result.finalHtmlDescription,
                descriptionForClassroom: result.finalPlainTextDescription,
                richtext: !!result.finalHtmlDescription,
                processingError: undefined
              };
            }
          }
          return assignment;
        });
        this.changeDetectorRef.markForCheck();
        this.loadingMessage = 'Preparing items for Classroom submission...';
      }),
      switchMap(driveProcessingResults => {
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
          return of({type: 'skipped' as const, reason: message, data: [...this.assignments]});
        }

        const updatedAssignmentsForClassroom = this.addContentAsMaterials(driveProcessingResults, itemsForClassroom);
        console.log(`[Orchestrator] Assignments fully prepared for ClassroomService: ${updatedAssignmentsForClassroom.length} items.`);
        if (updatedAssignmentsForClassroom.length > 0 && updatedAssignmentsForClassroom[0]) {
          console.log('[Orchestrator] Example of first assignment payload for ClassroomService (materials & description):',
            JSON.stringify(updatedAssignmentsForClassroom[0].materials, null, 2),
            JSON.stringify(updatedAssignmentsForClassroom[0].descriptionForClassroom, null, 2));
        }

        this.loadingMessage = `Submitting ${updatedAssignmentsForClassroom.length} assignment(s) to ${selectedContent.classroomIds.length} classroom(s)... This may take some time.`;
        this.changeDetectorRef.markForCheck();

        // Stage 2: Call ClassroomService.
        return this.classroom.assignContentToClassrooms(selectedContent.classroomIds, updatedAssignmentsForClassroom).pipe(
          map(classroomServiceResults => {
            console.log('[Orchestrator] Received final results from ClassroomService.assignContentToClassrooms.');
            this.assignments = this.assignments.map(assignment => {
              const finalResult = classroomServiceResults.find(cr => cr.associatedWithDeveloper?.id === assignment.associatedWithDeveloper?.id);
              return finalResult ? finalResult : assignment;
            });
            return {type: 'success' as const, response: classroomServiceResults, data: [...this.assignments]};
          }),
          catchError(classroomPipelineError => {
            const message = classroomPipelineError?.message || 'Unknown error during Classroom submission pipeline.';
            console.error(`[Orchestrator] Error from ClassroomService.assignContentToClassrooms pipeline: ${message}`, classroomPipelineError);
            this.assignments = this.assignments.map(a => {
              if (updatedAssignmentsForClassroom.some(itemSent => itemSent.associatedWithDeveloper?.id === a.associatedWithDeveloper?.id)) {
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
      catchError((pipelineError: any) => {
        const source = pipelineError?.source || 'Content Preparation Pipeline';
        const message = pipelineError?.error?.message || pipelineError?.message || 'An unknown error in content preparation.';
        console.error(`[Orchestrator] Error in main processing pipeline (source: ${source}): ${message}`, pipelineError);
        this.errorMessage = `Error during ${source}: ${message}`;
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
        this.assignments = [...finalResult.data];

        if (finalResult.type === 'success') {
          const itemsProcessedByClassroom = finalResult.response;
          const successfulSubmissions = itemsProcessedByClassroom.filter(r => !r.processingError).length;
          const failedSubmissions = itemsProcessedByClassroom.filter(r => r.processingError).length;

          if (failedSubmissions > 0) {
            this.errorMessage = `Submission to Classroom completed with ${failedSubmissions} item(s) failing. Check item details.`;
            if (successfulSubmissions > 0) {
              this.successMessage = `Successfully submitted ${successfulSubmissions} item(s) to classroom(s).`;
            } else {
              this.successMessage = null;
            }
          } else if (successfulSubmissions > 0) {
            this.successMessage = `Successfully submitted ${successfulSubmissions} item(s) to classroom(s).`;
            this.errorMessage = null;
          } else {
            this.errorMessage = "No items were successfully submitted to Classroom. Check selection or previous errors.";
            this.successMessage = null;
          }
        } else if (finalResult.type === 'skipped') {
          this.errorMessage = `Processing skipped: ${finalResult.reason}`;
          this.successMessage = null;
        } else if (finalResult.type === 'error') {
          this.errorMessage = `Processing Error (${finalResult.source || 'Unknown Stage'}): ${finalResult.error.message}`;
          this.successMessage = null;
        }
        this.changeDetectorRef.markForCheck();
      },
      error: (err) => {
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

  processAssignmentsAndCreateContent(
    courseName: string,
    itemsToProcess: ProcessedCourseWork[],
    allPackageImsccFiles: ImsccFile[] // Full list of unzipped files for context if needed by services
  ): Observable<ProcessingResult[]> {
    console.log(`[Orchestrator] processAssignmentsAndCreateContent: Starting content preparation for ${itemsToProcess.length} items.`);
    if (itemsToProcess.length === 0) {
      return of([]);
    }

    return from(itemsToProcess).pipe(
      concatMap((item, index) => {
        const itemLogPrefix = `ContentPrep Item ${index + 1}/${itemsToProcess.length} ("${item.title?.substring(0, 30)}..."):`;
        this.loadingMessage = `${itemLogPrefix} Preparing...`;
        this.changeDetectorRef.markForCheck();
        console.log(`${itemLogPrefix} Beginning content preparation.`);

        const assignmentName = item.title || 'Untitled Assignment';
        const driveTopicFolderName = item.associatedWithDeveloper?.topic || 'General';
        const itemId = item.associatedWithDeveloper?.id;

        if (!itemId) {
          console.error(`${itemLogPrefix} Critical error: Item is missing associatedDeveloper.id. Skipping content prep.`);
          return of({
            itemId: undefined, assignmentName, topicName: driveTopicFolderName,
            assignmentFolderId: 'ERROR_NO_ID', error: {message: 'Missing item ID for content prep', stage: 'Pre-flight (ContentPrep)'}
          } as ProcessingResult);
        }

        const qtiFileForService: ImsccFile | undefined = item.qtiFile?.[0];
        // Check if a main Doc should be created from the item's HTML description
        const shouldCreateDocFromDescription = !qtiFileForService && !!item.descriptionForDisplay && item.workType === 'ASSIGNMENT' && !!item.richtext;

        // This list will be used by replaceLocalLinksInHtml to understand original file references.
        const allOriginalLocalFilesInfo = (item.localFilesToUpload || []).map(ftu => {
          const originalUnzippedFile = this.unzippedFiles.find(uzf => uzf.name === ftu.file.name);
          return {
            ...ftu,
            file: { // Ensure ImsccFile structure with data
              name: ftu.file.name,
              mimeType: originalUnzippedFile?.mimeType || this.util.getMimeTypeFromExtension(ftu.file.name),
              data: originalUnzippedFile?.data || '' // Provide data from unzippedFiles
            }
          };
        });


        return this.drive.ensureAssignmentFolderStructure(courseName, driveTopicFolderName, assignmentName, itemId).pipe(
          switchMap(assignmentFolderId => {
            this.loadingMessage = `${itemLogPrefix} Folder ensured. Processing files...`;
            this.changeDetectorRef.markForCheck();

            // 1. Categorize local files from item.localFilesToUpload
            const htmlFileAssetsToConvert: Array<{file: ImsccFile; targetFileName: string}> = [];
            const otherLocalFilesToUploadInfo: Array<{file: ImsccFile; targetFileName: string}> = [];

            allOriginalLocalFilesInfo.forEach(ftu => {
              // Ensure ftu.file.data is populated from unzippedFiles if it was just a reference
              const originalUnzippedFile = this.unzippedFiles.find(uzf => uzf.name === ftu.file.name);
              const fileWithData = {
                ...ftu,
                file: {
                  ...ftu.file,
                  data: originalUnzippedFile?.data || ftu.file.data, // Prioritize fresh unzipped data
                  mimeType: originalUnzippedFile?.mimeType || ftu.file.mimeType
                }
              };

              if (fileWithData.file.mimeType === 'text/html' ||
                fileWithData.file.name.toLowerCase().endsWith('.html') ||
                fileWithData.file.name.toLowerCase().endsWith('.htm')) {
                htmlFileAssetsToConvert.push(fileWithData);
              } else {
                otherLocalFilesToUploadInfo.push(fileWithData);
              }
            });

            console.log(`${itemLogPrefix} HTML files to convert: ${htmlFileAssetsToConvert.length}, Other files to upload: ${otherLocalFilesToUploadInfo.length}`);

            // 2. Observable for uploading other (non-HTML) local files
            const uploadedOtherFiles$: Observable<DriveFile[]> = (otherLocalFilesToUploadInfo.length > 0 ?
              this.files.uploadLocalFiles(otherLocalFilesToUploadInfo, assignmentFolderId) : // Uses the enriched otherLocalFilesToUploadInfo
              of([])
            ).pipe(
              tap(uploaded => console.log(`${itemLogPrefix} ${uploaded.length} other local files uploaded.`)),
              catchError(uploadErr => {
                console.error(`${itemLogPrefix} Error uploading other files:`, uploadErr);
                return throwError(() => ({message: `Other file upload failed: ${uploadErr.message}`, stage: 'Other File Upload', details: uploadErr, itemId}));
              }),
              shareReplay(1)
            );

            // 3. Observable for converting HTML file assets to Google Docs
            const convertedHtmlDocs$: Observable<DriveFile[]> = htmlFileAssetsToConvert.length > 0 ?
              forkJoin(
                htmlFileAssetsToConvert.map(htmlFtu => {
                  const htmlContent = typeof htmlFtu.file.data === 'string' ? htmlFtu.file.data : '';
                  if (!htmlContent && typeof htmlFtu.file.data !== 'string') {
                    console.warn(`${itemLogPrefix} HTML file ${htmlFtu.file.name} has non-string data, attempting to decode if ArrayBuffer.`);
                    // This case should ideally be handled by unzipAndConvert to always yield string for text types
                  }
                  if (!htmlContent) {
                    console.warn(`${itemLogPrefix} HTML file ${htmlFtu.file.name} has no string content. Skipping conversion.`);
                    return of(null); // Will be filtered out
                  }
                  // Use targetFileName for the Doc title, as replaceLocalLinksInHtml expects this for matching.
                  const docTitle = htmlFtu.targetFileName || this.util.getFileNameWithoutExtension(this.util.getBasename(htmlFtu.file.name)) || 'Converted HTML Document';

                  console.log(`${itemLogPrefix} Converting HTML asset "${htmlFtu.file.name}" to Doc titled "${docTitle}".`);
                  return this.docs.createDocFromHtml(htmlContent, docTitle, itemId, assignmentFolderId).pipe(
                    catchError(docErr => {
                      console.error(`${itemLogPrefix} Error converting HTML file ${htmlFtu.file.name} to Doc:`, docErr);
                      return of(null); // Individual error, return null to filter out later
                    })
                  );
                })
              ).pipe(
                map(results => results.filter(r => r !== null) as DriveFile[]), // Filter out failed conversions
                tap(docs => console.log(`${itemLogPrefix} ${docs.length} HTML file assets converted to Docs.`)),
                catchError(htmlConversionErr => {
                  console.error(`${itemLogPrefix} Error during batch HTML asset to Doc conversion:`, htmlConversionErr);
                  return throwError(() => ({message: `Batch HTML to Docs failed: ${htmlConversionErr.message}`, stage: 'HTML Assets to Docs Batch', details: htmlConversionErr, itemId}));
                })
              ) :
              of([]); // No HTML files to convert

            // First, process all file uploads and HTML-to-Doc conversions for assets
            return forkJoin({
              uploadedOthers: uploadedOtherFiles$,
              convertedHtmls: convertedHtmlDocs$
            }).pipe(
              switchMap(({uploadedOthers, convertedHtmls}) => {
                this.loadingMessage = `${itemLogPrefix} Files processed. Handling description & primary content...`;
                this.changeDetectorRef.markForCheck();

                const allSupportingDriveFiles: DriveFile[] = [...uploadedOthers, ...convertedHtmls];
                console.log(`${itemLogPrefix} Total supporting Drive files (uploaded + converted HTML assets): ${allSupportingDriveFiles.length}`);
                allSupportingDriveFiles.forEach(df => console.log(`  - ${df.name} (ID: ${df.id})`));


                let currentHtmlDescriptionForProcessing = item.descriptionForDisplay;

                if (currentHtmlDescriptionForProcessing) {
                  console.log(`${itemLogPrefix} Replacing local links in main HTML description using all supporting files.`);
                  currentHtmlDescriptionForProcessing = this.replaceLocalLinksInHtml(
                    currentHtmlDescriptionForProcessing,
                    allSupportingDriveFiles, // Pass the combined list
                    allOriginalLocalFilesInfo, // Use the comprehensive list for matching original paths
                    itemLogPrefix
                  );
                }

            // 4. Primary Content Creation (QTI or description-based Doc)
                let primaryContentCreation$: Observable<DriveFile | Material | null>;

                if (qtiFileForService) {
                  console.log(`${itemLogPrefix} Creating QTI Form.`);
                  primaryContentCreation$ = this.qti.createFormFromQti(qtiFileForService, allPackageImsccFiles, assignmentName, itemId, assignmentFolderId)
                    .pipe(
                      tap(form => console.log(`${itemLogPrefix} QTI Form created/found.`)),
                      catchError(qtiErr => throwError(() => ({message: `QTI to Form failed: ${qtiErr.message}`, stage: 'QTI Processing', details: qtiErr, itemId})))
                    );
                } else if (shouldCreateDocFromDescription) {
                  console.log(`${itemLogPrefix} Creating Doc from main HTML description (after link replacement).`);
                  primaryContentCreation$ = this.docs.createDocFromHtml(currentHtmlDescriptionForProcessing || '', assignmentName, itemId, assignmentFolderId)
                    .pipe(
                      tap(doc => console.log(`${itemLogPrefix} Main HTML Doc created/found.`)),
                      catchError(docErr => throwError(() => ({message: `HTML to Doc (description) failed: ${docErr.message}`, stage: 'HTML to Doc (Description)', details: docErr, itemId})))
                    );
                } else {
                  console.log(`${itemLogPrefix} No primary QTI/Doc content to create from description. Description (if any) already processed for links.`);
                  primaryContentCreation$ = of(null);
                }

                // This inner forkJoin is just to structure the return if primaryContentCreation$ is also an observable
                // If primaryContentCreation$ is already resolved (e.g. of(null)), this map is simpler.
                return primaryContentCreation$.pipe(
                  map(createdPrimaryContentResult => {
                    const finalHtmlDesc = currentHtmlDescriptionForProcessing || '';
                    const finalPlainTextDesc = this.generatePlainText(finalHtmlDesc);

                      console.log(`${itemLogPrefix} Content preparation for item completed successfully.`);
                      return {
                        itemId, assignmentName, topicName: driveTopicFolderName, assignmentFolderId,
                          createdDoc: (createdPrimaryContentResult && 'mimeType' in createdPrimaryContentResult && (createdPrimaryContentResult as DriveFile).mimeType !== 'application/vnd.google-apps.form') ? createdPrimaryContentResult as DriveFile : undefined,
                          createdForm: (createdPrimaryContentResult && 'form' in createdPrimaryContentResult) ? createdPrimaryContentResult as Material : undefined,
                          uploadedFiles: allSupportingDriveFiles, // CRITICAL: This now includes all processed local files
                          error: null,
                          finalHtmlDescription: finalHtmlDesc,
                          finalPlainTextDescription: finalPlainTextDesc
                        } as ProcessingResult;
                    })
                );
              })
            );
          }),
          catchError((errorFromStage: any) => {
            const stage = errorFromStage?.stage || 'Content Preparation Sub-Stage';
            const message = errorFromStage?.message || 'Unknown error during content prep sub-stage.';
            console.error(`${itemLogPrefix} Error during content prep for item. Stage: ${stage}, Msg: ${message}`, errorFromStage.details);
            const originalDescription = item.descriptionForDisplay;
            return of({
              itemId, assignmentName, topicName: driveTopicFolderName,
              assignmentFolderId: errorFromStage?.assignmentFolderId || 'ERROR_CONTENT_PREP',
              error: {message, stage, details: errorFromStage.details || errorFromStage},
              finalHtmlDescription: originalDescription,
              finalPlainTextDescription: this.generatePlainText(originalDescription)
            } as ProcessingResult);
          })
        );
      }),
      toArray()
    );
  }

  private replaceLocalLinksInHtml(
    htmlContent: string,
    uploadedDriveFiles: DriveFile[], // This now contains ALL Drive files (uploaded + converted HTMLs)
    filesOriginallyQueuedForUpload: Array<{file: ImsccFile; targetFileName: string}>, // This is the original list from item.localFilesToUpload, used for context
    itemLogPrefix: string
  ): string {
    if (!htmlContent) return '';
    if (typeof DOMParser === 'undefined') {
      console.warn(`${itemLogPrefix} DOMParser not available. Skipping HTML link replacement.`);
      return htmlContent;
    }
    try {
      const parser = new DOMParser();
      const htmlDoc = parser.parseFromString(htmlContent, 'text/html');
      const body = htmlDoc.body || htmlDoc.documentElement;

      body.querySelectorAll('a[data-imscc-original-path]').forEach((anchor: Element) => {
        const htmlAnchor = anchor as HTMLAnchorElement;
        const originalPathAttribute = htmlAnchor.getAttribute('data-imscc-original-path');
        const displayTitleAttr = decode(htmlAnchor.getAttribute('data-imscc-display-title') || htmlAnchor.textContent || originalPathAttribute || 'Linked File');

        if (originalPathAttribute) {
          let matchedDriveFile: DriveFile | undefined;
          const normalizedOriginalPathAttr = normalizeSpacesString(this.util.tryDecodeURIComponent(originalPathAttribute).toLowerCase());

          // Find the original file info from the queued list that corresponds to this link's original path
          const originalQueuedFileInfo = filesOriginallyQueuedForUpload.find(ftu => {
            const decodedFullZipPath = this.util.tryDecodeURIComponent(ftu.file.name);
            const normalizedFtuNameInZip = normalizeSpacesString(decodedFullZipPath.toLowerCase());
            // Match if the full path in zip ends with the original path from HTML attribute, or vice-versa.
            // Or, if originalPathAttribute is just the basename.
            const ftuBasename = this.util.getBasename(normalizedFtuNameInZip).toLowerCase();
            return normalizedFtuNameInZip.endsWith(normalizedOriginalPathAttr) ||
              normalizedOriginalPathAttr.endsWith(normalizedFtuNameInZip) ||
              ftuBasename === normalizedOriginalPathAttr;
          });

          if (originalQueuedFileInfo) {
            // Now, find this file in the uploadedDriveFiles list.
            // The DriveFile name should match originalQueuedFileInfo.targetFileName (which was used as title for HTML-to-Doc)
            const targetFileNameToMatchLower = normalizeSpacesString(originalQueuedFileInfo.targetFileName.toLowerCase());
            matchedDriveFile = uploadedDriveFiles.find(udf =>
              normalizeSpacesString(udf.name?.toLowerCase() || '') === targetFileNameToMatchLower
            );
            if (!matchedDriveFile) { // Fallback: try matching against basename without extension if targetFileName included .html
              const targetBasenameNoExtLower = normalizeSpacesString(this.util.getFileNameWithoutExtension(this.util.getBasename(originalQueuedFileInfo.targetFileName)).toLowerCase());
              if (targetBasenameNoExtLower) {
                matchedDriveFile = uploadedDriveFiles.find(udf =>
                  normalizeSpacesString(this.util.getFileNameWithoutExtension(udf.name)?.toLowerCase() || '') === targetBasenameNoExtLower
                );
              }
            }
          }

          if (matchedDriveFile?.webViewLink) {
            htmlAnchor.href = matchedDriveFile.webViewLink;
            htmlAnchor.target = '_blank';
            htmlAnchor.rel = 'noopener noreferrer';
            htmlAnchor.textContent = displayTitleAttr;
            console.log(`${itemLogPrefix} Updated anchor for "${originalPathAttribute}" (Original Target: ${originalQueuedFileInfo?.targetFileName}) to Drive link: ${matchedDriveFile.webViewLink} (Matched Drive File: ${matchedDriveFile.name}).`);
          } else {
            console.warn(`${itemLogPrefix} No Drive file found for local link anchor: "${originalPathAttribute}" (Original Target: ${originalQueuedFileInfo?.targetFileName}). Original text: "${displayTitleAttr}"`);
            htmlAnchor.textContent = `[Broken Link: ${displayTitleAttr}]`;
            htmlAnchor.removeAttribute('href');
          }
          htmlAnchor.removeAttribute('data-imscc-original-path');
          htmlAnchor.removeAttribute('data-imscc-media-type');
          htmlAnchor.removeAttribute('data-imscc-display-title');
        }
      });

      body.querySelectorAll('span.imscc-local-file-placeholder[data-original-path]').forEach((span: Element) => {
        const htmlSpan = span as HTMLSpanElement;
        const originalPathAttribute = htmlSpan.getAttribute('data-original-path');
        const displayDefaultText = decode(htmlSpan.textContent || originalPathAttribute || 'Linked Content');

        if (originalPathAttribute) {
          let matchedDriveFile: DriveFile | undefined;
          const normalizedOriginalPathAttr = normalizeSpacesString(this.util.tryDecodeURIComponent(originalPathAttribute).toLowerCase());

          const originalQueuedFileInfo = filesOriginallyQueuedForUpload.find(ftu => {
            const decodedFullZipPath = this.util.tryDecodeURIComponent(ftu.file.name);
            const normalizedFtuNameInZip = normalizeSpacesString(decodedFullZipPath.toLowerCase());
            const ftuBasename = this.util.getBasename(normalizedFtuNameInZip).toLowerCase();
            return normalizedFtuNameInZip.endsWith(normalizedOriginalPathAttr) ||
              normalizedOriginalPathAttr.endsWith(normalizedFtuNameInZip) ||
              ftuBasename === normalizedOriginalPathAttr;
          });

          if (originalQueuedFileInfo) {
            const targetFileNameToMatchLower = normalizeSpacesString(originalQueuedFileInfo.targetFileName.toLowerCase());
            matchedDriveFile = uploadedDriveFiles.find(udf =>
              normalizeSpacesString(udf.name?.toLowerCase() || '') === targetFileNameToMatchLower
            );
            if (!matchedDriveFile) {
              const targetBasenameNoExtLower = normalizeSpacesString(this.util.getFileNameWithoutExtension(this.util.getBasename(originalQueuedFileInfo.targetFileName)).toLowerCase());
              if (targetBasenameNoExtLower) {
                matchedDriveFile = uploadedDriveFiles.find(udf =>
                  normalizeSpacesString(this.util.getFileNameWithoutExtension(udf.name)?.toLowerCase() || '') === targetBasenameNoExtLower
                );
              }
            }
          }

          if (matchedDriveFile?.webViewLink) {
            const newAnchor = htmlDoc.createElement('a');
            newAnchor.href = matchedDriveFile.webViewLink;
            newAnchor.target = '_blank';
            newAnchor.rel = 'noopener noreferrer';
            newAnchor.textContent = displayDefaultText;
            span.parentNode?.replaceChild(newAnchor, span);
            console.log(`${itemLogPrefix} Replaced SPAN placeholder for "${originalPathAttribute}" (Original Target: ${originalQueuedFileInfo?.targetFileName}) with Drive link (Matched Drive File: ${matchedDriveFile.name}).`);
          } else {
            console.warn(`${itemLogPrefix} No Drive file found for SPAN placeholder: "${originalPathAttribute}" (Original Target: ${originalQueuedFileInfo?.targetFileName}).`);
            htmlSpan.textContent = `[Broken Content: ${displayDefaultText}]`;
            htmlSpan.style.color = "red";
          }
        }
      });
      return htmlDoc.body.innerHTML;
    } catch (e) {
      console.error(`${itemLogPrefix} Error during HTML DOM manipulation for link replacement:`, e);
      return htmlContent;
    }
  }

  private generatePlainText(html: string | undefined | null): string {
    if (!html) return '';
    if (typeof document === 'undefined') {
      console.warn("`document` is not available. Cannot generate plain text from HTML.");
      return '';
    }
    try {
      const tempDiv = document.createElement('div');
      tempDiv.innerHTML = html;
      let text = (tempDiv.textContent || tempDiv.innerText || '').replace(/\s\s+/g, ' ').trim();
      const maxLen = 30000;
      if (text.length > maxLen) {
        text = text.substring(0, maxLen - 3) + "...";
        console.warn('[Orchestrator] Plain text description was truncated to fit Classroom API limits.');
      }
      return text;
    } catch (e) {
      console.error("[Orchestrator] Error generating plain text from HTML:", e);
      return "";
    }
  }

  private getMaterialKey(material: Material): string | null {
    if (material.driveFile?.driveFile?.id) return `drive-${material.driveFile.driveFile.id}`;
    if (material.youtubeVideo?.id) return `youtube-${material.youtubeVideo.id}`;
    if (material.link?.url) return `link-${this.util.tryDecodeURIComponent(material.link.url).toLowerCase()}`;
    if (material.form?.formUrl) return `form-${this.util.tryDecodeURIComponent(material.form.formUrl).toLowerCase()}`;
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
        console.warn('[Orchestrator] deduplicateMaterials: Encountered material with no identifiable key. Including as is.', material);
        uniqueMaterials.push(material);
      }
    }
    return uniqueMaterials;
  }

  addContentAsMaterials(
    driveProcessingResults: ProcessingResult[],
    itemsReadyForClassroom: ProcessedCourseWork[]
  ): ProcessedCourseWork[] {
    return itemsReadyForClassroom.map(item => {
      const updatedItem = {...item, materials: [...(item.materials || [])]};
      const result = driveProcessingResults.find(r => r.itemId === updatedItem.associatedWithDeveloper?.id);

      if (result && !result.error) {
        // Add the primary Google Doc (from HTML description) if it was created
        if (result.createdDoc?.id && result.createdDoc?.name) {
          updatedItem.materials.push({
            driveFile: {driveFile: {id: result.createdDoc.id, title: result.createdDoc.name}, shareMode: 'STUDENT_COPY'}
          });
        }
        // Add the Google Form (from QTI) if it was created
        if (result.createdForm?.form?.formUrl) {
          updatedItem.materials.push({
            link: {url: result.createdForm.form.formUrl, title: result.createdForm.form.title || result.assignmentName}
          });
        }

        // Add all other uploaded/converted files as materials
        // result.uploadedFiles now contains both directly uploaded files AND HTML files converted to Docs
        (result.uploadedFiles || []).forEach(uploadedFile => {
          if (uploadedFile?.id && uploadedFile?.name) {
            // Avoid re-adding the main description doc if it's also in uploadedFiles
            // (though current logic should keep them separate: createdDoc vs. uploadedFiles)
            if (!(result.createdDoc?.id === uploadedFile.id)) {
              updatedItem.materials.push({
                driveFile: {driveFile: {id: uploadedFile.id, title: uploadedFile.name}, shareMode: 'VIEW'}
              });
            }
          }
        });
        updatedItem.materials = this.deduplicateMaterials(updatedItem.materials);
      }
      return updatedItem;
    });
  }
}
