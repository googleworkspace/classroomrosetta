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

import { Component, inject, ViewChild, ChangeDetectorRef } from '@angular/core'; // Import ChangeDetectorRef
import JSZip from 'jszip';
import { Observable, from, of, throwError, forkJoin, EMPTY, Subject } from 'rxjs'; // Added Subject
import { map, switchMap, concatMap, catchError, tap, finalize, takeUntil,toArray } from 'rxjs/operators'; // Added takeUntil
import { CommonModule } from '@angular/common'; // Import CommonModule
import { MatButtonModule } from '@angular/material/button';
import { MatInputModule } from '@angular/material/input';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatIconModule } from '@angular/material/icon';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner'; // Import progress spinner
import { ClassroomService } from '../services/classroom/classroom.service';
import { ConverterService } from '../services/converter/converter.service';
import { DriveFolderService } from '../services/drive/drive.service';
// --- Import the refactored services ---
import {HtmlToDocsService} from '../services/html-to-docs/html-to-docs.service'; // Adjust path
import {FileUploadService} from '../services/file-upload/file-upload.service'; // Adjust path
import {QtiToFormsService} from '../services/qti-to-forms/qti-to-forms.service'; // Adjust path
import {UtilitiesService} from '../services/utilities/utilities.service';
// --- End refactored service imports ---

import { AuthService } from '../services/auth/auth.service';
import {
  ProcessedCourseWork,
  ProcessingResult,
  SubmissionData,
  Material,
  ImsccFile, // Ensure this is the correct interface used by QtiToFormsService
  DriveFile
} from '../interfaces/classroom-interface'; // Adjust path

// Import CourseworkDisplayComponent - Assuming it's standalone and path is correct
import { CourseworkDisplayComponent } from '../coursework-display/coursework-display.component';


// Helper function to escape special characters for use in RegExp
function escapeRegExp(string: string): string {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& means the whole matched string
}

// Interface for the structure of unzipped files within the component
interface UnzippedFile {
  name: string;
  data: string | ArrayBuffer; // This is what unzippedFiles will hold initially
  mimeType: string;
}

// Define a more structured error for ProcessingResult to align with ProcessedCourseWork
interface StagedProcessingError {
    message: string;
    stage?: 'Folder Creation' | 'File Upload' | 'Document Creation' | 'Form Creation' | 'Unknown';
    details?: string;
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
  // --- Component Properties ---
  selectedFile: File | null = null;
  unzippedFiles: UnzippedFile[] = []; // Holds files with data as string OR ArrayBuffer
  @ViewChild('fileInput') fileInput: any = null;
  assignments: ProcessedCourseWork[] = [];
  isProcessing: boolean = false;
  loadingMessage: string = '';
  errorMessage: string | null = null;
  successMessage: string | null = null;

  // --- Service Injection ---
  classroom = inject(ClassroomService);
  converter = inject(ConverterService);
  drive = inject(DriveFolderService);
  docs = inject(HtmlToDocsService);
  files = inject(FileUploadService);
  qti = inject(QtiToFormsService);
  auth = inject(AuthService);
  util = inject(UtilitiesService)

  private changeDetectorRef = inject(ChangeDetectorRef);


  // --- Component Methods ---

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
      this.unzippedFiles = []; // Reset unzippedFiles
      this.changeDetectorRef.markForCheck();
      try {
        await this.unzipAndConvert(this.selectedFile);
      } catch (error: any) {
        console.error('Error during file upload/unzip trigger:', error);
        this.errorMessage = `Error during setup: ${error?.message || error}`;
        this.isProcessing = false;
        this.loadingMessage = '';
        this.changeDetectorRef.markForCheck();
      }
    } else if (this.isProcessing) {
        console.warn('Processing already in progress.');
    } else {
      console.warn('No file selected for upload.');
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
              data = `data:${mimeType};base64,${base64Data}`; // Data is now a base64 string
            } else {
              data = await zipEntry.async('arraybuffer');
            }
            return { name: relativePath, data: data, mimeType: mimeType };
          })();
          filePromises.push(promise);
        }
      });

      const resolvedFiles = await Promise.all(filePromises);
      // this.unzippedFiles will now have image data as base64 strings, text as strings, others as ArrayBuffer
      this.unzippedFiles = resolvedFiles.filter(f => f !== null) as UnzippedFile[];
      console.log(`Unzipped files prepared: ${this.unzippedFiles.length} files.`, this.unzippedFiles);

      // Prepare files specifically for the ConverterService
      // ConverterService expects ImsccFile where data is string.
      // For images, this will be the base64 data URI. For other binary, it might need conversion or special handling in ConverterService.
      const imsccFilesForConverter: ImsccFile[] = this.unzippedFiles.map(uf => {
        let dataForConverter: string;
        if (typeof uf.data === 'string') {
          dataForConverter = uf.data; // Already string (text, XML, or base64 image from initial unzipping)
        } else if (uf.data instanceof ArrayBuffer) {
          // This case is for binary files that are NOT images and NOT text.
          // ConverterService might need to handle ArrayBuffer or expect a placeholder/different format.
          // For now, let's assume it expects a string, so we'll create a placeholder or attempt base64.
          // This part might need refinement based on how ConverterService handles generic binary files.
          console.warn(`File "${uf.name}" (mime: ${uf.mimeType}) is ArrayBuffer. ConverterService expects string. Attempting base64 for generic binary data.`);
          try {
            // Create a raw base64 string, not a data URI, as ConverterService might not expect data URI for non-image binary files.
            dataForConverter = this.util.arrayBufferToBase64(uf.data);
          }
          catch (e) {
            console.error(`Error converting ArrayBuffer for ${uf.name} to base64 for ConverterService`, e);
            dataForConverter = `[Binary data for ${uf.name} - conversion failed]`; // Placeholder
          }
        } else {
          console.error(`Unexpected data type for file ${uf.name} when preparing for ConverterService.`);
          dataForConverter = `[Error: Unexpected data type for ${uf.name}]`;
        }
        return { name: uf.name, data: dataForConverter, mimeType: uf.mimeType };
      });

      this.loadingMessage = 'Converting course structure...';
      this.changeDetectorRef.markForCheck();

      this.converter.convertImscc(imsccFilesForConverter)
        .pipe(
          finalize(() => {
            console.log('IMSCC conversion stream finalize.');
            if (!this.errorMessage) { // Only clear loading/set success if no error occurred during conversion
                 this.isProcessing = false;
                 this.loadingMessage = '';
                 this.successMessage = `Conversion complete. Found ${accumulatedAssignments.length} items. Ready for submission.`;
            }
            this.assignments = [...accumulatedAssignments];
            console.log('Final assignments structure after conversion:', this.assignments);
            this.changeDetectorRef.markForCheck();
          })
        )
        .subscribe({
          next: (assignment) => {
            accumulatedAssignments.push(assignment);
          },
          error: (error) => {
            console.error('Error during IMSCC conversion stream:', error);
            this.errorMessage = `Conversion Error: ${error?.message || error}`;
            this.isProcessing = false; // Ensure processing stops on error
            this.loadingMessage = '';
            this.changeDetectorRef.markForCheck();
          }
        });

    } catch (error: any) {
      console.error('Error during the unzipping or IMSCC conversion process:', error);
      this.errorMessage = `Unzip/Read Error: ${error?.message || error}`;
      this.assignments = [];
      this.isProcessing = false;
      this.loadingMessage = '';
      this.changeDetectorRef.markForCheck();
    }
  }

  process(selectedContent: SubmissionData) {
    if (this.isProcessing) {
        console.warn("Cannot start processing assignments while another operation is in progress.");
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

    console.log('Starting process function with selected content:', selectedContent);
    const token = this.auth.getGoogleAccessToken();

    if (!token) {
      this.errorMessage = "Processing aborted: No Google access token available. Please log in.";
      console.error(this.errorMessage);
      return;
    }

    let assignmentsToProcess = this.assignments.filter(assignment =>
      assignment.associatedWithDeveloper?.id &&
      selectedContent.assignmentIds.includes(assignment.associatedWithDeveloper.id)
    );

    if (assignmentsToProcess.length === 0) {
      console.warn("Processing skipped: No assignments selected or found matching the provided IDs.");
      this.errorMessage = "Could not find selected assignments to process.";
      return;
    }

    // Clear previous errors for items about to be processed
    const assignmentsToProcessIds = new Set(assignmentsToProcess.map(a => a.associatedWithDeveloper?.id));
    this.assignments = this.assignments.map(a => {
        if (a.associatedWithDeveloper?.id && assignmentsToProcessIds.has(a.associatedWithDeveloper.id)) {
          return {...a, processingError: undefined}; // Clear previous error
        }
        return a;
    });
    // Re-filter to ensure we have the cleaned items
    assignmentsToProcess = this.assignments.filter(assignment =>
      assignment.associatedWithDeveloper?.id &&
      selectedContent.assignmentIds.includes(assignment.associatedWithDeveloper.id)
    );


    const courseName = this.converter.coursename || 'Untitled Course';
    console.log(`Processing ${assignmentsToProcess.length} selected assignments for course: "${courseName}"`);
    console.log(`Target Classroom IDs: ${selectedContent.classroomIds.join(', ')}`);

    this.isProcessing = true;
    this.loadingMessage = `Processing ${assignmentsToProcess.length} assignment(s)... (Step 1: Drive Content)`;
    this.errorMessage = null;
    this.successMessage = null;
    this.changeDetectorRef.markForCheck();

    // Prepare allPackageFilesForServices for QtiToFormsService and other downstream services
    // This list should contain all files from the package, with data in string format where appropriate (e.g., base64 for images).
    const allPackageFilesForServices: ImsccFile[] = this.unzippedFiles.map(uf => {
      if (typeof uf.data === 'string') { // This includes base64 data URIs for images and text for XML/HTML
            return { name: uf.name, data: uf.data, mimeType: uf.mimeType };
        } else { // It's ArrayBuffer (for non-image, non-text binary files not converted to base64 in initial unzip)
          // QtiToFormsService uses this list primarily for path resolution of resources.
          // Images it needs will already be base64 strings in this.unzippedFiles.
          // For other binary files, a placeholder string for 'data' is fine.
          console.warn(`File "${uf.name}" (mime: ${uf.mimeType}) is ArrayBuffer. Providing placeholder data string for QtiToFormsService's allPackageFiles.`);
          return {name: uf.name, data: `[Binary data placeholder for ${uf.name}]`, mimeType: uf.mimeType};
        }
    });


    this.processAssignmentsAndCreateContent(courseName, assignmentsToProcess, token, allPackageFilesForServices).pipe(
      tap(driveProcessingResults => {
        console.log('Drive processing results (Uploads/Doc/Form Creation):', driveProcessingResults);
        // Update assignments with any errors from Drive processing
        this.assignments = this.assignments.map(assignment => {
          const result = driveProcessingResults.find(r => r.itemId === assignment.associatedWithDeveloper?.id);
          if (result && result.error) {
            const errorDetails = result.error as StagedProcessingError; // Cast to ensure type
            return {
              ...assignment,
              processingError: {
                message: errorDetails.message || 'Drive operation failed.',
                stage: errorDetails.stage || 'Drive Operation',
                details: errorDetails.details
              }
            };
          }
          return assignment;
        });
        this.changeDetectorRef.markForCheck();

        const itemErrors = driveProcessingResults.filter(r => !!r.error);
        if (itemErrors.length > 0) {
            console.warn(`Drive processing completed with ${itemErrors.length} item-level errors.`);
          // Potentially update UI to reflect partial success or specific item failures
        }
        this.loadingMessage = 'Associating content and submitting to Classroom...';
        this.changeDetectorRef.markForCheck();
      }),
      switchMap(driveProcessingResults => {
        // Filter out assignments that had errors during Drive processing
        const itemsForClassroom = this.assignments.filter(assignment =>
            selectedContent.assignmentIds.includes(assignment.associatedWithDeveloper?.id || '') &&
          !assignment.processingError // Only process items without prior errors
        );

        if (itemsForClassroom.length === 0 && assignmentsToProcess.length > 0) {
            this.errorMessage = assignmentsToProcess.length > 0 ? "All selected items failed during Drive content creation/upload." : "No items were selected for Classroom submission.";
            this.changeDetectorRef.markForCheck();
          // Use 'of' to return an observable that completes the stream but indicates error
            return of({ type: 'error' as const, source: 'Drive Processing', error: new Error(this.errorMessage), data: [...this.assignments] });
        }
        if (itemsForClassroom.length === 0) { // No items to process (e.g., none selected or all failed)
            return of({ type: 'skipped' as const, reason: 'No items to submit to Classroom', data: [...this.assignments] });
        }

        // Add created Drive content (Docs, Forms, Files) as materials to the assignments
        const updatedAssignmentsForClassroom = this.addContentAsMaterials(driveProcessingResults, itemsForClassroom);
        console.log('Assignments prepared for Classroom:', updatedAssignmentsForClassroom);

        if (!selectedContent.classroomIds || selectedContent.classroomIds.length === 0) {
          // This case should ideally be caught earlier, but good to have a safe return
          return of({ type: 'skipped' as const, reason: 'No classrooms selected', data: [...this.assignments] });
        }

        this.loadingMessage = `Submitting ${updatedAssignmentsForClassroom.length} assignment(s) to ${selectedContent.classroomIds.length} classroom(s)...`;
        this.changeDetectorRef.markForCheck();

        return this.classroom.assignContentToClassrooms(token, selectedContent.classroomIds, updatedAssignmentsForClassroom).pipe(
          map(classroomResults => {
            // Update this.assignments with the results from Classroom (e.g., links, classroom IDs)
            this.assignments = this.assignments.map(assignment => {
              const classroomResult = classroomResults.find(cr => cr.associatedWithDeveloper?.id === assignment.associatedWithDeveloper?.id);
              return classroomResult || assignment; // If a result exists, use it, otherwise keep original
            });
            return { type: 'success' as const, response: classroomResults, data: [...this.assignments] };
          }),
          catchError(classroomError => {
            const message = classroomError?.message || 'Unknown Classroom API error';
            // Mark items that were attempted for Classroom push as errored
            this.assignments = this.assignments.map(a => {
                if (itemsForClassroom.find(i => i.associatedWithDeveloper?.id === a.associatedWithDeveloper?.id)) {
                  // Only add Classroom error if no Drive error existed
                    if (!a.processingError) {
                        a.processingError = { message: `Classroom Batch Error: ${message}`, stage: 'Classroom Push' };
                    }
                }
                return a;
            });
            return throwError(() => ({ type: 'error' as const, source: 'assignContentToClassrooms', error: new Error(message), data: [...this.assignments] }));
          })
        );
      }),
      catchError((pipelineError: any) => { // Catch errors from the switchMap or earlier stages
        const source = pipelineError?.source || 'pipeline';
        const message = pipelineError?.error?.message || pipelineError?.message || 'An unknown error occurred.';
        this.errorMessage = `Error during ${source}: ${message}`;
        this.assignments = pipelineError.data || this.assignments; // Ensure assignments state is updated if error object carries it
        return of({ type: 'error' as const, source: source, error: new Error(message), data: [...this.assignments] });
      }),
      finalize(() => {
          console.log("Drive/Classroom processing pipeline finished.");
          this.isProcessing = false;
          this.loadingMessage = '';
          this.changeDetectorRef.markForCheck();
      })
    ).subscribe({
      next: (finalResult) => {
        // Update assignments with the final data from the stream
        this.assignments = [...finalResult.data];

        if (finalResult.type === 'success') {
          const itemsWithErrors = this.assignments.filter(a => a.processingError).length;
          if (itemsWithErrors > 0) {
            this.errorMessage = `Submission completed with ${itemsWithErrors} item(s) failing. Check item details.`;
            this.successMessage = `Successfully processed ${this.assignments.length - itemsWithErrors} out of ${selectedContent.assignmentIds.length} selected item(s).`;
          } else {
            this.successMessage = `Successfully submitted ${this.assignments.length} item(s) to classroom(s).`;
            this.errorMessage = null;
          }
        } else if (finalResult.type === 'skipped') {
          this.errorMessage = `Submission skipped: ${finalResult.reason}`;
          this.successMessage = null;
        } else if (finalResult.type === 'error') { // This is for errors caught by the outer catchError
           this.errorMessage = `Error during ${finalResult.source}: ${finalResult.error.message || 'Unknown error'}`;
           this.successMessage = null;
        }
        this.changeDetectorRef.markForCheck();
      },
      error: (err) => { // Fallback for unexpected errors not caught by the pipeline's catchError
        const errMsg = err?.error?.message || err?.message || err?.toString() || 'An unexpected error occurred.';
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
    allPackageImsccFiles: ImsccFile[] // This now receives ImsccFile[] with string data
  ): Observable<ProcessingResult[]> {

    console.log(`Starting Drive processing & content creation for ${itemsToProcess.length} items in course "${courseName}"`);

    if (itemsToProcess.length === 0) {
        return of([]);
    }

    return from(itemsToProcess).pipe(
      concatMap((item, index) => {
        this.loadingMessage = `Processing item ${index + 1}/${itemsToProcess.length}: ${item.title?.substring(0, 30)}... (Drive Content)`;
        this.changeDetectorRef.markForCheck();
        console.log(`Processing Item ${index + 1}/${itemsToProcess.length}: "${item.title || 'Untitled'}" (Type: ${item.workType})`);

        const assignmentName = item.title || 'Untitled Assignment';
        const topicName = item.associatedWithDeveloper?.topic || 'General';
        const originalAssignmentHtml = item.descriptionForDisplay;
        const itemId = item.associatedWithDeveloper?.id;

        // filesToUpload in ProcessedCourseWork already contains {file: ImsccFile, targetFileName: string}
        // where ImsccFile.data is string (base64 for images, text for others, or placeholder for unhandled binary)
        const filesToUploadForDrive = item.localFilesToUpload || [];

        const qtiFileArray = item.qtiFile; // This is ImsccFile[] where data is string
        const qtiFileForService: ImsccFile | undefined = (qtiFileArray && qtiFileArray.length > 0) ? qtiFileArray[0] : undefined;

        const shouldCreateDoc = !qtiFileForService && !!originalAssignmentHtml && item.workType === 'ASSIGNMENT' && !!item.richtext;

        if (!itemId) {
          console.error(`   Skipping item "${assignmentName}" due to missing associated developer ID.`);
          return of({
            itemId: undefined, assignmentName, topicName,
            assignmentFolderId: 'ERROR_NO_ID', createdDoc: undefined, createdForm: undefined,
            uploadedFiles: undefined,
            error: { message: 'Missing associated developer ID', stage: 'Pre-flight Check' }
          } as ProcessingResult);
        }

        console.log(`   Item ID: ${itemId}, Topic: "${topicName}", Files: ${filesToUploadForDrive.length}, Has QTI: ${!!qtiFileForService}, Create Doc: ${shouldCreateDoc}`);

        return this.drive.ensureAssignmentFolderStructure(
          courseName, topicName, assignmentName, itemId, accessToken
        ).pipe(
          switchMap(assignmentFolderId => {
            console.log(`   Drive Folder ensured/created for Item ID ${itemId}. Folder ID: ${assignmentFolderId}`);

            // FileUploadService expects ImsccFile where data is string (base64 for images) or ArrayBuffer.
            // Our this.unzippedFiles (source for localFilesToUpload) has images as base64 strings, others as ArrayBuffer or string.
            // We need to map item.localFilesToUpload (which has ImsccFile.data as string) back to a structure
            // FileUploadService can use, or adjust FileUploadService.
            // For now, assuming FileUploadService is robust or we adjust the input here.
            // Let's find the original UnzippedFile data for binary files.
            const filesForDriveUploadService = filesToUploadForDrive.map(ftu => {
              const originalUnzippedFile = this.unzippedFiles.find(uzf => uzf.name === ftu.file.name);
              if (originalUnzippedFile && originalUnzippedFile.data instanceof ArrayBuffer) {
                // Use the original ArrayBuffer for FileUploadService if it's binary
                return {file: {...ftu.file, data: originalUnzippedFile.data}, targetFileName: ftu.targetFileName};
              }
              // If it's already a string (e.g. base64 image data URI, or text file), FileUploadService should handle it.
              return ftu;
            });


            const uploadFiles$: Observable<DriveFile[]> = filesForDriveUploadService.length > 0
              ? this.files.uploadLocalFiles(filesForDriveUploadService, accessToken, assignmentFolderId).pipe(
                tap(uploadedFiles => console.log(`      Uploaded ${uploadedFiles.length} local file(s) for Item ID ${itemId}.`)),
                catchError(uploadError => {
                  console.error(`      ERROR uploading local files for Item ID ${itemId}:`, uploadError);
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
              this.loadingMessage = `Processing item ${index + 1}/${itemsToProcess.length}: ${item.title?.substring(0, 30)}... (Creating Form)`;
              this.changeDetectorRef.markForCheck();
              createContent$ = this.qti.createFormFromQti(
                qtiFileForService, // ImsccFile with string data (XML content)
                allPackageImsccFiles, // Full list, QTI service will find images (base64 strings)
                assignmentName,
                accessToken,
                itemId,
                assignmentFolderId
              ).pipe(
                tap(createdForm => console.log(`      Created Google Form for Item ID ${itemId}.`)),
                catchError(formError => {
                  console.error(`      ERROR creating Google Form for Item ID ${itemId}:`, formError);
                  return throwError(() => ({
                    message: formError.message || 'Form creation from QTI failed',
                    stage: 'Form Creation',
                    details: formError.toString(),
                    itemId
                  } as StagedProcessingError & { itemId: string }));
                })
              );
            } else if (shouldCreateDoc) {
              this.loadingMessage = `Processing item ${index + 1}/${itemsToProcess.length}: ${item.title?.substring(0, 30)}... (Creating Doc)`;
              this.changeDetectorRef.markForCheck();
              createContent$ = uploadFiles$.pipe(
                switchMap(uploadedDriveFiles => {
                  let modifiedHtml = originalAssignmentHtml; // This is already processed HTML
                  // Link replacement logic might need to be more robust if paths are complex
                  uploadedDriveFiles.forEach((uploadedFile, idx) => {
                    const originalFileRef = filesForDriveUploadService[idx];
                      if (originalFileRef && uploadedFile?.id && uploadedFile?.webViewLink) {
                          const originalFileName = originalFileRef.targetFileName;
                          try {
                              const decodedOriginalName = decodeURI(originalFileName);
                              const regex = new RegExp(`href=(["'])([^"']*(?:${escapeRegExp(decodedOriginalName)}|${escapeRegExp(originalFileName)}))\\1`, 'gi');
                              modifiedHtml = modifiedHtml.replace(regex, (match, quote, originalHref) => `href=${quote}${uploadedFile.webViewLink}${quote} target="_blank"`);
                          } catch (e) { console.error(`Error modifying HTML for ${originalFileName}:`, e); }
                      }
                  });
                  return this.docs.createDocFromHtml(
                    modifiedHtml, assignmentName, accessToken, itemId, assignmentFolderId
                  ).pipe(
                    tap(createdDoc => console.log(`      Created Google Doc for Item ID ${itemId}.`)),
                    catchError(docError => {
                      console.error(`      ERROR creating Google Doc for Item ID ${itemId}:`, docError);
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
            const stage = errorDetails?.stage || 'Folder Creation';
            const errorMessage = errorDetails?.message || errorDetails?.error?.message || 'Unknown error during item processing';
            const errorDetailsString = errorDetails?.details || errorDetails?.error?.toString() || errorDetails.toString();

            console.error(`   ERROR processing Item "${assignmentName}" (ID: ${itemId}): Failed at Stage: ${stage}`, errorMessage);
            return of({
              itemId: itemId, assignmentName, topicName,
              assignmentFolderId: 'ERROR', createdDoc: undefined, createdForm: undefined,
              uploadedFiles: undefined,
              error: { message: errorMessage, stage: stage, details: errorDetailsString }
            } as ProcessingResult);
          })
        );
      }),
      toArray()
    );
  }


  addContentAsMaterials(
    processingResults: ProcessingResult[],
    courseWorkItemsToUpdate: ProcessedCourseWork[]
  ): ProcessedCourseWork[] {
    console.log('Associating Drive content (Docs/Forms/Files) as Materials to CourseWork items...');
    const courseWorkMap = new Map<string, ProcessedCourseWork>();
    courseWorkItemsToUpdate.forEach(item => {
      if (item.associatedWithDeveloper?.id) {
        // Ensure materials array exists and is a fresh copy for modification
        courseWorkMap.set(item.associatedWithDeveloper.id, { ...item, materials: [...(item.materials || [])] });
      }
    });

    processingResults.forEach(result => {
      if (result.itemId) {
        const courseWorkItem = courseWorkMap.get(result.itemId);
        if (courseWorkItem) {
          if (result.error) {
            const errorObj = result.error as StagedProcessingError;
            courseWorkItem.processingError = { // Ensure this matches your ProcessedCourseWork interface
                message: errorObj.message || 'An error occurred during Drive operations.',
                stage: errorObj.stage || 'Drive Operation',
                details: errorObj.details || (typeof result.error === 'string' ? result.error : JSON.stringify(result.error))
            };
            console.warn(`   [Material Association] Item ID ${result.itemId} has processing error: ${errorObj.message} at stage ${errorObj.stage}`);
            return; // Skip adding materials if there was an error for this item
          }

          console.log(`   Updating materials for Item "${courseWorkItem.title}" (ID: ${result.itemId})`);
          if (!courseWorkItem.materials) courseWorkItem.materials = []; // Should be initialized above
          const addedMaterialNames: string[] = [];
          let primaryContentMaterial: Material | null = null;

          // Add created Google Doc as material (studentCopy)
          if (result.createdDoc?.id && result.createdDoc?.name) {
            const docMaterial: Material = { driveFile: { driveFile: { id: result.createdDoc.id, title: result.createdDoc.name }, shareMode: 'STUDENT_COPY' } };
            courseWorkItem.materials.push(docMaterial);
            addedMaterialNames.push(`"${result.createdDoc.name}" (Doc)`);
            primaryContentMaterial = docMaterial;
          }

          // Add created Google Form as material (link)
          if (result.createdForm?.form?.formUrl) {
            const formMaterial: Material = {link: {url: result.createdForm.form.formUrl, title: result.createdForm.form.title || result.assignmentName}};
            courseWorkItem.materials.push(formMaterial);
            addedMaterialNames.push(`"${result.createdForm.form.title || result.assignmentName}" (Form)`);
            primaryContentMaterial = formMaterial; // Prioritize Form if both Doc and Form exist for some reason
          }

          // Add uploaded files as materials (view only), excluding the main doc if it was one of the uploaded files
          const filesToAddAsMaterials = result.uploadedFiles || [];
          if (filesToAddAsMaterials.length > 0) {
            filesToAddAsMaterials.forEach(uploadedFile => {
              if (uploadedFile?.id && uploadedFile?.name) {
                // Avoid re-adding the main Google Doc if it was created from an HTML file that was also in localFilesToUpload
                if (!(result.createdDoc?.id === uploadedFile.id)) {
                    const fileMaterial: Material = { driveFile: { driveFile: { id: uploadedFile.id, title: uploadedFile.name }, shareMode: 'VIEW' } };
                  if (!courseWorkItem.materials) courseWorkItem.materials = []; // Should be initialized
                    courseWorkItem.materials.push(fileMaterial);
                    addedMaterialNames.push(`"${uploadedFile.name}"`);
                }
              }
            });
          }

          // Update description based on added materials
          if (addedMaterialNames.length > 0 && !courseWorkItem.descriptionForClassroom?.trim()) { // Only update if description is empty
              let materialDescription: string;
              if (primaryContentMaterial) {
                if (primaryContentMaterial.driveFile) { // It's a Doc
                      materialDescription = `Please review the attached document: ${addedMaterialNames[0]}.`;
                  } else if (primaryContentMaterial.link && primaryContentMaterial.link.url.includes('google.com/forms')) { // It's a Form
                      materialDescription = `Please complete the attached form: ${addedMaterialNames[0]}.`;
                  } else { // Other primary content (e.g. first uploaded file if no doc/form)
                       materialDescription = `Please see the attached content: ${addedMaterialNames[0]}.`;
                  }
                // Append other files if any
                  if (addedMaterialNames.length > 1) {
                    const otherFiles = addedMaterialNames.slice(1); // Get all other material names
                      if (otherFiles.length > 0) {
                         materialDescription += ` Additional file(s): ${otherFiles.join(', ')}.`;
                      }
                  }
              } else { // No primary Doc/Form, just uploaded files
                  if (addedMaterialNames.length === 1) materialDescription = `Please see the attached file: ${addedMaterialNames[0]}.`;
                  else materialDescription = `Please see the attached files (${addedMaterialNames.length}): ${addedMaterialNames.join(', ')}.`;
              }
              courseWorkItem.descriptionForClassroom = materialDescription;
          } else if (addedMaterialNames.length > 0) {
            console.log(`      Materials added for "${courseWorkItem.title}", but keeping existing classroom description.`);
          } else {
            console.log(`      No new materials were added for "${courseWorkItem.title}". Keeping original description.`);
          }

        } else {
          console.warn(`   [Material Association] Could not find matching CourseWork item in map for ProcessingResult with itemId: ${result.itemId}.`);
        }
      } else if (result.error) {
        console.error(`   [Material Association] Skipping material addition for item ID ${result.itemId || 'Unknown'} due to processing error:`, result.error);
      } else if (!result.itemId) {
        console.error(`   [Material Association] Skipping material addition for item "${result.assignmentName}" because it had no ID.`);
      }
    });

    console.log('Finished associating materials.');
    return Array.from(courseWorkMap.values());
  }

}
