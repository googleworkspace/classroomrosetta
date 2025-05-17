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

// classroom-interface.ts

export interface Material {
  link?: {url: string; title?: string;};
  driveFile?: {driveFile: {id: string; title: string}; shareMode: 'VIEW' | 'STUDENT_COPY' | 'EDIT' | 'DRIVE_FILE_UNSPECIFIED' };
  form?: {formUrl: string; title?: string;};
  youtubeVideo?: {id: string; title?: string; thumbnailUrl?: string;};
}

/**
 * Information associated with the original IMSCC item and resource.
 * Stored in the Coursework's `associatedWithDeveloper` field.
 */
export interface AssociatedInfo {
  id: string; // The identifier from the IMSCC <item> element
  resourceId?: string | null; // The identifier from the IMSCC <resource> element (can be null if resource wasn't found)
  topic?: string; // The Google Classroom Topic name derived from parent <item>s
  sourceHtmlFile?: ImsccFile | undefined; // The specific HTML file used as the primary content source (if applicable)
  sourceXmlFile?: ImsccFile | undefined; // The specific XML file used as the primary content source (if applicable, e.g., topic.xml, weblink.xml, qti.xml)
  sourceOtherFile?: ImsccFile | undefined; // The primary non-text/xml/html file used as the content source (if applicable)
  // You could potentially add more fields here to track other IMSCC metadata if needed
}


export interface CourseWorkBase {
  title?: string;
  materials?: Material[];
  /**
   * Type of course work.
   * 'ASSIGNMENT', 'SHORT_ANSWER_QUESTION', 'MULTIPLE_CHOICE_QUESTION' correspond to Google Classroom API CourseWork types.
   * 'MATERIAL' is an internal designation by the ConverterService to indicate that this item
   * should be created as a CourseWorkMaterial resource in Google Classroom using the
   * courses.courseWorkMaterials.create endpoint.
   */
  workType: 'ASSIGNMENT' | 'SHORT_ANSWER_QUESTION' | 'MULTIPLE_CHOICE_QUESTION' | 'MATERIAL';
  state: 'PUBLISHED' | 'DRAFT' | 'DELETED' | 'SCHEDULED';
  assignment?: {
    studentWorkFolder?: {id: string};
  };
  multipleChoiceQuestion?: {
    choices?: string[];
  };

  associatedWithDeveloper?: AssociatedInfo;
  description?: string; // Plain text description sent to Classroom API (now descriptionForClassroom)
  maxPoints?: number;
  topicId?: string;
  scheduledTime?: string;
  dueDate?: GoogleDate;
  dueTime?: TimeOfDay;
  submissionModificationMode?: 'MODIFIABLE_UNTIL_TURNED_IN' | 'MODIFIABLE' | 'LOCKED';
}

export interface ProcessedCourseWork extends CourseWorkBase {
  // description is used internally by the ConverterService and maps to the Classroom API's description field
  // descriptionForDisplay is the rich HTML content shown to the user *within* your application/UI before posting
  // descriptionForClassroom is the plain text (or limited formatted text) description sent to the Classroom API
  descriptionForDisplay: string; // Rich HTML content for UI display
  descriptionForClassroom: string; // Plain text content for Classroom API description field


  localFilesToUpload: Array<{file: ImsccFile; targetFileName: string}>;
  qtiFile?: ImsccFile[]; // Specific field for QTI XML file(s)
  htmlContent?: string; // Original raw HTML content extracted (if applicable) - might be redundant now?
  webLinkUrl?: string; // Extracted URL for WebLink resources
  richtext?: boolean; // Flag indicating if descriptionForDisplay contains rich text

  processingError?: {
    message: string;
    stage?: string;
    details?: any;
  };
  // Added fields to store results from Classroom API after creation
  classroomCourseWorkId?: string; // ID of the created CourseWork or CourseWorkMaterial
  classroomLink?: string;       // AlternateLink to view the item in Classroom
}

export interface Topic {
  courseId: string;
  topicId: string;
  name: string;
  updateTime: string;
}

export interface ListTopicsResponse {
  topic?: Topic[];
  nextPageToken?: string;
}

export interface ImsccFile {
  name: string;
  data: string | ArrayBuffer; // Data can be raw ArrayBuffer for binary files, or string for text/base64
  mimeType: string;
}

export interface GroupedCoursework {
  topic: string | undefined;
  items: ProcessedCourseWork[];
  allSelected: boolean;
  indeterminate: boolean;
}

export interface ProcessingResult {
  itemId: string | undefined; // IMSCC Item ID
  assignmentName: string; // Title of the created item
  topicName: string | undefined; // Google Classroom Topic name
  assignmentFolderId?: string; // ID of the Drive folder created for the assignment (if any)
  createdDoc?: DriveFile; // Info about a created Google Doc/Sheet/Slide (if any)
  createdForm?: Material; // Info about a created Google Form (if any)
  error?: any; // Error details if processing failed for this item
  uploadedFiles?: DriveFile[]; // Info about files uploaded to Drive and attached
  classroomCourseWorkId?: string; // ID of the created Classroom item
  classroomLink?: string;       // Link to the created Classroom item
}

export interface Classroom {
  id: string;
  name: string;
  section?: string;
  descriptionHeading?: string;
  room?: string;
  courseState: 'COURSE_STATE_UNSPECIFIED' | 'ACTIVE' | 'ARCHIVED' | 'PROVISIONED' | 'DECLINED' | 'SUSPENDED';
}

export interface ClassroomListResponse {
  courses: Classroom[];
  nextPageToken?: string;
}

export interface SubmissionData {
  classroomIds: string[];
  assignmentIds: string[]; // These would be Classroom Coursework IDs
}


export interface WebLinkXML {
  title: string;
  url: {
    href: string;
    target?: string;
    windowFeatures?: string;
  };
}

export interface FileObject {
  data: string;
  mimeType: string;
  name: string;
}


export interface GoogleClassroomCoursework {
  id?: string;
  title: string;
  description?: string;
  materials: Material[];
  workType?: 'ASSIGNMENT'; // API type 'ASSIGNMENT', 'SHORT_ANSWER_QUESTION', 'MULTIPLE_CHOICE_QUESTION'
  assigneeMode: 'ALL_STUDENTS';
  subAssignments?: GoogleClassroomCoursework[]; // This structure doesn't match the API, likely for internal grouping?
}

/**
 * Represents course work (e.g., assignment, question) in Google Classroom API v1.
 * Use this structure when creating CourseWork via courses.courseWork.create endpoint.
 * See: https://developers.google.com/classroom/reference/rest/v1/courses.courseWork
 */
export interface CourseWork {
  courseId?: string; // Identifier of the course. Read-only.
  id?: string; // Classroom-assigned identifier for this course work. Read-only.
  title?: string; // Title of this course work. Required on creation.
  description?: string; // Optional description. Supports limited HTML/Markdown.
  materials?: Material[]; // Materials attached to this course work.
  state?: 'PUBLISHED' | 'DRAFT' | 'DELETED' | 'SCHEDULED';
  alternateLink?: string; // URL to view in the Classroom UI. Read-only.
  creationTime?: string; // Timestamp of when this course work was created. Read-only.
  updateTime?: string; // Timestamp of the last update. Read-only.
  maxPoints?: number; // Maximum number of points possible for this course work.
  workType?: 'ASSIGNMENT' | 'SHORT_ANSWER_QUESTION' | 'MULTIPLE_CHOICE_QUESTION'; // The type of work.
  assigneeMode?: 'ALL_STUDENTS' | 'INDIVIDUAL_STUDENTS';
  creatorUserId?: string; // Identifier of the user who created the course work. Read-only.
  topicId?: string; // Identifier of the topic associated with this course work.
  scheduledTime?: string; // Optional timestamp for scheduled publication. If set, state must be 'DRAFT' (API will change to 'SCHEDULED').
  dueDate?: GoogleDate; // Optional due date.
  dueTime?: TimeOfDay; // Optional due time.
  submissionModificationMode?: 'MODIFIABLE_UNTIL_TURNED_IN' | 'MODIFIABLE' | 'LOCKED';
  assignment?: { studentWorkFolder?: {id: string}; }; // Additional details for assignments.
  multipleChoiceQuestion?: { choices?: string[]; }; // Additional details for multiple choice questions.
}

/**
 * Represents a material item in Google Classroom API v1.
 * Use this structure when creating CourseWorkMaterial via courses.courseWorkMaterials.create endpoint.
 * See: https://developers.google.com/classroom/reference/rest/v1/courses.courseWorkMaterials#CourseWorkMaterial
 */
export interface CourseWorkMaterial {
  courseId?: string; // Identifier of the course. Read-only.
  id?: string; // Classroom-assigned identifier for this material. Read-only.
  title?: string; // Title of this material. Required on creation.
  description?: string; // Optional description. Supports limited HTML/Markdown.
  materials?: Material[]; // Materials attached to this course work material.
  state?: 'PUBLISHED' | 'DRAFT' | 'DELETED' | 'SCHEDULED';
  alternateLink?: string; // URL to view in the Classroom UI. Read-only.
  creationTime?: string; // Timestamp of when this material was created. Read-only.
  updateTime?: string; // Timestamp of the last update. Read-only.
  creatorUserId?: string; // Identifier of the user who created the material. Read-only.
  topicId?: string; // Identifier of the topic associated with this material.
  scheduledTime?: string; // Optional timestamp for scheduled publication. If set, state must be 'DRAFT' (API will change to 'SCHEDULED').
  assigneeMode?: 'ALL_STUDENTS'; // Materials are always for all students. Read-only for creation.
  // individualStudentsOptions is not applicable to materials in the API structure
}


export interface GoogleDate {
  year?: number;
  month?: number;
  day?: number;
}

export interface TimeOfDay {
  hours?: number;
  minutes?: number;
  seconds?: number;
  nanos?: number;
}

export interface DriveFile {
  kind?: string;
  id: string; // Google Drive File ID
  name: string; // File name
  mimeType: string; // File MIME type
  webViewLink?: string; // URL to view the file in Drive
  appProperties?: { [key: string]: string; };
  parents?: string[];
  thumbnailLink?: string;
  contentHints?: { thumbnail?: { image: string, mimeType: string } };
}
