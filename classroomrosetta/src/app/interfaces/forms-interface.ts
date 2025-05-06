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

/**
 * Represents a Google Form resource.
 * See: https://developers.google.com/forms/api/reference/rest/v1/forms#Form
 */
export interface GoogleForm {
  formId: string; // The form ID. Read-only.
  info?: FormInfo; // Metadata about the form.
  items?: FormItem[]; // The items that make up the form. Read-only (use batchUpdate to modify).
  responderUri?: string; // The URL for respondents to view and submit the form. Read-only.
  linkedSheetId?: string; // The ID of the linked Google Sheet. Read-only.
  revisionId?: string; // The revision ID of the form. Read-only.
}

/**
* Metadata for a Google Form. Part of the GoogleForm resource.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#forminfo
*/
export interface FormInfo {
  title: string; // The title of the form. Required on creation.
  description?: string; // The description of the form.
  documentTitle?: string; // The title as it appears in Google Drive. Set on creation.
}

/**
* An item in a Google Form. Can be a question, section break, image, etc.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#item
*/
export interface FormItem {
  itemId?: string; // The unique ID for the item. Read-only.
  title?: string; // The title of the item (e.g., question text, image caption).
  description?: string; // Description text below the title.
  questionItem?: QuestionItem; // If the item is a question.
  imageItem?: ImageItem; // If the item is an image.
  videoItem?: VideoItem;
  pageBreakItem?: PageBreakItem;
  sectionHeaderItem?: SectionHeaderItem;
  textItem?: TextItem
  // Add other item types like questionGroupItem, pageBreakItem, textItem, videoItem if needed.
}

/**
 * Represents an image item in a form.
 * See: https://developers.google.com/forms/api/reference/rest/v1/forms#imageitem
 */
export interface ImageItem {
  image?: Image; // The image to display. Required.
}

/**
 * Data for an image.
 * See: https://developers.google.com/forms/api/reference/rest/v1/forms#image
 */
export interface Image {
  contentUri?: string; // A URI to the image, e.g., a "data:" URI for a base64 encoded image.
  altText?: string; // Alt text for the image.
  properties?: ImageProperties; // Properties of the image.
  sourceUri?: string;
}

/**
 * Properties of an image.
 * See: https://developers.google.com/forms/api/reference/rest/v1/forms#imageproperties
 */
export interface ImageProperties {
  alignment?: 'LEFT' | 'CENTER' | 'RIGHT'; // Alignment of the image.
  width?: number; // Width of the image in pixels.
}


/**
* A question item in a Google Form. Contained within a FormItem.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#questionitem
*/
export interface QuestionItem {
  question?: Question; // The question definition.
  image?: Image; // Optional image associated with the question.
}

/**
* Defines a question in a Google Form. Part of QuestionItem.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#question
*/
export interface Question {
  questionId?: string; // The unique ID for the question. Read-only.
  required?: boolean; // Whether the question must be answered.
  choiceQuestion?: ChoiceQuestion; // If it's a multiple choice, dropdown, checkbox question.
  textQuestion?: TextQuestion; // If it's a short answer or paragraph question.
  // Add other question types like dateQuestion, timeQuestion, fileUploadQuestion, scaleQuestion, gridQuestion if needed.
  grading?: Grading; // Optional grading information for quizzes.
}

/**
* Defines a choice-based question (multiple choice, checkbox, dropdown).
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#choicequestion
*/
export interface ChoiceQuestion {
  type: 'RADIO' | 'CHECKBOX' | 'DROP_DOWN'; // Type of choice question.
  options: Option[]; // List of choices. Required.
  shuffle?: boolean; // Whether to shuffle the option order.
}

/**
* Represents a single choice option.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#option
*/
export interface Option {
  value: string; // The text displayed for the option. Required.
  isOther?: boolean; // Whether this is the "Other" option.
  image?: Image; // Optional image associated with the option.
  goToAction?: 'NEXT_SECTION' | 'RESTART_FORM' | 'SUBMIT_FORM'; // Action on selecting this option.
  goToSectionId?: string; // Target section ID if goToAction is NEXT_SECTION.
}

/**
* Defines a text-based question (short answer, paragraph).
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#textquestion
*/
export interface TextQuestion {
  paragraph?: boolean; // True for paragraph, false for short answer.
}

/**
* Grading information for a question, turning the form into a quiz.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#grading
*/
export interface Grading {
  pointValue?: number; // Points the question is worth.
  correctAnswers?: CorrectAnswers; // The correct answers for automatic grading.
  generalFeedback?: Feedback; // Feedback shown regardless of answer.
  whenRight?: Feedback; // Feedback shown for correct answers.
  whenWrong?: Feedback; // Feedback shown for incorrect answers.
}

/**
* The correct answer(s) for a question.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#correctanswers
*/
export interface CorrectAnswers {
  answers: CorrectAnswer[]; // A list of correct answers. For ChoiceQuestions, usually one. For Checkbox, can be multiple.
}

/**
* A single correct answer value.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#correctanswer
*/
export interface CorrectAnswer {
  value: string; // The string value of the correct answer. Must match an Option's value for ChoiceQuestions.
}

/**
 * Feedback for an answer.
 * See: https://developers.google.com/forms/api/reference/rest/v1/forms#feedback
 */
export interface Feedback {
  text?: string;
  // material?: Material[]; // Material is defined in classroom-interface.ts. If needed, ensure it's imported.
  // For simplicity here, assuming feedback material is not immediately required.
}


// --- Batch Update Request Structures ---

/**
* Request body for the forms.batchUpdate method.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms/batchUpdate#request-body
*/
export interface BatchUpdateFormRequest {
  requests: FormRequest[]; // A list of updates to perform.
  includeFormInResponse?: boolean; // Whether to include the updated Form object in the response.
  writeControl?: WriteControl; // Optional write control parameters.
}

/**
* A single update request for the batchUpdate method.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#request
*/
export interface FormRequest {
  createItem?: CreateItemRequest; // Creates an item.
  updateSettings?: UpdateSettingsRequest; // Updates form settings (e.g., to make it a quiz).
  // Add other request types like updateItem, deleteItem, moveItem, updateFormInfo if needed.
}

/**
* Request to create a new item in a form.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#createitemrequest
*/
export interface CreateItemRequest {
  item: FormItem; // The item to create. itemId should be omitted.
  location?: Location; // The location to insert the item. If omitted, appended to the end.
}

/**
* Specifies where to insert or move an item.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#location
*/
export interface Location {
  index?: number; // The zero-based index.
  // fromEndOfSection?: boolean; // Alternative way to specify location (not typically needed for creation).
}

/**
* Response body for the forms.batchUpdate method.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms/batchUpdate#response-body
*/
export interface BatchUpdateFormResponse {
  form?: GoogleForm; // The updated form, if requested.
  replies?: Response[]; // Results for each request, in order.
  writeControl?: WriteControl;
}

/**
* Response for a single request within a batch update.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#response
*/
export interface Response {
  createItem?: CreateItemResponse;
  // Add other response types corresponding to request types.
}

/**
* Response for a CreateItemRequest.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#createitemresponse
*/
export interface CreateItemResponse {
  itemId?: string; // The ID of the created item.
  questionId?: string[]; // IDs of questions created (if item is a QuestionGroupItem).
  // Note: Google Forms API typically returns one questionId per createItem for a question.
  // This might be an array if the item created multiple questions (e.g. grid).
}

/**
* Optional write control parameters for requests.
* See: https://developers.google.com/forms/api/reference/rest/v1/forms#writecontrol
*/
export interface WriteControl {
  requiredRevisionId?: string; // Target revision ID for concurrency control.
  targetRevisionId?: string; // Desired revision ID after update.
}

/**
 * Request to update the settings of a Google Form.
 * See: https://developers.google.com/forms/api/reference/rest/v1/forms/batchUpdate#updatesettingsrequest
 */
export interface UpdateSettingsRequest {
  settings: FormSettings; // The new settings to apply.
  updateMask: string;     // A field mask to specify which settings to update. e.g., "quizSettings.isQuiz"
}

/**
 * Settings for a Google Form.
 * See: https://developers.google.com/forms/api/reference/rest/v1/forms#formsettings
 */
export interface FormSettings {
  quizSettings?: QuizSettings;
  // Add other form settings like collectEmail, responseReceipts, etc., if needed.
}

/**
 * Settings for a quiz.
 * See: https://developers.google.com/forms/api/reference/rest/v1/forms#quizsettings
 */
export interface QuizSettings {
  isQuiz?: boolean; // True if the form is a quiz, responses get grades.
  // Add other quiz settings like gradeImmediately, showCorrectAnswers, showPointValues if needed.
}


// Ensure these are exported if they are not already
export interface Question {
  // ... other question properties like choiceQuestion, textQuestion
  scaleQuestion?: ScaleQuestion;
  dateQuestion?: DateQuestion;
  timeQuestion?: TimeQuestion;
  rowQuestion?: RowQuestion; // For Grid questions
  // fileUploadQuestion?: FileUploadQuestion; // If you plan to support file uploads
  grading?: Grading; // Already likely there
  required?: boolean; // Already likely there
}

export interface ScaleQuestion {
  low: number;
  high: number;
  lowLabel?: string;
  highLabel?: string;
}

export interface DateQuestion {
  includeTime?: boolean;
  includeYear?: boolean; // FormApp uses setIncludesYear, so this is relevant
}

export interface TimeQuestion {
  duration?: boolean; // If true, use DURATION mode, else TIME_OF_DAY
}

// For GridItem and CheckboxGridItem
export interface Row {
    title: string;
}
export interface RowQuestion {
    rows: Row[];
    // Columns are typically handled by a ChoiceQuestion with type GRID or CHECKBOX_GRID
    // So, your existing ChoiceQuestion might need a 'type' like 'GRID' or 'CHECKBOX_GRID'
    // or you define columns explicitly here if your QTI parsing separates them.
    // For now, assuming columns are part of the choiceQuestion structure within the main Question.
}

export interface Video {
  youtubeUri: string;
  altText?: string; // Or properties for width, alignment if supported by Forms API
}

export interface VideoItem {
  video: Video;
  // title?: string; // Title for video item is part of the parent FormItem
  // description?: string; // Description for video item is part of the parent FormItem
}

export interface PageBreakItem {
  // Typically has no specific properties other than those on FormItem (title, description)
  // The presence of this object signifies a page break.
  // For example, in the FormRequest: { createItem: { item: { title: "Next Section", pageBreakItem: {} }, location: { ... } } }
  [key: string]: any; // Allow empty object {}
}

export interface SectionHeaderItem {
  // Similar to PageBreakItem, its presence indicates the type.
  // Title and description are on the parent FormItem.
  // For example: { createItem: { item: { title: "New Section Title", description: "Details about this section", sectionHeaderItem: {} }, location: { ... } } }
  [key: string]: any; // Allow empty object {}
}

export interface TextItem {
    // text: string; // The actual text content, if not using title/description of FormItem
    [key: string]: any; // Allow empty object if title/description of FormItem are used
}


