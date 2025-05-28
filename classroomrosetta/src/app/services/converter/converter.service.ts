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

import {Injectable, inject} from '@angular/core';
import {Observable, from, of, throwError, EMPTY, concat} from 'rxjs';
import {tap, catchError, concatMap, filter} from 'rxjs/operators';
import {
  ProcessedCourseWork,
  ImsccFile,
} from '../../interfaces/classroom-interface';
import {decode} from 'html-entities';
import {ImsccParsingHelperService} from './helper/imscc-parsing-helper.service';

@Injectable({
  providedIn: 'root'
})
export class ConverterService {
  private specialRefPrefixes = [
    '$IMS-CC-FILEBASE$',
    'IMS-CC-FILEBASE',
    '$CANVAS_OBJECT_REFERENCE$',
    'CANVAS_OBJECT_REFERENCE',
    '$WIKI_REFERENCE$'
  ];

  private readonly IMSCP_V1P1_NS = 'http://www.imsglobal.org/xsd/imscp_v1p1';
  private readonly D2L_V2P0_NS = 'http://desire2learn.com/xsd/d2lcp_v2p0';

  public coursename = '';
  private parsingHelper = inject(ImsccParsingHelperService);
  public skippedItemLog: {id?: string, title: string, reason: string}[] = [];

  private fileMap: Map<string, ImsccFile> = new Map();
  private manifestXmlDoc: XMLDocument | null = null;

  constructor() { }

  private getFileMapKey(path: string): string {
    if (!path) return '';
    return this.parsingHelper.correctCyrillicCPath(this.parsingHelper.tryDecodeURIComponent(path)).toLowerCase();
  }

  convertImscc(files: ImsccFile[]): Observable<ProcessedCourseWork> {
    console.log("Starting IMSCC conversion process (Main Thread)...");
    this.skippedItemLog = [];
    this.fileMap = new Map();
    this.manifestXmlDoc = null;

    files.forEach(file => {
      let processedFile = file;
      if (file.mimeType?.startsWith('image/') && this.parsingHelper.isArrayBuffer(file.data)) {
        try {
          const byteNumbers = new Array(file.data.byteLength);
          const byteArray = new Uint8Array(file.data);
          for (let i = 0; i < file.data.byteLength; i++) {byteNumbers[i] = byteArray[i];}
          const binaryString = String.fromCharCode.apply(null, byteNumbers);
          const base64String = btoa(binaryString);
          processedFile = {...file, data: `data:${file.mimeType};base64,${base64String}`};
        } catch (e) {
          console.warn(`Could not convert ArrayBuffer to base64 for image ${file.name}`, e);
        }
      }
      else if (this.parsingHelper.isArrayBuffer(file.data) && !file.mimeType?.startsWith('image/') && !file.mimeType?.startsWith('video/') && !file.mimeType?.startsWith('audio/')) {
        try {
          const textDecoder = new TextDecoder('utf-8');
          const textData = textDecoder.decode(file.data);
          processedFile = {...file, data: textData};
        } catch (e) {
          console.warn(`Could not decode ArrayBuffer as text for file ${file.name}`, e);
        }
      }
      const normalizedFileNameKey = this.getFileMapKey(processedFile.name);
      this.fileMap.set(normalizedFileNameKey, processedFile);
    });

    const manifestFileKey = this.getFileMapKey('imsmanifest.xml');
    const manifestFile = this.fileMap.get(manifestFileKey);

    if (!manifestFile || typeof manifestFile.data !== 'string') {
      console.error('IMS Manifest file (imsmanifest.xml) not found or data is not a string AFTER file processing.');
      return throwError(() => new Error('imsmanifest.xml not found or data is not a string.'));
    }

    let organization: Element | null = null;
    let rootItems: Element[] = [];
    let resourcesElement: Element | null = null;
    let directResources: Element[] = [];

    try {
      const parser = new DOMParser();
      let doc = parser.parseFromString(manifestFile.data, "application/xml");
      let parseError = doc.querySelector('parsererror');

      if (parseError) {
        console.warn("XML Parsing Error (application/xml), attempting fallback:", parseError.textContent);
        let tempDoc = parser.parseFromString(manifestFile.data, "text/xml");
        parseError = tempDoc.querySelector('parsererror');
        if (!parseError) {
          doc = tempDoc;
          console.warn("Parsed manifest using text/xml fallback.");
        } else {
          console.warn("Fallback XML Parsing Error (text/xml), attempting BOM removal:", parseError.textContent);
          const cleanedData = manifestFile.data.charCodeAt(0) === 0xFEFF ? manifestFile.data.substring(1) : manifestFile.data;
          tempDoc = parser.parseFromString(cleanedData, "application/xml");
          parseError = tempDoc.querySelector('parsererror');
          if (!parseError) {
            doc = tempDoc;
            console.warn("Parsed manifest after removing BOM.");
          } else {
            console.error("Final XML Parsing Failed after all fallbacks:", parseError.textContent);
            throw new Error('Failed to parse imsmanifest.xml after fallbacks. Check manifest structure.');
          }
        }
      }
      this.manifestXmlDoc = doc;
      this.coursename = this.parsingHelper.extractManifestTitle(this.manifestXmlDoc) || 'Untitled Course';
      console.log(`Extracted course name: ${this.coursename}`);
      const organizationsNode = this.manifestXmlDoc.getElementsByTagNameNS(this.IMSCP_V1P1_NS, 'organizations')[0]
        || this.manifestXmlDoc.getElementsByTagName('organizations')[0];

      if (organizationsNode) {
        organization = organizationsNode.getElementsByTagNameNS(this.IMSCP_V1P1_NS, 'organization')[0]
          || organizationsNode.getElementsByTagName('organization')[0];
        if (organization) {
          rootItems = Array.from(organization.children).filter(
            (node): node is Element => node instanceof Element && node.localName === 'item'
          );
          if (rootItems.length === 0) console.warn('No <item> elements found within the organization.');
        } else console.warn('No <organization> element found within <organizations>.');
      } else {
        resourcesElement = this.manifestXmlDoc.getElementsByTagNameNS(this.IMSCP_V1P1_NS, 'resources')[0]
          || this.manifestXmlDoc.getElementsByTagName('resources')[0];
        if (resourcesElement) {
          directResources = Array.from(resourcesElement.children).filter(
            (node): node is Element => node instanceof Element && node.localName === 'resource'
          );
          if (directResources.length === 0) console.warn('Found <resources> element, but it contains no <resource> children.');
        } else {
          console.error('Manifest contains neither <organizations> nor <resources> elements. Cannot process.');
          return throwError(() => new Error('No <organizations> or <resources> found in manifest'));
        }
      }
    } catch (error) {
      console.error('Error processing IMSCC package manifest:', error);
      const message = error instanceof Error ? error.message : String(error);
      return throwError(() => new Error(`Failed to process IMSCC manifest: ${message}`));
    }

    let processingStream: Observable<ProcessedCourseWork>;
    if (rootItems.length > 0) {
      processingStream = this.processImsccItemsStream(rootItems, undefined);
    } else if (directResources.length > 0) {
      processingStream = from(directResources).pipe(
        concatMap(resource => {
          try {
            const title = resource.getAttribute('title') || this.parsingHelper.extractTitleFromMetadata(resource) || 'Untitled Resource';
            const identifier = resource.getAttribute('identifier') || `resource_${Math.random().toString(36).substring(2)}`;
            return this.processResource(resource, title, identifier, undefined);
          } catch (error) {
            console.error(`Error processing direct resource (ID: ${resource.getAttribute('identifier') || 'unknown'}):`, error);
            this.skippedItemLog.push({id: resource.getAttribute('identifier') || undefined, title: resource.getAttribute('title') || 'Unknown direct resource', reason: `Error during processing: ${error instanceof Error ? error.message : String(error)}`});
            return EMPTY;
          }
        }),
        filter((result): result is ProcessedCourseWork => result !== null)
      );
    } else {
      console.warn('No root items or direct resources found to process. Conversion will yield no results.');
      processingStream = EMPTY;
    }

    return processingStream.pipe(
      tap(item => console.log(` -> Emitting processed item: "${item.title}" (Type: ${item.workType}, ID: ${item.associatedWithDeveloper?.id}, GDoc Candidate: ${!!item.richtext})`)),
      catchError(err => {
        console.error("Error during IMSCC content processing stream:", err);
        const wrappedError = err instanceof Error ? err : new Error(String(err));
        return throwError(() => new Error(`Error processing IMSCC content stream: ${wrappedError.message}`));
      })
    );
  }

  private processImsccItemsStream(
    items: Element[],
    parentTopic?: string
  ): Observable<ProcessedCourseWork> {
    if (!items || items.length === 0) return EMPTY;

    return from(items).pipe(
      concatMap((item: Element) => {
        try {
          const identifier = item.getAttribute('identifier') || `item_${Math.random().toString(36).substring(2)}`;
          const titleElement = item.querySelector(':scope > title');
          const rawTitle = titleElement?.textContent?.trim() || this.parsingHelper.extractTitleFromMetadata(item) || 'Untitled Item';
          const identifierRef = item.getAttribute('identifierref');
          const childItems = Array.from(item.children).filter(
            (node): node is Element => node instanceof Element && node.localName === 'item'
          );

          const sanitizedTopicName = this.parsingHelper.sanitizeTopicName(rawTitle);
          let resourceObservable: Observable<ProcessedCourseWork | null> = EMPTY;

          if (identifierRef) {
            const resourceSelector = `resource[identifier="${identifierRef}"]`;
            const resource = this.manifestXmlDoc?.querySelector(resourceSelector) ||
              Array.from(this.manifestXmlDoc?.getElementsByTagName('resource') || []).find(r => r.getAttribute('identifier') === identifierRef);

            if (resource) {
              resourceObservable = this.processResource(resource, rawTitle, identifier, parentTopic);
            } else {
              console.warn(`   Resource not found for identifierref: ${identifierRef} (Item: "${rawTitle}"). This item might be a folder or a broken link.`);
              this.skippedItemLog.push({id: identifier, title: rawTitle, reason: `Resource not found for ref: ${identifierRef}`});
            }
          } else {
            if (childItems.length === 0) {
              this.skippedItemLog.push({id: identifier, title: rawTitle, reason: 'No resource reference and no child items'});
            }
          }

          let childItemsObservable: Observable<ProcessedCourseWork> = EMPTY;
          if (childItems.length > 0) {
            childItemsObservable = this.processImsccItemsStream(childItems, sanitizedTopicName);
          }

          return concat(resourceObservable, childItemsObservable).pipe(
            filter((result): result is ProcessedCourseWork => result !== null)
          );

        } catch (error) {
          const itemIdentifier = item.getAttribute('identifier') || 'unknown_item';
          const itemTitle = item.querySelector(':scope > title')?.textContent?.trim() || 'Untitled Item';
          console.error(`Error processing individual item (ID: ${itemIdentifier}, Title: ${itemTitle}):`, error);
          this.skippedItemLog.push({id: itemIdentifier, title: itemTitle, reason: `Error during item processing: ${error instanceof Error ? error.message : String(error)}`});
          return EMPTY;
        }
      }),
      catchError(err => {
        console.error("Error in processImsccItems stream:", err);
        return throwError(() => err);
      })
    );
  }

  private processResource(
    resource: Element,
    itemTitle: string,
    imsccIdentifier: string,
    parentTopic?: string
  ): Observable<ProcessedCourseWork | null> {
    const resourceIdentifier = resource.getAttribute('identifier');
    const resourceType = resource.getAttribute('type');
    const resourceHref = resource.getAttribute('href');
    const baseHref = resource.getAttribute('xml:base');

    const resourceOwnTitle = this.parsingHelper.extractTitleFromMetadata(resource) || resource.getAttribute('title');
    const finalTitle = resourceOwnTitle || itemTitle;

    if (!this.manifestXmlDoc) {
      console.error(`   [Converter] processResource called before manifestXmlDoc is set.`);
      this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: 'Internal Error: Manifest not parsed'});
      return of(null);
    }

    const d2lMaterialType = resource.getAttributeNS(this.D2L_V2P0_NS, 'material_type');
    if (d2lMaterialType === 'orgunitconfig') {
      this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: 'D2L orgunitconfig'});
      return of(null);
    }

    let courseworkBase: Partial<ProcessedCourseWork> & {convertToGoogleDoc?: boolean} = {
      title: finalTitle,
      state: 'DRAFT',
      materials: [],
      localFilesToUpload: [],
      associatedWithDeveloper: {
        id: imsccIdentifier,
        resourceId: resourceIdentifier,
        topic: parentTopic,
      },
      descriptionForDisplay: '',
      descriptionForClassroom: '',
      richtext: false,
      workType: 'ASSIGNMENT', // Default workType
      convertToGoogleDoc: false // Default to false
    };

    let primaryResourceFile: ImsccFile | null = null;
    let resolvedPrimaryHref: string | null = null;
    let primaryFileXmlDoc: XMLDocument | null = null;

    let primaryFilePathOrUrl = resourceHref;
    if (!primaryFilePathOrUrl) {
      const firstFileElement = Array.from(resource.children).find(node => node instanceof Element && node.localName === 'file') as Element | undefined;
      primaryFilePathOrUrl = firstFileElement?.getAttribute('href') || null;
    }

    // Logic to resolve primaryResourceFile and primaryFileXmlDoc (if applicable)
    // This part is complex and remains largely the same as your original code.
    // For brevity, I'm assuming this part correctly populates primaryResourceFile and primaryFileXmlDoc.
    // START: Simplified primary file resolution for focus (replace with your full logic)
    if (primaryFilePathOrUrl && !primaryFilePathOrUrl.match(/^https?:\/\//i) && !this.specialRefPrefixes.some(prefix => primaryFilePathOrUrl!.startsWith(prefix))) {
      resolvedPrimaryHref = this.parsingHelper.resolveRelativePath(baseHref, this.parsingHelper.tryDecodeURIComponent(primaryFilePathOrUrl));
      if (resolvedPrimaryHref) {
        primaryResourceFile = this.fileMap.get(this.getFileMapKey(resolvedPrimaryHref)) || null;
      }
      if (!primaryResourceFile) {
        const resolvedRawHref = this.parsingHelper.resolveRelativePath(baseHref, primaryFilePathOrUrl);
        if (resolvedRawHref && resolvedRawHref !== resolvedPrimaryHref) {
          primaryResourceFile = this.fileMap.get(this.getFileMapKey(resolvedRawHref)) || null;
          if (primaryResourceFile) resolvedPrimaryHref = resolvedRawHref;
        }
      }

      if (!primaryResourceFile) {
        console.warn(`   [Converter] Referenced file not found in package: ${primaryFilePathOrUrl} (Resolved attempts: ${resolvedPrimaryHref})`);
      } else {
        if ((primaryResourceFile.name.toLowerCase().endsWith('.xml') || primaryResourceFile.mimeType?.includes('xml')) && typeof primaryResourceFile.data === 'string') {
          try {
            const parser = new DOMParser();
            const cleanXmlData = primaryResourceFile.data.charCodeAt(0) === 0xFEFF ? primaryResourceFile.data.substring(1) : primaryResourceFile.data;
            primaryFileXmlDoc = parser.parseFromString(cleanXmlData, "application/xml");
            if (primaryFileXmlDoc.querySelector('parsererror')) {
              console.warn(`   [Converter] XML parsing error for ${primaryResourceFile.name}.`);
              primaryFileXmlDoc = null;
            }
          } catch (e) {
            console.error(`   [Converter] Exception parsing primary file XML for ${primaryResourceFile.name}:`, e);
            primaryFileXmlDoc = null;
          }
        }
      }
    } else if (primaryFilePathOrUrl && primaryFilePathOrUrl.match(/^https?:\/\//i)) {
      resolvedPrimaryHref = primaryFilePathOrUrl;
    } else if (primaryFilePathOrUrl && this.specialRefPrefixes.some(prefix => primaryFilePathOrUrl!.startsWith(prefix))) {
      const matchedPrefix = this.specialRefPrefixes.find(p => primaryFilePathOrUrl!.startsWith(p));
      let pathAfterPrefix = primaryFilePathOrUrl!.substring(matchedPrefix!.length);
      if (pathAfterPrefix.startsWith('/')) pathAfterPrefix = pathAfterPrefix.substring(1);

      if (matchedPrefix === '$IMS-CC-FILEBASE$') {
        const resolvedPathFromRoot = this.parsingHelper.resolveRelativePath("", pathAfterPrefix);
        if (resolvedPathFromRoot) {
          const cleanedResolvedPath = resolvedPathFromRoot.split(/[?#]/)[0];
          primaryResourceFile = this.fileMap.get(this.getFileMapKey(cleanedResolvedPath)) || null;
          if (primaryResourceFile) {
            resolvedPrimaryHref = primaryResourceFile.name;
          } else {
            const pathAfterPrefixCleanedForFallback = pathAfterPrefix.split(/[?#]/)[0];
            const decodedFileName = this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleanedForFallback).toLowerCase();
            for (const keyFromMap of this.fileMap.keys()) {
              if (keyFromMap.endsWith(decodedFileName)) {
                primaryResourceFile = this.fileMap.get(keyFromMap)!;
                resolvedPrimaryHref = primaryResourceFile.name;
                break;
              }
            }
            if (!primaryResourceFile) {
              const directKey = this.getFileMapKey(pathAfterPrefixCleanedForFallback);
              primaryResourceFile = this.fileMap.get(directKey) || null;
              if (primaryResourceFile) resolvedPrimaryHref = primaryResourceFile.name;
            }
          }
        }
      } else { // For other special prefixes like $CANVAS_OBJECT_REFERENCE$ or $WIKI_REFERENCE$
        const pathAfterPrefixCleanedForNonCC = pathAfterPrefix.split(/[?#]/)[0];
        // For these, the resolvedPrimaryHref might remain the special prefixed URL,
        // or if a file is found directly by the cleaned path, it becomes the file name.
        resolvedPrimaryHref = primaryFilePathOrUrl; // Default to the original special link
        const directKey = this.getFileMapKey(pathAfterPrefixCleanedForNonCC);
        const foundFile = this.fileMap.get(directKey) || null;
        if (foundFile) {
          primaryResourceFile = foundFile;
          resolvedPrimaryHref = primaryResourceFile.name; // Update to actual file name if found
        }
      }

      // If a primaryResourceFile was determined from a special prefix, try to parse if XML
      if (primaryResourceFile) {
        if ((primaryResourceFile.name.toLowerCase().endsWith('.xml') || primaryResourceFile.mimeType?.includes('xml')) && typeof primaryResourceFile.data === 'string') {
          try {
            const parser = new DOMParser();
            const cleanXmlData = primaryResourceFile.data.charCodeAt(0) === 0xFEFF ? primaryResourceFile.data.substring(1) : primaryResourceFile.data;
            primaryFileXmlDoc = parser.parseFromString(cleanXmlData, "application/xml");
            if (primaryFileXmlDoc.querySelector('parsererror')) primaryFileXmlDoc = null;
          } catch (e) {primaryFileXmlDoc = null;}
        }
      }
    }
    // END: Simplified primary file resolution

    const isStandardQti = (resourceType === 'imsqti_xmlv1p2/xml' || resourceType === 'imsqti_xmlv1p2p1/imsqti_asiitem_xmlv1p2p1' || resourceType?.startsWith('application/vnd.ims.qti') || resourceType?.startsWith('assessment/x-bb-qti') || ((primaryResourceFile?.name?.toLowerCase().endsWith('.xml') || resolvedPrimaryHref?.toLowerCase().endsWith('.xml')) && resourceType?.toLowerCase().includes('qti')));
    const isD2lQuiz = d2lMaterialType === 'd2lquiz';
    const isDiscussionTopic = (primaryResourceFile && primaryFileXmlDoc && this.parsingHelper.isTopicXml(primaryResourceFile, primaryFileXmlDoc)) ||
      resourceType?.toLowerCase().includes('discussiontopic') ||
      resourceType?.toLowerCase().startsWith('imsdt');


    if (isStandardQti || isD2lQuiz) {
      courseworkBase.workType = 'ASSIGNMENT';
      if (primaryResourceFile && primaryFileXmlDoc) {
        courseworkBase.qtiFile = [primaryResourceFile];
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
      } else if (primaryResourceFile && (primaryResourceFile.name.toLowerCase().endsWith('.zip') || primaryResourceFile.mimeType === 'application/zip')) {
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile, targetFileName: primaryResourceFile.name.split('/').pop() || primaryResourceFile.name});
        courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
      } else if (primaryResourceFile) {
        console.warn(`   [Converter] QTI/Assessment "${finalTitle}" main file (${primaryResourceFile.name}) not XML/ZIP. Attaching as general file.`);
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile, targetFileName: primaryResourceFile.name.split('/').pop() || primaryResourceFile.name});
        courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
      }
      else {
        console.warn(`   Skipping QTI/Assessment resource "${finalTitle}" (ID: ${resourceIdentifier}): No valid primary file found.`);
        this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: 'QTI/Assessment - No valid primary file'});
        return of(null);
      }
    }
    else if (primaryResourceFile && primaryFileXmlDoc && this.parsingHelper.isWebLinkXml(primaryResourceFile, primaryFileXmlDoc)) {
      const extractedUrl = this.parsingHelper.extractWebLinkUrl(primaryResourceFile!, primaryFileXmlDoc!);
      if (extractedUrl) {
        courseworkBase.webLinkUrl = extractedUrl;
        courseworkBase.workType = 'ASSIGNMENT'; // Or MATERIAL depending on preference
        if (!courseworkBase.materials?.some(m => m.link?.url === extractedUrl)) {
          courseworkBase.materials?.push({link: {url: extractedUrl}});
        }
        if (!courseworkBase.descriptionForClassroom) courseworkBase.descriptionForClassroom = `Please follow this link: ${finalTitle}`;
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
        courseworkBase.convertToGoogleDoc = true; // Web links are candidates for GDoc conversion
      } else {
        console.warn(`   [Converter] WebLink XML "${finalTitle}" found but could not extract URL. Attaching XML file.`);
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile!, targetFileName: primaryResourceFile!.name.split('/').pop() || primaryResourceFile!.name});
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
        courseworkBase.workType = 'MATERIAL';
      }
    }
    else if (isDiscussionTopic) {
      courseworkBase.workType = 'SHORT_ANSWER_QUESTION'; // Classroom type for discussions
      let contentHtml: string | null = null;
      let contentSourceFilePath: string | null = null;

      if (primaryResourceFile && primaryFileXmlDoc && this.parsingHelper.isTopicXml(primaryResourceFile, primaryFileXmlDoc)) {
        contentHtml = this.parsingHelper.extractTopicDescriptionHtml(primaryResourceFile, primaryFileXmlDoc);
        contentSourceFilePath = primaryResourceFile.name;
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
        if (!contentHtml) console.warn(`   [Converter] Discussion Topic XML "${finalTitle}" description is empty.`);
      } else if (primaryResourceFile && (primaryResourceFile.mimeType === 'text/html' || primaryResourceFile.name.toLowerCase().endsWith('.html')) && typeof primaryResourceFile.data === 'string') {
        // This case handles if a discussion topic resource directly points to an HTML file instead of an XML topic file
        contentHtml = primaryResourceFile.data;
        contentSourceFilePath = primaryResourceFile.name;
        courseworkBase.associatedWithDeveloper!.sourceHtmlFile = primaryResourceFile;
      } else {
        console.warn(`   [Converter] Discussion Topic "${finalTitle}" (ID: ${resourceIdentifier}): Could not find primary content from XML or direct HTML.`);
      }

      if (contentHtml && contentHtml.trim() !== '') {
        const processedHtml = this.processHtmlContent(contentSourceFilePath || '', contentHtml);
        courseworkBase.descriptionForDisplay = processedHtml.descriptionForDisplay;
        courseworkBase.richtext = processedHtml.richtext;
        courseworkBase.localFilesToUpload?.push(...processedHtml.referencedFiles); // Adds images, etc.
        courseworkBase.descriptionForClassroom = processedHtml.descriptionForClassroom;

        // *** MODIFICATION START: Treat discussion prompt HTML as a file for GDoc conversion ***
        if (courseworkBase.richtext) { // Only if the prompt itself is rich
          const promptHtmlFileName = `${this.parsingHelper.sanitizeTopicName(finalTitle)}_prompt.html`;
          const promptHtmlFile: ImsccFile = {
            name: promptHtmlFileName,
            data: processedHtml.descriptionForDisplay, // Use the fully processed HTML
            mimeType: 'text/html'
          };
          // Add this generated HTML file to the list of files to upload
          if (!courseworkBase.localFilesToUpload?.find(f => f.file.name === promptHtmlFileName)) {
            courseworkBase.localFilesToUpload?.push({file: promptHtmlFile, targetFileName: promptHtmlFile.name});
          }
          courseworkBase.convertToGoogleDoc = true; // Mark for potential GDoc conversion

          // Optional: Adjust description to point to the attached prompt,
          // if you don't want the full prompt in the description field anymore.
          // courseworkBase.descriptionForClassroom = `Please see the attached document for the discussion prompt: ${promptHtmlFile.name}.`;
          // courseworkBase.descriptionForDisplay = `<p>Please see the attached document for the discussion prompt: <a href="${promptHtmlFile.name}" data-imscc-local-media-type="file">${promptHtmlFile.name}</a></p><hr/>${processedHtml.descriptionForDisplay}`;
        }
        // *** MODIFICATION END ***


        const plainTextLength = courseworkBase.descriptionForClassroom?.replace(/\s/g, '').length || 0;
        const displayPlainTextLength = (courseworkBase.descriptionForDisplay?.replace(/<[^>]+>/g, '').trim() || '').length;
        if (plainTextLength < 10 && displayPlainTextLength > 0 && !courseworkBase.convertToGoogleDoc) { // Avoid override if we just set it
          courseworkBase.descriptionForClassroom = `Discussion Prompt: "${finalTitle}". See details below.`;
        } else if (plainTextLength > 0 && courseworkBase.descriptionForClassroom.trim().toLowerCase() === finalTitle.trim().toLowerCase() && displayPlainTextLength > 0 && displayPlainTextLength !== finalTitle.trim().length && !courseworkBase.convertToGoogleDoc) {
          courseworkBase.descriptionForClassroom = `Discussion Prompt: ${finalTitle}. See details below.`;
        } else if (!courseworkBase.descriptionForClassroom?.trim() && displayPlainTextLength > 0 && !courseworkBase.convertToGoogleDoc) {
          courseworkBase.descriptionForClassroom = `Discussion Prompt: ${finalTitle}. See formatted content below.`;
        }


      } else { // Fallback if no contentHtml was extracted
        courseworkBase.descriptionForDisplay = `<p>${finalTitle}</p>`;
        courseworkBase.descriptionForClassroom = `Discussion: ${finalTitle}`;
        courseworkBase.richtext = true;
        // If primaryResourceFile (e.g. the empty topic.xml) exists, attach it
        if (primaryResourceFile && courseworkBase.localFilesToUpload && !courseworkBase.localFilesToUpload.some(f => f.file.name === primaryResourceFile!.name)) {
          courseworkBase.localFilesToUpload.push({file: primaryResourceFile, targetFileName: primaryResourceFile.name.split('/').pop() || primaryResourceFile.name});
          courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile; // Or sourceXmlFile if appropriate
        }
      }
    }
    else if (primaryResourceFile && (primaryResourceFile.mimeType === 'text/html' || primaryResourceFile.name.toLowerCase().endsWith('.html'))) {
      // This handles standalone HTML files
      courseworkBase.workType = 'ASSIGNMENT'; // Or MATERIAL
      const htmlSourcePath = primaryResourceFile.name;
      if (typeof primaryResourceFile.data === 'string') {
        const processedHtml = this.processHtmlContent(htmlSourcePath, primaryResourceFile.data);
        courseworkBase.descriptionForDisplay = processedHtml.descriptionForDisplay;
        courseworkBase.descriptionForClassroom = processedHtml.descriptionForClassroom || `Please review the content: ${finalTitle}`;
        courseworkBase.richtext = processedHtml.richtext;
        courseworkBase.localFilesToUpload?.push(...processedHtml.referencedFiles);
        courseworkBase.associatedWithDeveloper!.sourceHtmlFile = primaryResourceFile;
        courseworkBase.convertToGoogleDoc = true; // HTML files are candidates for GDoc conversion
      } else {
        const targetFileName = primaryResourceFile.name.split('/').pop() || primaryResourceFile.name;
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile, targetFileName: targetFileName});
        courseworkBase.descriptionForClassroom = `Please see the attached HTML file: ${targetFileName}`;
        courseworkBase.workType = 'MATERIAL';
        courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
        console.warn(`   [Converter] Primary HTML file "${finalTitle}" data was not string. Attaching file.`);
      }
    }
    else if (resolvedPrimaryHref && (resolvedPrimaryHref.startsWith('http://') || resolvedPrimaryHref.startsWith('https://') || this.specialRefPrefixes.some(prefix => resolvedPrimaryHref!.startsWith(prefix)))) {
      // Handles direct web links or special LMS links that resolve to URLs
      courseworkBase.workType = 'MATERIAL';
      const linkUrl = resolvedPrimaryHref;
      if (!courseworkBase.materials?.some(m => m.link?.url === linkUrl)) {
        courseworkBase.materials?.push({link: {url: linkUrl}});
      }
      if (!courseworkBase.descriptionForClassroom) {
        let cleanLink = linkUrl;
        this.specialRefPrefixes.forEach(prefix => cleanLink = cleanLink.replace(prefix, ''));
        courseworkBase.descriptionForClassroom = `Link: ${finalTitle}${cleanLink ? ` (${cleanLink})` : ''}`;
      }
      // Convert to Google Doc if it's a public URL and not backed by a local file from the package
      if ((resolvedPrimaryHref.startsWith('http://') || resolvedPrimaryHref.startsWith('https://')) && !primaryResourceFile) {
        courseworkBase.convertToGoogleDoc = true;
      }
    }
    else if (primaryResourceFile) {
      // General file attachment
      courseworkBase.workType = 'MATERIAL';
      const targetFileName = primaryResourceFile.name.split('/').pop() || primaryResourceFile.name;
      if (!courseworkBase.localFilesToUpload?.some(f => f.file.name === primaryResourceFile!.name)) {
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile, targetFileName: targetFileName});
      }
      if (!courseworkBase.descriptionForClassroom) courseworkBase.descriptionForClassroom = `Please see the attached file: ${targetFileName}`;
      courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
    }
    else {
      // No primary content found or unhandled type
      console.warn(`   Skipping Resource "${finalTitle}" (ID: ${resourceIdentifier}, Type: ${resourceType || 'N/A'}): Could not determine primary file/link, or it's an unhandled type with no content.`);
      this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: `Unhandled resource type or no primary file/link (${resourceType || 'N/A'})`});
      return of(null);
    }

    // Dependency processing (remains the same)
    const dependencyElements = Array.from(resource.children).filter((node): node is Element => node instanceof Element && node.localName === 'dependency');
    dependencyElements.forEach(dep => {
      const depIdRef = dep.getAttribute('identifierref');
      if (!depIdRef) return;
      const depRes = this.manifestXmlDoc?.querySelector(`resource[identifier="${depIdRef}"]`) || Array.from(this.manifestXmlDoc?.getElementsByTagName('resource') || []).find(r => r.getAttribute('identifier') === depIdRef);
      if (depRes) {
        const depHref = depRes.getAttribute('href');
        const depBaseHref = depRes.getAttribute('xml:base');
        let resolvedDepHrefAttempt: string | null = null;
        let depFile: ImsccFile | null = null;

        if (depHref) {
          resolvedDepHrefAttempt = this.parsingHelper.resolveRelativePath(depBaseHref || baseHref, this.parsingHelper.tryDecodeURIComponent(depHref));
          if (resolvedDepHrefAttempt) {
            depFile = this.fileMap.get(this.getFileMapKey(resolvedDepHrefAttempt)) || null;
          }
          if (!depFile) {
            const resolvedRawDepHref = this.parsingHelper.resolveRelativePath(depBaseHref || baseHref, depHref);
            if (resolvedRawDepHref && resolvedRawDepHref !== resolvedDepHrefAttempt) {
              depFile = this.fileMap.get(this.getFileMapKey(resolvedRawDepHref)) || null;
              if (depFile) resolvedDepHrefAttempt = resolvedRawDepHref;
            }
          }
        }

        if (depFile && depFile.name !== primaryResourceFile?.name &&
          !depFile.mimeType?.startsWith('image/') &&
          !depFile.mimeType?.startsWith('video/') &&
          !depFile.mimeType?.startsWith('text/') && // Allow text/html, text/xml etc. to be handled by main logic if they were primary
          !depFile.mimeType?.includes('xml') &&
          !depFile.mimeType?.includes('html') // Avoid re-adding primary HTML/XML as dependency if it was already processed
        ) {
          const targetFileName = depFile.name.split('/').pop() || depFile.name;
          if (courseworkBase.localFilesToUpload && !courseworkBase.localFilesToUpload.some(f => f.file.name === depFile!.name)) {
            courseworkBase.localFilesToUpload.push({file: depFile, targetFileName: targetFileName});
          }
        } else if (!depFile && resolvedDepHrefAttempt && (resolvedDepHrefAttempt.startsWith('http://') || resolvedDepHrefAttempt.startsWith('https://') || this.specialRefPrefixes.some(prefix => resolvedDepHrefAttempt!.startsWith(prefix)))) {
          const linkUrl = resolvedDepHrefAttempt;
          if (courseworkBase.materials && !courseworkBase.materials.some(m => m.link?.url === linkUrl)) {
            courseworkBase.materials.push({link: {url: linkUrl}});
          }
        } else if (!depFile) {
          // console.warn(`   [Converter] Dependency with identifierref "${depIdRef}" (href: ${depHref || 'N/A'}) could not be resolved to a file or link.`);
        }
      } else {
        // console.warn(`   [Converter] Dependency identifierref "${depIdRef}" does not reference a resource.`);
      }
    });


    const hasContent = !!courseworkBase.descriptionForClassroom?.trim() ||
      !!courseworkBase.descriptionForDisplay?.trim() ||
      !!courseworkBase.qtiFile ||
      (courseworkBase.materials && courseworkBase.materials.length > 0) ||
      (courseworkBase.localFilesToUpload && courseworkBase.localFilesToUpload.length > 0);

    if (!hasContent) {
      console.warn(`   Skipping Resource "${finalTitle}" (ID: ${resourceIdentifier}): Resulted in no processable content after all checks.`);
      this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: 'No processable content found in resource'});
      return of(null);
    }

    return of(courseworkBase as ProcessedCourseWork);
  }


  // processHtmlContent method (remains the same as your original code)
  // For brevity, I'm not repeating it here. It should be included in the actual file.
  private processHtmlContent(
    htmlSourcePath: string,
    htmlString: string
  ): {
    descriptionForDisplay: string;
    descriptionForClassroom: string;
    referencedFiles: Array<{file: ImsccFile; targetFileName: string}>;
    externalLinks: string[];
    richtext: boolean;
  } {
    if (!htmlString) {
      console.warn(`[processHtmlContent] No raw HTML data provided for source: ${htmlSourcePath}`);
      return {descriptionForDisplay: '', descriptionForClassroom: '', referencedFiles: [], externalLinks: [], richtext: false};
    }

    const parser = new DOMParser();
    let cleanHtmlData = htmlString.charCodeAt(0) === 0xFEFF ? htmlString.substring(1) : htmlString;
    cleanHtmlData = this.parsingHelper.preProcessHtmlForDisplay(cleanHtmlData);

    const htmlDoc = parser.parseFromString(cleanHtmlData, 'text/html');
    const contentElement = htmlDoc.body || htmlDoc.documentElement;

    if (!contentElement) {
      console.warn(`[processHtmlContent] Could not parse HTML body or document element for ${htmlSourcePath}.`);
      const errorMsg = `Error: Could not parse HTML content in ${htmlSourcePath}. The original file may need to be attached manually.`;
      return {
        descriptionForDisplay: `<p><i>${errorMsg}</i></p>`,
        descriptionForClassroom: errorMsg,
        referencedFiles: [],
        externalLinks: [],
        richtext: true
      };
    }

    const referencedFiles: Array<{file: ImsccFile; targetFileName: string}> = [];
    const externalLinks: string[] = [];

    let containsRichElements = contentElement.querySelector('img, table, ul, ol, h1, h2, h3, h4, h5, h6, blockquote, pre, code, strong, em, u, s, sub, sup, p, div, span[style], video, iframe') !== null;
    if (!containsRichElements && contentElement.innerHTML.includes('<br')) containsRichElements = true;
    if (!containsRichElements && contentElement.children.length > 0) {
      const simpleTextLength = (contentElement.textContent || '').replace(/\s/g, '').length;
      const htmlLength = contentElement.innerHTML.replace(/\s/g, '').length;
      if (htmlLength > simpleTextLength + 10) { // A bit arbitrary, but catches some formatting
        containsRichElements = true;
      }
    }


    Array.from(contentElement.querySelectorAll('a, img, video')).forEach((el: Element) => {
      const isLink = el.tagName.toUpperCase() === 'A';
      const isImage = el.tagName.toUpperCase() === 'IMG';
      const isVideo = el.tagName.toUpperCase() === 'VIDEO';

      if (isLink || isImage) {
        const attributeName = isLink ? 'href' : 'src';
        const originalRef = el.getAttribute(attributeName);

        if (!originalRef || originalRef.trim() === '' || originalRef === '#') {
          if (isLink && el.textContent?.trim()) {
            const textNode = htmlDoc.createTextNode(el.textContent);
            el.parentNode?.replaceChild(textNode, el);
          } else {el.remove();}
          return;
        }
        if (originalRef.match(/^https?:\/\//i)) {
          if (isLink && !externalLinks.includes(originalRef)) externalLinks.push(originalRef);
          return;
        }
        if (originalRef.match(/^mailto:/i) || originalRef.match(/^tel:/i)) return;
        if (originalRef.match(/^javascript:/i)) {
          if (el.parentNode) {
            const textNode = htmlDoc.createTextNode(el.textContent || 'Removed Scripted Link');
            el.parentNode.replaceChild(textNode, el);
          } else {el.remove();}
          return;
        }
        if (originalRef.match(/^#/i)) { // Internal page anchors
          if (el.parentNode) {
            // Keep the anchor text, but remove the link for now if it's just an internal page anchor
            // Or, one could attempt to map these to Classroom features if applicable (e.g. if content becomes one long doc)
            const textNode = htmlDoc.createTextNode(el.textContent || 'Internal Anchor');
            el.parentNode.replaceChild(textNode, el);
          } else {el.remove();}
          return;
        }
        if (isImage && originalRef.match(/^data:image/i)) return; // Keep data URIs for images

        let file: ImsccFile | null = null;
        let pathForLogging: string = `Original ref: "${originalRef}"`;
        let potentialFileKey = '';
        let resolvedPathForLookup: string | null = null;

        const matchedPrefix = this.specialRefPrefixes.find(p => originalRef.startsWith(p));
        let pathAfterPrefixCleaned = '';
        let baseForResolutionInLoop = htmlSourcePath; // Base for relative paths is the HTML file's own path

        if (matchedPrefix === '$IMS-CC-FILEBASE$') {
          baseForResolutionInLoop = ""; // $IMS-CC-FILEBASE$ is relative to package root
          let pathPart = originalRef.substring(matchedPrefix.length);
          pathForLogging = `($IMS-CC-FILEBASE$) Original path part: "${pathPart}" from "${originalRef}"`;

          let decodedPathPartForResolve = this.parsingHelper.tryDecodeURIComponent(pathPart);
          resolvedPathForLookup = this.parsingHelper.resolveRelativePath(baseForResolutionInLoop, decodedPathPartForResolve);

          if (resolvedPathForLookup) {
            resolvedPathForLookup = resolvedPathForLookup.split(/[?#]/)[0]; // Remove query params/fragments for lookup
            pathForLogging += ` | Resolved to (A): "${resolvedPathForLookup}"`;

            potentialFileKey = this.getFileMapKey(resolvedPathForLookup);
            pathForLogging += ` | Key (A1): "${potentialFileKey}"`;
            file = this.fileMap.get(potentialFileKey) || null;

            if (!file) { // Try common variations if not found
              const pathVariant = resolvedPathForLookup.replace(/ - /g, "_-_").replace(/ /g, "_");
              if (pathVariant !== resolvedPathForLookup) {
                const variantKey = this.getFileMapKey(pathVariant);
                pathForLogging += ` | Key (A2 - variant): "${variantKey}" (from pathVariant: "${pathVariant}")`;
                file = this.fileMap.get(variantKey) || null;
                if (file) resolvedPathForLookup = file.name; // Update if found with variant
              }
            }
          } else { // Resolution from root failed, try using the decoded path part directly
            pathForLogging += ` | Path part resolution from root failed. Trying pathPart as is.`;
            pathAfterPrefixCleaned = decodedPathPartForResolve.split(/[?#]/)[0];
            pathForLogging += ` | Resolved to (B - direct decoded pathPart): "${pathAfterPrefixCleaned}"`;
            potentialFileKey = this.getFileMapKey(pathAfterPrefixCleaned);
            pathForLogging += ` | Key (B1): "${potentialFileKey}"`;
            file = this.fileMap.get(potentialFileKey) || null;
            if (file) resolvedPathForLookup = file.name;
            else {
              // Specific Schoology/Canvas-like fallback if pathPart started with ../resources/ or similar
              if (pathPart.startsWith('../resources/')) {
                const schoologyFilename = pathPart.substring('../resources/'.length).split(/[?#]/)[0];
                const constructedSchoologyPath = 'resources/' + schoologyFilename;
                potentialFileKey = this.getFileMapKey(constructedSchoologyPath);
                pathForLogging += ` | Key (Schoology Specific Fallback): "${potentialFileKey}" for path "${constructedSchoologyPath}"`;
                file = this.fileMap.get(potentialFileKey) || null;
                if (file) resolvedPathForLookup = file.name;
              }
            }
          }

          // Final fallback: search by filename only if path resolution failed
          if (!file) {
            const pathSegmentForFilename = resolvedPathForLookup || this.parsingHelper.tryDecodeURIComponent(pathPart.split(/[?#]/)[0]);
            const filenameOnly = pathSegmentForFilename.split('/').pop();
            if (filenameOnly) {
              const filenameKeyNormalized = this.getFileMapKey(filenameOnly);
              pathForLogging += ` | FilenameKeyA (norm): "${filenameKeyNormalized}"`;
              for (const [keyFromMap, mappedFile] of this.fileMap.entries()) {
                if (keyFromMap.endsWith('/' + filenameKeyNormalized) || keyFromMap === filenameKeyNormalized) {
                  file = mappedFile;
                  resolvedPathForLookup = file.name; // Use the full path from map
                  break;
                }
              }
              if (!file) { // Try filename variant
                const filenameVariant = filenameOnly.replace(/ - /g, "_-_").replace(/ /g, "_");
                if (filenameVariant !== filenameOnly) {
                  const filenameVariantKey = this.getFileMapKey(filenameVariant);
                  pathForLogging += ` | FilenameKeyB (variant): "${filenameVariantKey}" (from filenameVariant: "${filenameVariant}")`;
                  for (const [keyFromMap, mappedFile] of this.fileMap.entries()) {
                    if (keyFromMap.endsWith('/' + filenameVariantKey) || keyFromMap === filenameVariantKey) {
                      file = mappedFile;
                      resolvedPathForLookup = file.name;
                      break;
                    }
                  }
                }
              }
            }
          }
          // console.log(`   [processHtmlContent] $IMS-CC-FILEBASE$ Attempted to match file for link "${originalRef}". Found: ${file?.name}. Path for logging: "${pathForLogging}"`);

        } else if (matchedPrefix === '$WIKI_REFERENCE$' || matchedPrefix === 'WIKI_REFERENCE' || matchedPrefix === '$CANVAS_OBJECT_REFERENCE$' || matchedPrefix === 'CANVAS_OBJECT_REFERENCE') {
          let pathPart = originalRef.substring(matchedPrefix.length);
          if (pathPart.startsWith('/')) pathPart = pathPart.substring(1);
          pathAfterPrefixCleaned = pathPart.split(/[?#]/)[0]; // Remove query params/fragments
          pathForLogging = `(${matchedPrefix}) Path part: "${pathAfterPrefixCleaned}"`;

          const objectIdOrSlug = pathAfterPrefixCleaned.startsWith('pages/') && (matchedPrefix.includes('WIKI')) ?
            pathAfterPrefixCleaned.substring('pages/'.length) :
            pathAfterPrefixCleaned.split('/').pop() || pathAfterPrefixCleaned;

          pathForLogging += ` | Extracted objectIdOrSlug: "${objectIdOrSlug}"`;

          if (this.manifestXmlDoc) {
            let targetResourceHref: string | null = null;
            let foundResourceFromManifest: Element | undefined = undefined;

            // Attempt 1: objectIdOrSlug is an item identifier, find its resource
            const itemEl = Array.from(this.manifestXmlDoc.getElementsByTagName('item')).find(itm => itm.getAttribute('identifier') === objectIdOrSlug);
            if (itemEl) {
              const itemRef = itemEl.getAttribute('identifierref');
              if (itemRef) {
                foundResourceFromManifest = Array.from(this.manifestXmlDoc.getElementsByTagName('resource')).find(r => r.getAttribute('identifier') === itemRef);
                if (foundResourceFromManifest) pathForLogging += ` | Found item by ID "${objectIdOrSlug}", then resource by ref "${itemRef}"`;
              } else { // Item might directly link (less common for these types but check)
                targetResourceHref = itemEl.getAttribute('href');
                if (targetResourceHref) pathForLogging += ` | Found item by ID "${objectIdOrSlug}" with direct href (unusual for this prefix type)`;
              }
            }
            // Attempt 2: objectIdOrSlug is a resource identifier directly
            if (!foundResourceFromManifest && !targetResourceHref) {
              foundResourceFromManifest = Array.from(this.manifestXmlDoc.getElementsByTagName('resource')).find(r => r.getAttribute('identifier') === objectIdOrSlug);
              if (foundResourceFromManifest) pathForLogging += ` | Found resource directly by ID "${objectIdOrSlug}"`;
            }

            if (foundResourceFromManifest) {
              targetResourceHref = foundResourceFromManifest.getAttribute('href');
              if (!targetResourceHref || targetResourceHref.endsWith('/')) { // Check for <file> child if resource href is missing/directory
                const fileEl = foundResourceFromManifest.querySelector('file');
                if (fileEl?.getAttribute('href')) {
                  targetResourceHref = fileEl.getAttribute('href');
                  pathForLogging += ` | Used href from <file> child: "${targetResourceHref}"`;
                }
              }
            }

            if (targetResourceHref) {
              pathForLogging += ` | Manifest href to lookup: "${targetResourceHref}"`;
              // Resolve this href as if it were from the root or a known base
              const decodedTargetHref = this.parsingHelper.tryDecodeURIComponent(targetResourceHref);
              const resolvedFromManifestHref = this.parsingHelper.resolveRelativePath("", decodedTargetHref.split(/[?#]/)[0]); // Assume relative to root

              if (resolvedFromManifestHref) {
                potentialFileKey = this.getFileMapKey(resolvedFromManifestHref);
                file = this.fileMap.get(potentialFileKey) || null;
                if (file) resolvedPathForLookup = file.name;
                else if (matchedPrefix.includes('WIKI')) { // Wiki specific fallback
                  const wikiContentRelative = `wiki_content/${resolvedFromManifestHref.split('/').pop()}`;
                  potentialFileKey = this.getFileMapKey(wikiContentRelative);
                  file = this.fileMap.get(potentialFileKey) || null;
                  if (file) {
                    resolvedPathForLookup = file.name;
                    pathForLogging += ` | Fallback to wiki_content relative: ${potentialFileKey}`;
                  }
                }
              }
            }
          }
          // Fallback to common patterns if manifest lookup fails for WIKI
          if (!file && matchedPrefix.includes('WIKI')) {
            const decodedSlugForPattern = this.parsingHelper.tryDecodeURIComponent(objectIdOrSlug.replace(/-/g, ' '));
            const commonPatternPath = `wiki_content/${decodedSlugForPattern}.html`;
            pathForLogging += ` | Trying WIKI pattern: "${commonPatternPath}"`;
            potentialFileKey = this.getFileMapKey(commonPatternPath);
            file = this.fileMap.get(potentialFileKey) || null;
            if (file) resolvedPathForLookup = file.name;
          }
          // if(file) console.log(`   [processHtmlContent] Matched ${matchedPrefix} to: "${file?.name}" (Path for logging: ${pathForLogging})`);

        }
        else if (originalRef.startsWith('/content/enforced/') || originalRef.startsWith('/content/group/')) { // D2L specific paths
          pathForLogging = `(D2L Content Link) ${originalRef}`;
          try {
            const contentMarker = originalRef.startsWith('/content/enforced/') ? '/content/enforced/' : '/content/group/';
            const pathAfterContentMarker = originalRef.substring(originalRef.indexOf(contentMarker) + contentMarker.length);
            const firstSlashIndex = pathAfterContentMarker.indexOf('/');
            let actualContentPath = pathAfterContentMarker;
            if (firstSlashIndex !== -1) { // Remove the org unit ID part
              actualContentPath = pathAfterContentMarker.substring(firstSlashIndex + 1);
            }
            actualContentPath = actualContentPath.split('?')[0]; // Remove query params

            if (actualContentPath) {
              const decodedContentPath = this.parsingHelper.tryDecodeURIComponent(actualContentPath);
              pathForLogging += ` -> Extracted D2L path: ${decodedContentPath}`;
              // D2L paths are usually directly from root of content, so base is ""
              resolvedPathForLookup = this.parsingHelper.resolveRelativePath("", decodedContentPath);
              if (resolvedPathForLookup) {
                potentialFileKey = this.getFileMapKey(resolvedPathForLookup);
                file = this.fileMap.get(potentialFileKey) || null;
              }

              if (!file) { // Fallback to filename only if full path not found
                const filenameOnly = decodedContentPath.split('/').pop() || "";
                if (filenameOnly && filenameOnly !== decodedContentPath) { // Ensure it's not the same path
                  pathForLogging += ` / filename only: ${filenameOnly}`;
                  // Search globally for this filename
                  for (const [keyFromMap, mappedFile] of this.fileMap.entries()) {
                    if (keyFromMap.endsWith('/' + this.getFileMapKey(filenameOnly)) || this.getFileMapKey(mappedFile.name.split('/').pop() || "") === this.getFileMapKey(filenameOnly)) {
                      file = mappedFile;
                      resolvedPathForLookup = file.name;
                      break;
                    }
                  }
                }
              }
            }
            // if (!file) {
            //   console.warn(`   [processHtmlContent] D2L content link processing did not find a file for: ${originalRef}. Path for logging: ${pathForLogging}`);
            // }
          } catch (e) {
            console.error(`   [processHtmlContent] Error processing D2L content link URL ${originalRef}:`, e);
          }
        }
        else { // General relative path or other unhandled special prefix
          if (!matchedPrefix) { // Standard relative path
            pathAfterPrefixCleaned = originalRef.split(/[?#]/)[0];
            // baseForResolutionInLoop is already htmlSourcePath (directory of current HTML file)
            pathForLogging = `(Relative Path) ${this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleaned)} from base ${baseForResolutionInLoop}`;
          } else { // An unhandled special prefix, assume it might be root-relative
            let pathPart = originalRef.substring(matchedPrefix.length);
            if (pathPart.startsWith('/')) pathPart = pathPart.substring(1);
            pathAfterPrefixCleaned = pathPart.split(/[?#]/)[0];
            baseForResolutionInLoop = ""; // Assume relative to package root for unknown prefixes
            pathForLogging = `(${matchedPrefix} - root relative assumption) ${this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleaned)}`;
          }

          const decodedPathSegment = this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleaned);
          resolvedPathForLookup = this.parsingHelper.resolveRelativePath(baseForResolutionInLoop, decodedPathSegment);

          if (resolvedPathForLookup) {
            potentialFileKey = this.getFileMapKey(resolvedPathForLookup);
            file = this.fileMap.get(potentialFileKey) || null;
            if (file) pathForLogging += ` | Resolved to: ${resolvedPathForLookup}`;
          }

          // If not found with decoded, try with raw path (if different) - sometimes decoding isn't desired
          if (!file && pathAfterPrefixCleaned !== decodedPathSegment && baseForResolutionInLoop === htmlSourcePath) {
            const resolvedRawPath = this.parsingHelper.resolveRelativePath(baseForResolutionInLoop, pathAfterPrefixCleaned);
            if (resolvedRawPath && resolvedRawPath !== resolvedPathForLookup) {
              let alternativeFileKey = this.getFileMapKey(resolvedRawPath);
              file = this.fileMap.get(alternativeFileKey) || null;
              if (file) {
                pathForLogging += ` | Resolved with raw path to: ${resolvedRawPath}`;
                resolvedPathForLookup = resolvedRawPath; // Update to the path that worked
              }
            }
          }
          // Fallback: search by filename only if path resolution failed
          if (!file) {
            const filenameOnly = (resolvedPathForLookup || decodedPathSegment).split('/').pop();
            if (filenameOnly) {
              for (const [keyFromMap, mappedFile] of this.fileMap.entries()) {
                if (keyFromMap.endsWith('/' + this.getFileMapKey(filenameOnly)) || this.getFileMapKey(mappedFile.name.split('/').pop() || "") === this.getFileMapKey(filenameOnly)) {
                  file = mappedFile;
                  resolvedPathForLookup = file.name; // Use the full path from map
                  pathForLogging += ` | Found by filename fallback: ${file.name}`;
                  break;
                }
              }
            }
          }
        }


        if (file) {
          const targetFileName = file.name.split('/').pop() || file.name;
          // Use the resolved path for the anchor, which should be the correct key in fileMap
          const pathForAnchor = resolvedPathForLookup || file.name;


          if (isImage) {
            if (file.mimeType?.startsWith('image/') && typeof file.data === 'string' && file.data.startsWith('data:image')) {
              el.setAttribute('src', file.data); // Use base64 data directly if available
            } else {
              // For images that will be uploaded, replace <img> with a placeholder link for now.
              // The UI or subsequent processing will need to handle these.
              // Or, if you can predict the final Drive URL, you could use it here.
              // For now, point to the local file name that will be uploaded.
              const altText = el.getAttribute('alt') || targetFileName || 'image';
              const newAnchor = htmlDoc.createElement('a');
              newAnchor.href = pathForAnchor; // This href will be relative to the package
              newAnchor.textContent = `Image: ${decode(altText)} (${targetFileName})`;
              newAnchor.setAttribute('data-imscc-local-media-type', 'image');
              newAnchor.setAttribute('data-imscc-original-path', pathForAnchor); // Store original resolved path
              el.parentNode?.replaceChild(newAnchor, el);
              if (!referencedFiles.some(rf => rf.file.name === file!.name)) {
                referencedFiles.push({file, targetFileName: targetFileName});
              }
            }
          } else if (isLink) {
            // Update the href to point to the resolved local file path
            // This path is relative to the package root.
            el.setAttribute('href', pathForAnchor);
            el.setAttribute('data-imscc-original-path', pathForAnchor); // Store original resolved path
            if (file.mimeType?.startsWith('video/')) {
              el.setAttribute('data-imscc-local-media-type', 'video');
              if (!el.textContent?.trim()) el.textContent = `Video: ${targetFileName}`;
            } else if (file.mimeType?.startsWith('image/')) {
              el.setAttribute('data-imscc-local-media-type', 'image');
              if (!el.textContent?.trim()) el.textContent = `Image: ${targetFileName}`;
            } else {
              el.setAttribute('data-imscc-local-media-type', 'file');
              if (!el.textContent?.trim()) el.textContent = targetFileName;
            }
            if (!referencedFiles.some(rf => rf.file.name === file!.name)) {
              referencedFiles.push({file, targetFileName: targetFileName});
            }
          }
        } else { // File not found after all attempts
          console.warn(`   [processHtmlContent] Local file referenced in <${el.tagName}> not found: ${originalRef} (Attempted lookup with path(s): ${pathForLogging})`);
          const span = htmlDoc.createElement('span');
          span.style.cssText = "color: red; border: 1px dashed red; padding: 2px 5px; display: inline-block; font-style: italic;";
          span.textContent = `[Broken Link: ${decode(el.textContent || originalRef)}]`;
          if (el.parentNode) {
            el.parentNode.replaceChild(span, el);
          } else {
            el.remove(); // Should not happen if document is valid
          }
        }

      } else if (isVideo) { // Handling <video> tag
        let elementsToSearchForSrc = Array.from((el as HTMLVideoElement).querySelectorAll('source'));
        if (elementsToSearchForSrc.length === 0 && (el as HTMLVideoElement).hasAttribute('src')) {
          // If no <source> tags, but <video> has a src attribute directly
          const videoSrcAttr = (el as HTMLVideoElement).getAttribute('src');
          if (videoSrcAttr) {
            const tempSource = htmlDoc.createElement('source');
            tempSource.setAttribute('src', videoSrcAttr);
            elementsToSearchForSrc.push(tempSource);
          }
        }

        let localVideoFile: ImsccFile | null = null;
        let firstSourceRefForPlaceholder: string | null = null;
        let resolvedVideoSrcPathForAnchor: string | null = null;

        for (const sourceEl of elementsToSearchForSrc) {
          const originalSrc = sourceEl.getAttribute('src');
          if (!originalSrc) continue;
          if (!firstSourceRefForPlaceholder) firstSourceRefForPlaceholder = originalSrc;

          if (originalSrc.match(/^https?:\/\//i) || originalSrc.match(/^data:video/i)) {
            // External or data URI video, leave as is or handle as external link
            // For simplicity, we might choose to not alter these here or add to externalLinks if it's a full URL
            continue;
          }

          let currentSourceFile: ImsccFile | null = null;
          // Resolve video source similar to images/links
          // (Simplified for brevity, use the same detailed logic as above for 'originalRef')
          const matchedPrefixVideo = this.specialRefPrefixes.find(p => originalSrc.startsWith(p));
          let pathAfterPrefixCleanedVideo = '';
          let baseForResolutionInLoopVideo = htmlSourcePath;

          if (matchedPrefixVideo) {
            let pathPart = originalSrc.substring(matchedPrefixVideo.length);
            if (pathPart.startsWith('/')) pathPart = pathPart.substring(1);
            pathAfterPrefixCleanedVideo = pathPart.split(/[?#]/)[0];
            baseForResolutionInLoopVideo = ""; // Assume root relative for special prefixes
          } else if (originalSrc.startsWith('/content/enforced/') || originalSrc.startsWith('/content/group/')) {
            // D2L path handling
            const contentMarker = originalSrc.startsWith('/content/enforced/') ? '/content/enforced/' : '/content/group/';
            const pathAfterContentMarker = originalSrc.substring(originalSrc.indexOf(contentMarker) + contentMarker.length);
            const firstSlashIndex = pathAfterContentMarker.indexOf('/');
            pathAfterPrefixCleanedVideo = (firstSlashIndex !== -1) ? pathAfterContentMarker.substring(firstSlashIndex + 1) : pathAfterContentMarker;
            pathAfterPrefixCleanedVideo = pathAfterPrefixCleanedVideo.split('?')[0];
            baseForResolutionInLoopVideo = "";
          }
          else { // Standard relative path
            pathAfterPrefixCleanedVideo = originalSrc.split(/[?#]/)[0];
          }

          const decodedRelativePath = this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleanedVideo);
          let resolvedPath = this.parsingHelper.resolveRelativePath(baseForResolutionInLoopVideo, decodedRelativePath);

          if (resolvedPath) {
            currentSourceFile = this.fileMap.get(this.getFileMapKey(resolvedPath)) || null;
          }
          // Add fallbacks for raw path and filename only if needed, similar to image/link logic

          if (currentSourceFile && currentSourceFile.mimeType?.startsWith('video/')) {
            localVideoFile = currentSourceFile;
            resolvedVideoSrcPathForAnchor = resolvedPath || localVideoFile.name; // Use the path that worked
            if (!referencedFiles.some(rf => rf.file.name === localVideoFile!.name)) {
              referencedFiles.push({file: localVideoFile, targetFileName: localVideoFile.name.split('/').pop() || localVideoFile.name});
            }
            break; // Found a suitable local video source
          }
        }

        const videoTitle = decode((el as HTMLVideoElement).getAttribute('title') || localVideoFile?.name.split('/').pop() || firstSourceRefForPlaceholder || 'Video');
        if (localVideoFile && resolvedVideoSrcPathForAnchor) {
          const newAnchor = htmlDoc.createElement('a');
          newAnchor.href = resolvedVideoSrcPathForAnchor; // Path relative to package
          newAnchor.textContent = `Video: ${videoTitle}`;
          newAnchor.setAttribute('data-imscc-local-media-type', 'video');
          newAnchor.setAttribute('data-imscc-original-path', resolvedVideoSrcPathForAnchor);
          el.parentNode?.replaceChild(newAnchor, el);
          containsRichElements = true;
        } else {
          // If no local video file found, replace <video> with a placeholder or link to the first source
          const refToShow = this.parsingHelper.tryDecodeURIComponent(firstSourceRefForPlaceholder || "unknown video source");
          const span = htmlDoc.createElement('span');
          span.style.cssText = "color: #555; border: 1px dashed #ccc; padding: 2px 5px; display: inline-block; font-style: italic;";
          if (firstSourceRefForPlaceholder && firstSourceRefForPlaceholder.match(/^https?:\/\//i)) {
            const externalVideoLink = htmlDoc.createElement('a');
            externalVideoLink.href = firstSourceRefForPlaceholder;
            externalVideoLink.textContent = `External Video: ${videoTitle}`;
            externalVideoLink.target = "_blank";
            span.appendChild(externalVideoLink);
            if (!externalLinks.includes(firstSourceRefForPlaceholder)) externalLinks.push(firstSourceRefForPlaceholder);
          } else {
            span.textContent = `[Video: ${videoTitle} - local source not found. Original ref: ${refToShow}]`;
          }
          el.parentNode?.replaceChild(span, el);
          // console.warn(`   [processHtmlContent] Local video source(s) not found for video element. First original ref for placeholder: ${refToShow}`);
        }
      }
    });

    // Process <iframe> elements
    Array.from(contentElement.querySelectorAll('iframe')).forEach((iframeEl: Element) => {
      const iframeSrc = iframeEl.getAttribute('src');
      const iframeTitle = iframeEl.getAttribute('title') || 'Embedded Content';

      if (iframeSrc) {
        const newAnchor = htmlDoc.createElement('a');
        newAnchor.href = iframeSrc; // Keep original src
        newAnchor.textContent = `Embedded Content: ${decode(iframeTitle)} (Link: ${iframeSrc})`;
        newAnchor.target = '_blank'; // Open in new tab
        newAnchor.rel = 'noopener noreferrer';

        if (iframeSrc.match(/^https?:\/\//i) && !externalLinks.includes(iframeSrc)) {
          externalLinks.push(iframeSrc);
        }
        // Replace iframe with a paragraph containing the link
        const p = htmlDoc.createElement('p');
        p.appendChild(document.createTextNode('[An iframe was present here, linking to: '));
        p.appendChild(newAnchor);
        p.appendChild(document.createTextNode(']'));
        p.style.border = "1px solid #ccc";
        p.style.padding = "10px";
        p.style.backgroundColor = "#f9f9f9";


        if (iframeEl.parentNode) {
          iframeEl.parentNode.replaceChild(p, iframeEl);
        } else { // Should not happen in valid HTML structure
          contentElement.appendChild(p);
        }
        containsRichElements = true;
      } else {
        // If iframe has no src, remove it or replace with placeholder
        const placeholderText = htmlDoc.createTextNode(`[Unsupported embedded content: ${decode(iframeTitle || 'Untitled Iframe')}]`);
        if (iframeEl.parentNode) {
          iframeEl.parentNode.replaceChild(placeholderText, iframeEl);
        }
        // console.warn(`   [processHtmlContent] Removed <iframe> with no src attribute (title: ${decode(iframeTitle || 'N/A')}).`);
      }
    });


    // Serialize the modified HTML back to a string
    const descriptionForDisplay = decode(contentElement.innerHTML); // Decode entities one last time for consistency

    // Generate plain text description for Classroom API
    const tempDiv = htmlDoc.createElement('div');
    tempDiv.innerHTML = descriptionForDisplay; // Use the fully processed HTML for text extraction

    // Improve plain text extraction: try to respect line breaks from <p>, <br>, <li>
    let classroomDesc = "";
    function extractTextWithLineBreaks(node: Node) {
      if (node.nodeType === Node.TEXT_NODE) {
        classroomDesc += node.textContent;
      } else if (node.nodeType === Node.ELEMENT_NODE) {
        const el = node as Element;
        const tagName = el.tagName.toLowerCase();
        if (tagName === 'br' || tagName === 'p' || tagName === 'div' || tagName === 'li' || tagName.match(/^h[1-6]$/)) {
          if (classroomDesc.length > 0 && !classroomDesc.endsWith('\n')) {
            classroomDesc += '\n'; // Add newline before these block elements if not already there
          }
        }
        for (let i = 0; i < node.childNodes.length; i++) {
          extractTextWithLineBreaks(node.childNodes[i]);
        }
        if (tagName === 'p' || tagName === 'div' || tagName === 'li' || tagName.match(/^h[1-6]$/)) {
          if (!classroomDesc.endsWith('\n')) classroomDesc += '\n'; // Add newline after
        }
      }
    }
    extractTextWithLineBreaks(tempDiv);
    classroomDesc = classroomDesc.replace(/\n\s*\n/g, '\n').replace(/\s+/g, ' ').trim(); // Normalize whitespace and multiple newlines


    const maxDescLength = 25000; // Classroom API limit (approx)
    if (classroomDesc.length > maxDescLength) {
      let truncated = classroomDesc.substring(0, maxDescLength - 20); // Leave room for "..." and buffer
      const lastSensibleBreak = Math.max(truncated.lastIndexOf('.'), truncated.lastIndexOf('!'), truncated.lastIndexOf('?'));
      if (lastSensibleBreak > maxDescLength * 0.8) { // Only truncate at sentence end if it's reasonably far
        truncated = truncated.substring(0, lastSensibleBreak + 1);
      } else { // Otherwise, just cut
        truncated = truncated.substring(0, maxDescLength - 3);
      }
      classroomDesc = truncated + "...";
    }

    return {
      descriptionForDisplay,
      descriptionForClassroom: classroomDesc,
      referencedFiles,
      externalLinks,
      richtext: containsRichElements || referencedFiles.length > 0 || externalLinks.length > 0
    };
  }

}
