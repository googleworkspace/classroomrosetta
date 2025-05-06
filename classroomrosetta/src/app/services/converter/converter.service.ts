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
import {Observable, from, of, throwError, EMPTY, concat, Subject} from 'rxjs';
import {map, tap, catchError, concatMap, filter, mergeMap, toArray} from 'rxjs/operators';
import {
  ProcessedCourseWork,
  ImsccFile,
} from '../../interfaces/classroom-interface'; // Adjust path
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

  private fileMap: Map<string, ImsccFile> = new Map(); // Map for quick file lookup
  private manifestXmlDoc: XMLDocument | null = null; // Store parsed manifest

  constructor() { }

  convertImscc(files: ImsccFile[]): Observable<ProcessedCourseWork> {
    console.log("Starting IMSCC conversion process (Main Thread)...");
    this.skippedItemLog = [];
    this.fileMap = new Map(); // Reset map for a new conversion
    this.manifestXmlDoc = null; // Reset manifest doc

    // --- 1. Process Files and Build Map ---
    const processedFiles = files.map(file => {
       let processedFile = file;
      // Process image files to base64 data URLs if they are ArrayBuffers
      if (file.mimeType?.startsWith('image/') && this.parsingHelper.isArrayBuffer(file.data)) {
        try {
          const byteNumbers = new Array(file.data.byteLength);
          const byteArray = new Uint8Array(file.data);
          for (let i = 0; i < file.data.byteLength; i++) { byteNumbers[i] = byteArray[i]; }
          const binaryString = String.fromCharCode.apply(null, byteNumbers);
          const base64String = btoa(binaryString);
          processedFile = {...file, data: `data:${file.mimeType};base64,${base64String}`};
        } catch (e) {
          console.warn(`Could not convert ArrayBuffer to base64 for image ${file.name}`, e);
          // processedFile remains original
        }
      }
      // For text-based files, attempt decoding ArrayBuffer to string
      else if (this.parsingHelper.isArrayBuffer(file.data) && !file.mimeType?.startsWith('image/')) {
        try {
          const textDecoder = new TextDecoder('utf-8'); // Assuming UTF-8 is common
          const textData = textDecoder.decode(file.data);
          console.log(`Decoded ArrayBuffer as text for file: ${file.name}`);
          processedFile = {...file, data: textData};
        } catch (e) {
          console.warn(`Could not decode ArrayBuffer as text for file ${file.name}`, e);
          // processedFile remains original
        }
      }

      // Add the processed file to the map using a normalized key
      // Keys should be lowercase, decoded, and have cyrillic 'c' corrected
      const normalizedFileName = this.parsingHelper.correctCyrillicCPath(this.parsingHelper.tryDecodeURIComponent(processedFile.name)).toLowerCase();
      this.fileMap.set(normalizedFileName, processedFile);

      return processedFile; // Return the (potentially modified) file for completeness if needed elsewhere
    });

    const manifestFile = this.fileMap.get(this.parsingHelper.correctCyrillicCPath(this.parsingHelper.tryDecodeURIComponent('imsmanifest.xml')).toLowerCase());

     if (!manifestFile || typeof manifestFile.data !== 'string') {
         console.error('IMS Manifest file (imsmanifest.xml) not found or data is not a string AFTER file processing.');
         return throwError(() => new Error('imsmanifest.xml not found or data is not a string.'));
     }


    // --- 2. Parse Manifest ---
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
      this.manifestXmlDoc = doc; // Store the parsed manifest

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

    // --- 3. Process Items or Direct Resources ---
    let processingStream: Observable<ProcessedCourseWork>;
    if (rootItems.length > 0) {
      // Pass parentTopic, xmlDoc and fileMap are accessed via 'this'
      processingStream = this.processImsccItemsStream(rootItems, undefined);
    } else if (directResources.length > 0) {
      processingStream = from(directResources).pipe(
        concatMap(resource => {
          try {
            const title = resource.getAttribute('title') || this.parsingHelper.extractTitleFromMetadata(resource) || 'Untitled Resource';
            const identifier = resource.getAttribute('identifier') || `resource_${Math.random().toString(36).substring(2)}`;
            // Pass parentTopic, xmlDoc and fileMap are accessed via 'this'
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
      tap(item => console.log(` -> Emitting processed item: "${item.title}" (Type: ${item.workType}, ID: ${item.associatedWithDeveloper?.id})`)),
      catchError(err => {
        console.error("Error during IMSCC content processing stream:", err);
        const wrappedError = err instanceof Error ? err : new Error(String(err));
        return throwError(() => new Error(`Error processing IMSCC content stream: ${wrappedError.message}`));
      })
    );
  }

  private processImsccItemsStream(
    items: Element[],
    parentTopic?: string // xmlDoc and fileMap are class properties
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
            // Access manifestXmlDoc via 'this'
            const resource = this.manifestXmlDoc?.querySelector(resourceSelector) ||
                             Array.from(this.manifestXmlDoc?.getElementsByTagName('resource') || []).find(r => r.getAttribute('identifier') === identifierRef);

            if (resource) {
               // Pass parentTopic, access xmlDoc and fileMap via 'this'
              resourceObservable = this.processResource(resource, rawTitle, identifier, parentTopic);
            } else {
              console.warn(`   Resource not found for identifierref: ${identifierRef} (Item: "${rawTitle}"). This item might be a folder or a broken link.`);
              this.skippedItemLog.push({id: identifier, title: rawTitle, reason: `Resource not found for ref: ${identifierRef}`});
            }
          } else {
            if (childItems.length === 0) {
              console.log(`   Item "${rawTitle}" (ID: ${identifier}) has no identifierref and no child items. Skipping as a standalone processed item.`);
              this.skippedItemLog.push({id: identifier, title: rawTitle, reason: 'No resource reference and no child items'});
            } else {
              console.log(`   Item "${rawTitle}" (ID: ${identifier}) is a container/topic. Processing sub-items.`);
            }
          }

          let childItemsObservable: Observable<ProcessedCourseWork> = EMPTY;
          if (childItems.length > 0) {
             // Pass sanitizedTopicName as the new parentTopic
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
    itemTitle: string, // Title from the referring <item>
    imsccIdentifier: string, // ID from the referring <item>
    parentTopic?: string // xmlDoc and fileMap are class properties
  ): Observable<ProcessedCourseWork | null> {
    const resourceIdentifier = resource.getAttribute('identifier');
    const resourceType = resource.getAttribute('type');
    const resourceHref = resource.getAttribute('href'); // Primary file/link of the resource
    const baseHref = resource.getAttribute('xml:base'); // Base path for resolving resource href

    // Prioritize title from resource metadata, then resource attribute, then item title
    const resourceOwnTitle = this.parsingHelper.extractTitleFromMetadata(resource) || resource.getAttribute('title');
    const finalTitle = resourceOwnTitle || itemTitle;

    console.log(`   [Converter] Processing Resource "${finalTitle}" (ID: ${resourceIdentifier}, Type: ${resourceType || 'N/A'}) referenced by Item ID: ${imsccIdentifier}`);

    if (!this.manifestXmlDoc) { // Should not happen based on convertImscc logic, but safeguard
         console.error(`   [Converter] processResource called before manifestXmlDoc is set.`);
         this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: 'Internal Error: Manifest not parsed'});
         return of(null);
    }

    // Skip known ignorable D2L resource types
    const d2lMaterialType = resource.getAttributeNS(this.D2L_V2P0_NS, 'material_type');
    if (d2lMaterialType === 'orgunitconfig') {
      console.log(`   Skipping D2L orgunitconfig resource: "${finalTitle}" (ID: ${resourceIdentifier})`);
      this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: 'D2L orgunitconfig'});
      return of(null);
    }

    // Initialize the base coursework object
    let courseworkBase: Partial<ProcessedCourseWork> = {
      title: finalTitle,
      state: 'DRAFT', // Always import as draft
      materials: [],
      localFilesToUpload: [],
      associatedWithDeveloper: {
        id: imsccIdentifier, // The item identifier
        resourceId: resourceIdentifier, // The resource identifier
        topic: parentTopic, // The Google Classroom topic name
      },
      descriptionForDisplay: '', // HTML content for internal display
      descriptionForClassroom: '', // Plain text for Classroom API description
      richtext: false, // Does descriptionForDisplay contain rich text?
      workType: 'ASSIGNMENT' // Default type, overridden below
    };

    // --- Identify the primary file/link referenced by the resource ---
    let primaryResourceFile: ImsccFile | null = null;
    let resolvedPrimaryHref: string | null = null;
    let primaryFileXmlDoc: XMLDocument | null = null;

    let primaryFilePathOrUrl = resourceHref;
    if (!primaryFilePathOrUrl) {
      const firstFileElement = Array.from(resource.children).find(node => node instanceof Element && node.localName === 'file') as Element | undefined;
      primaryFilePathOrUrl = firstFileElement?.getAttribute('href') || null;
      if (primaryFilePathOrUrl) {
        console.log(`   [Converter] No resource href, using href from first <file> element: ${primaryFilePathOrUrl}`);
      }
    }

    // Resolve path and find the corresponding ImsccFile object if it's a local file
    if (primaryFilePathOrUrl && !primaryFilePathOrUrl.match(/^https?:\/\//i) && !this.specialRefPrefixes.some(prefix => primaryFilePathOrUrl!.startsWith(prefix))) {
      resolvedPrimaryHref = this.parsingHelper.resolveRelativePath(baseHref, primaryFilePathOrUrl);
      if (resolvedPrimaryHref) {
        // Use the file map for efficient lookup
        const normalizedResolvedPath = this.parsingHelper.correctCyrillicCPath(this.parsingHelper.tryDecodeURIComponent(resolvedPrimaryHref)).toLowerCase();
        primaryResourceFile = this.fileMap.get(normalizedResolvedPath) || null;

        if (!primaryResourceFile) {
          console.warn(`   [Converter] Referenced file not found in package: ${primaryFilePathOrUrl} (Resolved: ${resolvedPrimaryHref})`);
        } else {
          console.log(`   [Converter] Found primary resource file: ${primaryResourceFile.name}`);
          // If it's an XML file with string data, try parsing it immediately
          if ((primaryResourceFile.name.toLowerCase().endsWith('.xml') || primaryResourceFile.mimeType?.includes('xml')) && typeof primaryResourceFile.data === 'string') {
            try {
              console.log(`   [Converter] Attempting to parse primary file XML: ${primaryResourceFile.name} (MIME: ${primaryResourceFile.mimeType})`);
              const parser = new DOMParser();
              const cleanXmlData = primaryResourceFile.data.charCodeAt(0) === 0xFEFF ? primaryResourceFile.data.substring(1) : primaryResourceFile.data;
              primaryFileXmlDoc = parser.parseFromString(cleanXmlData, "application/xml");
              if (primaryFileXmlDoc.querySelector('parsererror')) {
                console.warn(`   [Converter] XML parsing error for ${primaryResourceFile.name}.`);
                primaryFileXmlDoc = null;
              } else {
                console.log(`   [Converter] Successfully parsed XML for ${primaryResourceFile.name}.`);
              }
            } catch (e) {
              console.error(`   [Converter] Exception parsing primary file XML for ${primaryResourceFile.name}:`, e);
              primaryFileXmlDoc = null;
            }
          }
        }
      } else {
        console.warn(`   [Converter] Could not resolve primary resource href: ${primaryFilePathOrUrl}`);
      }
    } else if (primaryFilePathOrUrl && primaryFilePathOrUrl.match(/^https?:\/\//i)) {
      resolvedPrimaryHref = primaryFilePathOrUrl;
      console.log(`   [Converter] Primary resource is an external URL: ${resolvedPrimaryHref}`);
    } else if (primaryFilePathOrUrl && this.specialRefPrefixes.some(prefix => primaryFilePathOrUrl!.startsWith(prefix))) {
      console.log(`   [Converter] Primary resource uses a special prefix: ${primaryFilePathOrUrl}. Treating as external or unhandled link.`);
      resolvedPrimaryHref = primaryFilePathOrUrl;
    } else {
      console.warn(`   [Converter] No primary href or file found/resolvable for resource ID: ${resourceIdentifier}`);
    }


    // --- Determine the coursework type and content ---
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
        console.log(`   [Converter] Identified QTI/Assessment: "${finalTitle}" (Resource ID: ${resourceIdentifier}). Attached QTI file.`);
      } else if (primaryResourceFile) {
        console.warn(`   [Converter] QTI/Assessment "${finalTitle}" (Resource ID: ${resourceIdentifier}) main file (${primaryResourceFile.name}) could not be parsed as XML or was not XML. Attaching as general file.`);
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile, targetFileName: primaryResourceFile.name.split('/').pop() || primaryResourceFile.name});
        courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
      } else {
        console.warn(`   Skipping QTI/Assessment resource "${finalTitle}" (ID: ${resourceIdentifier}): No valid primary file found.`);
        this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: 'QTI/Assessment - No valid primary file'});
        return of(null);
      }
    }
    else if (primaryResourceFile && primaryFileXmlDoc && this.parsingHelper.isWebLinkXml(primaryResourceFile, primaryFileXmlDoc)) {
      const extractedUrl = this.parsingHelper.extractWebLinkUrl(primaryResourceFile!, primaryFileXmlDoc!);
      if (extractedUrl) {
        courseworkBase.webLinkUrl = extractedUrl;
        courseworkBase.workType = 'ASSIGNMENT'; // Or 'MATERIAL'? Depends on how you want to represent weblinks
        if (!courseworkBase.materials?.some(m => m.link?.url === extractedUrl)) {
          courseworkBase.materials?.push({link: {url: extractedUrl}});
        }
        if (!courseworkBase.descriptionForClassroom) courseworkBase.descriptionForClassroom = `Please follow this link: ${finalTitle}`;
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
        console.log(`   [Converter] Identified WebLink: "${finalTitle}" -> ${extractedUrl} (Resource ID: ${resourceIdentifier}).`);
      } else {
        console.warn(`   [Converter] WebLink XML "${finalTitle}" (Resource ID: ${resourceIdentifier}) found but could not extract URL. Attaching XML file.`);
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile!, targetFileName: primaryResourceFile!.name.split('/').pop() || primaryResourceFile!.name});
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
        courseworkBase.workType = 'MATERIAL';
      }
    }
    else if (isDiscussionTopic) {
      courseworkBase.workType = 'SHORT_ANSWER_QUESTION';
      console.log(`   [Converter] Identified Discussion Topic: "${finalTitle}" (Resource ID: ${resourceIdentifier}).`);

      let contentHtml: string | null = null;
      let contentSourceFilePath: string | null = null; // Use the file name/path as the source path

      if (primaryResourceFile && primaryFileXmlDoc && this.parsingHelper.isTopicXml(primaryResourceFile, primaryFileXmlDoc)) {
        contentHtml = this.parsingHelper.extractTopicDescriptionHtml(primaryResourceFile, primaryFileXmlDoc);
        contentSourceFilePath = primaryResourceFile.name;
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
        console.log(`   [Converter] Discussion Topic "${finalTitle}": Main file is topic XML.`);
        if (!contentHtml) {
          console.warn(`   [Converter] ...but extracted HTML description from XML is empty.`);
        }
      } else if (primaryResourceFile && (primaryResourceFile.mimeType === 'text/html' || primaryResourceFile.name.toLowerCase().endsWith('.html')) && typeof primaryResourceFile.data === 'string') {
        contentHtml = primaryResourceFile.data;
        contentSourceFilePath = primaryResourceFile.name;
        courseworkBase.associatedWithDeveloper!.sourceHtmlFile = primaryResourceFile;
        console.log(`   [Converter] Discussion Topic "${finalTitle}": Main file is HTML (as fallback content source).`);
      } else {
        console.warn(`   [Converter] Discussion Topic "${finalTitle}" (Resource ID: ${resourceIdentifier}): Could not find primary content in XML <text> or main HTML file.`);
      }

      // Now, process the found contentHtml (if any)
      if (contentHtml && contentHtml.trim() !== '') {
        console.log(`   [Converter] Processing HTML content for Discussion Topic "${finalTitle}". Length: ${contentHtml.length}`);
        // Pass the source file path and the HTML string
        const processedHtml = this.processHtmlContent(contentSourceFilePath || '', contentHtml); // Access fileMap via 'this'

        courseworkBase.descriptionForDisplay = processedHtml.descriptionForDisplay;
        courseworkBase.richtext = processedHtml.richtext;
        courseworkBase.localFilesToUpload?.push(...processedHtml.referencedFiles);
        processedHtml.externalLinks.forEach(linkUrl => {
          if (!courseworkBase.materials?.some(m => m.link?.url === linkUrl)) courseworkBase.materials?.push({link: {url: linkUrl}});
        });

        courseworkBase.descriptionForClassroom = processedHtml.descriptionForClassroom;

        // Refined fallback logic for Classroom description
        const plainTextLength = courseworkBase.descriptionForClassroom?.replace(/\s/g, '').length || 0;
        const displayPlainTextLength = (courseworkBase.descriptionForDisplay?.replace(/<[^>]+>/g, '').trim() || '').length;

        if (plainTextLength < 10 && displayPlainTextLength > 0) {
           courseworkBase.descriptionForClassroom = `Discussion Prompt: "${finalTitle}". See details below.`;
        } else if (plainTextLength > 0 && courseworkBase.descriptionForClassroom.trim().toLowerCase() === finalTitle.trim().toLowerCase() && displayPlainTextLength > 0 && displayPlainTextLength !== finalTitle.trim().length) {
            // Only add "See details below" hint if there's actual different content in the formatted version
            courseworkBase.descriptionForClassroom = `Discussion Prompt: ${finalTitle}. See details below.`;
        } else if (!courseworkBase.descriptionForClassroom.trim() && displayPlainTextLength > 0) {
             // If plain text is empty, but display HTML has content, provide a generic hint
             courseworkBase.descriptionForClassroom = `Discussion Prompt: ${finalTitle}. See formatted content below.`;
        }
         // Otherwise, use the extracted plain text directly (including if it's identical to title but has no extra HTML formatting)


        console.log(`   [Converter] Final descriptionForClassroom for "${finalTitle}": ${courseworkBase.descriptionForClassroom}`);
        console.log(`   [Converter] Final descriptionForDisplay for "${finalTitle}": ${courseworkBase.descriptionForDisplay.substring(0, 100)}...`);

      } else {
        console.log(`   [Converter] Discussion Topic "${finalTitle}" (Resource ID: ${resourceIdentifier}): No extractable HTML content found.`);
        courseworkBase.descriptionForDisplay = `<p>${finalTitle}</p>`;
        courseworkBase.descriptionForClassroom = `Discussion: ${finalTitle}`;
        courseworkBase.richtext = true;
        if (primaryResourceFile && courseworkBase.localFilesToUpload && !courseworkBase.localFilesToUpload.some(f => f.file.name === primaryResourceFile!.name)) {
          console.log(`   [Converter] Attaching primary file ${primaryResourceFile.name} as fallback for discussion topic.`);
          courseworkBase.localFilesToUpload.push({file: primaryResourceFile, targetFileName: primaryResourceFile.name.split('/').pop() || primaryResourceFile.name});
          courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
        }
      }
    }
    else if (primaryResourceFile && (primaryResourceFile.mimeType === 'text/html' || primaryResourceFile.name.toLowerCase().endsWith('.html'))) {
      courseworkBase.workType = 'ASSIGNMENT'; // Or 'MATERIAL'? Depends on common use cases. 'ASSIGNMENT' is a safe default.
      const htmlSourcePath = primaryResourceFile.name;
      if (typeof primaryResourceFile.data === 'string') {
        console.log(`   [Converter] Identified HTML file: "${finalTitle}" (Resource ID: ${resourceIdentifier}).`);
        const processedHtml = this.processHtmlContent(htmlSourcePath, primaryResourceFile.data); // Pass htmlSourcePath and data
        courseworkBase.descriptionForDisplay = processedHtml.descriptionForDisplay;
        courseworkBase.descriptionForClassroom = processedHtml.descriptionForClassroom || `Please review the content: ${finalTitle}`;
        courseworkBase.richtext = processedHtml.richtext;
        courseworkBase.localFilesToUpload?.push(...processedHtml.referencedFiles);
        processedHtml.externalLinks.forEach(linkUrl => {
          if (!courseworkBase.materials?.some(m => m.link?.url === linkUrl)) courseworkBase.materials?.push({link: {url: linkUrl}});
        });
        courseworkBase.associatedWithDeveloper!.sourceHtmlFile = primaryResourceFile;
        console.log(`   [Converter] Processed HTML content for "${finalTitle}". Description length: ${courseworkBase.descriptionForClassroom.length}`);
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
      console.log(`   [Converter] Identified External Link/Special Ref: "${finalTitle}" -> ${linkUrl}.`);
    }
    else if (primaryResourceFile) {
      courseworkBase.workType = 'MATERIAL';
      const targetFileName = primaryResourceFile.name.split('/').pop() || primaryResourceFile.name;
      if (!courseworkBase.localFilesToUpload?.some(f => f.file.name === primaryResourceFile!.name)) {
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile, targetFileName: targetFileName});
      }
      if (!courseworkBase.descriptionForClassroom) courseworkBase.descriptionForClassroom = `Please see the attached file: ${targetFileName}`;
      courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
      console.log(`   [Converter] Identified General File Material: "${finalTitle}" (Resource ID: ${resourceIdentifier}). File: ${primaryResourceFile.name}.`);
    }
    else {
      console.warn(`   Skipping Resource "${finalTitle}" (ID: ${resourceIdentifier}, Type: ${resourceType || 'N/A'}): Could not determine primary file/link, or it's an unhandled type with no content.`);
      this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: `Unhandled resource type or no primary file/link (${resourceType || 'N/A'})`});
      return of(null);
    }

    // --- Process Dependencies ---
    const dependencyElements = Array.from(resource.children).filter((node): node is Element => node instanceof Element && node.localName === 'dependency');
    dependencyElements.forEach(dep => {
      const depIdRef = dep.getAttribute('identifierref');
      if (!depIdRef) return;

      // Access manifestXmlDoc via 'this'
      const depRes = this.manifestXmlDoc?.querySelector(`resource[identifier="${depIdRef}"]`) || Array.from(this.manifestXmlDoc?.getElementsByTagName('resource') || []).find(r => r.getAttribute('identifier') === depIdRef);
      if (depRes) {
        const depHref = depRes.getAttribute('href');
        const depBaseHref = depRes.getAttribute('xml:base');
        const resolvedDepHref = depHref ? this.parsingHelper.resolveRelativePath(depBaseHref || baseHref, depHref) : null;

        if (resolvedDepHref) {
          // Use file map for lookup
          const normalizedResolvedDepPath = this.parsingHelper.correctCyrillicCPath(this.parsingHelper.tryDecodeURIComponent(resolvedDepHref)).toLowerCase();
          const depFile = this.fileMap.get(normalizedResolvedDepPath) || null;

          if (depFile && typeof depFile.data !== 'string') { // Check if data is ArrayBuffer (binary)
            // Explicitly check types to exclude those typically processed differently
            if (!depFile.mimeType?.startsWith('text/') &&
              !depFile.mimeType?.startsWith('application/xml') &&
              !depFile.mimeType?.startsWith('image/')) {
                const targetFileName = depFile.name.split('/').pop() || depFile.name;
                if (courseworkBase.localFilesToUpload && !courseworkBase.localFilesToUpload.some(f => f.file.name === depFile!.name)) {
                  courseworkBase.localFilesToUpload.push({file: depFile, targetFileName: targetFileName});
                  console.log(`   [Converter] Added dependency file for upload: ${depFile.name}`);
                }
              } else {
                console.log(`   [Converter] Dependency file ${depFile.name} is text/xml/image type (string data), skipping direct attachment via dependency.`);
              }
          } else if (!depFile && (resolvedDepHref.startsWith('http://') || resolvedDepHref.startsWith('https://') || this.specialRefPrefixes.some(prefix => resolvedDepHref!.startsWith(prefix)))) {
            const linkUrl = resolvedDepHref;
            if (courseworkBase.materials && !courseworkBase.materials.some(m => m.link?.url === linkUrl)) {
              courseworkBase.materials.push({link: {url: linkUrl}});
              console.log(`   [Converter] Added dependency link to materials: ${linkUrl}`);
            }
          } else if (depFile && typeof depFile.data === 'string') {
             console.log(`   [Converter] Dependency file ${depFile.name} has string data, skipping direct attachment via dependency.`);
          } else {
            console.warn(`   [Converter] Dependency with identifierref "${depIdRef}" could not be resolved to a file or link.`);
          }
        } else {
          console.warn(`   [Converter] Could not resolve href for dependency with identifierref "${depIdRef}".`);
        }
      } else {
        console.warn(`   [Converter] Dependency identifierref "${depIdRef}" does not reference a resource.`);
      }
    });


    // Final check
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

    console.log(`   [Converter] Successfully processed resource "${finalTitle}".`);
    return of(courseworkBase as ProcessedCourseWork);
  }


  private processHtmlContent(
    htmlSourcePath: string, // Path of the source file (for relative link resolution)
    htmlString: string // The raw HTML string content
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

    let containsRichElements = contentElement.querySelector('img, table, ul, ol, h1, h2, h3, h4, h5, h6, blockquote, pre, code, strong, em, u, s, sub, sup, p, div, span[style]') !== null;
    if (!containsRichElements && contentElement.innerHTML.includes('<br')) containsRichElements = true;
     if (!containsRichElements && contentElement.children.length > 0) {
        const simpleTextLength = (contentElement.textContent || '').replace(/\s/g, '').length;
        const htmlLength = contentElement.innerHTML.replace(/\s/g, '').length;
         if (htmlLength > simpleTextLength + 10) {
              containsRichElements = true;
         }
    }


    Array.from(contentElement.querySelectorAll('a, img')).forEach((el: Element) => {
      const isLink = el.tagName.toUpperCase() === 'A';
      const isImage = el.tagName.toUpperCase() === 'IMG';
      const attributeName = isLink ? 'href' : (isImage ? 'src' : null);

      if (!attributeName) return;

      const originalRefValue = el.getAttribute(attributeName);
      if (!originalRefValue || originalRefValue.trim() === '' || originalRefValue === '#') {
        el.remove();
        return;
      }

      if (originalRefValue.match(/^https?:\/\//i)) {
        if (isLink && !externalLinks.includes(originalRefValue)) {
          externalLinks.push(originalRefValue);
        } else if (isImage) {
        }
        return;
      }

      if (originalRefValue.match(/^data:image/i) && isImage) {
        return;
      }

      if (originalRefValue.match(/^mailto:/i) || originalRefValue.match(/^tel:/i) || originalRefValue.match(/^javascript:/i)) {
        return;
      }
      if (originalRefValue.match(/^#/i)) {
        if (el.parentNode) {
          const textNode = htmlDoc.createTextNode(el.textContent || '');
          el.parentNode.replaceChild(textNode, el);
        } else {
          el.remove();
        }
        return;
      }

      const matchedPrefix = this.specialRefPrefixes.find(prefix => originalRefValue.startsWith(prefix));
      let pathPartForResolution = originalRefValue;
      let baseForResolution = htmlSourcePath;

      if (matchedPrefix) {
        pathPartForResolution = originalRefValue.substring(matchedPrefix.length);
        if (pathPartForResolution.startsWith('/')) pathPartForResolution = pathPartForResolution.substring(1);
        baseForResolution = "";
        console.log(`   [processHtmlContent] Handling special prefix "${matchedPrefix}" in "${originalRefValue}". Path part: "${pathPartForResolution}"`);
      }

      const queryOrHashIndex = pathPartForResolution.search(/[?#]/);
      if (queryOrHashIndex !== -1) {
        pathPartForResolution = pathPartForResolution.substring(0, queryOrHashIndex);
      }

      const decodedPathPart = this.parsingHelper.tryDecodeURIComponent(pathPartForResolution);
      const resolvedPath = this.parsingHelper.resolveRelativePath(baseForResolution, decodedPathPart);

      if (resolvedPath) {
        // Use file map for lookup
        const normalizedResolvedPath = this.parsingHelper.correctCyrillicCPath(this.parsingHelper.tryDecodeURIComponent(resolvedPath)).toLowerCase();
        const file = this.fileMap.get(normalizedResolvedPath) || null;

        if (file) {
          const targetFileName = file.name.split('/').pop() || file.name;

          if (isImage && file.mimeType?.startsWith('image/') && typeof file.data === 'string' && file.data.startsWith('data:image')) {
            // Image data is already base64, keep it as is.
            // This case should ideally be handled in the initial file processing map if needed.
            // If data URLs are handled there, this branch won't be needed here for adding to referencedFiles.
             // el.setAttribute('src', file.data); // Keep this line to ensure data URL is on the element
             console.log(`   [processHtmlContent] Found data URL image for ${file.name}. Keeping.`);
          } else if (isImage) {
            const altText = el.getAttribute('alt') || targetFileName || 'image';
            const span = htmlDoc.createElement('span');
            span.style.cssText = "color: #555; border: 1px dashed #ccc; padding: 2px 5px; display: inline-block; font-style: italic;";
            span.textContent = `[Image: ${decode(altText)} - will be attached separately]`;
            el.parentNode?.replaceChild(span, el);
            if (!referencedFiles.some(rf => rf.file.name === file!.name)) {
              referencedFiles.push({file, targetFileName: targetFileName});
              console.log(`   [processHtmlContent] Replaced local image ${file.name} with placeholder and added to attachments.`);
            }
          } else if (isLink) {
            const linkText = el.textContent?.trim() || targetFileName || file.name;
            const span = htmlDoc.createElement('span');
            span.style.cssText = "color: #555; border: 1px dashed #ccc; padding: 2px 5px; display: inline-block; font-style: italic;";
            span.textContent = `${decode(linkText)} [Attached File: ${targetFileName}]`;
            el.parentNode?.replaceChild(span, el);
            if (!referencedFiles.some(rf => rf.file.name === file!.name)) {
              referencedFiles.push({file, targetFileName: targetFileName});
              console.log(`   [processHtmlContent] Replaced local file link ${file.name} with placeholder and added to attachments.`);
            }
          }
        } else {
          console.warn(`   [processHtmlContent] Local file referenced in HTML not found: ${originalRefValue} (Resolved: ${resolvedPath})`);
          const span = htmlDoc.createElement('span');
          span.style.cssText = "color: red; border: 1px dashed red; padding: 2px 5px; display: inline-block; font-style: italic; text-decoration: line-through;";
          span.textContent = `[Broken Link: ${originalRefValue}]`;
          if (el.parentNode) {
            el.parentNode.replaceChild(span, el);
          } else {
            el.remove();
          }
        }
      } else {
        console.warn(`   [processHtmlContent] Could not resolve local path referenced in HTML: ${originalRefValue}`);
        const span = htmlDoc.createElement('span');
        span.style.cssText = "color: red; border: 1px dashed red; padding: 2px 5px; display: inline-block; font-style: italic; text-decoration: line-through;";
        span.textContent = `[Broken Link: ${originalRefValue}]`;
        if (el.parentNode) {
          el.parentNode.replaceChild(span, el);
        } else {
          el.remove();
        }
      }
    });

    const descriptionForDisplay = decode(contentElement.innerHTML);

    let classroomDesc = (contentElement.innerText || contentElement.textContent || '').replace(/\s+/g, ' ').trim();

    const maxDescLength = 300;
    if (classroomDesc.length > maxDescLength) {
      classroomDesc = classroomDesc.substring(0, maxDescLength - 3) + "...";
    }
    console.log(`   [processHtmlContent] Extracted plain text for classroomDesc: ${classroomDesc.substring(0, 100)}...`);


    return {
      descriptionForDisplay,
      descriptionForClassroom: classroomDesc,
      referencedFiles,
      externalLinks,
      richtext: containsRichElements || referencedFiles.length > 0 || externalLinks.length > 0
    };
  }
}
