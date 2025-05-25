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

  private fileMap: Map<string, ImsccFile> = new Map(); // Map for quick file lookup
  private manifestXmlDoc: XMLDocument | null = null; // Store parsed manifest

  constructor() { }

  // Helper function to get a normalized key for fileMap lookup
  private getFileMapKey(path: string): string {
    if (!path) return '';
    return this.parsingHelper.correctCyrillicCPath(this.parsingHelper.tryDecodeURIComponent(path)).toLowerCase();
  }


  convertImscc(files: ImsccFile[]): Observable<ProcessedCourseWork> {
    console.log("Starting IMSCC conversion process (Main Thread)...");
    this.skippedItemLog = [];
    this.fileMap = new Map(); // Reset map for a new conversion
    this.manifestXmlDoc = null; // Reset manifest doc

    // --- 1. Process Files and Build Map ---
    // Iterate over each file provided in the input array.
    files.forEach(file => {
      let processedFile = file;
      // If the file is an image and its data is an ArrayBuffer, convert it to a base64 data URL.
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
      // If the file is not an image, video, or audio, and its data is an ArrayBuffer, decode it as UTF-8 text.
      else if (this.parsingHelper.isArrayBuffer(file.data) && !file.mimeType?.startsWith('image/') && !file.mimeType?.startsWith('video/') && !file.mimeType?.startsWith('audio/')) {
        try {
          const textDecoder = new TextDecoder('utf-8');
          const textData = textDecoder.decode(file.data);
          // console.log(`Decoded ArrayBuffer as text for file: ${file.name}`); // Reduced verbosity
          processedFile = {...file, data: textData};
        } catch (e) {
          console.warn(`Could not decode ArrayBuffer as text for file ${file.name}`, e);
        }
      }
      // Normalize the file name to use as a key in the fileMap.
      const normalizedFileNameKey = this.getFileMapKey(processedFile.name);
      this.fileMap.set(normalizedFileNameKey, processedFile);
    });

    // Attempt to find the IMS manifest file (imsmanifest.xml) in the fileMap.
    const manifestFileKey = this.getFileMapKey('imsmanifest.xml');
    const manifestFile = this.fileMap.get(manifestFileKey);

    // If the manifest file is not found or its data is not a string, log an error and return an error Observable.
    if (!manifestFile || typeof manifestFile.data !== 'string') {
      console.error('IMS Manifest file (imsmanifest.xml) not found or data is not a string AFTER file processing.');
      return throwError(() => new Error('imsmanifest.xml not found or data is not a string.'));
    }

    let organization: Element | null = null;
    let rootItems: Element[] = [];
    let resourcesElement: Element | null = null;
    let directResources: Element[] = [];

    // --- 2. Parse Manifest XML ---
    try {
      const parser = new DOMParser();
      let doc = parser.parseFromString(manifestFile.data, "application/xml");
      let parseError = doc.querySelector('parsererror');

      // Handle potential XML parsing errors with fallbacks (text/xml, BOM removal).
      if (parseError) {
        console.warn("XML Parsing Error (application/xml), attempting fallback:", parseError.textContent);
        let tempDoc = parser.parseFromString(manifestFile.data, "text/xml");
        parseError = tempDoc.querySelector('parsererror');
        if (!parseError) {
          doc = tempDoc;
          console.warn("Parsed manifest using text/xml fallback.");
        } else {
          console.warn("Fallback XML Parsing Error (text/xml), attempting BOM removal:", parseError.textContent);
          // Remove Byte Order Mark (BOM) if present and try parsing again.
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
      this.manifestXmlDoc = doc; // Store the successfully parsed XML document.

      // Extract the course name from the manifest.
      this.coursename = this.parsingHelper.extractManifestTitle(this.manifestXmlDoc) || 'Untitled Course';
      console.log(`Extracted course name: ${this.coursename}`);

      // Find the <organizations> element, then the <organization> element.
      const organizationsNode = this.manifestXmlDoc.getElementsByTagNameNS(this.IMSCP_V1P1_NS, 'organizations')[0]
        || this.manifestXmlDoc.getElementsByTagName('organizations')[0];

      if (organizationsNode) {
        organization = organizationsNode.getElementsByTagNameNS(this.IMSCP_V1P1_NS, 'organization')[0]
          || organizationsNode.getElementsByTagName('organization')[0];
        if (organization) {
          // Get all top-level <item> elements within the organization.
          rootItems = Array.from(organization.children).filter(
            (node): node is Element => node instanceof Element && node.localName === 'item'
          );
          if (rootItems.length === 0) console.warn('No <item> elements found within the organization.');
        } else console.warn('No <organization> element found within <organizations>.');
      } else {
        // If no <organizations> found, look for a <resources> element directly.
        resourcesElement = this.manifestXmlDoc.getElementsByTagNameNS(this.IMSCP_V1P1_NS, 'resources')[0]
          || this.manifestXmlDoc.getElementsByTagName('resources')[0];
        if (resourcesElement) {
          // Get all <resource> elements within <resources>.
          directResources = Array.from(resourcesElement.children).filter(
            (node): node is Element => node instanceof Element && node.localName === 'resource'
          );
          if (directResources.length === 0) console.warn('Found <resources> element, but it contains no <resource> children.');
        } else {
          // If neither <organizations> nor <resources> are found, it's an error.
          console.error('Manifest contains neither <organizations> nor <resources> elements. Cannot process.');
          return throwError(() => new Error('No <organizations> or <resources> found in manifest'));
        }
      }
    } catch (error) {
      console.error('Error processing IMSCC package manifest:', error);
      const message = error instanceof Error ? error.message : String(error);
      return throwError(() => new Error(`Failed to process IMSCC manifest: ${message}`));
    }

    // --- 3. Determine Processing Strategy and Initiate Stream ---
    let processingStream: Observable<ProcessedCourseWork>;
    if (rootItems.length > 0) {
      // If root items exist (typical structure), process them.
      processingStream = this.processImsccItemsStream(rootItems, undefined);
    } else if (directResources.length > 0) {
      // If only direct resources exist (flat structure), process them.
      processingStream = from(directResources).pipe(
        concatMap(resource => {
          try {
            const title = resource.getAttribute('title') || this.parsingHelper.extractTitleFromMetadata(resource) || 'Untitled Resource';
            const identifier = resource.getAttribute('identifier') || `resource_${Math.random().toString(36).substring(2)}`;
            return this.processResource(resource, title, identifier, undefined);
          } catch (error) {
            console.error(`Error processing direct resource (ID: ${resource.getAttribute('identifier') || 'unknown'}):`, error);
            this.skippedItemLog.push({id: resource.getAttribute('identifier') || undefined, title: resource.getAttribute('title') || 'Unknown direct resource', reason: `Error during processing: ${error instanceof Error ? error.message : String(error)}`});
            return EMPTY; // Skip this resource on error
          }
        }),
        filter((result): result is ProcessedCourseWork => result !== null) // Filter out null results (skipped items)
      );
    } else {
      // If no items or resources are found, return an empty stream.
      console.warn('No root items or direct resources found to process. Conversion will yield no results.');
      processingStream = EMPTY;
    }

    // Return the processing stream, logging each emitted item and handling errors.
    return processingStream.pipe(
      tap(item => console.log(` -> Emitting processed item: "${item.title}" (Type: ${item.workType}, ID: ${item.associatedWithDeveloper?.id}, GDoc Candidate: ${!!item.richtext})`)),
      catchError(err => {
        console.error("Error during IMSCC content processing stream:", err);
        const wrappedError = err instanceof Error ? err : new Error(String(err));
        return throwError(() => new Error(`Error processing IMSCC content stream: ${wrappedError.message}`));
      })
    );
  }

  /**
   * Recursively processes a list of IMSCC <item> elements from the manifest.
   * @param items - Array of <item> XML elements to process.
   * @param parentTopic - Optional name of the parent topic/folder for these items.
   * @returns An Observable emitting ProcessedCourseWork objects.
   */
  private processImsccItemsStream(
    items: Element[],
    parentTopic?: string
  ): Observable<ProcessedCourseWork> {
    if (!items || items.length === 0) return EMPTY; // Return empty if no items to process.

    // Use `from` to create an Observable from the array of items,
    // then `concatMap` to process them sequentially.
    return from(items).pipe(
      concatMap((item: Element) => {
        try {
          const identifier = item.getAttribute('identifier') || `item_${Math.random().toString(36).substring(2)}`;
          const titleElement = item.querySelector(':scope > title'); // Get the direct child <title>
          const rawTitle = titleElement?.textContent?.trim() || this.parsingHelper.extractTitleFromMetadata(item) || 'Untitled Item';
          const identifierRef = item.getAttribute('identifierref'); // Reference to a <resource>
          const childItems = Array.from(item.children).filter(
            (node): node is Element => node instanceof Element && node.localName === 'item'
          ); // Get child <item> elements for recursion.

          const sanitizedTopicName = this.parsingHelper.sanitizeTopicName(rawTitle);

          let resourceObservable: Observable<ProcessedCourseWork | null> = EMPTY;

          // If the item references a resource, process that resource.
          if (identifierRef) {
            const resourceSelector = `resource[identifier="${identifierRef}"]`;
            // Find the resource in the manifest by its identifier.
            const resource = this.manifestXmlDoc?.querySelector(resourceSelector) ||
              Array.from(this.manifestXmlDoc?.getElementsByTagName('resource') || []).find(r => r.getAttribute('identifier') === identifierRef);

            if (resource) {
              resourceObservable = this.processResource(resource, rawTitle, identifier, parentTopic);
            } else {
              // Log a warning if the referenced resource is not found.
              console.warn(`   Resource not found for identifierref: ${identifierRef} (Item: "${rawTitle}"). This item might be a folder or a broken link.`);
              this.skippedItemLog.push({id: identifier, title: rawTitle, reason: `Resource not found for ref: ${identifierRef}`});
            }
          } else {
            // If no identifierref, this item might be a folder/topic or an empty item.
            if (childItems.length === 0) {
              // console.log(`   Item "${rawTitle}" (ID: ${identifier}) has no identifierref and no child items. Skipping as a standalone processed item.`); // Reduced verbosity
              this.skippedItemLog.push({id: identifier, title: rawTitle, reason: 'No resource reference and no child items'});
            } else {
              // This item acts as a container/topic for its child items.
              // console.log(`   Item "${rawTitle}" (ID: ${identifier}) is a container/topic. Processing sub-items.`); // Reduced verbosity
            }
          }

          // Recursively process child items. The current item's title becomes the parentTopic for its children.
          let childItemsObservable: Observable<ProcessedCourseWork> = EMPTY;
          if (childItems.length > 0) {
            childItemsObservable = this.processImsccItemsStream(childItems, sanitizedTopicName);
          }

          // Concatenate the observable for the current item's resource (if any)
          // with the observable for its child items.
          return concat(resourceObservable, childItemsObservable).pipe(
            filter((result): result is ProcessedCourseWork => result !== null) // Filter out nulls (skipped items)
          );

        } catch (error) {
          const itemIdentifier = item.getAttribute('identifier') || 'unknown_item';
          const itemTitle = item.querySelector(':scope > title')?.textContent?.trim() || 'Untitled Item';
          console.error(`Error processing individual item (ID: ${itemIdentifier}, Title: ${itemTitle}):`, error);
          this.skippedItemLog.push({id: itemIdentifier, title: itemTitle, reason: `Error during item processing: ${error instanceof Error ? error.message : String(error)}`});
          return EMPTY; // Skip this item on error
        }
      }),
      catchError(err => {
        // Catch any errors that occur within the stream processing.
        console.error("Error in processImsccItems stream:", err);
        return throwError(() => err); // Re-throw the error.
      })
    );
  }


  /**
   * Processes a single <resource> element from the IMSCC manifest.
   * Determines the type of resource (QTI, HTML, weblink, discussion, file) and extracts relevant information.
   * @param resource - The <resource> XML element.
   * @param itemTitle - The title of the <item> that references this resource.
   * @param imsccIdentifier - The identifier of the <item> (used for tracking).
   * @param parentTopic - Optional parent topic name.
   * @returns An Observable emitting a single ProcessedCourseWork object or null if skipped.
   */
  private processResource(
    resource: Element,
    itemTitle: string,
    imsccIdentifier: string,
    parentTopic?: string
  ): Observable<ProcessedCourseWork | null> {
    const resourceIdentifier = resource.getAttribute('identifier');
    const resourceType = resource.getAttribute('type'); // e.g., 'imsqti_xmlv1p2/xml', 'webcontent', 'imsdt_xmlv1p1'
    const resourceHref = resource.getAttribute('href'); // Primary file/link for the resource
    const baseHref = resource.getAttribute('xml:base'); // Base path for resolving relative hrefs

    // Prefer title from resource metadata, then resource attribute, then item title.
    const resourceOwnTitle = this.parsingHelper.extractTitleFromMetadata(resource) || resource.getAttribute('title');
    const finalTitle = resourceOwnTitle || itemTitle;

    // console.log(`   [Converter] Processing Resource "${finalTitle}" (ID: ${resourceIdentifier}, Type: ${resourceType || 'N/A'}) referenced by Item ID: ${imsccIdentifier}`); // Reduced verbosity

    // Ensure manifestXmlDoc is available (should always be at this stage).
    if (!this.manifestXmlDoc) {
      console.error(`   [Converter] processResource called before manifestXmlDoc is set.`);
      this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: 'Internal Error: Manifest not parsed'});
      return of(null);
    }

    // Skip D2L-specific configuration resources.
    const d2lMaterialType = resource.getAttributeNS(this.D2L_V2P0_NS, 'material_type');
    if (d2lMaterialType === 'orgunitconfig') {
      // console.log(`   Skipping D2L orgunitconfig resource: "${finalTitle}" (ID: ${resourceIdentifier})`); // Reduced verbosity
      this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: 'D2L orgunitconfig'});
      return of(null);
    }

    // Initialize a base structure for the coursework item.
    let courseworkBase: Partial<ProcessedCourseWork> & {convertToGoogleDoc?: boolean} = { // Added convertToGoogleDoc here
      title: finalTitle,
      state: 'DRAFT', // Default state for new coursework
      materials: [],
      localFilesToUpload: [],
      associatedWithDeveloper: { // Metadata for tracking origin
        id: imsccIdentifier,
        resourceId: resourceIdentifier,
        topic: parentTopic,
      },
      descriptionForDisplay: '', // HTML description for display (e.g., in an iframe)
      descriptionForClassroom: '', // Plain text or simplified description for Classroom
      richtext: false, // Whether descriptionForDisplay contains rich HTML
      workType: 'ASSIGNMENT', // Default work type
      convertToGoogleDoc: false // Default to false
    };

    let primaryResourceFile: ImsccFile | null = null;
    let resolvedPrimaryHref: string | null = null; // The fully resolved path or URL of the primary file/link
    let primaryFileXmlDoc: XMLDocument | null = null; // Parsed XML document if the primary file is XML

    // --- Determine and Load Primary Resource File/Link ---
    // 1. Check resource 'href' attribute.
    let primaryFilePathOrUrl = resourceHref;
    if (!primaryFilePathOrUrl) {
      // 2. If no resource 'href', check the 'href' of the first <file> child element.
      const firstFileElement = Array.from(resource.children).find(node => node instanceof Element && node.localName === 'file') as Element | undefined;
      primaryFilePathOrUrl = firstFileElement?.getAttribute('href') || null;
      // if (primaryFilePathOrUrl) console.log(`   [Converter] No resource href, using href from first <file> element: ${primaryFilePathOrUrl}`); // Reduced verbosity
    }

    // If a path/URL is found and it's not an absolute HTTP(S) URL or a special prefixed one, try to resolve it as a local file.
    if (primaryFilePathOrUrl && !primaryFilePathOrUrl.match(/^https?:\/\//i) && !this.specialRefPrefixes.some(prefix => primaryFilePathOrUrl!.startsWith(prefix))) {
      resolvedPrimaryHref = this.parsingHelper.resolveRelativePath(baseHref, this.parsingHelper.tryDecodeURIComponent(primaryFilePathOrUrl));
      if (resolvedPrimaryHref) {
        primaryResourceFile = this.fileMap.get(this.getFileMapKey(resolvedPrimaryHref)) || null;
      }
      // Fallback: try resolving without full URI decoding if the first attempt failed.
      if (!primaryResourceFile) {
        const resolvedRawHref = this.parsingHelper.resolveRelativePath(baseHref, primaryFilePathOrUrl);
        if (resolvedRawHref && resolvedRawHref !== resolvedPrimaryHref) {
          // console.log(`   [Converter] Primary file decoded path lookup failed for "${resolvedPrimaryHref}", trying less decoded path "${resolvedRawHref}"`); // Reduced verbosity
          primaryResourceFile = this.fileMap.get(this.getFileMapKey(resolvedRawHref)) || null;
          if (primaryResourceFile) resolvedPrimaryHref = resolvedRawHref;
        }
      }

      if (!primaryResourceFile) {
        console.warn(`   [Converter] Referenced file not found in package: ${primaryFilePathOrUrl} (Resolved attempts: ${resolvedPrimaryHref})`);
      } else {
        // console.log(`   [Converter] Found primary resource file: ${primaryResourceFile.name} (MIME: ${primaryResourceFile.mimeType})`); // Reduced verbosity
        // If the primary file is XML, try to parse it.
        if ((primaryResourceFile.name.toLowerCase().endsWith('.xml') || primaryResourceFile.mimeType?.includes('xml')) && typeof primaryResourceFile.data === 'string') {
          try {
            // console.log(`   [Converter] Attempting to parse primary file XML: ${primaryResourceFile.name}`); // Reduced verbosity
            const parser = new DOMParser();
            const cleanXmlData = primaryResourceFile.data.charCodeAt(0) === 0xFEFF ? primaryResourceFile.data.substring(1) : primaryResourceFile.data; // Remove BOM
            primaryFileXmlDoc = parser.parseFromString(cleanXmlData, "application/xml");
            if (primaryFileXmlDoc.querySelector('parsererror')) {
              console.warn(`   [Converter] XML parsing error for ${primaryResourceFile.name}.`);
              primaryFileXmlDoc = null;
            } else {
              // console.log(`   [Converter] Successfully parsed XML for ${primaryResourceFile.name}.`); // Reduced verbosity
            }
          } catch (e) {
            console.error(`   [Converter] Exception parsing primary file XML for ${primaryResourceFile.name}:`, e);
            primaryFileXmlDoc = null;
          }
        }
      }
    } else if (primaryFilePathOrUrl && primaryFilePathOrUrl.match(/^https?:\/\//i)) {
      // If it's an absolute HTTP(S) URL, use it directly.
      resolvedPrimaryHref = primaryFilePathOrUrl;
      // console.log(`   [Converter] Primary resource is an external URL: ${resolvedPrimaryHref}`); // Reduced verbosity
    } else if (primaryFilePathOrUrl && this.specialRefPrefixes.some(prefix => primaryFilePathOrUrl!.startsWith(prefix))) {
      // Handle special prefixes (e.g., $IMS-CC-FILEBASE$).
      // console.log(`   [Converter] Primary resource uses a special prefix: ${primaryFilePathOrUrl}. Handling as potentially resolvable file or link.`); // Reduced verbosity
      const matchedPrefix = this.specialRefPrefixes.find(p => primaryFilePathOrUrl!.startsWith(p));
      let pathAfterPrefix = primaryFilePathOrUrl!.substring(matchedPrefix!.length);
      if (pathAfterPrefix.startsWith('/')) pathAfterPrefix = pathAfterPrefix.substring(1); // Remove leading slash if present
      const pathAfterPrefixCleaned = pathAfterPrefix.split(/[?#]/)[0]; // Remove query params/fragments

      if (matchedPrefix === '$IMS-CC-FILEBASE$') {
        // For $IMS-CC-FILEBASE$, try matching the end of file paths in the map.
        const decodedFileName = this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleaned).toLowerCase();
        for (const keyFromMap of this.fileMap.keys()) {
          if (keyFromMap.endsWith(decodedFileName)) {
            primaryResourceFile = this.fileMap.get(keyFromMap)!;
            resolvedPrimaryHref = primaryResourceFile.name;
            // console.log(`   [Converter] Found IMS-CC-FILEBASE primary file "${primaryResourceFile.name}" by endsWith "${decodedFileName}"`); // Reduced verbosity
            break;
          }
        }
        // Fallback to direct key lookup if endsWith match fails.
        if (!primaryResourceFile) {
          const directKey = this.getFileMapKey(pathAfterPrefixCleaned);
          primaryResourceFile = this.fileMap.get(directKey) || null;
          if (primaryResourceFile) resolvedPrimaryHref = primaryResourceFile.name;
        }
      } else {
        // For other special prefixes, treat as a potential link but also try to find a local file.
        resolvedPrimaryHref = primaryFilePathOrUrl; // Default to the full prefixed path as a link
        const directKey = this.getFileMapKey(pathAfterPrefixCleaned);
        primaryResourceFile = this.fileMap.get(directKey) || null;
        if (primaryResourceFile) resolvedPrimaryHref = primaryResourceFile.name; // If file found, use its name
      }

      if (primaryResourceFile) {
        // console.log(`   [Converter] Found primary resource file via special prefix: ${primaryResourceFile.name}`); // Reduced verbosity
        // If the found file is XML, try to parse it.
        if ((primaryResourceFile.name.toLowerCase().endsWith('.xml') || primaryResourceFile.mimeType?.includes('xml')) && typeof primaryResourceFile.data === 'string') {
          try {
            const parser = new DOMParser();
            const cleanXmlData = primaryResourceFile.data.charCodeAt(0) === 0xFEFF ? primaryResourceFile.data.substring(1) : primaryResourceFile.data;
            primaryFileXmlDoc = parser.parseFromString(cleanXmlData, "application/xml");
            if (primaryFileXmlDoc.querySelector('parsererror')) primaryFileXmlDoc = null;
          } catch (e) {primaryFileXmlDoc = null;}
        }
      } else {
        // console.log(`   [Converter] Special prefix resource "${primaryFilePathOrUrl}" not found as a local file. Will be treated as a link if applicable.`); // Reduced verbosity
      }

    } else {
      // console.warn(`   [Converter] No primary href or file found/resolvable for resource ID: ${resourceIdentifier}`); // Reduced verbosity
    }

    // --- Determine Resource Type and Process Accordingly ---
    const isStandardQti = (resourceType === 'imsqti_xmlv1p2/xml' || resourceType === 'imsqti_xmlv1p2p1/imsqti_asiitem_xmlv1p2p1' || resourceType?.startsWith('application/vnd.ims.qti') || resourceType?.startsWith('assessment/x-bb-qti') || ((primaryResourceFile?.name?.toLowerCase().endsWith('.xml') || resolvedPrimaryHref?.toLowerCase().endsWith('.xml')) && resourceType?.toLowerCase().includes('qti')));
    const isD2lQuiz = d2lMaterialType === 'd2lquiz';
    const isDiscussionTopic = (primaryResourceFile && primaryFileXmlDoc && this.parsingHelper.isTopicXml(primaryResourceFile, primaryFileXmlDoc)) ||
      resourceType?.toLowerCase().includes('discussiontopic') ||
      resourceType?.toLowerCase().startsWith('imsdt');


    if (isStandardQti || isD2lQuiz) {
      // --- QTI Assessment / Quiz ---
      courseworkBase.workType = 'ASSIGNMENT'; // QTI usually implies an assignment
      if (primaryResourceFile && primaryFileXmlDoc) {
        // If primary file is parsed XML (likely the QTI XML itself)
        courseworkBase.qtiFile = [primaryResourceFile]; // Attach for later QTI processing
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
        // console.log(`   [Converter] Identified QTI/Assessment: "${finalTitle}" (Resource ID: ${resourceIdentifier}). Attached QTI file.`); // Reduced verbosity
      } else if (primaryResourceFile && (primaryResourceFile.name.toLowerCase().endsWith('.zip') || primaryResourceFile.mimeType === 'application/zip')) {
        // If primary file is a ZIP (often used for QTI packages)
        // console.log(`   [Converter] Identified QTI/Assessment (ZIP): "${finalTitle}" (Resource ID: ${resourceIdentifier}). Attaching ZIP file.`); // Reduced verbosity
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile, targetFileName: primaryResourceFile.name.split('/').pop() || primaryResourceFile.name});
        courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
      } else if (primaryResourceFile) {
        // If primary file exists but isn't recognized XML or ZIP, attach as a general file.
        console.warn(`   [Converter] QTI/Assessment "${finalTitle}" main file (${primaryResourceFile.name}) not XML/ZIP. Attaching as general file.`);
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile, targetFileName: primaryResourceFile.name.split('/').pop() || primaryResourceFile.name});
        courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
      }
      else {
        // If no valid primary file for QTI, skip this resource.
        console.warn(`   Skipping QTI/Assessment resource "${finalTitle}" (ID: ${resourceIdentifier}): No valid primary file found.`);
        this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: 'QTI/Assessment - No valid primary file'});
        return of(null);
      }
    }
    else if (primaryResourceFile && primaryFileXmlDoc && this.parsingHelper.isWebLinkXml(primaryResourceFile, primaryFileXmlDoc)) {
      // --- Web Link (defined in its own XML file, e.g., by D2L) ---
      const extractedUrl = this.parsingHelper.extractWebLinkUrl(primaryResourceFile!, primaryFileXmlDoc!);
      if (extractedUrl) {
        courseworkBase.webLinkUrl = extractedUrl; // Store the extracted URL
        courseworkBase.workType = 'ASSIGNMENT'; // Or 'MATERIAL', depending on desired outcome
        if (!courseworkBase.materials?.some(m => m.link?.url === extractedUrl)) {
          courseworkBase.materials?.push({link: {url: extractedUrl}});
        }
        if (!courseworkBase.descriptionForClassroom) courseworkBase.descriptionForClassroom = `Please follow this link: ${finalTitle}`;
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
        courseworkBase.convertToGoogleDoc = true; // Mark as candidate for GDoc conversion
        // console.log(`   [Converter] Identified WebLink: "${finalTitle}" -> ${extractedUrl} (Resource ID: ${resourceIdentifier}).`); // Reduced verbosity
      } else {
        // If URL extraction fails, attach the XML file itself.
        console.warn(`   [Converter] WebLink XML "${finalTitle}" found but could not extract URL. Attaching XML file.`);
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile!, targetFileName: primaryResourceFile!.name.split('/').pop() || primaryResourceFile!.name});
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
        courseworkBase.workType = 'MATERIAL';
      }
    }
    else if (isDiscussionTopic) {
      // --- Discussion Topic ---
      courseworkBase.workType = 'SHORT_ANSWER_QUESTION'; // Map discussions to short answer questions
      // console.log(`   [Converter] Identified Discussion Topic: "${finalTitle}" (Resource ID: ${resourceIdentifier}).`); // Reduced verbosity

      let contentHtml: string | null = null;
      let contentSourceFilePath: string | null = null; // Path of the file from which HTML was extracted

      // Check if the primary file is the discussion topic's XML definition.
      if (primaryResourceFile && primaryFileXmlDoc && this.parsingHelper.isTopicXml(primaryResourceFile, primaryFileXmlDoc)) {
        contentHtml = this.parsingHelper.extractTopicDescriptionHtml(primaryResourceFile, primaryFileXmlDoc);
        contentSourceFilePath = primaryResourceFile.name;
        courseworkBase.associatedWithDeveloper!.sourceXmlFile = primaryResourceFile;
        // console.log(`   [Converter] Discussion Topic "${finalTitle}": Main file is topic XML.`); // Reduced verbosity
        if (!contentHtml) console.warn(`   [Converter] Discussion Topic XML "${finalTitle}" description is empty.`);
      } else if (primaryResourceFile && (primaryResourceFile.mimeType === 'text/html' || primaryResourceFile.name.toLowerCase().endsWith('.html')) && typeof primaryResourceFile.data === 'string') {
        // Fallback: If primary file is HTML, use its content.
        contentHtml = primaryResourceFile.data;
        contentSourceFilePath = primaryResourceFile.name;
        courseworkBase.associatedWithDeveloper!.sourceHtmlFile = primaryResourceFile;
        // console.log(`   [Converter] Discussion Topic "${finalTitle}": Main file is HTML (as fallback content source).`); // Reduced verbosity
      } else {
        console.warn(`   [Converter] Discussion Topic "${finalTitle}" (ID: ${resourceIdentifier}): Could not find primary content.`);
      }

      if (contentHtml && contentHtml.trim() !== '') {
        // console.log(`   [Converter] Processing HTML content for Discussion Topic "${finalTitle}". Length: ${contentHtml.length}`); // Reduced verbosity
        const processedHtml = this.processHtmlContent(contentSourceFilePath || '', contentHtml);
        courseworkBase.descriptionForDisplay = processedHtml.descriptionForDisplay;
        courseworkBase.richtext = processedHtml.richtext;
        courseworkBase.localFilesToUpload?.push(...processedHtml.referencedFiles);
        // REMOVED: Do not add external links from HTML content as materials
        // processedHtml.externalLinks.forEach(linkUrl => {
        //   if (!courseworkBase.materials?.some(m => m.link?.url === linkUrl)) courseworkBase.materials?.push({link: {url: linkUrl}});
        // });
        courseworkBase.descriptionForClassroom = processedHtml.descriptionForClassroom;

        // Heuristics to improve classroom description for discussions if the extracted text is too short or just the title.
        const plainTextLength = courseworkBase.descriptionForClassroom?.replace(/\s/g, '').length || 0;
        const displayPlainTextLength = (courseworkBase.descriptionForDisplay?.replace(/<[^>]+>/g, '').trim() || '').length;
        if (plainTextLength < 10 && displayPlainTextLength > 0) {
          courseworkBase.descriptionForClassroom = `Discussion Prompt: "${finalTitle}". See details below.`;
        } else if (plainTextLength > 0 && courseworkBase.descriptionForClassroom.trim().toLowerCase() === finalTitle.trim().toLowerCase() && displayPlainTextLength > 0 && displayPlainTextLength !== finalTitle.trim().length) {
          courseworkBase.descriptionForClassroom = `Discussion Prompt: ${finalTitle}. See details below.`;
        } else if (!courseworkBase.descriptionForClassroom.trim() && displayPlainTextLength > 0) {
          courseworkBase.descriptionForClassroom = `Discussion Prompt: ${finalTitle}. See formatted content below.`;
        }
        // console.log(`   [Converter] Final descriptionForClassroom for "${finalTitle}": ${courseworkBase.descriptionForClassroom.substring(0,100)}...`); // Reduced verbosity
        // console.log(`   [Converter] Final descriptionForDisplay for "${finalTitle}": ${courseworkBase.descriptionForDisplay.substring(0, 100)}...`); // Reduced verbosity
      } else {
        // If no HTML content found for discussion, use a generic description.
        // console.log(`   [Converter] Discussion Topic "${finalTitle}" (Resource ID: ${resourceIdentifier}): No extractable HTML content found.`); // Reduced verbosity
        courseworkBase.descriptionForDisplay = `<p>${finalTitle}</p>`; // Basic display
        courseworkBase.descriptionForClassroom = `Discussion: ${finalTitle}`;
        courseworkBase.richtext = true;
        // If there was a primary file (e.g., an empty topic.xml), attach it.
        if (primaryResourceFile && courseworkBase.localFilesToUpload && !courseworkBase.localFilesToUpload.some(f => f.file.name === primaryResourceFile!.name)) {
          courseworkBase.localFilesToUpload.push({file: primaryResourceFile, targetFileName: primaryResourceFile.name.split('/').pop() || primaryResourceFile.name});
          courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
        }
      }
    }
    else if (primaryResourceFile && (primaryResourceFile.mimeType === 'text/html' || primaryResourceFile.name.toLowerCase().endsWith('.html'))) {
      // --- HTML Page ---
      courseworkBase.workType = 'ASSIGNMENT'; // Or 'MATERIAL'
      const htmlSourcePath = primaryResourceFile.name;
      if (typeof primaryResourceFile.data === 'string') {
        // console.log(`   [Converter] Identified HTML file: "${finalTitle}" (Resource ID: ${resourceIdentifier}).`); // Reduced verbosity
        const processedHtml = this.processHtmlContent(htmlSourcePath, primaryResourceFile.data);
        courseworkBase.descriptionForDisplay = processedHtml.descriptionForDisplay;
        courseworkBase.descriptionForClassroom = processedHtml.descriptionForClassroom || `Please review the content: ${finalTitle}`;
        courseworkBase.richtext = processedHtml.richtext;
        courseworkBase.localFilesToUpload?.push(...processedHtml.referencedFiles);
        // REMOVED: Do not add external links from HTML content as materials
        // processedHtml.externalLinks.forEach(linkUrl => {
        //  if (!courseworkBase.materials?.some(m => m.link?.url === linkUrl)) courseworkBase.materials?.push({link: {url: linkUrl}});
        // });
        courseworkBase.associatedWithDeveloper!.sourceHtmlFile = primaryResourceFile;
        courseworkBase.convertToGoogleDoc = true; // HTML content page is a candidate
        // console.log(`   [Converter] Processed HTML content for "${finalTitle}". Description length: ${courseworkBase.descriptionForClassroom.length}`); // Reduced verbosity
      } else {
        // If HTML file data isn't a string (e.g., was an ArrayBuffer that failed to decode), attach as a file.
        const targetFileName = primaryResourceFile.name.split('/').pop() || primaryResourceFile.name;
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile, targetFileName: targetFileName});
        courseworkBase.descriptionForClassroom = `Please see the attached HTML file: ${targetFileName}`;
        courseworkBase.workType = 'MATERIAL';
        courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
        console.warn(`   [Converter] Primary HTML file "${finalTitle}" data was not string. Attaching file.`);
      }
    }
    else if (resolvedPrimaryHref && (resolvedPrimaryHref.startsWith('http://') || resolvedPrimaryHref.startsWith('https://') || this.specialRefPrefixes.some(prefix => resolvedPrimaryHref!.startsWith(prefix)))) {
      // --- External Link or Special Prefixed Link (not resolved to a local file) ---
      courseworkBase.workType = 'MATERIAL';
      const linkUrl = resolvedPrimaryHref;
      if (!courseworkBase.materials?.some(m => m.link?.url === linkUrl)) {
        courseworkBase.materials?.push({link: {url: linkUrl}});
      }
      if (!courseworkBase.descriptionForClassroom) {
        let cleanLink = linkUrl;
        // Remove prefixes for a cleaner display in the description.
        this.specialRefPrefixes.forEach(prefix => cleanLink = cleanLink.replace(prefix, ''));
        courseworkBase.descriptionForClassroom = `Link: ${finalTitle}${cleanLink ? ` (${cleanLink})` : ''}`;
      }
      // Mark true web links (not special prefixes that might have resolved to local files, though that's handled by `primaryResourceFile` check)
      if ((resolvedPrimaryHref.startsWith('http://') || resolvedPrimaryHref.startsWith('https://')) && !primaryResourceFile) {
        courseworkBase.convertToGoogleDoc = true;
      }
      // console.log(`   [Converter] Identified External Link/Special Ref: "${finalTitle}" -> ${linkUrl}.`); // Reduced verbosity
    }
    else if (primaryResourceFile) {
      // --- General File (PDF, DOCX, etc.) ---
      courseworkBase.workType = 'MATERIAL';
      const targetFileName = primaryResourceFile.name.split('/').pop() || primaryResourceFile.name;
      if (!courseworkBase.localFilesToUpload?.some(f => f.file.name === primaryResourceFile!.name)) {
        courseworkBase.localFilesToUpload?.push({file: primaryResourceFile, targetFileName: targetFileName});
      }
      if (!courseworkBase.descriptionForClassroom) courseworkBase.descriptionForClassroom = `Please see the attached file: ${targetFileName}`;
      courseworkBase.associatedWithDeveloper!.sourceOtherFile = primaryResourceFile;
      // console.log(`   [Converter] Identified General File Material: "${finalTitle}" (Resource ID: ${resourceIdentifier}). File: ${primaryResourceFile.name} (MIME: ${primaryResourceFile.mimeType}).`); // Reduced verbosity
    }
    else {
      // --- Unhandled or Skipped Resource ---
      console.warn(`   Skipping Resource "${finalTitle}" (ID: ${resourceIdentifier}, Type: ${resourceType || 'N/A'}): Could not determine primary file/link, or it's an unhandled type with no content.`);
      this.skippedItemLog.push({id: imsccIdentifier, title: finalTitle, reason: `Unhandled resource type or no primary file/link (${resourceType || 'N/A'})`});
      return of(null); // Skip this resource
    }

    // --- Process Dependencies ---
    const dependencyElements = Array.from(resource.children).filter((node): node is Element => node instanceof Element && node.localName === 'dependency');
    dependencyElements.forEach(dep => {
      const depIdRef = dep.getAttribute('identifierref');
      if (!depIdRef) return; // Skip if no identifierref

      const depRes = this.manifestXmlDoc?.querySelector(`resource[identifier="${depIdRef}"]`) || Array.from(this.manifestXmlDoc?.getElementsByTagName('resource') || []).find(r => r.getAttribute('identifier') === depIdRef);
      if (depRes) {
        const depHref = depRes.getAttribute('href');
        const depBaseHref = depRes.getAttribute('xml:base'); // Base for this dependency's href
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
              // console.log(`   [Converter] Dependency file decoded path lookup failed for "${resolvedDepHrefAttempt}", trying less decoded path "${resolvedRawDepHref}"`); // Reduced verbosity
              depFile = this.fileMap.get(this.getFileMapKey(resolvedRawDepHref)) || null;
              if (depFile) resolvedDepHrefAttempt = resolvedRawDepHref;
            }
          }
        }

        if (depFile && depFile.name !== primaryResourceFile?.name &&
          !depFile.mimeType?.startsWith('image/') &&
          !depFile.mimeType?.startsWith('video/') &&
          !depFile.mimeType?.startsWith('text/') &&
          !depFile.mimeType?.includes('xml')
        ) {
          const targetFileName = depFile.name.split('/').pop() || depFile.name;
          if (courseworkBase.localFilesToUpload && !courseworkBase.localFilesToUpload.some(f => f.file.name === depFile!.name)) {
            courseworkBase.localFilesToUpload.push({file: depFile, targetFileName: targetFileName});
            // console.log(`   [Converter] Added non-media/non-text dependency file for upload: ${depFile.name}`); // Reduced verbosity
          }
        } else if (!depFile && resolvedDepHrefAttempt && (resolvedDepHrefAttempt.startsWith('http://') || resolvedDepHrefAttempt.startsWith('https://') || this.specialRefPrefixes.some(prefix => resolvedDepHrefAttempt!.startsWith(prefix)))) {
          const linkUrl = resolvedDepHrefAttempt;
          if (courseworkBase.materials && !courseworkBase.materials.some(m => m.link?.url === linkUrl)) {
            courseworkBase.materials.push({link: {url: linkUrl}}); // Keep dependency links as materials
            // console.log(`   [Converter] Added dependency link to materials: ${linkUrl}`); // Reduced verbosity
          }
        } else if (!depFile) { // Only log if depFile is null, not if it was primary/media/etc.
          console.warn(`   [Converter] Dependency with identifierref "${depIdRef}" (href: ${depHref || 'N/A'}) could not be resolved to a file or link.`);
        }
      } else {
        console.warn(`   [Converter] Dependency identifierref "${depIdRef}" does not reference a resource.`);
      }
    });

    // --- Final Check for Content ---
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

    // console.log(`   [Converter] Successfully processed resource "${finalTitle}".`); // Reduced verbosity
    return of(courseworkBase as ProcessedCourseWork); // Return the processed coursework item.
  }


  /**
   * Processes HTML string content to extract a display version, a classroom-friendly version,
   * referenced local files (images, videos, other docs), and external links.
   * It modifies the HTML to replace local file links with placeholders or updated paths.
   * @param htmlSourcePath - The original path of the HTML file (used for resolving relative links).
   * @param htmlString - The raw HTML string.
   * @returns An object containing processed descriptions, files, links, and a richtext flag.
   */
  private processHtmlContent(
    htmlSourcePath: string, // e.g., 'web_content/page.html' or 'topic_files/discussion.html'
    htmlString: string
  ): {
    descriptionForDisplay: string;    // HTML for rich display (e.g., in an iframe/WebView)
    descriptionForClassroom: string;  // Plain text or simplified HTML for Classroom description
    referencedFiles: Array<{file: ImsccFile; targetFileName: string}>; // Local files to upload
    externalLinks: string[];          // External HTTP(S) links found
    richtext: boolean;                // True if descriptionForDisplay contains significant HTML structure
  } {
    if (!htmlString) {
      console.warn(`[processHtmlContent] No raw HTML data provided for source: ${htmlSourcePath}`);
      return {descriptionForDisplay: '', descriptionForClassroom: '', referencedFiles: [], externalLinks: [], richtext: false};
    }

    const parser = new DOMParser();
    // Remove BOM and perform initial cleanup before parsing.
    let cleanHtmlData = htmlString.charCodeAt(0) === 0xFEFF ? htmlString.substring(1) : htmlString;
    cleanHtmlData = this.parsingHelper.preProcessHtmlForDisplay(cleanHtmlData);

    const htmlDoc = parser.parseFromString(cleanHtmlData, 'text/html');
    const contentElement = htmlDoc.body || htmlDoc.documentElement; // Prefer <body>, fallback to root

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

    // --- Determine if HTML contains rich elements ---
    let containsRichElements = contentElement.querySelector('img, table, ul, ol, h1, h2, h3, h4, h5, h6, blockquote, pre, code, strong, em, u, s, sub, sup, p, div, span[style], video, iframe') !== null; // Added iframe
    if (!containsRichElements && contentElement.innerHTML.includes('<br')) containsRichElements = true;
    if (!containsRichElements && contentElement.children.length > 0) {
      const simpleTextLength = (contentElement.textContent || '').replace(/\s/g, '').length;
      const htmlLength = contentElement.innerHTML.replace(/\s/g, '').length;
      if (htmlLength > simpleTextLength + 10) {
        containsRichElements = true;
      }
    }

    // --- Process <a>, <img>, and <video> elements ---
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
        if (originalRef.match(/^#/i)) {
          if (el.parentNode) {
            const textNode = htmlDoc.createTextNode(el.textContent || 'Internal Link');
            el.parentNode.replaceChild(textNode, el);
          } else {el.remove();}
          return;
        }
        if (isImage && originalRef.match(/^data:image/i)) return;

        let file: ImsccFile | null = null;
        let pathForLogging: string = originalRef;
        let potentialFileKey = '';

        const matchedPrefix = this.specialRefPrefixes.find(p => originalRef.startsWith(p));
        let pathAfterPrefixCleaned = '';
        let baseForResolutionInLoop = htmlSourcePath;

        if (matchedPrefix === '$IMS-CC-FILEBASE$') {
          let pathPart = originalRef.substring(matchedPrefix.length);
          if (pathPart.startsWith('/')) pathPart = pathPart.substring(1);
          pathAfterPrefixCleaned = pathPart.split(/[?#]/)[0];
          const decodedFileName = this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleaned).toLowerCase();
          pathForLogging = `(IMS-CC-FILEBASE) endsWith: ${decodedFileName}`;
          for (const keyFromMap of this.fileMap.keys()) {
            if (keyFromMap.endsWith(decodedFileName)) {
              file = this.fileMap.get(keyFromMap)!;
              // console.log(`   [processHtmlContent] Found IMS-CC-FILEBASE file "${file.name}" for link "${originalRef}" by endsWith "${decodedFileName}"`); // Reduced verbosity
              break;
            }
          }
          if (!file) {
            const directKey = this.getFileMapKey(pathAfterPrefixCleaned);
            pathForLogging += ` / direct: ${directKey}`;
            file = this.fileMap.get(directKey) || null;
            // if (file) console.log(`   [processHtmlContent] Found IMS-CC-FILEBASE file "${file.name}" for link "${originalRef}" by direct full path match.`); // Reduced verbosity
          }
        } else if (matchedPrefix === '$WIKI_REFERENCE$') {
          let pathPart = originalRef.substring(matchedPrefix.length);
          if (pathPart.startsWith('/')) pathPart = pathPart.substring(1);
          pathAfterPrefixCleaned = pathPart.split(/[?#]/)[0];
          const pageSlugOrIdRaw = pathAfterPrefixCleaned;
          pathForLogging = `(WIKI_REFERENCE) slug/ID: ${this.parsingHelper.tryDecodeURIComponent(pageSlugOrIdRaw)}`;
          const slugPartDecoded = this.parsingHelper.tryDecodeURIComponent(pageSlugOrIdRaw.replace(/^pages\//i, ''));
          const commonPatternPath = `wiki_content/${slugPartDecoded}.html`;
          file = this.fileMap.get(this.getFileMapKey(commonPatternPath)) || null;

          if (file) {
            // console.log(`   [processHtmlContent] Found WIKI file "${file.name}" for link "${originalRef}" by slug pattern "${commonPatternPath}"`); // Reduced verbosity
          } else if (this.manifestXmlDoc) {
            const idToLookup = this.parsingHelper.tryDecodeURIComponent(pageSlugOrIdRaw.replace(/^pages\//i, ''));
            const resourceById = Array.from(this.manifestXmlDoc.getElementsByTagName('resource'))
              .find(r => r.getAttribute('identifier') === idToLookup && (r.getAttribute('type') === 'webcontent' || r.getAttribute('href')?.toLowerCase().endsWith('.html')));
            if (resourceById) {
              const resourceHref = resourceById.getAttribute('href');
              if (resourceHref) {
                pathForLogging += ` -> manifest href: ${resourceHref}`;
                file = this.fileMap.get(this.getFileMapKey(resourceHref)) || null;
                // if (file) console.log(`   [processHtmlContent] Found WIKI file "${file.name}" for link "${originalRef}" by ID "${idToLookup}" to manifest href "${resourceHref}"`); // Reduced verbosity
              }
            }
          }
        }
        else if (originalRef.startsWith('/content/enforced/') || originalRef.startsWith('/content/group/')) {
          // console.log(`   [processHtmlContent] Detected potential D2L content link: ${originalRef}`); // Reduced verbosity
          pathForLogging = `(D2L Content Link) ${originalRef}`;
          try {
            const contentMarker = originalRef.startsWith('/content/enforced/') ? '/content/enforced/' : '/content/group/';
            const pathAfterContentMarker = originalRef.substring(originalRef.indexOf(contentMarker) + contentMarker.length);
            const firstSlashIndex = pathAfterContentMarker.indexOf('/');
            let actualContentPath = pathAfterContentMarker;
            if (firstSlashIndex !== -1) {
              actualContentPath = pathAfterContentMarker.substring(firstSlashIndex + 1);
            }
            actualContentPath = actualContentPath.split('?')[0];

            if (actualContentPath) {
              const decodedContentPath = this.parsingHelper.tryDecodeURIComponent(actualContentPath);
              pathForLogging += ` -> Extracted D2L path: ${decodedContentPath}`;
              potentialFileKey = this.getFileMapKey(decodedContentPath);
              file = this.fileMap.get(potentialFileKey) || null;

              if (file) {
                // console.log(`   [processHtmlContent] Found D2L file by extracted path "${decodedContentPath}"`); // Reduced verbosity
              } else {
                const filenameOnly = decodedContentPath.split('/').pop() || "";
                if (filenameOnly && filenameOnly !== decodedContentPath) {
                  pathForLogging += ` / filename only: ${filenameOnly}`;
                  potentialFileKey = this.getFileMapKey(filenameOnly);
                  file = this.fileMap.get(potentialFileKey) || null;
                  // if (file) console.log(`   [processHtmlContent] Found D2L file by filename only "${filenameOnly}"`); // Reduced verbosity
                }
              }
            }
            if (!file) {
              console.warn(`   [processHtmlContent] D2L content link processing did not find a file for: ${originalRef}. Path for logging: ${pathForLogging}`);
            }
          } catch (e) {
            console.error(`   [processHtmlContent] Error processing D2L content link URL ${originalRef}:`, e);
          }
        }
        else {
          if (!matchedPrefix) {
            pathAfterPrefixCleaned = originalRef.split(/[?#]/)[0];
            baseForResolutionInLoop = htmlSourcePath;
            pathForLogging = `(Relative Path) ${this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleaned)}`;
          } else {
            let pathPart = originalRef.substring(matchedPrefix.length);
            if (pathPart.startsWith('/')) pathPart = pathPart.substring(1);
            pathAfterPrefixCleaned = pathPart.split(/[?#]/)[0];
            baseForResolutionInLoop = "";
            pathForLogging = `(${matchedPrefix} - root relative) ${this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleaned)}`;
          }

          const decodedPathSegment = this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleaned);
          let resolvedPath = this.parsingHelper.resolveRelativePath(baseForResolutionInLoop, decodedPathSegment);

          if (resolvedPath) {
            potentialFileKey = this.getFileMapKey(resolvedPath);
            file = this.fileMap.get(potentialFileKey) || null;
            if (file) pathForLogging = resolvedPath;
          }

          if (!file && pathAfterPrefixCleaned !== decodedPathSegment) {
            const resolvedRawPath = this.parsingHelper.resolveRelativePath(baseForResolutionInLoop, pathAfterPrefixCleaned);
            if (resolvedRawPath && resolvedRawPath !== resolvedPath) {
              let alternativeFileKey = this.getFileMapKey(resolvedRawPath);
              // console.log(`   [processHtmlContent] (Fallback Relative/Special) Decoded path lookup failed for "${originalRef}", trying non-decoded path "${resolvedRawPath}" (key: ${alternativeFileKey})`); // Reduced verbosity
              file = this.fileMap.get(alternativeFileKey) || null;
              if (file) pathForLogging = resolvedRawPath;
            }
          }
        }

        if (file) {
          const targetFileName = file.name.split('/').pop() || file.name;
          const resolvedOriginalPath = this.parsingHelper.resolveRelativePath(baseForResolutionInLoop, pathAfterPrefixCleaned) || pathAfterPrefixCleaned;

          if (isImage) {
            if (file.mimeType?.startsWith('image/') && typeof file.data === 'string' && file.data.startsWith('data:image')) {
              el.setAttribute('src', file.data);
            } else {
              const altText = el.getAttribute('alt') || targetFileName || 'image';
              const newAnchor = htmlDoc.createElement('a');
              newAnchor.href = resolvedOriginalPath;
              newAnchor.textContent = `Image: ${decode(altText)}`;
              newAnchor.setAttribute('data-imscc-local-media-type', 'image');
              newAnchor.setAttribute('data-imscc-original-path', resolvedOriginalPath);
              el.parentNode?.replaceChild(newAnchor, el);
              if (!referencedFiles.some(rf => rf.file.name === file!.name)) {
                referencedFiles.push({file, targetFileName: targetFileName});
              }
            }
          } else if (isLink) {
            const newAnchor = htmlDoc.createElement('a');
            newAnchor.href = resolvedOriginalPath;
            newAnchor.textContent = el.textContent?.trim() || targetFileName || file.name;
            newAnchor.setAttribute('data-imscc-original-path', resolvedOriginalPath);
            if (file.mimeType?.startsWith('video/')) {
              newAnchor.setAttribute('data-imscc-local-media-type', 'video');
              newAnchor.textContent = `Video: ${el.textContent?.trim() || targetFileName}`;
            } else if (file.mimeType?.startsWith('image/')) {
              newAnchor.setAttribute('data-imscc-local-media-type', 'image');
              newAnchor.textContent = `Image: ${el.textContent?.trim() || targetFileName}`;
            } else {
              newAnchor.setAttribute('data-imscc-local-media-type', 'file');
            }
            el.parentNode?.replaceChild(newAnchor, el);
            if (!referencedFiles.some(rf => rf.file.name === file!.name)) {
              referencedFiles.push({file, targetFileName: targetFileName});
            }
            // console.log(`   [processHtmlContent] Replaced <a> to local file ${file.name} with new anchor pointing to relative path: ${resolvedOriginalPath}`); // Reduced verbosity
          }
        } else {
          console.warn(`   [processHtmlContent] Local file referenced in <${el.tagName}> not found: ${originalRef} (Attempted lookup with path(s): ${pathForLogging})`);
          const span = htmlDoc.createElement('span');
          span.style.cssText = "color: red; border: 1px dashed red; padding: 2px 5px; display: inline-block; font-style: italic; text-decoration: line-through;";
          span.textContent = `[Broken Link: ${originalRef}]`;
          if (el.parentNode) {
            el.parentNode.replaceChild(span, el);
          } else {
            el.remove();
          }
        }

      } else if (isVideo) {
        let elementsToSearchForSrc = Array.from((el as HTMLVideoElement).querySelectorAll('source'));
        let localVideoFile: ImsccFile | null = null;
        let firstSourceRefForPlaceholder: string | null = null;
        let resolvedVideoSrcPathForAnchor: string | null = null;

        for (const sourceEl of elementsToSearchForSrc) {
          const originalSrc = sourceEl.getAttribute('src');
          if (!originalSrc) continue;
          if (!firstSourceRefForPlaceholder) firstSourceRefForPlaceholder = originalSrc;

          if (originalSrc.match(/^https?:\/\//i) || originalSrc.match(/^data:video/i)) {
            continue;
          }

          let currentSourceFile: ImsccFile | null = null;
          const matchedPrefixVideo = this.specialRefPrefixes.find(p => originalSrc.startsWith(p));
          let pathAfterPrefixCleanedVideo = '';
          let baseForResolutionInLoopVideo = htmlSourcePath;

          if (matchedPrefixVideo) {
            let pathPart = originalSrc.substring(matchedPrefixVideo.length);
            if (pathPart.startsWith('/')) pathPart = pathPart.substring(1);
            pathAfterPrefixCleanedVideo = pathPart.split(/[?#]/)[0];
            baseForResolutionInLoopVideo = "";
          } else if (originalSrc.startsWith('/content/enforced/') || originalSrc.startsWith('/content/group/')) {
            const contentMarker = originalSrc.startsWith('/content/enforced/') ? '/content/enforced/' : '/content/group/';
            const pathAfterContentMarker = originalSrc.substring(originalSrc.indexOf(contentMarker) + contentMarker.length);
            const firstSlashIndex = pathAfterContentMarker.indexOf('/');
            pathAfterPrefixCleanedVideo = (firstSlashIndex !== -1) ? pathAfterContentMarker.substring(firstSlashIndex + 1) : pathAfterContentMarker;
            pathAfterPrefixCleanedVideo = pathAfterPrefixCleanedVideo.split('?')[0];
            baseForResolutionInLoopVideo = "";
          }
          else {
            pathAfterPrefixCleanedVideo = originalSrc.split(/[?#]/)[0];
          }

          const decodedRelativePath = this.parsingHelper.tryDecodeURIComponent(pathAfterPrefixCleanedVideo);
          let resolvedPath = this.parsingHelper.resolveRelativePath(baseForResolutionInLoopVideo, decodedRelativePath);
          if (resolvedPath) {
            currentSourceFile = this.fileMap.get(this.getFileMapKey(resolvedPath)) || null;
          }
          if (!currentSourceFile && pathAfterPrefixCleanedVideo !== decodedRelativePath) {
            const resolvedRawPath = this.parsingHelper.resolveRelativePath(baseForResolutionInLoopVideo, pathAfterPrefixCleanedVideo);
            if (resolvedRawPath && resolvedRawPath !== resolvedPath) {
              currentSourceFile = this.fileMap.get(this.getFileMapKey(resolvedRawPath)) || null;
              if (currentSourceFile) resolvedPath = resolvedRawPath;
            }
          }

          if (currentSourceFile && currentSourceFile.mimeType?.startsWith('video/')) {
            localVideoFile = currentSourceFile;
            resolvedVideoSrcPathForAnchor = resolvedPath;
            // console.log(`   [processHtmlContent] Identified local video ${localVideoFile.name} from <video><source src="${originalSrc}">. Resolved relative path for anchor: ${resolvedVideoSrcPathForAnchor}`); // Reduced verbosity
            if (!referencedFiles.some(rf => rf.file.name === localVideoFile!.name)) {
              referencedFiles.push({file: localVideoFile, targetFileName: localVideoFile.name.split('/').pop() || localVideoFile.name});
            }
            break;
          }
        }

        const videoTitle = decode((el as HTMLVideoElement).getAttribute('title') || localVideoFile?.name.split('/').pop() || 'Video');
        if (localVideoFile && resolvedVideoSrcPathForAnchor) {
          const newAnchor = htmlDoc.createElement('a');
          newAnchor.href = resolvedVideoSrcPathForAnchor;
          newAnchor.textContent = `Video: ${videoTitle}`;
          newAnchor.setAttribute('data-imscc-local-media-type', 'video');
          newAnchor.setAttribute('data-imscc-original-path', resolvedVideoSrcPathForAnchor);
          el.parentNode?.replaceChild(newAnchor, el);
          containsRichElements = true;
          // console.log(`   [processHtmlContent] Replaced <video> tag with anchor: ${newAnchor.outerHTML}`); // Reduced verbosity
        } else {
          const refToShow = firstSourceRefForPlaceholder || "unknown source";
          const span = htmlDoc.createElement('span');
          span.style.cssText = "color: #555; border: 1px dashed #ccc; padding: 2px 5px; display: inline-block; font-style: italic;";
          span.textContent = `[Video: ${videoTitle} - local source not found or not a video file. Original ref: ${this.parsingHelper.tryDecodeURIComponent(refToShow)}]`;
          el.parentNode?.replaceChild(span, el);
          console.warn(`   [processHtmlContent] Local video source(s) not found for video element. First original ref for placeholder: ${refToShow}`);
        }
      }
    });

    // --- Process <iframe> elements ---
    Array.from(contentElement.querySelectorAll('iframe')).forEach((iframeEl: Element) => {
      const iframeSrc = iframeEl.getAttribute('src');
      const iframeTitle = iframeEl.getAttribute('title');

      if (iframeSrc) {
        const newAnchor = htmlDoc.createElement('a');
        newAnchor.href = iframeSrc;
        // Set the anchor text to the iframe's title, or the src as a fallback
        newAnchor.textContent = decode(iframeTitle || iframeSrc);
        newAnchor.target = '_blank'; // Open in new tab
        newAnchor.rel = 'noopener noreferrer'; // Security best practice

        // Add to externalLinks if it's a new link and it's a valid HTTP/HTTPS URL
        if (iframeSrc.match(/^https?:\/\//i) && !externalLinks.includes(iframeSrc)) {
          externalLinks.push(iframeSrc);
        }

        // Create a paragraph to wrap the link for better block display
        const p = htmlDoc.createElement('p');
        p.appendChild(newAnchor); // Add the anchor (with title as text) to the paragraph

        if (iframeEl.parentNode) {
          iframeEl.parentNode.replaceChild(p, iframeEl);
        } else {
          // Fallback if iframe has no parent (should not happen in valid HTML structure)
          contentElement.appendChild(p);
        }

        containsRichElements = true; // Replacing iframe with a link is a structural change
        console.log(`   [processHtmlContent] Replaced <iframe> (src: ${iframeSrc}) with anchor link using title: "${newAnchor.textContent}".`);
      } else {
        // If iframe has no src, it's likely not useful, remove it or replace with placeholder
        const placeholderText = htmlDoc.createTextNode(`[Unsupported embedded content: ${decode(iframeTitle || 'Untitled Iframe')}]`);
        if (iframeEl.parentNode) {
          iframeEl.parentNode.replaceChild(placeholderText, iframeEl);
        }
        console.warn(`   [processHtmlContent] Removed <iframe> with no src attribute (title: ${decode(iframeTitle || 'N/A')}).`);
      }
    });


    const descriptionForDisplay = decode(contentElement.innerHTML);
    const tempDiv = htmlDoc.createElement('div');
    tempDiv.innerHTML = descriptionForDisplay;
    let classroomDesc = (tempDiv.textContent || tempDiv.innerText || '').replace(/\s+/g, ' ').trim();

    const maxDescLength = 25000;
    if (classroomDesc.length > maxDescLength) {
      let truncated = classroomDesc.substring(0, maxDescLength - 20);
      const lastPeriod = truncated.lastIndexOf('.');
      if (lastPeriod > 0) {
        truncated = truncated.substring(0, lastPeriod + 1);
      } else {
        truncated = truncated.substring(0, maxDescLength - 3);
      }
      classroomDesc = truncated + "...";
    }
    // console.log(`   [processHtmlContent] Extracted plain text for classroomDesc: ${classroomDesc.substring(0, 100)}...`); // Reduced verbosity

    return {
      descriptionForDisplay,
      descriptionForClassroom: classroomDesc,
      referencedFiles,
      externalLinks,
      richtext: containsRichElements || referencedFiles.length > 0 || externalLinks.length > 0
    };
  }
}
