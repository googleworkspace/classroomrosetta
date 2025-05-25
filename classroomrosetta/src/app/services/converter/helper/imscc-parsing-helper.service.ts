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

import {Injectable} from '@angular/core';
import {decode as htmlEntitiesDecode} from 'html-entities';
import {ImsccFile} from '../../../interfaces/classroom-interface'; // Adjust path

@Injectable({
  providedIn: 'root'
})
export class ImsccParsingHelperService {

  // --- XML Namespaces ---
  private readonly IMSCP_V1P1_NS = 'http://www.imsglobal.org/xsd/imscp_v1p1';
  private readonly IMSMD_V1P2_NS = 'http://www.imsglobal.org/xsd/imsmd_rootv1p2p1';
  private readonly LOMIMSCC_V1P1_NS = 'http://ltsc.ieee.org/xsd/imsccv1p1/LOM/manifest';
  private readonly LOMIMSCC_V1P3_NS = 'http://ltsc.ieee.org/xsd/imsccv1p3/LOM/manifest';
  private readonly WEBLINK_V1P2_NS = "http://www.imsglobal.org/xsd/imsccv1p2/imswl_v1p2";
  private readonly DISCUSSIONTOPIC_V1P1_NS = "http://www.imsglobal.org/xsd/imsccv1p1/imsdt_v1p1";


  constructor() { }

  public isArrayBuffer(data: string | ArrayBuffer): data is ArrayBuffer {
    return data instanceof ArrayBuffer;
  }

  public tryDecodeURIComponent(uriComponent: string): string {
    if (!uriComponent) return '';
    let decoded = uriComponent;
    try {
      // Attempt to decode multiple times for cases like double encoding
      for (let i = 0; i < 5; i++) {
        const previouslyDecoded = decoded;
        decoded = decodeURIComponent(decoded.replace(/\+/g, ' '));
        if (decoded === previouslyDecoded) break; // Stop if no change
      }
    } catch (e) {
      // If decodeURIComponent fails, fall back to HTML entity decoding only
      return htmlEntitiesDecode(uriComponent);
    }
    // Final HTML entity decode after URI decoding
    return htmlEntitiesDecode(decoded);
  }

  public correctCyrillicCPath(path: string | null): string {
    if (!path) return '';
    const cyrillicCPrefix = '\u0441ontent/'; // Cyrillic 'с'
    if (path.startsWith(cyrillicCPrefix)) {
      return 'content/' + path.substring(cyrillicCPrefix.length);
    }
    return path;
  }

  public getDirectory(pathStr: string | null | undefined): string {
    if (!pathStr) return "";
    const normalized = pathStr.trim().replace(/\\/g, '/'); // Normalize slashes
    if (normalized.endsWith('/')) {
      return normalized.replace(/\/$/, ''); // Remove trailing slash if it's a directory path
    }
    const lastSlash = normalized.lastIndexOf('/');
    if (lastSlash === -1) {
      return ""; // No directory part (root level)
    }
    return normalized.substring(0, lastSlash); // Extract directory part
  }

  public resolveRelativePath(
    basePathInput: string | null | undefined,
    decodedRelativePath: string | null | undefined
  ): string | null {
    if (typeof decodedRelativePath !== 'string') return null;

    const normRelativePath = decodedRelativePath.trim().replace(/\\/g, '/');
    // Base path should be the directory containing the file, not the file itself
    const actualBaseDir = this.getDirectory(basePathInput);

    if (!normRelativePath) return actualBaseDir || null; // If relative path is empty, return base directory

    // Handle absolute paths from the package root
    if (normRelativePath.startsWith('/')) {
      const parts = normRelativePath.substring(1).split('/');
      const resolvedParts: string[] = [];
      for (const part of parts) {
        if (part === '..') {
          if (resolvedParts.length > 0) resolvedParts.pop(); // Go up one level
        } else if (part !== '.' && part !== '') {
          resolvedParts.push(part); // Add path segment
        }
      }
      return resolvedParts.join('/');
    }

    // Handle relative paths
    const baseParts = actualBaseDir ? actualBaseDir.split('/') : [];
    const relativeParts = normRelativePath.split('/');
    let combinedParts = [...baseParts];

    for (const part of relativeParts) {
      if (part === '..') {
        if (combinedParts.length > 0) combinedParts.pop(); // Go up one level
      } else if (part !== '.' && part !== '') {
        combinedParts.push(part); // Add path segment
      }
    }
    return combinedParts.join('/');
  }

  public sanitizeTopicName(name: string): string {
    if (!name) return 'Untitled Topic';
    // Replace slashes, ampersands, and remove most non-alphanumeric characters
    let sanitized = name.replace(/\//g, '-').replace(/&/g, 'and');
    sanitized = sanitized.replace(/[^a-zA-Z0-9\s\-_.,()']/g, '');
    sanitized = sanitized.trim();
    // Truncate to Google Classroom's topic name limit (100 chars)
    if (sanitized.length > 100) {
      sanitized = sanitized.substring(0, 97) + '...';
    }
    return sanitized || 'Untitled Topic'; // Fallback if sanitization results in empty string
  }

  public extractTitleFromMetadata(element: Element): string | null {
    try {
      // Find the metadata element (namespace aware)
      const metadataElement = element.getElementsByTagNameNS(this.IMSCP_V1P1_NS, 'metadata')[0] || element.getElementsByTagName('metadata')[0];
      if (!metadataElement) return null;

      let lomElement: Element | null = null;
      let lomNamespace: string | null = null;

      // Try finding LOM element with known namespaces first
      const namespacesToTry = [this.IMSMD_V1P2_NS, this.LOMIMSCC_V1P3_NS, this.LOMIMSCC_V1P1_NS];
      for (const ns of namespacesToTry) {
        lomElement = metadataElement.getElementsByTagNameNS(ns, 'lom')[0] as Element | null;
        if (lomElement) {
          lomNamespace = ns;
          break;
        }
      }
      // Fallback to finding LOM element without specific namespace
      if (!lomElement) {
        lomElement = metadataElement.getElementsByTagName('lom')[0] as Element | null;
        if (lomElement) lomNamespace = '*'; // Indicate generic namespace found
      }
      if (!lomElement) return null;

      // Find the 'general' element within LOM (namespace aware)
      let generalElement: Element | null;
      if (lomNamespace && lomNamespace !== '*') {
        generalElement = lomElement.getElementsByTagNameNS(lomNamespace!, 'general')[0] as Element | null;
      } else {
        generalElement = lomElement.getElementsByTagName('general')[0] as Element | null;
      }
      if (!generalElement) return null;

      // Find the 'title' element within 'general' (namespace aware)
      let titleElement: Element | null;
      if (lomNamespace && lomNamespace !== '*') {
        titleElement = generalElement.getElementsByTagNameNS(lomNamespace!, 'title')[0] as Element | null;
      } else {
        titleElement = generalElement.getElementsByTagName('title')[0] as Element | null;
      }
      if (!titleElement) return null;

      // Find the actual title string within 'title' (can be 'langstring' or 'string')
      let stringElement: Element | null = null;
      if (lomNamespace === this.IMSMD_V1P2_NS || lomNamespace === '*') { // IMSMD uses langstring
        stringElement = titleElement.getElementsByTagNameNS(this.IMSMD_V1P2_NS, 'langstring')[0] as Element | null ||
          titleElement.getElementsByTagName('langstring')[0] as Element | null;
      }
      if (!stringElement) { // Try 'string' element as fallback or for other LOM versions
        stringElement = (lomNamespace && lomNamespace !== '*') ?
          titleElement.getElementsByTagNameNS(lomNamespace!, 'string')[0] as Element | null :
          titleElement.getElementsByTagName('string')[0] as Element | null;
      }

      // If no specific string element, use the text content of the title element itself
      if (!stringElement) {
        const directTitle = titleElement.textContent?.trim();
        return directTitle || null;
      }
      return stringElement.textContent?.trim() || null; // Return the text content of the string element
    } catch (error) {
      console.error('Error extracting title from resource metadata:', error);
      return null;
    }
  }

  public extractManifestTitle(xmlDoc: XMLDocument): string | null {
    if (!xmlDoc || !xmlDoc.documentElement) return null;
    // Similar logic to extractTitleFromMetadata, but applied to the root metadata
    try {
      const metadataElement = xmlDoc.documentElement.getElementsByTagNameNS(this.IMSCP_V1P1_NS, 'metadata')[0] || xmlDoc.documentElement.getElementsByTagName('metadata')[0];
      if (!metadataElement) return null;

      let lomElement: Element | null = null;
      let lomNamespace: string | null = null;

      const namespacesToTry = [this.IMSMD_V1P2_NS, this.LOMIMSCC_V1P3_NS, this.LOMIMSCC_V1P1_NS];
      for (const ns of namespacesToTry) {
        lomElement = metadataElement.getElementsByTagNameNS(ns, 'lom')[0] as Element | null;
        if (lomElement) {
          lomNamespace = ns;
          break;
        }
      }
      if (!lomElement) {
        lomElement = metadataElement.getElementsByTagName('lom')[0] as Element | null;
        if (lomElement) lomNamespace = '*';
      }
      if (!lomElement) return null;

      const generalElement = (lomNamespace && lomNamespace !== '*') ?
        lomElement.getElementsByTagNameNS(lomNamespace, 'general')[0] as Element | null :
        lomElement.getElementsByTagName('general')[0] as Element | null;
      if (!generalElement) return null;

      const titleElement = (lomNamespace && lomNamespace !== '*') ?
        generalElement.getElementsByTagNameNS(lomNamespace, 'title')[0] as Element | null :
        generalElement.getElementsByTagName('title')[0] as Element | null;
      if (!titleElement) return null;

      let stringElement: Element | null = null;
      if (lomNamespace === this.IMSMD_V1P2_NS || lomNamespace === '*') {
        stringElement = titleElement.getElementsByTagNameNS(this.IMSMD_V1P2_NS, 'langstring')[0] as Element | null ||
          titleElement.getElementsByTagName('langstring')[0] as Element | null;
      }
      if (!stringElement && lomNamespace) {
        stringElement = (lomNamespace !== '*') ?
          titleElement.getElementsByTagNameNS(lomNamespace, 'string')[0] as Element | null :
          titleElement.getElementsByTagName('string')[0] as Element | null;
      }

      if (!stringElement) {
        const directTitle = titleElement.textContent?.trim();
        return directTitle || null;
      }
      return stringElement.textContent?.trim() || null;
    } catch (error) {
      console.error('An unexpected error occurred during manifest title extraction:', error);
      return null;
    }
  }

  public isWebLinkXml(file: ImsccFile, xmlDoc?: XMLDocument | null): boolean {
    let docToCheck: XMLDocument | null = xmlDoc || null;
    const fileName = file?.name || 'unknown file';
    if (!docToCheck) {
      if (!file || !(file.mimeType?.includes('xml')) || typeof file.data !== 'string') { // Looser check for XML mime type
        return false;
      }
      try {
        const parser = new DOMParser();
        const cleanData = file.data.charCodeAt(0) === 0xFEFF ? file.data.substring(1) : file.data;
        const parsedDoc: XMLDocument = parser.parseFromString(cleanData, "application/xml"); // Use application/xml
        if (!parsedDoc || !parsedDoc.documentElement || parsedDoc.querySelector('parsererror')) return false;
        docToCheck = parsedDoc;
      } catch (e) {return false;}
    }
    if (!docToCheck) return false;
    // Check for webLink element using namespace or tag name
    return !!(docToCheck.getElementsByTagNameNS(this.WEBLINK_V1P2_NS, 'webLink')[0] || docToCheck.getElementsByTagName('webLink')[0]);
  }

  public extractWebLinkUrl(file: ImsccFile, xmlDoc: XMLDocument): string | null {
    // Find the url element (namespace aware) and get its href attribute
    const urlElement = xmlDoc.getElementsByTagNameNS(this.WEBLINK_V1P2_NS, 'url')[0] || xmlDoc.getElementsByTagName('url')[0];
    return urlElement?.getAttribute('href') || null;
  }

  // This function extracts the *description* from the discussion XML file (e.g., topic_1.xml)
  public isTopicXml(file: ImsccFile, xmlDoc?: XMLDocument | null): boolean {
    let docToCheck: XMLDocument | null = xmlDoc || null;
    const fileName = file?.name || 'unknown file';
    // console.log(`[ParsingHelper] Checking isTopicXml for: ${fileName}`); // Verbose logging
    if (!docToCheck) {
      // If no pre-parsed doc, check if the file looks like XML and parse it
      if (!file || !(file.mimeType?.includes('xml')) || typeof file.data !== 'string') {
        // console.log(`[ParsingHelper] isTopicXml: Returning false (no doc, invalid file/data) for ${fileName}`);
        return false;
      }
      try {
        // console.log(`[ParsingHelper] isTopicXml: Attempting to parse file data for ${fileName}`);
        const parser = new DOMParser();
        const cleanData = file.data.charCodeAt(0) === 0xFEFF ? file.data.substring(1) : file.data; // Remove BOM if present
        const parsedDoc: XMLDocument = parser.parseFromString(cleanData, "application/xml"); // Use stricter parser
        if (!parsedDoc || !parsedDoc.documentElement || parsedDoc.querySelector('parsererror')) {
          console.warn(`[ParsingHelper] isTopicXml: XML parsing failed or found parsererror for ${fileName}`);
          return false;
        }
        docToCheck = parsedDoc;
      } catch (e) {
        console.error(`[ParsingHelper] isTopicXml: Exception parsing XML for ${fileName}`, e);
        return false;
      }
    }
    if (!docToCheck) {
      // console.warn(`[ParsingHelper] isTopicXml: docToCheck is null after parsing attempt for ${fileName}`);
      return false;
    }

    // Check for the <topic> element (namespace aware)
    const topicElement = docToCheck.getElementsByTagNameNS(this.DISCUSSIONTOPIC_V1P1_NS, 'topic')[0] || docToCheck.getElementsByTagName('topic')[0];
    if (!topicElement) {
      // console.log(`[ParsingHelper] isTopicXml: No <topic> element found in ${fileName}`);
      return false;
    }

    // Check for the <text> element within <topic> (namespace aware)
    const textElement = topicElement.getElementsByTagNameNS(this.DISCUSSIONTOPIC_V1P1_NS, 'text')[0] || topicElement.getElementsByTagName('text')[0];
    if (!textElement) {
      // console.log(`[ParsingHelper] isTopicXml: No <text> element found within <topic> in ${fileName}`);
      return false;
    }

    // console.log(`[ParsingHelper] isTopicXml: Found <topic> and <text> elements in ${fileName}. Returning true.`);
    return true;
  }

  // This function extracts the *description* from the discussion XML file (e.g., topic_1.xml)
  // The main content is typically in a separate HTML file linked by a different resource.
  public extractTopicDescriptionHtml(file: ImsccFile, xmlDoc: XMLDocument): string | null {
    const fileName = file?.name || 'unknown file';
    console.log(`[ParsingHelper] Attempting extractTopicDescriptionHtml for: ${fileName}`);

    // Find the <topic> element (namespace aware)
    const topicElement = xmlDoc.getElementsByTagNameNS(this.DISCUSSIONTOPIC_V1P1_NS, 'topic')[0] || xmlDoc.getElementsByTagName('topic')[0];
    if (!topicElement) {
      console.warn(`[ParsingHelper] extractTopicDescriptionHtml: No <topic> element found in XML for file: ${fileName}`);
      return null;
    }

    // Find the <text> element within <topic> (namespace aware)
    const textElement = topicElement.getElementsByTagNameNS(this.DISCUSSIONTOPIC_V1P1_NS, 'text')[0] || topicElement.getElementsByTagName('text')[0];
    // Check if the text element exists and has the correct texttype attribute
    if (textElement && textElement.getAttribute('texttype') === 'text/html') {
      const rawTextContent = textElement.textContent || "";
      console.log(`[ParsingHelper] Raw textContent from <text> for topic "${fileName}": ${rawTextContent.substring(0, 250)}...`);

      // Decode HTML entities (e.g., < becomes <)
      // Note: Sometimes this content is double-encoded, the tryDecodeURIComponent might be useful here
      // const decodedHtmlContent = this.tryDecodeURIComponent(rawTextContent); // Or just simple htmlEntitiesDecode
      const decodedHtmlContent = htmlEntitiesDecode(rawTextContent);

      console.log(`[ParsingHelper] Decoded HTML content (Description) for topic "${fileName}": ${decodedHtmlContent.substring(0, 250)}...`);

      const trimmedContent = decodedHtmlContent.trim();

      // The description text is typically just HTML fragments, maybe with <p> or <div>.
      // We should keep the HTML structure for display, but remove comments and handle basic sanitization.
      // Removing ALL HTML tags (as the old check implied) was incorrect for keeping structure.
      // Remove HTML comments (<-- ... -->)
      const contentWithoutComments = trimmedContent.replace(/<!--[\s\S]*?-->/g, '').trim();


      // Check if there is any non-whitespace content left
      if (contentWithoutComments.replace(/\s/g, '') !== '') {
        console.log(`[ParsingHelper] Final extracted HTML Description for topic "${fileName}" (has textual content): ${contentWithoutComments.substring(0, 150)}...`);
        // Apply basic display preprocessing
        return this.preProcessHtmlForDisplay(contentWithoutComments);
      }

      // Log a warning if no significant text content was found
      console.warn(`[ParsingHelper] <text> element in topic "${fileName}" has no significant textual content after decoding, trimming, and removing comments. Content without comments was: "${contentWithoutComments.substring(0, 150)}..."`);
      return null;
    }
    // Log a warning if the <text> element or the texttype attribute is missing/incorrect
    console.warn(`[ParsingHelper] No <text texttype="text/html"> element found in topic "${fileName}", or texttype attribute missing/incorrect.`);
    return null;
  }

  /**
   * Extracts the content from a generic HTML file.
   * This function is intended to be used for the *main content* HTML file
   * associated with a discussion topic, identified via a resource link in manifest.
   * @param file The ImsccFile object representing the HTML file.
   * @returns The HTML content string, or null if the file is not suitable or empty.
   */
  public extractHtmlFileContent(file: ImsccFile): string | null {
    if (!file || typeof file.data !== 'string') {
      console.warn(`[ParsingHelper] extractHtmlFileContent: Invalid file or data type for file: ${file?.name}`);
      return null;
    }

    // Optionally check mime type, though .html extension is a strong indicator
    // if (!file.mimeType?.includes('html')) {
    //     console.warn(`[ParsingHelper] extractHtmlFileContent: File does not appear to be HTML: ${file.name}, MimeType: ${file.mimeType}`);
    //     // Depending on strictness, you might still process it if data is string
    // }

    const rawContent = file.data;
    console.log(`[ParsingHelper] Extracting HTML content from file: ${file.name}`);

    // HTML files themselves should ideally not be HTML-entity encoded at the top level,
    // but content *within* them might be. Standard DOM parsing handles this.
    // However, Common Cartridge can sometimes wrap entire file contents in XML elements
    // which might lead to odd encoding issues. Let's apply htmlEntitiesDecode just in case
    // it was incorrectly encoded at a higher level before being stored in file.data.
    const decodedContent = htmlEntitiesDecode(rawContent);

    // Apply display preprocessing (remove artifacts, fix img links etc.)
    const processedContent = this.preProcessHtmlForDisplay(decodedContent);

    // Remove HTML comments (<-- ... -->) - important for potentially hidden content
    const contentWithoutComments = processedContent.replace(/<!--[\s\S]*?-->/g, '').trim();


    // Check if there is any non-whitespace content left after processing
    if (contentWithoutComments.replace(/\s/g, '') === '') {
      console.warn(`[ParsingHelper] Extracted HTML content from ${file.name} is empty or only whitespace after processing.`);
      return null;
    }

    console.log(`[ParsingHelper] Successfully extracted HTML content from file: ${file.name} (first 250 chars): ${contentWithoutComments.substring(0, 250)}...`);
    return contentWithoutComments;
  }


  public preProcessHtmlForDisplay(html: string): string {
    let processedHtml = html;
    // Remove <a> tags that only wrap an <img> tag pointing to the same image source (common artifact from some LMS exports)
    const imageLinkRegex = /<a\s+[^>]*?href=(["'])([^"']*?\.(png|jpe?g|gif|bmp|svg|webp))\1[^>]*?>(?:\s*<br\s*\/?>\s*)?(<img\s+[^>]*?src=(["'])(?:[^"'>]*\/)?\2\5[^>]*?>)(?:\s*<br\s*\/?>\s*)?<\/a>/gi;
    // Corrected regex slightly to handle paths like 'content/images/image.png'
    const correctedImageLinkRegex = /<a\s+[^>]*?href=(["'])([^"']*?\.(png|jpe?g|gif|bmp|svg|webp))(?:\?[\s\S]*?)?\1[^>]*?>(?:\s*<br\s*\/?>\s*)?(<img\s+[^>]*?src=(["'])(?:[^"'>]*\/)?\2\5[^>]*?>)(?:\s*<br\s*\/?>\s*)?<\/a>/gi;

    processedHtml = processedHtml.replace(correctedImageLinkRegex, (match, _q1, _href, _ext, imgTag) => {
      // console.log("Removed image link wrapper:", match); // Log what's being removed
      return imgTag; // Keep only the <img> tag
    });


    // Attempt to close unclosed <a> tags heuristically (might not cover all cases)
    // This is tricky and potentially fragile. Add logging to see if it's necessary/working.
    // console.log("Attempting to close unclosed <a> tags...");
    processedHtml = processedHtml.replace(/(<a\s+[^>]*?>[^<>]*(?:<br\s*\/?>)?)(?=\s*<\/(?:div|p|h[1-6]|ul|ol|li|body|html|table|tbody|tr|td)>)/gi, '$1</a>'); // Before closing block elements
    processedHtml = processedHtml.replace(/(<a\s+[^>]*?>[^<]*?)(?=\s*<a\s)/gi, '$1</a>'); // Before another opening <a>
    processedHtml = processedHtml.replace(/(<a\s+[^>]*?>\s*(?:<br\s*\/?>)?\s*)(<input\s)/gi, '$1</a>$2'); // Before an <input>
    processedHtml = processedHtml.replace(/(<a\s+[^>]*?>[^<]*?)(?=\s*(?:<\/(?:div|p|h[1-6]|ul|ol|li|body|html|table|tbody|tr|td)>)?\s*$)/gi, '$1</a>'); // At the end of the string or before end of document/block

    // Remove empty <a> tags (after attempting to close)
    processedHtml = processedHtml.replace(/<a\s+[^>]*?>\s*<\/a>/gi, '');

    // Normalize <br> tags to be self-closing
    processedHtml = processedHtml.replace(/<br\s*(?!\/)\s*>/gi, '<br />');


    // Remove zero-width space character often found in HTML exports
    processedHtml = processedHtml.replace(/\u200B/g, '');


    // Remove common empty or near-empty block elements that might clutter display
    processedHtml = processedHtml.replace(/<p>\s*(?: ?)*\s*<\/p>/gi, ''); // Empty paragraphs
    processedHtml = processedHtml.replace(/<div>\s*(?: ?)*\s*<\/div>/gi, ''); // Empty divs
    // Add more specific cleaning based on observed output if needed


    return processedHtml;
  }
}
