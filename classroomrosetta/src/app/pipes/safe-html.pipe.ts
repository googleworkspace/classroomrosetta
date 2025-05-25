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

import {Pipe, PipeTransform} from '@angular/core';
import {DomSanitizer, SafeHtml} from '@angular/platform-browser';

@Pipe({
  name: 'safeHtml',
  standalone: true
})
export class SafeHtmlPipe implements PipeTransform {

  constructor(private sanitizer: DomSanitizer) { }

  /**
   * Transforms a string containing HTML into a SafeHtml object,
   * which tells Angular that the HTML is trusted and should not be sanitized.
   * IMPORTANT: Only use this with HTML content that you know is safe to render.
   * Binding untrusted HTML can lead to XSS vulnerabilities.
   * @param value The HTML string to trust.
   * @returns A SafeHtml object that can be bound to [innerHTML].
   */
  transform(value: string | null | undefined): SafeHtml | null {
    if (value === null || value === undefined || value === '') {
      return null;
    }
    return this.sanitizer.bypassSecurityTrustHtml(value);
  }
}
