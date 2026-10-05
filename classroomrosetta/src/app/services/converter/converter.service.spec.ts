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

import { TestBed } from '@angular/core/testing';
import { firstValueFrom, toArray } from 'rxjs';
import { ConverterService } from './converter.service';
import { ImsccFile, ProcessedCourseWork } from '../../interfaces/classroom-interface';

describe('ConverterService', () => {
  let service: ConverterService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(ConverterService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });

  it('should convert Canvas assignments when LearningModules organization has no child items', async () => {
    const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<manifest identifier="canvas_export_test" xmlns="http://www.imsglobal.org/xsd/imscp_v1p1">
  <metadata>
    <schema>IMS Common Cartridge</schema>
    <schemaversion>1.1.0</schemaversion>
  </metadata>
  <organizations>
    <organization identifier="org_1" default="true">
      <item identifier="LearningModules"/>
    </organization>
  </organizations>
  <resources>
    <resource identifier="res_week11" type="associatedcontent/imscc_xmlv1p1/learning-application-resource" href="ge280e4651228f6bccfc68bfb9c7a1e5e/assignment_settings.xml">
      <file href="ge280e4651228f6bccfc68bfb9c7a1e5e/assignment_settings.xml"/>
      <file href="ge280e4651228f6bccfc68bfb9c7a1e5e/week-11-october-20-24.html"/>
    </resource>
    <resource identifier="res_course_settings" type="associatedcontent/imscc_xmlv1p1/learning-application-resource" href="course_settings/canvas_export.txt">
      <file href="course_settings/canvas_export.txt"/>
    </resource>
    <resource identifier="res_banner" type="webcontent" href="web_resources/banner.png">
      <file href="web_resources/banner.png"/>
    </resource>
  </resources>
</manifest>`;

    const assignmentSettingsXml = `<?xml version="1.0" encoding="UTF-8"?>
<assignment xmlns="http://canvas.instructure.com/xsd/cccv1p0" identifier="ge280e4651228f6bccfc68bfb9c7a1e5e">
  <title>Week 11: October 20-24</title>
  <due_at>2025-10-24T23:59:00Z</due_at>
  <points_possible>100.0</points_possible>
  <workflow_state>published</workflow_state>
  <submission_types>online_upload</submission_types>
</assignment>`;

    const assignmentHtml = `<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
  <title>Assignment: Week 11: October 20-24</title>
</head>
<body>
  <table style="width: 53.3954%;" border="1">
    <tbody>
      <tr>
        <td><strong>Date</strong></td>
        <td><strong>Monday, October 20</strong></td>
      </tr>
      <tr>
        <td>Lesson</td>
        <td>3.3 Solving Quadratics Using Square Roots</td>
      </tr>
      <tr>
        <td>Objective</td>
        <td>Students will solve a quadratic equation using square roots.</td>
      </tr>
      <tr>
        <td>Agenda</td>
        <td>
          <ul>
            <li>Bellringer</li>
            <li><a href="$IMS-CC-FILEBASE$/Uploaded%20Media/WS%203.3.pdf">WS 3.3.pdf</a></li>
          </ul>
        </td>
      </tr>
    </tbody>
  </table>
</body>
</html>`;

    const files: ImsccFile[] = [
      { name: 'imsmanifest.xml', data: manifestXml, mimeType: 'text/xml' },
      { name: 'ge280e4651228f6bccfc68bfb9c7a1e5e/assignment_settings.xml', data: assignmentSettingsXml, mimeType: 'text/xml' },
      { name: 'ge280e4651228f6bccfc68bfb9c7a1e5e/week-11-october-20-24.html', data: assignmentHtml, mimeType: 'text/html' },
      { name: 'course_settings/canvas_export.txt', data: 'course_settings_data', mimeType: 'text/plain' },
      { name: 'web_resources/banner.png', data: new ArrayBuffer(8), mimeType: 'image/png' },
      { name: 'web_resources/Uploaded Media/WS 3.3.pdf', data: 'pdf_content', mimeType: 'application/pdf' },
    ];

    const results: ProcessedCourseWork[] = await firstValueFrom(service.convertImscc(files).pipe(toArray()));

    expect(results.length).toBe(1);
    const item = results[0];

    expect(item.workType).toBe('ASSIGNMENT');
    expect(item.title).toBe('Week 11: October 20-24');
    expect(item.maxPoints).toBe(100);
    expect(item.state).toBe('PUBLISHED');
    expect(item.dueDate).toBeUndefined();
    expect(item.dueTime).toBeUndefined();
    expect(item.descriptionForDisplay).toContain('3.3 Solving Quadratics Using Square Roots');
    expect(item.descriptionForDisplay).toContain('Students will solve a quadratic equation using square roots');

    // Verify attached worksheet file from $IMS-CC-FILEBASE$
    expect(item.localFilesToUpload.length).toBe(1);
    expect(item.localFilesToUpload[0].targetFileName).toBe('WS 3.3.pdf');
  });

  it('should fallback to HTML title tag when assignment_settings.xml has no title', async () => {
    const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<manifest identifier="canvas_export_test" xmlns="http://www.imsglobal.org/xsd/imscp_v1p1">
  <organizations>
    <organization identifier="org_1" default="true">
      <item identifier="LearningModules"/>
    </organization>
  </organizations>
  <resources>
    <resource identifier="res_week12" type="associatedcontent/imscc_xmlv1p1/learning-application-resource" href="assign_dir/assignment_settings.xml">
      <file href="assign_dir/assignment_settings.xml"/>
      <file href="assign_dir/week-12.html"/>
    </resource>
  </resources>
</manifest>`;

    const assignmentSettingsXml = `<?xml version="1.0" encoding="UTF-8"?>
<assignment xmlns="http://canvas.instructure.com/xsd/cccv1p0" identifier="assign_dir">
  <points_possible>50.0</points_possible>
</assignment>`;

    const assignmentHtml = `<html>
<head>
  <title>Assignment: Week 12: October 27-31</title>
</head>
<body>
  <p>Assignment content for week 12</p>
</body>
</html>`;

    const files: ImsccFile[] = [
      { name: 'imsmanifest.xml', data: manifestXml, mimeType: 'text/xml' },
      { name: 'assign_dir/assignment_settings.xml', data: assignmentSettingsXml, mimeType: 'text/xml' },
      { name: 'assign_dir/week-12.html', data: assignmentHtml, mimeType: 'text/html' },
    ];

    const results: ProcessedCourseWork[] = await firstValueFrom(service.convertImscc(files).pipe(toArray()));

    expect(results.length).toBe(1);
    const item = results[0];
    expect(item.workType).toBe('ASSIGNMENT');
    expect(item.title).toBe('Week 12: October 27-31');
    expect(item.maxPoints).toBe(50);
  });

  it('should convert multiple Canvas assignments in a package', async () => {
    const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<manifest identifier="canvas_export_test" xmlns="http://www.imsglobal.org/xsd/imscp_v1p1">
  <organizations>
    <organization identifier="org_1" default="true">
      <item identifier="LearningModules"/>
    </organization>
  </organizations>
  <resources>
    <resource identifier="res_1" type="associatedcontent/imscc_xmlv1p1/learning-application-resource" href="assign1/assignment_settings.xml">
      <file href="assign1/assignment_settings.xml"/>
      <file href="assign1/unit-1.html"/>
    </resource>
    <resource identifier="res_2" type="associatedcontent/imscc_xmlv1p1/learning-application-resource" href="assign2/assignment_settings.xml">
      <file href="assign2/assignment_settings.xml"/>
      <file href="assign2/unit-2.html"/>
    </resource>
  </resources>
</manifest>`;

    const settings1 = `<assignment xmlns="http://canvas.instructure.com/xsd/cccv1p0"><title>Unit 1</title><points_possible>10</points_possible></assignment>`;
    const html1 = `<html><head><title>Unit 1</title></head><body><p>Lesson 1</p></body></html>`;

    const settings2 = `<assignment xmlns="http://canvas.instructure.com/xsd/cccv1p0"><title>Unit 2</title><points_possible>20</points_possible></assignment>`;
    const html2 = `<html><head><title>Unit 2</title></head><body><p>Lesson 2</p></body></html>`;

    const files: ImsccFile[] = [
      { name: 'imsmanifest.xml', data: manifestXml, mimeType: 'text/xml' },
      { name: 'assign1/assignment_settings.xml', data: settings1, mimeType: 'text/xml' },
      { name: 'assign1/unit-1.html', data: html1, mimeType: 'text/html' },
      { name: 'assign2/assignment_settings.xml', data: settings2, mimeType: 'text/xml' },
      { name: 'assign2/unit-2.html', data: html2, mimeType: 'text/html' },
    ];

    const results: ProcessedCourseWork[] = await firstValueFrom(service.convertImscc(files).pipe(toArray()));

    expect(results.length).toBe(2);
    expect(results[0].title).toBe('Unit 1');
    expect(results[0].maxPoints).toBe(10);
    expect(results[1].title).toBe('Unit 2');
    expect(results[1].maxPoints).toBe(20);
  });

  it('should extract course title from CC v1.2 lomimscc metadata and handle empty resources cleanly', async () => {
    const manifestXml = `<?xml version="1.0" encoding="UTF-8"?>
<manifest xmlns="http://www.imsglobal.org/xsd/imsccv1p2/imscp_v1p1" identifier="cctd0001" xmlns:lomimscc="http://ltsc.ieee.org/xsd/imsccv1p2/LOM/manifest">
  <metadata>
    <schema>IMS Common Cartridge</schema>
    <schemaversion>1.2.0</schemaversion>
    <lomimscc:lom>
      <lomimscc:general>
        <lomimscc:title>
          <lomimscc:string>Home</lomimscc:string>
        </lomimscc:title>
      </lomimscc:general>
    </lomimscc:lom>
  </metadata>
  <organizations>
    <organization identifier="org" structure="rooted-hierarchy">
      <item identifier="root"/>
    </organization>
  </organizations>
  <resources/>
</manifest>`;

    const files: ImsccFile[] = [
      { name: 'imsmanifest.xml', data: manifestXml, mimeType: 'text/xml' }
    ];

    const results: ProcessedCourseWork[] = await firstValueFrom(service.convertImscc(files).pipe(toArray()));

    expect(service.coursename).toBe('Home');
    expect(results.length).toBe(0);
  });
});

