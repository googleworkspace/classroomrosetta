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

import {Component, Input, Output, EventEmitter, OnChanges, SimpleChanges, ChangeDetectionStrategy, ChangeDetectorRef, OnInit} from '@angular/core';
import {CommonModule} from '@angular/common';
import {MatExpansionModule} from '@angular/material/expansion';
import {MatCheckboxModule} from '@angular/material/checkbox';
import {MatListModule} from '@angular/material/list';
import {MatIconModule} from '@angular/material/icon';
import {MatButtonModule} from '@angular/material/button';
import {MatDividerModule} from '@angular/material/divider';
import {ClassroomMultiselectComponent} from '../classroom-select/classroom-multiselect.component';

import {ProcessedCourseWork, GroupedCoursework, SubmissionData} from '../interfaces/classroom-interface'; // Adjust path

import {SafeHtmlPipe} from '../pipes/safe-html.pipe';

@Component({
  selector: 'app-coursework-display',
  standalone: true,
  imports: [
    CommonModule,
    MatExpansionModule,
    MatCheckboxModule,
    MatListModule,
    MatIconModule,
    MatButtonModule,
    MatDividerModule,
    ClassroomMultiselectComponent,
    SafeHtmlPipe
  ],
  templateUrl: './coursework-display.component.html',
  styleUrl: './coursework-display.component.scss',
  changeDetection: ChangeDetectionStrategy.OnPush
})
export class CourseworkDisplayComponent implements OnChanges, OnInit {

  @Input() assignments: ProcessedCourseWork[] = [];

  @Output() submitSelection = new EventEmitter<SubmissionData>();

  groupedCoursework: GroupedCoursework[] = [];
  selection = new Map<string, boolean>();
  isAllSelected = false;
  isIndeterminate = false;

  selectedClassroomIds: string[] = [];

  constructor(private cdRef: ChangeDetectorRef) { }

  ngOnInit(): void {
    if (this.assignments.length > 0) {
      this.groupAndInitialize();
      this.updateMasterCheckboxState();
    }
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['assignments']) {
      this.groupAndInitialize();
      this.updateMasterCheckboxState();
      this.cdRef.markForCheck();
    }
  }

  private groupAndInitialize(): void {
    const groups = new Map<string | undefined, ProcessedCourseWork[]>();
    this.selection.clear();
    this.assignments.forEach(item => {
      const itemId = item.associatedWithDeveloper?.id;
      if (!itemId) {
        console.warn('Coursework item missing ID, cannot be selected:', item.title);
        return;
      }

      const topic = item.associatedWithDeveloper?.topic;
      const group = groups.get(topic);
      if (group) {
        group.push(item);
      } else {
        groups.set(topic, [item]);
      }
      this.selection.set(itemId, false);
    });

    this.groupedCoursework = Array.from(groups.entries()).map(([topic, items]) => ({
      topic: topic ?? 'General',
      items: items,
      allSelected: false,
      indeterminate: false
    })).sort((a, b) => (a.topic ?? '').localeCompare(b.topic ?? ''));

    this.groupedCoursework.forEach(group => this.updateTopicCheckboxState(group));
  }

  toggleItemSelection(itemId: string, group: GroupedCoursework): void {
    const currentState = this.selection.get(itemId);
    this.selection.set(itemId, !currentState);
    this.updateTopicCheckboxState(group);
    this.updateMasterCheckboxState();
    this.cdRef.markForCheck();
  }

  toggleTopicSelection(group: GroupedCoursework): void {
    const targetState = !group.allSelected;
    group.items.forEach(item => {
      const itemId = item.associatedWithDeveloper?.id;
      if (itemId) {
        this.selection.set(itemId, targetState);
      }
    });
    this.updateTopicCheckboxState(group);
    this.updateMasterCheckboxState();
    this.cdRef.markForCheck();
  }

  toggleAllSelection(): void {
    const targetState = !this.isAllSelected;
    this.selection.forEach((_, key) => {
      this.selection.set(key, targetState);
    });
    this.groupedCoursework.forEach(group => this.updateTopicCheckboxState(group));
    this.updateMasterCheckboxState();
    this.cdRef.markForCheck();
  }

  private updateTopicCheckboxState(group: GroupedCoursework): void {
    let allSelectedInGroup = true;
    let noneSelectedInGroup = true;

    if (group.items.length === 0) {
      group.allSelected = false;
      group.indeterminate = false;
      return;
    }

    for (const item of group.items) {
      const itemId = item.associatedWithDeveloper?.id;
      if (itemId) {
        if (this.selection.get(itemId)) {
          noneSelectedInGroup = false;
        } else {
          allSelectedInGroup = false;
        }
      } else {
        allSelectedInGroup = false;
      }
      if (!allSelectedInGroup && !noneSelectedInGroup) {
        break;
      }
    }

    group.allSelected = allSelectedInGroup;
    group.indeterminate = !allSelectedInGroup && !noneSelectedInGroup;
  }

  private updateMasterCheckboxState(): void {
    let allSelectedOverall = true;
    let noneSelectedOverall = true;

    if (this.selection.size === 0) {
      this.isAllSelected = false;
      this.isIndeterminate = false;
      return;
    }

    this.selection.forEach(isSelected => {
      if (isSelected) {
        noneSelectedOverall = false;
      } else {
        allSelectedOverall = false;
      }
    });

    this.isAllSelected = allSelectedOverall;
    this.isIndeterminate = !allSelectedOverall && !noneSelectedOverall;
  }

  handleClassroomSelection(selectedIds: string[]): void {
    this.selectedClassroomIds = selectedIds;
    this.cdRef.markForCheck();
  }

  onSubmit(): void {
    const selectedAssignmentIds = Array.from(this.selection.entries())
      .filter(([_, isSelected]) => isSelected)
      .map(([id, _]) => id);

    const submissionData: SubmissionData = {
      classroomIds: this.selectedClassroomIds,
      assignmentIds: selectedAssignmentIds
    };

    console.log('Submitting data:', submissionData);
    this.submitSelection.emit(submissionData);
  }

  getSelectedAssignmentCount(): number {
    return Array.from(this.selection.values()).filter(isSelected => isSelected).length;
  }

  getMaterialIcon(material: any): string {
    if (material.link) return 'link';
    if (material.driveFile) return 'folder_open';
    if (material.form) return 'assignment';
    if (material.youtubeVideo) return 'play_circle_outline';
    return 'attachment';
  }

  getMaterialTitle(material: any): string {
    return material.link?.title || material.link?.url ||
      material.driveFile?.driveFile?.title ||
      material.form?.title ||
      material.youtubeVideo?.title ||
      'Attached Material';
  }

  getMaterialLink(material: any): string | undefined {
    return material.link?.url ||
      material.driveFile?.driveFile?.alternateLink ||
      material.form?.formUrl ||
      material.youtubeVideo?.alternateLink;
  }

  trackByGroupTopic(index: number, group: GroupedCoursework): string | undefined {
    return group.topic;
  }

  trackByItemId(index: number, item: ProcessedCourseWork): string {
    const itemId = item.associatedWithDeveloper?.id;
    return itemId ?? `item-index-${index}`;
  }
}
