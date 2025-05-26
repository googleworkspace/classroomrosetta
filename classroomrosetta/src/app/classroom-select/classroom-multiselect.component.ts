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

import {Component, OnInit, OnDestroy, Input, Output, EventEmitter, SimpleChanges, OnChanges, inject, ChangeDetectionStrategy, ChangeDetectorRef} from '@angular/core';
import {Subscription, EMPTY} from 'rxjs';
import {catchError} from 'rxjs/operators';
import {FormsModule} from '@angular/forms';
import {MatSelectModule} from '@angular/material/select'; // Use MatSelectModule
import {MatFormFieldModule} from '@angular/material/form-field'; // Needed for mat-form-field wrapper
import {MatProgressSpinnerModule} from '@angular/material/progress-spinner';
import {Classroom} from '../interfaces/classroom-interface'; // Adjust path if needed
import {ClassroomService} from '../services/classroom/classroom.service'; // Adjust path if needed
import {AuthService} from '../services/auth/auth.service'; // Adjust path if needed

@Component({
  selector: 'app-classroom-multiselect',
  standalone: true,
  imports: [
    FormsModule,
    MatSelectModule,
    MatFormFieldModule,
    MatProgressSpinnerModule,
  ],
  templateUrl: './classroom-multiselect.component.html',
  styleUrls: ['./classroom-multiselect.component.scss'],
  changeDetection: ChangeDetectionStrategy.OnPush
})
export class ClassroomMultiselectComponent implements OnInit, OnDestroy, OnChanges {

  classroomService = inject(ClassroomService);
  auth = inject(AuthService);
  private cdRef = inject(ChangeDetectorRef);

  @Input() authToken: string = '';
  @Output() selectionChange = new EventEmitter<string[]>();

  allClassrooms: Classroom[] = [];
  selectedClassroomIds: string[] = [];
  isLoading = false;
  errorMessage: string | null = null;
  private classroomSubscription: Subscription | null = null;

  constructor() {
    this.authToken = this.auth.getGoogleAccessToken() || '';
  }

  ngOnChanges(changes: SimpleChanges): void {
    if (changes['authToken'] && changes['authToken'].currentValue) {
      this.fetchClassrooms();
    } else if (changes['authToken'] && !changes['authToken'].currentValue) {
      this.resetState();
      this.errorMessage = 'Authentication token is required to load classrooms.';
      this.cdRef.markForCheck();
    }
  }

  ngOnInit(): void {
    if (this.authToken && this.allClassrooms.length === 0) {
      this.fetchClassrooms();
    } else if (!this.authToken) {
      this.resetState();
      this.errorMessage = 'Authentication token is required. Login again.';
      this.cdRef.markForCheck();
    }
  }

  ngOnDestroy(): void {
    if (this.classroomSubscription) {
      this.classroomSubscription.unsubscribe();
    }
  }

  fetchClassrooms(): void {
    if (this.isLoading || !this.authToken) {
      return;
    }

    this.isLoading = true;
    this.errorMessage = null;
    this.cdRef.markForCheck();

    if (this.classroomSubscription) {
      this.classroomSubscription.unsubscribe();
    }

    this.classroomSubscription = this.classroomService.getActiveClassrooms().pipe(
      catchError(error => {
        console.error('Error fetching classrooms:', error);
        this.errorMessage = error?.message || 'Failed to load classrooms.';
        this.isLoading = false;
        this.allClassrooms = [];
        this.selectedClassroomIds = [];
        this.emitSelection();
        this.cdRef.markForCheck();
        return EMPTY;
      })
    ).subscribe({
      next: (classrooms) => {
        this.allClassrooms = classrooms;
        this.isLoading = false;
        // Optional: Preserve selection if items still exist after refresh
        // this.selectedClassroomIds = this.selectedClassroomIds.filter(id =>
        //    this.allClassrooms.some(c => c.id === id)
        // );
        this.cdRef.markForCheck();
      },
      error: (err) => {
        console.error('Subscription received error (should have been caught):', err);
        if (this.isLoading) {
          this.isLoading = false;
          this.cdRef.markForCheck();
        }
      },
      complete: () => {
        if (this.isLoading) {
          this.isLoading = false;
          this.cdRef.markForCheck();
          this.cdRef.detectChanges();
        }
      }
    });
  }

  private resetState(): void {
    if (this.classroomSubscription) {
      this.classroomSubscription.unsubscribe();
      this.classroomSubscription = null;
    }
    this.allClassrooms = [];
    this.selectedClassroomIds = [];
    this.isLoading = false;
    this.errorMessage = null;
    this.emitSelection();
  }


  /**
   * Called when the mat-select selection changes.
   * Emits the current selection and triggers change detection.
   */
  onSelectionChange(): void {
    this.emitSelection();
    this.cdRef.markForCheck(); // Mark for check as the selection array has changed
  }


  /**
   * Emits the current array of selected classroom IDs via the selectionChange Output.
   */
  private emitSelection(): void {
    // Ensure emitting a copy if downstream modifications are possible
    this.selectionChange.emit([...this.selectedClassroomIds]);
  }

  /**
   * TrackBy function for the @for loop iterating over mat-options.
   * Helps Angular optimize rendering.
   */
  trackByClassroomId(index: number, item: Classroom): string {
    const itemId = item?.id;
    return itemId ?? `classroom-index-${index}`;
  }
}
