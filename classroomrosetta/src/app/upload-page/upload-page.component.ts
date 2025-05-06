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

import { Component, inject, OnInit  } from '@angular/core';
import { MenuComponent } from '../menu/menu.component';
import { AuthService } from '../services/auth/auth.service';
import { FileUploadComponent } from '../file-upload/file-upload.component';


@Component({
  selector: 'app-upload-page',
  imports: [MenuComponent, FileUploadComponent],
  templateUrl: './upload-page.component.html',
  styleUrl: './upload-page.component.scss'
})
export class UploadPageComponent implements OnInit{
  auth = inject(AuthService)
  loggedin = false;
  ngOnInit(): void {
    this.auth.user$.subscribe(user=>{
      this.loggedin =!!user
    })
  }
}
