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

import { Component, inject, OnInit } from '@angular/core';
import { AuthService } from '../services/auth/auth.service';
import {MatButtonModule} from '@angular/material/button';
import {User} from '@angular/fire/auth';

@Component({
  selector: 'app-profile-login',
  imports: [MatButtonModule],
  templateUrl: './profile-login.component.html',
  styleUrl: './profile-login.component.scss'
})
export class ProfileLoginComponent implements OnInit {
  authService = inject(AuthService)
  user: User|null = null
  ngOnInit(){
    this.authService.user$.subscribe(user=>{
      this.user = user
    })
  }

  async onGoogleSignIn(): Promise<void> {
    try {
      await this.authService.signInWithGoogle();
    } catch (error) {
      console.error('Google Sign-In error:', error);
    }
  }
  async onGoogleSignOut(): Promise<void> {
    try {
      await this.authService.googleLogout();
    } catch (error) {
      console.error('Google Sign-Out error:', error);
    }
  }

}
