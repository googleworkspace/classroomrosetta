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

import { Injectable, inject, OnDestroy } from '@angular/core';
import {
  user, // Observable stream of the current user
  User, // Firebase User interface
  Auth, // Firebase Auth instance
  signOut, // Firebase sign out function
  GoogleAuthProvider, // Google Auth provider
  signInWithPopup, // Sign in method
  browserSessionPersistence, // Persistence type
  UserCredential, // Type for sign-in result
  setPersistence // Function to set persistence
} from '@angular/fire/auth';
import { Observable, Subscription, BehaviorSubject } from 'rxjs';

// Define scopes required for Google APIs
const SCOPES = [
  "https://www.googleapis.com/auth/classroom.courses.readonly",
  "https://www.googleapis.com/auth/userinfo.profile",
  "https://www.googleapis.com/auth/userinfo.email",
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/forms",
  "https://www.googleapis.com/auth/classroom.topics",
  "https://www.googleapis.com/auth/classroom.coursework.me",
  "https://www.googleapis.com/auth/classroom.courseworkmaterials",
  "https://www.googleapis.com/auth/classroom.coursework.students",
  "https://www.googleapis.com/auth/script.external_request"
];

// Key for storing the Google OAuth Access Token in session storage
const GOOGLE_ACCESS_TOKEN_KEY = 'googleOAuthAccessToken';

@Injectable({
  providedIn: 'root'
})
export class AuthService implements OnDestroy {

  private auth = inject(Auth);
  // BehaviorSubject holds the current user state, allowing synchronous access and observable stream
  private userSubject = new BehaviorSubject<User | null>(null);
  // Public observable for components to subscribe to user state changes
  user$: Observable<User | null> = this.userSubject.asObservable();
  private userSubscription: Subscription;

  // Stores the *current* Google OAuth access token in memory.
  // This token expires (usually after 1 hour) and is NOT automatically refreshed by this service.
  private googleAccessToken: string | null = null;

  constructor() {
    console.log("AuthService: Initializing...");

    // Log the injected Auth instance *before* trying to use it
    console.log("AuthService: Injected Auth instance:", this.auth);

    // Attempt to set Firebase persistence right away in the constructor.
    try {
      console.log("AuthService: Attempting setPersistence in constructor...");
      setPersistence(this.auth, browserSessionPersistence)
        .then(() => {
          console.log("AuthService: Firebase persistence successfully set to session storage in constructor.");
        })
        .catch((error) => {
          console.error("AuthService: Error setting persistence via promise in constructor:", error);
          console.log("AuthService: Auth object state at time of error:", this.auth);
        });
    } catch (syncError) {
      console.error("AuthService: Synchronous error calling setPersistence in constructor:", syncError);
      console.log("AuthService: Auth object state at time of sync error:", this.auth);
    }

    // Attempt to load the Google access token from session storage on service initialization
    this.loadTokenFromStorage();

    // Subscribe to Firebase auth state changes
    this.userSubscription = user(this.auth).subscribe(firebaseUser => {
      console.log("AuthService: Firebase Auth state changed:", firebaseUser ? firebaseUser.uid : 'No user');
      this.userSubject.next(firebaseUser);

      if (!firebaseUser) {
        this.clearGoogleToken();
        console.log("AuthService: User logged out or Firebase session ended. Cleared Google Access Token.");
      } else {
        if (!this.googleAccessToken) {
          this.loadTokenFromStorage();
          if (this.googleAccessToken) {
            console.log("AuthService: Loaded Google Access Token from storage after auth state confirmation.");
            }
        }
        if (!this.googleAccessToken) {
          console.warn(`AuthService: Firebase user ${firebaseUser.uid} exists, but Google Access Token is missing. Forcing logout.`);
          this.googleLogout().catch(err => {
            console.error("AuthService: Error during forced logout:", err);
          });
        } else {
          console.log("AuthService: Firebase user present and Google Access Token available.");
        }
      }
    });
  }

  // Clean up the Firebase auth subscription when the service is destroyed
  ngOnDestroy(): void {
    console.log("AuthService: Destroying - Unsubscribing from auth state changes.");
    this.userSubscription?.unsubscribe();
  }

  /**
   * Gets the currently logged-in Firebase user synchronously.
   * @returns {User | null} The current Firebase user or null.
   */
  get currentUser(): User | null {
    return this.userSubject.getValue();
  }

  /**
   * Initiates the Google Sign-In process using a popup.
   * Stores the obtained Google OAuth access token in memory and session storage.
   * **Crucially, this function MUST be called again by the application
   * if a Google API call fails due to an expired/invalid token (e.g., 401 error)
   * in order to obtain a NEW access token.** Firebase does NOT auto-refresh this token.
   * @returns {Promise<User | null>} Resolves with the Firebase User object on success, or null/throws on error.
   */
  async signInWithGoogle(): Promise<User | null> {
    const provider = new GoogleAuthProvider();
    SCOPES.forEach(scope => {
      provider.addScope(scope);
    });

    console.log("AuthService: Attempting Google Sign-In via popup...");
    try {
      // Ensure persistence is set BEFORE the sign-in attempt for reliability.
      await setPersistence(this.auth, browserSessionPersistence);
      console.log("AuthService: Ensured session persistence is set before popup.");

      const result: UserCredential = await signInWithPopup(this.auth, provider);
      const firebaseUser = result.user;
      console.log("AuthService: Google Sign-In successful for Firebase user:", firebaseUser.uid);

      const credential = GoogleAuthProvider.credentialFromResult(result);
      if (credential?.accessToken) {
        this.googleAccessToken = credential.accessToken;
        this.saveTokenToStorage(this.googleAccessToken);
        console.log("AuthService: New Google Access Token obtained and stored.");
      } else {
         console.warn("AuthService: Could not retrieve Google credential or access token from sign-in result.");
        this.clearGoogleToken();
      }
      return firebaseUser;

    } catch (error: any) {
      console.error("AuthService: Error signing in with Google: ", error.code, error.message);
      if (error.code === 'auth/unsupported-persistence-type') {
        console.error("AuthService: Browser doesn't support session persistence (e.g., private mode?). Sign-in likely failed.");
      }
      this.clearGoogleToken();
      throw error; // Rethrow
    }
  }

  /**
   * Signs the current user out of Firebase.
   * Clears the Google OAuth access token from memory and session storage.
   */
  async googleLogout(): Promise<void> {
    const currentUserId = this.currentUser?.uid;
    console.log("AuthService: Attempting Firebase Logout for user:", currentUserId ?? 'N/A');
    this.clearGoogleToken(); // Clear token first
    try {
      await signOut(this.auth);
      console.log("AuthService: Firebase Sign out successful.");
    } catch (error: any) {
      console.error('AuthService: Firebase Logout error:', error.code, error.message);
      throw error;
    }
  }

  /**
   * Retrieves the *currently stored* Google OAuth access token (from memory or session storage).
   * **IMPORTANT:** This token might be expired! Handle 401 errors and call `signInWithGoogle()` to refresh.
   * @returns {string | null} The stored access token, or null if none is available.
   */
  getGoogleAccessToken(): string | null {
    if (!this.googleAccessToken) {
      this.loadTokenFromStorage();
    }
    return this.googleAccessToken;
  }

  // --- Private Helper Methods ---

  /**
   * Saves the Google OAuth access token to session storage.
   * @param {string} token The token to save.
   */
  private saveTokenToStorage(token: string): void {
    try {
      sessionStorage.setItem(GOOGLE_ACCESS_TOKEN_KEY, token);
      // console.log("AuthService: Google Access Token saved to session storage."); // Reduce log noise
    } catch (e) {
      console.error("AuthService: Failed to save token to session storage.", e);
    }
  }

  /**
   * Attempts to load the Google OAuth access token from session storage into memory.
   */
  private loadTokenFromStorage(): void {
    try {
      const storedToken = sessionStorage.getItem(GOOGLE_ACCESS_TOKEN_KEY);
      if (storedToken) {
        this.googleAccessToken = storedToken;
      }
    } catch (e) {
      console.error("AuthService: Failed to load token from session storage.", e);
      this.googleAccessToken = null;
    }
  }

  /**
   * Clears the Google OAuth access token from memory and session storage.
   */
  private clearGoogleToken(): void {
    const wasPresent = !!this.googleAccessToken;
    this.googleAccessToken = null;
    try {
      sessionStorage.removeItem(GOOGLE_ACCESS_TOKEN_KEY);
      if (wasPresent) {
        console.log("AuthService: Google Access Token cleared from memory and session storage.");
      }
    } catch (e) {
      console.error("AuthService: Failed to remove token from session storage.", e);
    }
  }
}


// --- Interfaces ---
export interface Profile {
  family_name: string;
  given_name: string;
  granted_scopes?: string;
  id: string;
  name: string;
  picture: string;
}

export interface UserA {
  family_name: string;
  given_name: string;
  id: string;
  name: string;
  picture: string;
  email: string;
  uid: string;
  id_token: string;
  over_18?: boolean;
  [key: string]: string | boolean | undefined;
}
