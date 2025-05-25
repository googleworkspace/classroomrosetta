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

import {Injectable, inject, OnDestroy} from '@angular/core';
import {
  user, // Observable stream of the current user
  User, // Firebase User interface
  Auth, // Firebase Auth instance
  signOut, // Firebase sign out function
  GoogleAuthProvider, // Google Auth provider
  signInWithPopup, // Sign in method
  browserSessionPersistence, // Persistence type
  UserCredential, // Type for sign-in result
  setPersistence, // Function to set persistence
  OAuthCredential // Type for OAuth credential
} from '@angular/fire/auth';
import {Observable, Subscription, BehaviorSubject} from 'rxjs';

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

// Keys for storing Google OAuth Access Token and its expiration time in session storage
const GOOGLE_ACCESS_TOKEN_KEY = 'googleOAuthAccessToken';
const GOOGLE_ACCESS_TOKEN_EXPIRES_AT_KEY = 'googleOAuthAccessTokenExpiresAt';

// Buffer time in seconds before actual expiry to attempt proactive refresh
const TOKEN_REFRESH_BUFFER_SECONDS = 60; // Refresh 1 minute before actual expiry
const DEFAULT_TOKEN_EXPIRY_SECONDS = 3600; // Default to 1 hour (3600 seconds)

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
  private googleAccessToken: string | null = null;
  // Stores the timestamp (in milliseconds) when the Google OAuth access token expires.
  private googleAccessTokenExpiresAt: number | null = null;
  // Timer for scheduling proactive token refresh
  private tokenRefreshTimer: any = null;

  constructor() {
    console.log("AuthService: Initializing...");
    console.log("AuthService: Injected Auth instance:", this.auth);

    try {
      console.log("AuthService: Attempting setPersistence in constructor...");
      if (!this.auth) {
        console.error("AuthService: Auth instance is null/undefined in constructor before setPersistence call.");
      } else {
        setPersistence(this.auth, browserSessionPersistence)
          .then(() => {
            console.log("AuthService: Firebase persistence successfully set to session storage in constructor.");
          })
          .catch((error) => {
            console.warn("AuthService: Initial attempt to set persistence in constructor failed (will be re-attempted on sign-in):", error);
            console.log("AuthService: Auth object state at time of initial persistence warning:", this.auth);
          });
      }
    } catch (syncError) {
      console.error("AuthService: Synchronous error during initial persistence setup in constructor:", syncError);
      console.log("AuthService: Auth object state at time of sync error:", this.auth);
    }

    // Attempt to load the Google access token and its expiry from session storage on service initialization
    this.loadTokenFromStorage();

    // Subscribe to Firebase auth state changes
    this.userSubscription = user(this.auth).subscribe(firebaseUser => {
      console.log("AuthService: Firebase Auth state changed:", firebaseUser ? firebaseUser.uid : 'No user');
      this.userSubject.next(firebaseUser);

      if (!firebaseUser) {
        this.clearGoogleToken();
        console.log("AuthService: User logged out or Firebase session ended. Cleared Google Access Token and refresh timer.");
      } else {
        if (!this.googleAccessToken) {
          this.loadTokenFromStorage();
        }

        if (this.googleAccessToken && this.googleAccessTokenExpiresAt) {
          console.log("AuthService: Firebase user present and Google Access Token available/loaded.");
          if (!this.tokenRefreshTimer && Date.now() < this.googleAccessTokenExpiresAt - (TOKEN_REFRESH_BUFFER_SECONDS * 1000)) {
            this.scheduleTokenRefresh();
          }
        } else {
          console.warn(`AuthService: Firebase user ${firebaseUser.uid} is authenticated, but Google Access Token is missing or expired. Application may need to prompt for Google Sign-In.`);
        }
      }
    });
  }

  ngOnDestroy(): void {
    console.log("AuthService: Destroying - Unsubscribing from auth state changes and clearing timers.");
    this.userSubscription?.unsubscribe();
    if (this.tokenRefreshTimer) {
      clearTimeout(this.tokenRefreshTimer);
      this.tokenRefreshTimer = null;
    }
  }

  get currentUser(): User | null {
    return this.userSubject.getValue();
  }

  async signInWithGoogle(): Promise<User | null> {
    const provider = new GoogleAuthProvider();
    SCOPES.forEach(scope => {
      provider.addScope(scope);
    });

    console.log("AuthService: Attempting Google Sign-In via popup...");
    try {
      await setPersistence(this.auth, browserSessionPersistence);
      console.log("AuthService: Ensured session persistence is set before popup.");

      const result: UserCredential = await signInWithPopup(this.auth, provider);
      const firebaseUser = result.user;
      console.log("AuthService: Google Sign-In successful for Firebase user:", firebaseUser.uid);

      const credential = GoogleAuthProvider.credentialFromResult(result) as OAuthCredential | null;

      if (credential?.accessToken) {
        this.googleAccessToken = credential.accessToken;

        // We use the default expiry time (1 hour).
        const expiresInSeconds = DEFAULT_TOKEN_EXPIRY_SECONDS;
        console.log(`AuthService: Using default token expiry: ${expiresInSeconds}s`);

        this.googleAccessTokenExpiresAt = Date.now() + expiresInSeconds * 1000;
        this.saveTokenToStorage(this.googleAccessToken, this.googleAccessTokenExpiresAt);
        this.scheduleTokenRefresh(); // Schedule refresh for the new token
        console.log("AuthService: New Google Access Token obtained, stored, and refresh scheduled. Expires at:", new Date(this.googleAccessTokenExpiresAt).toISOString());
      } else {
        console.warn("AuthService: Could not retrieve Google credential or access token from sign-in result.");
        this.clearGoogleToken();
      }
      return firebaseUser;

    } catch (error: any) {
      console.error("AuthService: Error signing in with Google: ", error.code, error.message);
      this.clearGoogleToken();
      throw error;
    }
  }

  async googleLogout(): Promise<void> {
    const currentUserId = this.currentUser?.uid;
    console.log("AuthService: Attempting Firebase Logout for user:", currentUserId ?? 'N/A');
    this.clearGoogleToken();
    try {
      await signOut(this.auth);
      console.log("AuthService: Firebase Sign out successful.");
    } catch (error: any) {
      console.error('AuthService: Firebase Logout error:', error.code, error.message);
      throw error;
    }
  }

  getGoogleAccessToken(): string | null {
    if (!this.googleAccessToken) {
      this.loadTokenFromStorage();
    }

    if (this.isTokenLikelyExpired(TOKEN_REFRESH_BUFFER_SECONDS / 2)) {
      console.warn("AuthService: getGoogleAccessToken() - Token is missing, expired, or about to expire. Returning null.");
      if (this.googleAccessTokenExpiresAt && Date.now() >= this.googleAccessTokenExpiresAt) {
        this.clearGoogleToken();
      }
      return null;
    }
    return this.googleAccessToken;
  }

  isTokenLikelyExpired(bufferSeconds: number = TOKEN_REFRESH_BUFFER_SECONDS): boolean {
    if (!this.googleAccessToken || !this.googleAccessTokenExpiresAt) {
      return true;
    }
    return Date.now() >= (this.googleAccessTokenExpiresAt - bufferSeconds * 1000);
  }

  // --- Private Helper Methods ---

  private saveTokenToStorage(token: string, expiresAt: number): void {
    try {
      sessionStorage.setItem(GOOGLE_ACCESS_TOKEN_KEY, token);
      sessionStorage.setItem(GOOGLE_ACCESS_TOKEN_EXPIRES_AT_KEY, expiresAt.toString());
    } catch (e) {
      console.error("AuthService: Failed to save token and/or expiry to session storage.", e);
    }
  }

  private loadTokenFromStorage(): void {
    try {
      const storedToken = sessionStorage.getItem(GOOGLE_ACCESS_TOKEN_KEY);
      const storedExpiresAtString = sessionStorage.getItem(GOOGLE_ACCESS_TOKEN_EXPIRES_AT_KEY);

      if (storedToken && storedExpiresAtString) {
        this.googleAccessToken = storedToken;
        this.googleAccessTokenExpiresAt = parseInt(storedExpiresAtString, 10);

        if (this.isTokenLikelyExpired(0)) {
          console.warn("AuthService: Loaded Google Access Token from storage is expired. Clearing it.");
          this.clearGoogleToken();
        } else {
          console.log("AuthService: Google Access Token and expiry loaded from session storage. Expires at:", new Date(this.googleAccessTokenExpiresAt).toISOString());
          this.scheduleTokenRefresh();
        }
      } else {
        this.googleAccessToken = null;
        this.googleAccessTokenExpiresAt = null;
      }
    } catch (e) {
      console.error("AuthService: Failed to load token and/or expiry from session storage.", e);
      this.clearGoogleToken();
    }
  }

  private clearGoogleToken(): void {
    const wasPresent = !!this.googleAccessToken;
    this.googleAccessToken = null;
    this.googleAccessTokenExpiresAt = null;

    if (this.tokenRefreshTimer) {
      clearTimeout(this.tokenRefreshTimer);
      this.tokenRefreshTimer = null;
      console.log("AuthService: Cleared proactive token refresh timer.");
    }

    try {
      sessionStorage.removeItem(GOOGLE_ACCESS_TOKEN_KEY);
      sessionStorage.removeItem(GOOGLE_ACCESS_TOKEN_EXPIRES_AT_KEY);
      if (wasPresent) {
        console.log("AuthService: Google Access Token and expiry cleared from memory and session storage.");
      }
    } catch (e) {
      console.error("AuthService: Failed to remove token and/or expiry from session storage.", e);
    }
  }

  private scheduleTokenRefresh(): void {
    if (this.tokenRefreshTimer) {
      clearTimeout(this.tokenRefreshTimer);
      this.tokenRefreshTimer = null;
    }

    if (!this.googleAccessToken || !this.googleAccessTokenExpiresAt) {
      console.log("AuthService: Cannot schedule token refresh, token or expiry missing.");
      return;
    }

    const refreshDelay = this.googleAccessTokenExpiresAt - Date.now() - (TOKEN_REFRESH_BUFFER_SECONDS * 1000);

    if (refreshDelay > 0) {
      this.tokenRefreshTimer = setTimeout(async () => {
        console.log("AuthService: Proactive refresh timer triggered. Attempting to refresh Google Access Token...");
        await this.attemptProactiveTokenRefresh();
      }, refreshDelay);
      console.log(`AuthService: Proactive token refresh scheduled in ${Math.round(refreshDelay / 1000)}s.`);
    } else {
      if (Date.now() >= this.googleAccessTokenExpiresAt) {
        console.warn("AuthService: Token is already expired. Proactive refresh not scheduled. User re-authentication needed.");
      } else {
        console.log("AuthService: Token is within the refresh buffer or past its ideal proactive refresh time. Refresh will occur on demand or if user re-signs in.");
      }
    }
  }

  private async attemptProactiveTokenRefresh(): Promise<void> {
    console.log("AuthService: Attempting proactive Google Access Token refresh...");
    if (!this.currentUser) {
      console.log("AuthService: No Firebase user currently signed in. Skipping proactive refresh.");
      return;
    }
    try {
      await this.signInWithGoogle();
      console.log("AuthService: Proactive token refresh attempt completed (see signInWithGoogle logs for outcome).");
    } catch (error) {
      console.warn("AuthService: Proactive token refresh attempt failed.", error);
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
