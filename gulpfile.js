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

// Import necessary Gulp and Node.js modules
const { series, src, dest, parallel } = require('gulp'); // Added parallel for API enabling
const { exec } = require('child_process'); // Destructure exec directly
const clean = require('gulp-clean');
const fs = require("fs"); // Filesystem module, now used

// Load environment variables from .env file
require('dotenv').config();

// --- General Configuration ---
const PROJECT_NAME = process.env.PROJECT_NAME;
const APP_NAME = process.env.APP_NAME;
const DIST_PATH = process.env.DIST_PATH; // Output folder name in Angular app's dist (e.g., project-name)
const FIREBASE_PROJECT_PATH = process.env.FIREBASE_PROJECT_PATH || './firebase'; // Path to firebase project dir

// --- Apps Script Configuration ---
const APPS_SCRIPT_DIR = process.env.APPS_SCRIPT_DIR || 'appsscript'; // Directory for your Apps Script files
const APPS_SCRIPT_PROJECT_TITLE = process.env.APPS_SCRIPT_PROJECT_TITLE || `${PROJECT_NAME || 'My'}FormsHelperAPI`;

// --- Google Cloud Configuration ---
const GCLOUD_PROJECT_ID = process.env.GCLOUD_PROJECT_ID; // Project ID for gcloud commands

// --- Error Handling for Missing Core Environment Variables ---
if (!PROJECT_NAME || !APP_NAME || !DIST_PATH) {
  console.error("Error: Missing required Angular build environment variables (PROJECT_NAME, APP_NAME, DIST_PATH).");
  console.error("Please ensure your .env file is correctly configured.");
  process.exit(1);
}

// --- Path Constants ---
const angularAppPath = `./${APP_NAME}`;
const angularDistPath = `${angularAppPath}/dist/${DIST_PATH}`;
const deployPath = './deploy';
const deployClientPath = `${deployPath}/client`;
const claspJsonPath = `${APPS_SCRIPT_DIR}/.clasp.json`;
const appsscriptJsonPath = `${APPS_SCRIPT_DIR}/appsscript.json`;

// =================================================================================================
// == UTILITY FUNCTION FOR EXECUTING COMMANDS
// =================================================================================================
/**
 * Executes a shell command and returns a Promise.
 * @param {string} command - The command to execute.
 * @param {object} options - Options for child_process.exec (e.g., cwd).
 * @param {string} taskName - A descriptive name for the task for logging.
 * @returns {Promise<void>}
 */
function executeCommand(command, options = {}, taskName = 'Command') {
  return new Promise((resolve, reject) => {
    exec(command, options, (err, stdout, stderr) => {
      if (stdout) console.log(`[${taskName} - STDOUT]:\n${stdout}`);
      // Log stderr as errors only if err is present, otherwise as info/warnings
      if (stderr) {
          if (err) console.error(`[${taskName} - STDERR]:\n${stderr}`);
          else console.log(`[${taskName} - STDERR/Info]:\n${stderr}`); // Clasp often uses stderr for info
      }
      if (err) {
        console.error(`[${taskName} - ERROR]: Execution failed for command: ${command}`, err);
        reject(err);
      } else {
        console.log(`[${taskName}]: Successfully executed: ${command}`);
        resolve();
      }
    });
  });
}

// =================================================================================================
// == ANGULAR & FIREBASE TASKS
// =================================================================================================

function cleanDeployDirectory() {
  return src(deployPath, { read: false, allowEmpty: true })
    .pipe(clean({ force: true }));
}

function buildClientApp() {
  console.log(`Building Angular app: ${APP_NAME}...`);
  const command = `"ng" build`; // Consider --configuration production
  return executeCommand(command, { cwd: angularAppPath }, 'Angular Build');
}

function copyClientOutput() {
  console.log(`Copying client files from ${angularDistPath} to ${deployClientPath}...`);
  return src([`${angularDistPath}/**/*`], { encoding: false, nodir: true })
    .pipe(dest(deployClientPath));
}

function deployToFirebaseHosting() {
  console.log(`Deploying to Firebase hosting for project: ${PROJECT_NAME}...`);
  const command = `firebase deploy --only hosting --force`;
  return executeCommand(command, { cwd: FIREBASE_PROJECT_PATH }, 'Firebase Deploy');
}

// =================================================================================================
// == GOOGLE APPS SCRIPT (CLASP) TASKS
// =================================================================================================

function ensureAppsScriptDirExists(cb) {
  if (!fs.existsSync(APPS_SCRIPT_DIR)) {
    console.log(`Creating Apps Script directory: ${APPS_SCRIPT_DIR}`);
    fs.mkdirSync(APPS_SCRIPT_DIR, { recursive: true });
  }
  cb();
}

function ensureAppScriptManifest(cb) {
  if (!fs.existsSync(appsscriptJsonPath)) {
    console.log(`Creating default manifest file: ${appsscriptJsonPath}`);
    const defaultManifest = {
      timeZone: process.env.APPS_SCRIPT_TIMEZONE || "America/New_York",
      dependencies: {},
      exceptionLogging: "STACKDRIVER",
      runtimeVersion: "V8",
      executionApi: {
        access: "DOMAIN"
      }
    };
    if (!fs.existsSync(APPS_SCRIPT_DIR)) {
      fs.mkdirSync(APPS_SCRIPT_DIR, { recursive: true });
    }
    fs.writeFileSync(appsscriptJsonPath, JSON.stringify(defaultManifest, null, 2));
    console.log(`${appsscriptJsonPath} created successfully. IMPORTANT: Review and customize oauthScopes, webapp.access, and executionApi.access settings!`);
  } else {
    console.log(`Manifest file ${appsscriptJsonPath} already exists. IMPORTANT: Ensure oauthScopes, webapp, and executionApi settings are correctly configured for your API!`);
  }
  cb();
}

function initClaspProject() {
  if (fs.existsSync(claspJsonPath)) {
    console.log(`Clasp project already initialized at ${claspJsonPath}. Skipping creation.`);
    return Promise.resolve();
  }
  console.log(`Initializing new Clasp project: ${APPS_SCRIPT_PROJECT_TITLE} in ${APPS_SCRIPT_DIR}`);
  const command = `clasp create --type standalone --title "${APPS_SCRIPT_PROJECT_TITLE}" --rootDir .`;
  return executeCommand(command, { cwd: APPS_SCRIPT_DIR }, 'Clasp Init')
    .catch(err => {
        console.error(`Ensure 'clasp' is installed globally (npm install -g @google/clasp) and you are logged in (clasp login).`);
        throw err; // Re-throw to stop Gulp series
    });
}

function pushAppScriptCode() {
  console.log(`Pushing code (including appsscript.json) to Apps Script project from ${APPS_SCRIPT_DIR}...`);
  console.log(`This makes the latest code available for both Web App deployments and Apps Script API (scripts.run) execution.`);
  const command = `clasp push --force`;
  return executeCommand(command, { cwd: APPS_SCRIPT_DIR }, 'Clasp Push');
}

function deployAppScript() {
  console.log('Deploying Apps Script project as a Web App (creating/updating /exec URL)...');
  // This command creates a new versioned Web App deployment.
  // The deployment URL will be output by clasp.
  const command = `clasp deploy --description "Automated Gulp Web App deployment"`;
  return executeCommand(command, { cwd: APPS_SCRIPT_DIR }, 'Clasp Deploy (Web App)');
}

// =================================================================================================
// == GOOGLE CLOUD API MANAGEMENT TASKS
// =================================================================================================

/**
 * Enables specified Google Cloud APIs for the configured project.
 */
async function enableGoogleCloudAPIs() {
  if (!GCLOUD_PROJECT_ID) {
    console.error('Error: GCLOUD_PROJECT_ID environment variable is not set.');
    console.error('Please set it in your .env file or ensure gcloud is configured with a default project.');
    return Promise.reject(new Error('GCLOUD_PROJECT_ID not set.'));
  }

  console.log(`Enabling Google Cloud APIs for project: ${GCLOUD_PROJECT_ID}...`);
  console.log('This includes enabling "script.googleapis.com" which is required for using the Apps Script API (scripts.run).');

  const apisToEnable = [
    'drive.googleapis.com',        // Google Drive API
    'compute.googleapis.com',      // Compute Engine API
    'classroom.googleapis.com',    // Google Classroom API
    'forms.googleapis.com',        // Google Forms API
    'identitytoolkit.googleapis.com', // Identity Toolkit API
    'script.googleapis.com',       // Apps Script API (for scripts.run and managing projects)
    'iamcredentials.googleapis.com', // IAM Service Account Credentials API
    'sts.googleapis.com'           // Security Token Service API
  ];

  for (const api of apisToEnable) {
    const command = `gcloud services enable ${api} --project=${GCLOUD_PROJECT_ID}`;
    try {
      await executeCommand(command, {}, `Enable API: ${api}`);
    } catch (error) {
      console.error(`Failed to enable API: ${api}. Halting further API enabling.`);
      throw error;
    }
  }
  console.log('All specified Google Cloud APIs have been processed.');
}


// =================================================================================================
// == GULP TASK EXPORTS
// =================================================================================================

exports.clean = cleanDeployDirectory;
exports.ngBuild = buildClientApp;
exports.copy = copyClientOutput;
exports.deployFirebase = deployToFirebaseHosting;

exports.build = series(
  cleanDeployDirectory,
  buildClientApp,
  copyClientOutput,
  deployToFirebaseHosting
);

// Initializes local clasp setup, ensures manifest (for webapp and executionApi)
exports.claspInit = series(ensureAppsScriptDirExists, ensureAppScriptManifest, initClaspProject);

// Pushes code to Google Drive, making it available for `scripts.run` (via HEAD) and for new web app deployments
exports.claspPush = series(ensureAppsScriptDirExists, ensureAppScriptManifest, pushAppScriptCode); // ensure manifest before push

// Creates/updates a versioned Web App deployment (generates /exec URL)
exports.claspDeployWebApp = series(ensureAppsScriptDirExists, ensureAppScriptManifest, pushAppScriptCode, deployAppScript); // push latest code before deploying webapp

// Full sequence: ensure local setup, push code, then deploy a web app version.
// After this, your script is callable via its /exec URL AND via scripts.run (using the latest pushed code).
exports.manageAppScript = series(
  ensureAppsScriptDirExists,
  ensureAppScriptManifest,   // Creates/verifies manifest with webapp and executionApi settings
  initClaspProject,          // Initializes clasp if needed
  pushAppScriptCode,         // Pushes current code (makes it executable via scripts.run)
  deployAppScript            // Deploys a new Web App version (/exec URL)
);

exports.enableCloudApis = enableGoogleCloudAPIs;


console.log("Gulpfile loaded. Tasks: gulp build, gulp manageAppScript, gulp claspInit, gulp claspPush, gulp claspDeployWebApp, gulp enableCloudApis, ...");
