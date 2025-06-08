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

// Import necessary Gulp and Node.js modules
const { series, src, dest, parallel } = require('gulp');
const { exec } = require('child_process');
const clean = require('gulp-clean');
const fs = require("fs");
const path = require("path"); // For path joining

// Load environment variables from .env file
// We will call config() again later if we modify .env mid-process
require('dotenv').config();

// --- General Configuration ---
const PROJECT_NAME = process.env.PROJECT_NAME;
const APP_NAME = process.env.APP_NAME;
const DIST_PATH = process.env.DIST_PATH;
const FIREBASE_PROJECT_PATH = process.env.FIREBASE_PROJECT_PATH || './firebase'; // Root dir for Firebase setup
const ENV_FILE_PATH = process.env.ENV_FILE_PATH || '.env'; // Path to the .env file

// --- Apps Script Configuration ---
const APPS_SCRIPT_DIR = process.env.APPS_SCRIPT_DIR || 'appsscript';
const APPS_SCRIPT_PROJECT_TITLE = process.env.APPS_SCRIPT_PROJECT_TITLE || `${PROJECT_NAME || 'My'}FormsHelperAPI`;
const APPS_SCRIPT_API_ENV_VAR = 'APPS_SCRIPT_EXECUTION_API_URL'; // Key in .env for the API URL

// --- Google Cloud Configuration ---
const GCLOUD_PROJECT_ID = process.env.GCLOUD_PROJECT_ID; // Used for gcloud commands and can be Firebase project ID

// --- Error Handling for Missing Core Environment Variables ---
if (!PROJECT_NAME || !APP_NAME || !DIST_PATH) {
  console.error("Error: Missing required Angular build environment variables (PROJECT_NAME, APP_NAME, DIST_PATH).");
  console.error("Please ensure your .env file is correctly configured with these values.");
  process.exit(1);
}

// --- Path Constants ---
const angularAppPath = `./${APP_NAME}`;
const angularDistPath = `${angularAppPath}/dist/${DIST_PATH}`; // Source of Angular build output
const angularEnvDevPath = path.join(angularAppPath, 'src', 'environments', 'environment.development.ts');
const angularEnvProdPath = path.join(angularAppPath, 'src', 'environments', 'environment.ts');

// Firebase Hosting specific paths
const firebaseJsonConfigPath = path.join(FIREBASE_PROJECT_PATH, 'firebase.json');
const firebaseRcPath = path.join(FIREBASE_PROJECT_PATH, '.firebaserc'); // Path for .firebaserc
const firebaseHostingPublicDirName = 'public/browser'; // The value for "public" in firebase.json
const firebaseFinalDeployPath = path.join(FIREBASE_PROJECT_PATH, firebaseHostingPublicDirName); // Full path for client files

// Apps Script paths
const claspJsonPath = path.join(APPS_SCRIPT_DIR, '.clasp.json');
const appsscriptJsonPath = path.join(APPS_SCRIPT_DIR, 'appsscript.json');


// =================================================================================================
// == UTILITY FUNCTION FOR EXECUTING COMMANDS
// =================================================================================================
/**
 * Executes a shell command and returns a Promise.
 * @param {string} command - The command to execute.
 * @param {object} options - Options for child_process.exec (e.g., cwd).
 * @param {string} taskName - A descriptive name for the task for logging.
 * @returns {Promise<{stdout: string, stderr: string}>} Resolves with stdout and stderr
 */
function executeCommand(command, options = {}, taskName = 'Command') {
  return new Promise((resolve, reject) => {
    exec(command, options, (err, stdout, stderr) => {
      const logPrefix = `[${taskName}]`;
      if (stdout) console.log(`${logPrefix} STDOUT:\n${stdout}`);
      if (stderr) {
        if (err) console.error(`${logPrefix} STDERR:\n${stderr}`);
        else console.log(`${logPrefix} STDERR (Info):\n${stderr}`);
      }
      if (err) {
        console.error(`${logPrefix} ERROR: Execution failed for command: ${command}`, err);
        reject(err);
      } else {
        console.log(`${logPrefix} OK: Successfully executed: ${command}`);
        resolve({ stdout, stderr });
      }
    });
  });
}

// =================================================================================================
// == .ENV FILE MANAGEMENT UTILITY
// =================================================================================================
/**
 * Updates or adds a key-value pair in the .env file.
 * Creates the file if it doesn't exist.
 * @param {string} key - The environment variable key.
 * @param {string} value - The value for the environment variable.
 */
async function updateEnvFile(key, value) {
  console.log(`Updating ${ENV_FILE_PATH}: Setting ${key}=${value}`);
  let envFileContent = "";
  if (fs.existsSync(ENV_FILE_PATH)) {
    envFileContent = fs.readFileSync(ENV_FILE_PATH, 'utf8');
  }

  const lines = envFileContent.split('\n');
  let keyFound = false;
  const newLines = lines.map(line => {
    if (line.startsWith(`${key}=`)) {
      keyFound = true;
      return `${key}=${value}`;
    }
    return line;
  }).filter(line => line.trim() !== ''); // Remove empty lines that might result from split

  if (!keyFound) {
    newLines.push(`${key}=${value}`);
  }

  fs.writeFileSync(ENV_FILE_PATH, newLines.join('\n') + '\n'); // Add trailing newline
  console.log(`${ENV_FILE_PATH} updated successfully with ${key}.`);

  console.log('Reloading .env variables into process.env...');
  require('dotenv').config({ override: true, path: path.resolve(process.cwd(), ENV_FILE_PATH) });
}


// =================================================================================================
// == ANGULAR ENVIRONMENT FILE MANAGEMENT
// =================================================================================================
/**
 * Generates the content for Angular environment files.
 * @param {boolean} isProduction - True for production environment.ts, false for development.
 * @returns {string} The content of the environment file.
 */
function generateAngularEnvironmentContent(isProduction) {
  const firebaseConfigKeys = [
    'FIREBASE_API_KEY', 'FIREBASE_AUTH_DOMAIN', 'FIREBASE_PROJECT_ID',
    'FIREBASE_STORAGE_BUCKET', 'FIREBASE_MESSAGING_SENDER_ID', 'FIREBASE_APP_ID', 'FIREBASE_CLIENT_ID'
  ];
  for (const key of firebaseConfigKeys) {
    if (!process.env[key]) {
      console.warn(`Warning: Environment variable ${key} is not set in .env. It will be undefined in Angular's environment.ts.`);
    }
  }

  const appsScriptApiUrl = process.env[APPS_SCRIPT_API_ENV_VAR] || '';
  if (!appsScriptApiUrl) {
    console.warn(`Warning: ${APPS_SCRIPT_API_ENV_VAR} is not set in .env. formsItemsApi will be empty in Angular's environment.ts.`);
  }

  return `// Auto-generated by Gulp
// ${isProduction ? 'Production' : 'Development'} Environment

export const environment = {
  production: ${isProduction},
  firebaseConfig: {
    apiKey: "${process.env.FIREBASE_API_KEY || ''}",
    authDomain: "${process.env.FIREBASE_AUTH_DOMAIN || ''}",
    projectId: "${process.env.FIREBASE_PROJECT_ID || ''}",
    storageBucket: "${process.env.FIREBASE_STORAGE_BUCKET || ''}",
    messagingSenderId: "${process.env.FIREBASE_MESSAGING_SENDER_ID || ''}",
    appId: "${process.env.FIREBASE_APP_ID || ''}",
    clientId: "${process.env.FIREBASE_CLIENT_ID || ''}"
  },
  formsItemsApi: "${appsScriptApiUrl}"
  // Add other environment-specific variables here if needed
};
`;
}

async function updateAngularDevelopmentEnvironment() {
  const content = generateAngularEnvironmentContent(false);
  fs.writeFileSync(angularEnvDevPath, content, 'utf8');
  console.log(`Updated Angular environment file: ${angularEnvDevPath}`);
}

async function updateAngularProductionEnvironment() {
  const content = generateAngularEnvironmentContent(true);
  fs.writeFileSync(angularEnvProdPath, content, 'utf8');
  console.log(`Updated Angular environment file: ${angularEnvProdPath}`);
}

const updateAngularEnvironments = parallel(updateAngularDevelopmentEnvironment, updateAngularProductionEnvironment);

// =================================================================================================
// == DIRECTORY ENSURING TASKS
// =================================================================================================
function ensureAppsScriptDir(cb) {
  if (!fs.existsSync(APPS_SCRIPT_DIR)) {
    console.log(`Creating Apps Script directory: ${APPS_SCRIPT_DIR}`);
    fs.mkdirSync(APPS_SCRIPT_DIR, { recursive: true });
  }
  cb();
}

function ensureFirebaseProjectDir(cb) {
  if (!fs.existsSync(FIREBASE_PROJECT_PATH)) {
    console.log(`Creating Firebase project directory: ${FIREBASE_PROJECT_PATH}`);
    fs.mkdirSync(FIREBASE_PROJECT_PATH, { recursive: true });
  }
  cb();
}

// =================================================================================================
// == ANGULAR & FIREBASE TASKS
// =================================================================================================

function cleanFirebasePublicDirectory() {
  console.log(`Cleaning Firebase public directory: ${firebaseFinalDeployPath}`);
  return src(firebaseFinalDeployPath, { read: false, allowEmpty: true })
    .pipe(clean({ force: true }));
}

function buildClientApp() {
  console.log(`Building Angular app: ${APP_NAME}... (using configuration: production)`);
  const command = `"ng" build --configuration production`;
  return executeCommand(command, { cwd: angularAppPath }, 'Angular Build');
}

function copyClientOutputToFirebase() {
  console.log(`Copying client files from ${angularDistPath} to ${firebaseFinalDeployPath}...`);
  // Ensure the target directory exists before copying
  if (!fs.existsSync(firebaseFinalDeployPath)) {
    fs.mkdirSync(firebaseFinalDeployPath, { recursive: true });
    console.log(`Created directory: ${firebaseFinalDeployPath}`);
  }
  return src([`${angularDistPath}/**/*`], { encoding: false, nodir: true })
    .pipe(dest(firebaseFinalDeployPath));
}

async function ensureFirebaseRc() {
  // This task relies on FIREBASE_PROJECT_PATH already existing.
  const firebaseProjectId = process.env.FIREBASE_PROJECT_ID || GCLOUD_PROJECT_ID;
  if (!firebaseProjectId) {
    console.error('Error: FIREBASE_PROJECT_ID or GCLOUD_PROJECT_ID must be set in .env to create .firebaserc');
    return Promise.reject(new Error('Firebase project ID not set for .firebaserc.'));
  }

  if (!fs.existsSync(firebaseRcPath)) {
    console.log(`Creating default .firebaserc file: ${firebaseRcPath}`);
    const firebaseRcContent = {
      projects: {
        default: firebaseProjectId
      }
    };
    fs.writeFileSync(firebaseRcPath, JSON.stringify(firebaseRcContent, null, 2));
    console.log(`${firebaseRcPath} created successfully, pointing default to project: ${firebaseProjectId}.`);
  } else {
    console.log(`.firebaserc file ${firebaseRcPath} already exists.`);
  }
}

async function ensureFirebaseJson() {
  // This task relies on FIREBASE_PROJECT_PATH already existing.
  if (!fs.existsSync(firebaseJsonConfigPath)) {
    console.log(`Creating default Firebase configuration file: ${firebaseJsonConfigPath}`);
    const firebaseConfig = {
      hosting: {
        public: firebaseHostingPublicDirName, // e.g., "public/browser"
        ignore: [
          "firebase.json",
          "**/.*",
          "**/node_modules/**"
        ],
        rewrites: [
          {
            source: "**",
            destination: "/index.html"
          }
        ]
      }
    };
    fs.writeFileSync(firebaseJsonConfigPath, JSON.stringify(firebaseConfig, null, 2));
    console.log(`${firebaseJsonConfigPath} created successfully.`);
  } else {
    console.log(`Firebase configuration file ${firebaseJsonConfigPath} already exists.`);
  }
}

function deployToFirebaseHosting() {
  const firebaseProjectId = process.env.FIREBASE_PROJECT_ID || GCLOUD_PROJECT_ID;
  if (!firebaseProjectId) {
    console.error('Error: Firebase project ID not found. Set FIREBASE_PROJECT_ID or GCLOUD_PROJECT_ID in .env');
    return Promise.reject(new Error('Firebase project ID not set for deployment.'));
  }
  console.log(`Deploying to Firebase hosting for project: ${firebaseProjectId}... location: ${FIREBASE_PROJECT_PATH}`);
  const command = `firebase deploy --only hosting --project ${firebaseProjectId} --force`;
  return executeCommand(command, { cwd: FIREBASE_PROJECT_PATH }, 'Firebase Deploy');
}

// =================================================================================================
// == GOOGLE APPS SCRIPT (CLASP) TASKS
// =================================================================================================

function ensureAppScriptManifest(cb) {
  if (!fs.existsSync(appsscriptJsonPath)) {
    console.log(`Creating default manifest file: ${appsscriptJsonPath}`);
    const defaultManifest = {
      timeZone: process.env.APPS_SCRIPT_TIMEZONE || "America/New_York",
      dependencies: {
        // Example: libraries: [{userSymbol:"LibName", libraryId:"SCRIPT_ID_OF_LIBRARY", version:"DEVELOPMENT_OR_VERSION_NUMBER"}]
        // Example: enabledAdvancedServices: [{userSymbol: "Drive", serviceId: "drive", version: "v2"}]
      },
      exceptionLogging: "STACKDRIVER",
      runtimeVersion: "V8",
      webapp: { // Essential for `clasp deploy` to create a web app /exec URL
        access: "ANYONE_ANONYMOUS", // Or "ANYONE", "DOMAIN", "MYSELF"
        executeAs: "USER_ACCESSING" // Or "USER_DEPLOYING"
      },
      executionApi: { // Essential for script.googleapis.com/:run
        access: "ANYONE" // Or "DOMAIN", "MYSELF" (MYSELF is usually too restrictive for APIs called by web apps)
      }
    };
    fs.writeFileSync(appsscriptJsonPath, JSON.stringify(defaultManifest, null, 2));
    console.log(`${appsscriptJsonPath} created successfully. IMPORTANT: Review and customize access levels, oauthScopes (in dependencies/enabledAdvancedServices), webapp, and executionApi settings!`);
  } else {
    console.log(`Manifest file ${appsscriptJsonPath} already exists. Ensure it's configured correctly.`);
  }
  cb();
}

async function initClaspProject() {
  if (fs.existsSync(claspJsonPath)) {
    console.log(`Clasp project already initialized at ${claspJsonPath}. Checking scriptId...`);
    try {
      const claspConfig = JSON.parse(fs.readFileSync(claspJsonPath, 'utf8'));
      if (claspConfig.scriptId) {
        console.log(`Found scriptId: ${claspConfig.scriptId} in ${claspJsonPath}.`);
        return Promise.resolve();
        } else {
          console.warn(`${claspJsonPath} exists but does not contain a scriptId. Will attempt to create/re-init.`);
        }
    } catch (err) {
      console.warn(`Error reading ${claspJsonPath}, will attempt to create/re-init. Error: ${err.message}`);
    }
  }

  console.log(`Initializing new Clasp project: ${APPS_SCRIPT_PROJECT_TITLE} in ${APPS_SCRIPT_DIR}`);
  const command = `clasp create --type standalone --title "${APPS_SCRIPT_PROJECT_TITLE}" --rootDir .`;
  try {
    await executeCommand(command, { cwd: APPS_SCRIPT_DIR }, 'Clasp Create');
    console.log(`Clasp project created. ${claspJsonPath} should now contain the scriptId.`);
    if (!fs.existsSync(claspJsonPath)) {
      throw new Error(`${claspJsonPath} was not created after 'clasp create'. Check Clasp setup and permissions.`);
    }
  } catch (err) {
    console.error(`Clasp create failed. Ensure 'clasp' is installed globally (npm install -g @google/clasp) and you are logged in (clasp login).`);
    throw err;
  }
}

async function getScriptIdFromClaspJson() {
  if (!fs.existsSync(claspJsonPath)) {
    console.error(`Error: ${claspJsonPath} not found. Run 'gulp claspInit' first.`);
    throw new Error(`${claspJsonPath} not found.`);
  }
  try {
    const claspConfig = JSON.parse(fs.readFileSync(claspJsonPath, 'utf8'));
    if (!claspConfig.scriptId) {
      throw new Error(`scriptId not found in ${claspJsonPath}.`);
    }
    console.log(`Retrieved scriptId: ${claspConfig.scriptId}`);
    return claspConfig.scriptId;
  } catch (err) {
    console.error(`Error reading or parsing ${claspJsonPath}: ${err.message}`);
    throw err;
  }
}

function pushAppScriptCode() {
  console.log(`Pushing code from ${APPS_SCRIPT_DIR} to Google Apps Script project...`);
  const command = `clasp push --force`;
  return executeCommand(command, { cwd: APPS_SCRIPT_DIR }, 'Clasp Push');
}

function deployAppScript() {
  console.log('Creating a new versioned deployment for Apps Script project (Web App / API)...');
  const command = `clasp deploy --description "Automated Gulp deployment: ${new Date().toISOString()}"`;
  return executeCommand(command, { cwd: APPS_SCRIPT_DIR }, 'Clasp Deploy');
}

// =================================================================================================
// == GOOGLE CLOUD API MANAGEMENT TASKS
// =================================================================================================

async function enableGoogleCloudAPIs() {
  const effectiveGcloudProjectId = GCLOUD_PROJECT_ID || process.env.FIREBASE_PROJECT_ID;
  if (!effectiveGcloudProjectId) {
    console.warn('Warning: GCLOUD_PROJECT_ID or FIREBASE_PROJECT_ID environment variable is not set. API enabling will be skipped.');
    return Promise.resolve();
  }

  console.log(`Enabling Google Cloud APIs for project: ${effectiveGcloudProjectId}...`);
  const apisToEnable = [
    'drive.googleapis.com', 'compute.googleapis.com', 'classroom.googleapis.com',
    'forms.googleapis.com', 'identitytoolkit.googleapis.com', 'script.googleapis.com',
    'iamcredentials.googleapis.com', 'sts.googleapis.com', 'firebase.googleapis.com',
    'cloudresourcemanager.googleapis.com' // Useful for project interactions, like permissions checks
  ];

  for (const api of apisToEnable) {
    const command = `gcloud services enable ${api} --project=${effectiveGcloudProjectId}`;
    try {
      await executeCommand(command, {}, `Enable API: ${api}`);
    } catch (error) {
      console.error(`Failed to enable API: ${api}. Check gcloud configuration and permissions.`);
      throw error;
    }
  }
  console.log('All specified Google Cloud APIs have been processed for enabling.');
}


// =================================================================================================
// == GULP TASK EXPORTS AND SEQUENCES
// =================================================================================================

// Individual tasks
exports.clean = cleanFirebasePublicDirectory; // Renamed and targets correct path
exports.ngBuild = series(updateAngularEnvironments, buildClientApp);
exports.copy = copyClientOutputToFirebase; // Renamed and targets correct path
exports.deployFirebase = deployToFirebaseHosting;
exports.enableCloudApis = enableGoogleCloudAPIs;
exports.ensureFirebaseJson = ensureFirebaseJson;
exports.ensureFirebaseRc = ensureFirebaseRc;

// Apps Script tasks
exports.claspInit = series(ensureAppsScriptDir, ensureAppScriptManifest, initClaspProject);
exports.claspPush = series(ensureAppsScriptDir, ensureAppScriptManifest, pushAppScriptCode);
exports.claspDeployWebApp = series(exports.claspPush, deployAppScript);

// Combined sequences
const setupAppsScript = series(
  exports.claspInit,
  exports.claspPush,
  deployAppScript
);
exports.setupAppsScript = setupAppsScript;

const prepareFirebaseDir = series(
  ensureFirebaseProjectDir,
  ensureFirebaseRc,
  ensureFirebaseJson
);
exports.prepareFirebaseDir = prepareFirebaseDir;

const buildAndDeployClient = series(
  prepareFirebaseDir, // Ensures .firebaserc and firebase.json are set up first
  updateAngularEnvironments,
  cleanFirebasePublicDirectory, // Uses updated name and targets correct path
  buildClientApp,
  // copyClientOutputToFirebase,   // Uses updated name and targets correct path
  deployToFirebaseHosting
);
exports.buildAndDeployClient = buildAndDeployClient;


// Grand sequence
exports.fullSetupAndDeploy = series(
  enableGoogleCloudAPIs,
  ensureAppsScriptDir, // Ensure Apps Script dir exists before setup
  setupAppsScript,
  buildAndDeployClient
);

// Default task
exports.default = exports.fullSetupAndDeploy;


console.log("\n---------------------------------------------------------------------");
console.log("Gulpfile loaded. Ensure your .env file has all necessary variables:");
console.log("  PROJECT_NAME, APP_NAME, DIST_PATH");
console.log("  GCLOUD_PROJECT_ID (used by gcloud, can be same as FIREBASE_PROJECT_ID)");
console.log("  FIREBASE_PROJECT_ID (used for Firebase specific configs, can be same as GCLOUD_PROJECT_ID)");
console.log("  FIREBASE_API_KEY, FIREBASE_AUTH_DOMAIN, etc. (for Angular env files)");
console.log("  FIREBASE_CLIENT_ID (for Angular env files)");
console.log("  Optional: APPS_SCRIPT_DIR, APPS_SCRIPT_PROJECT_TITLE, APPS_SCRIPT_TIMEZONE, ENV_FILE_PATH");
console.log("---------------------------------------------------------------------\n");
console.log("Key tasks available:");
console.log("  gulp fullSetupAndDeploy (Default: Enables APIs, Sets up Apps Script, Builds & Deploys Client)");
console.log("  gulp buildAndDeployClient (Manages Firebase files, Angular envs, builds, deploys to Firebase)");
console.log("  gulp prepareFirebaseDir (Ensures Firebase project dir, .firebaserc, and firebase.json)");
console.log("  gulp ensureFirebaseRc   (Creates .firebaserc if it doesn't exist)");
console.log("  gulp ensureFirebaseJson (Creates firebase.json if it doesn't exist)");
console.log("  gulp enableCloudApis    (Enables configured Google Cloud APIs)");
console.log("  gulp setupAppsScript    (Initializes Clasp, gets API URL, pushes & deploys Apps Script)");
console.log("  gulp claspInit          (Initializes local Clasp project, ensures manifest)");
console.log("  gulp claspPush          (Pushes Apps Script code to Google Drive)");
console.log("  gulp claspDeployWebApp  (Pushes and then creates a new Apps Script deployment version)");
console.log("  gulp updateAngularEnvironments (Injects .env values into Angular environment files)");
console.log("\nTo run a task: gulp <taskName>");
console.log("Example: gulp fullSetupAndDeploy\n");

