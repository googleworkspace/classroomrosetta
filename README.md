# IMSCC to Google Classroom Converter - Project Synopsis

## Overview

This project is designed to facilitate the migration of course content from learning management systems (LMS) that export in the IMS Common Cartridge (IMSCC) format to Google Classroom. It aims to automate the conversion of various IMSCC package components into compatible Google Classroom assignments, materials, and structures.

## Core Functionality: IMSCC Parsing Helper Service

A key component of this system is the **IMSCC Parsing Helper Service**. This service is responsible for the initial processing and interpretation of the unzipped IMSCC package contents. Its primary functions include:

* **Manifest Parsing:** Reading and interpreting the `imsmanifest.xml` file, which serves as the blueprint for the course structure and its resources.
* **Metadata Extraction:** Extracting crucial metadata from the manifest and associated resource files. This includes:
    * Course titles.
    * Module/Item titles (using LOM - Learning Object Metadata - standards where available, and handling various LOM versions and namespaces).
    * Resource-specific information.
* **Content Extraction:**
    * Identifying and extracting HTML content from discussion topics (`imsdt_v1p1` schema), which often contains the core instructional text.
    * Extracting URLs from web link resources (`imswl_v1p2` schema).
* **Resource Identification & Handling:**
    * Identifying different types of resources within the package, such as:
        * Discussion Topics
        * Web Links
        * QTI (Question & Test Interoperability) assessments/quizzes (though the detailed conversion is handled by other services).
        * General web content and files.
    * Resolving relative file paths specified in the manifest and HTML content to locate actual files within the package.
* **Data Cleaning and Preparation:**
    * Decoding URI components and HTML entities to ensure content is correctly interpreted.
    * Sanitizing names (e.g., for topics, assignments) to be compatible with Google Classroom requirements.
    * Pre-processing HTML content to normalize certain tags (e.g., `<br>`) and clean up common artifacts (e.g., redundant links wrapping images) before further conversion or display.
* **Utility Functions:** Providing various helper functions for path manipulation, MIME type detection (in conjunction with other services), and handling character encoding issues (like BOM or Cyrillic path corrections).

## Goal

The ultimate goal is to provide a streamlined workflow for educators to transfer their existing course materials from IMSCC-compliant platforms into Google Classroom, minimizing manual effort and preserving as much of the original course structure and content fidelity as possible. The `ImsccParsingHelperService` lays the groundwork for this by making the complex IMSCC structure understandable and usable by subsequent conversion and API interaction services.


## Installation

A GCP project is needed to run this project. Billing needs to be applied to work with Firebase. Firebase hosting is used to host the application which rund client side.

1. Install Firebase Cli
```
npm install -g firebase-tools
```

2. Create a directory
```
mkdir classroomrosetta
cd classroomrosetta
```

3. Pull the github project
```
git clone https://github.com/xxxxxxx
```

4. Initialize Node
```
npm install
```

5. Initialize Firebase
```
firebase init
```

6. Configure the .env file and replace as much as is available.
```
PROJECT_NAME=<Project name>
GCLOUD_PROJECT_ID=<Project name>
APP_NAME=<App name for angular>
DIST_PATH=<Distribution path>
FIREBASE_API_KEY=<Firebase API key from the firebase console>
FIREBASE_AUTH_DOMAIN=<firebase domain>
FIREBASE_PROJECT_ID=<Project name>
FIREBASE_STORAGE_BUCKET=<Firebase storage Bucket>
FIREBASE_MESSAGING_SENDER_ID=<Firebase messaging sender id>
FIREBASE_APP_ID=<Firebase app id>
FIREBASE_CLIENT_ID=<Firebase client id>
APPS_SCRIPT_EXECUTION_API_URL=<Deployed Apps Script API URL>
```

7. Deploy using the Gulp Command
```
gulp fullSetupAndDeploy
```

8. Deploy the Apps Script file as an API. The project will need the OAuth consent screent completed for internal use and the project number associated in the settings. Replace the APPS_SCRIPT_EXECUTION_API_URL with the API url.

9. Run the Gulp deployment again.
```
gulp fullSetupAndDeploy
```
