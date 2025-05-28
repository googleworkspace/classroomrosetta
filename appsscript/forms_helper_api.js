/**
 * Creates items in a Google Form using a multi-pass approach to get Google-hosted image URLs.
 * For images associated with questions:
 * 1. A temporary ImageItem is created to get a processed image URL from FormApp.
 * 2. This processed URL is then used to fetch the blob for the actual question's image.
 * 3. The temporary ImageItem is deleted.
 * Question titles and descriptions are set directly. Image filenames are not used in descriptions.
 *
 * @param {string} formId The ID of the Google Form to modify.
 * @param {Array<Object>} requestsPayload An array of request objects, where each object defines an item to create.
 * @return {Object} A response object indicating success or failure.
 */
function createFormItemsInGoogleForm(formId, requestsPayload) {
  const functionName = "createFormItemsInGoogleForm"; // Using the function name from your provided base
  let responseMessage = { success: false, message: '', createdItems: 0, errors: [] };
  let createdItemCount = 0; // Ensure this is declared to be accessible in Pass 4 success block

  console.log(`[${functionName}] Execution started. Form ID: ${formId}, Number of requests: ${requestsPayload ? requestsPayload.length : 'N/A'}`);

  try {
    if (!formId) throw new Error("Missing 'formId' argument.");
    if (!requestsPayload || !Array.isArray(requestsPayload)) throw new Error("Missing or invalid 'requestsPayload' array argument.");

    const form = FormApp.openById(formId);
    // Work with a deep copy of the requests payload to modify image source URIs if needed by the strategy
    const processedRequestsPayload = JSON.parse(JSON.stringify(requestsPayload));

    const imageProcessInfo = [];
    // { originalRequestIndex: number, itemPath: string[], tempItemTitle: string, tempItemId: string,
    //   originalSourceUri: string, originalAltText: string, processedUriFailed?: boolean, finalSourceUriForQuestion?: string }


    // --- Pass 1: Create temporary ImageItems for all images to get their contentUris ---
    console.log(`[${functionName}] Pass 1: Creating temporary image items...`);
    processedRequestsPayload.forEach((request, index) => {
      if (request.createItem && request.createItem.item) {
        const itemData = request.createItem.item;
        let imageDetailPath = null;
        let imageDetails = null;

        if (itemData.questionItem && itemData.questionItem.image && itemData.questionItem.image.sourceUri) {
          imageDetailPath = ['questionItem', 'image'];
          imageDetails = itemData.questionItem.image;
        } else if (itemData.imageItem && itemData.imageItem.image && itemData.imageItem.image.sourceUri) {
          imageDetailPath = ['imageItem', 'image'];
          imageDetails = itemData.imageItem.image;
        }

        if (itemData.questionItem && itemData.questionItem.question && imageDetails && imageDetailPath && imageDetailPath[0] === 'questionItem') {
          const tempItemTitle = `TEMP_IMG_FOR_REQ_INDEX_${index}_${Utilities.getUuid().substring(0, 8)}`;
          const currentImageInfo = {
            originalRequestIndex: index,
            itemPath: imageDetailPath,
            tempItemTitle: tempItemTitle,
            tempItemId: null,
            originalSourceUri: imageDetails.sourceUri,
            originalAltText: imageDetails.altText,
            processedUriFailed: false,
            finalSourceUriForQuestion: imageDetails.sourceUri
          };
          imageProcessInfo.push(currentImageInfo);

          try {
            console.log(`[${functionName}] Fetching blob with backoff for temp image: ${tempItemTitle} from ${imageDetails.sourceUri}`);
            const imageBlob = fetchWithExponentialBackoff(imageDetails.sourceUri, { muteHttpExceptions: true }, 3, 1000).getBlob(); // Added muteHttpExceptions here for consistency

            console.log(`[${functionName}] Adding temp image: ${tempItemTitle}`);
            const tempImageItem = form.addImageItem().setImage(imageBlob).setTitle(tempItemTitle);
            currentImageInfo.tempItemId = tempImageItem.getId();
          } catch (e) {
            console.error(`[${functionName}] Error creating temp image (index ${index}, title ${tempItemTitle}): ${e.message}`);
            responseMessage.errors.push(`Temp image creation failed (index ${index}, source ${imageDetails.sourceUri}): ${e.message}`);
            currentImageInfo.processedUriFailed = true;
          }
        }
      }
    });

    // --- Pass 2: Fetch Form content via REST API to get contentUris for temporary images ---
    const imagesThatHadTempItemCreated = imageProcessInfo.filter(info => !info.processedUriFailed && info.tempItemId);
    if (imagesThatHadTempItemCreated.length > 0) {
      console.log(`[${functionName}] Pass 2: Fetching form content via REST API for ${imagesThatHadTempItemCreated.length} processed image URLs...`);
      try {
        const token = ScriptApp.getOAuthToken();
        const fieldsForGet = 'items(title,itemId,imageItem(image(contentUri)))';
        const formsApiUrl = `https://forms.googleapis.com/v1/forms/${formId}?fields=${encodeURIComponent(fieldsForGet)}`;

        const params = { method: "get", headers: { "Authorization": "Bearer " + token }, muteHttpExceptions: true };
        // Using fetchWithExponentialBackoff for this API call as well for robustness
        const formApiResponse = fetchWithExponentialBackoff(formsApiUrl, params, 3, 1500);
        const responseCode = formApiResponse.getResponseCode();
        const formJsonText = formApiResponse.getContentText();

        if (responseCode !== 200) throw new Error(`Forms API GET failed with status ${responseCode}: ${formJsonText}`);
        const formJson = JSON.parse(formJsonText);

        console.log(`[${functionName}] Successfully fetched form content. Processing ${formJson.items ? formJson.items.length : 0} items from API response.`);

        if (formJson.items) {
          formJson.items.forEach(apiItem => {
            const info = imagesThatHadTempItemCreated.find(ipi => ipi.tempItemTitle === apiItem.title);

            if (info && apiItem.imageItem && apiItem.imageItem.image && apiItem.imageItem.image.contentUri) {
              const processedContentUri = apiItem.imageItem.image.contentUri;
              console.log(`[${functionName}] Found contentUri for temp item "${apiItem.title}": ${processedContentUri}. Verifying...`);

              if (isSecureImageUrl(processedContentUri)) {
                console.log(`[${functionName}] URL is secure. Storing for final use with question.`);
                info.finalSourceUriForQuestion = processedContentUri;
              } else {
                console.warn(`[${functionName}] URL "${processedContentUri}" for temp item "${info.tempItemTitle}" is NOT secure. Original URI will be used if secure, otherwise image removed for this question.`);
                info.processedUriFailed = true;
                if (!isSecureImageUrl(info.originalSourceUri)) {
                  console.warn(`[${functionName}] Original URL "${info.originalSourceUri}" for question ${info.originalRequestIndex} is also NOT secure. Image will be removed.`);
                }
              }
            } else if (info) {
                console.warn(`[${functionName}] Could not find contentUri for temp item "${info.tempItemTitle}" (ID: ${info.tempItemId}) in API response. Original URI will be used if secure.`);
                info.processedUriFailed = true;
            }
          });
        }
      } catch (e) {
        console.error(`[${functionName}] Pass 2 Error (Fetching/Processing contentUris): ${e.message}. Stack: ${e.stack}`);
        responseMessage.errors.push(`Error fetching/processing image contentUris: ${e.message}`);
        imagesThatHadTempItemCreated.forEach(info => info.processedUriFailed = true);
      }
    }

    // --- Pass 3: Update original payload with processed URIs (or remove image if insecure) and delete temporary items ---
    console.log(`[${functionName}] Pass 3: Updating payload and deleting ${imageProcessInfo.filter(info => info.tempItemId).length} temporary image items...`);
    imageProcessInfo.forEach(info => {
      let targetImageObjectContainer = processedRequestsPayload[info.originalRequestIndex].createItem.item;
      let parentOfImage = targetImageObjectContainer;
      let keyForImage = null;

      // Traverse to the image object itself, and keep track of its parent and key
      for (let i = 0; i < info.itemPath.length; i++) {
        if (!targetImageObjectContainer) break;
        if (i < info.itemPath.length - 1) {
          parentOfImage = targetImageObjectContainer; // parent is one level up
          targetImageObjectContainer = targetImageObjectContainer[info.itemPath[i]];
        } else { // last part of path is the key for the image object itself
          keyForImage = info.itemPath[i]; // e.g. 'image'
          targetImageObjectContainer = targetImageObjectContainer[info.itemPath[i]]; // This is the image obj {sourceUri, altText}
        }
      }

      const targetImageObject = targetImageObjectContainer; // now refers to the actual image object e.g. {sourceUri: ..., altText: ...}

      if (targetImageObject) {
        if (info.finalSourceUriForQuestion && isSecureImageUrl(info.finalSourceUriForQuestion)) {
          targetImageObject.sourceUri = info.finalSourceUriForQuestion;
        } else if (isSecureImageUrl(info.originalSourceUri)) {
          targetImageObject.sourceUri = info.originalSourceUri;
           console.warn(`[${functionName}] Using original secure URI for question ${info.originalRequestIndex} as processed URI was not secure or failed.`);
        } else {
          console.warn(`[${functionName}] Both original and processed URIs for question ${info.originalRequestIndex} are insecure or unavailable. Removing image from request.`);
          responseMessage.errors.push(`Image for item (index ${info.originalRequestIndex}, title "${processedRequestsPayload[info.originalRequestIndex].createItem.item.title}") removed due to insecure/unavailable URL.`);

          // Correctly delete the image property (e.g., 'image') from its parent object (e.g., 'questionItem')
          if (parentOfImage && keyForImage && parentOfImage[keyForImage]) {
            // Example: if info.itemPath was ['questionItem', 'image'], parentOfImage is questionItem, keyForImage is 'image'
            // So we delete parentOfImage['image']
            let pathParent = processedRequestsPayload[info.originalRequestIndex].createItem.item;
            for (let k = 0; k < info.itemPath.length - 1; k++) {
              pathParent = pathParent[info.itemPath[k]];
            }
            delete pathParent[info.itemPath[info.itemPath.length - 1]];

          } else {
            console.warn(`[${functionName}] Could not properly locate parent to delete image property for request index ${info.originalRequestIndex}. Image might still be in payload.`);
          }
        }
      }

      if (info.tempItemId) {
        try {
          const itemToDelete = form.getItemById(info.tempItemId);
          if (itemToDelete) {
            form.deleteItem(itemToDelete);
            console.log(`[${functionName}] Deleted temporary image item ID ${info.tempItemId} (temp title: ${info.tempItemTitle})`);
          } else {
            console.warn(`[${functionName}] Could not find temporary item with ID ${info.tempItemId} (temp title: ${info.tempItemTitle}) for deletion.`);
          }
        } catch (e) {
          console.error(`[${functionName}] Error deleting temporary image item ID ${info.tempItemId}: ${e.message}`);
        }
      }
    });

    // --- Pass 4: Batch Create Final Items via Forms REST API using the processedRequestsPayload ---
    if (processedRequestsPayload.length > 0) {
        console.log(`[${functionName}] Pass 4: Sending batchUpdate to Forms API with ${processedRequestsPayload.length} requests...`);
        try {
            const token = ScriptApp.getOAuthToken();
            const batchUpdateUrl = `https://forms.googleapis.com/v1/forms/${formId}:batchUpdate`;
            const batchUpdatePayload = {
                requests: processedRequestsPayload,
                includeFormInResponse: false
            };
            const batchUpdateOptions = {
                method: "post",
                contentType: "application/json",
                headers: { "Authorization": "Bearer " + token },
                payload: JSON.stringify(batchUpdatePayload),
              muteHttpExceptions: true // Important for fetchWithExponentialBackoff to handle non-2xx codes
            };

          // MODIFIED: Use fetchWithExponentialBackoff for the batch update
          const maxRetriesForBatchUpdate = 3;
          const initialDelayForBatchUpdate = 2000; // 2 seconds
          console.log(`[${functionName}] Attempting batch update with exponential backoff (max ${maxRetriesForBatchUpdate} retries, initial delay ${initialDelayForBatchUpdate}ms).`);

          const batchApiResponse = fetchWithExponentialBackoff(
            batchUpdateUrl,
            batchUpdateOptions,
            maxRetriesForBatchUpdate,
            initialDelayForBatchUpdate
          );

            const batchResponseCode = batchApiResponse.getResponseCode();
            const batchResponseText = batchApiResponse.getContentText();
          // Attempt to parse JSON, but handle cases where it might not be JSON (e.g., network error before JSON response)
          let batchResponseJson = {};
          try {
            batchResponseJson = JSON.parse(batchResponseText);
          } catch (parseError) {
            console.warn(`[${functionName}] BatchUpdate response was not valid JSON: ${batchResponseText.substring(0, 200)}`);
            // If parsing fails but we have a non-2xx code, the error handling below will still use batchResponseText.
            // If it was a 2xx code but not JSON, that's an unexpected API behavior.
          }


          console.log(`[${functionName}] BatchUpdate API Response Code after retries: ${batchResponseCode}`);

            if (batchResponseCode >= 200 && batchResponseCode < 300) {
                responseMessage.success = true;
              // Safely access replies and filter
              const replies = batchResponseJson.replies || [];
              createdItemCount = replies.filter(reply => reply.createItem && reply.createItem.itemId).length;

              // Fallback for counting if replies structure is not as expected or empty, but overall success
              if (replies.length === 0 && processedRequestsPayload.length > 0) {
                // This is a less precise fallback, assuming all sent requests were successful if API returned 2xx
                // It might be better to rely solely on `replies` if present.
                console.warn(`[${functionName}] Batch update returned 2xx but no 'replies' array. Estimating created items based on payload length. This might be inaccurate.`);
                createdItemCount = processedRequestsPayload.filter(req => {
                  // A more robust filter would check if the original request was for an image that didn't fail processing
                  // For now, just count all successfully processed (non-failed image) items.
                  // This requires checking imageProcessInfo for the corresponding request.
                  const originalReqIndex = processedRequestsPayload.indexOf(req);
                  const imgInfo = imageProcessInfo.find(ipi => ipi.originalRequestIndex === originalReqIndex);
                  if (imgInfo) { // If there was image processing info for this request
                    return !imgInfo.processedUriFailed; // Count if image processing didn't fail
                  }
                       return true; // Count if it wasn't an image item or had no processing info
                     }).length;
              }

                responseMessage.createdItems = createdItemCount;
                responseMessage.message = `Batch update successful. Processed approximately ${createdItemCount} items.`;

              replies.forEach((reply, i) => {
                const reqTitle = processedRequestsPayload[i]?.createItem?.item?.title || `Request ${i + 1}`;
                if (!(reply.createItem && reply.createItem.itemId)) {
                    // Log if an individual item in a successful batch didn't confirm creation
                    console.warn(`[${functionName}] Individual item "${reqTitle}" in batch (index ${i}) may have had issues or no specific success confirmation in reply, despite overall 2xx.`);
                  }
                });
            } else {
              // Error handling for non-2xx responses after retries
                responseMessage.success = false;
              const errorObj = batchResponseJson.error || {};
              const errorDetails = errorObj.details || [];
              let detailedErrorMessages = errorDetails.map(detail =>
                detail.fieldViolations ? detail.fieldViolations.map(fv => `Field: ${fv.field}, Desc: ${fv.description}`).join('; ')
                  : (detail.message || JSON.stringify(detail))
              ).join('; ');

              const errorMessage = errorObj.message || batchResponseText; // Use text if no structured error message
              const errorCode = errorObj.code || batchResponseCode; // Use HTTP code if no API error code
              responseMessage.message = `Batch update failed after retries with status ${errorCode}: ${errorMessage}. ${detailedErrorMessages ? 'Details: ' + detailedErrorMessages : ''}`;
                responseMessage.errors.push(responseMessage.message);
              console.error(`[${functionName}] BatchUpdate API Error after retries:`, responseMessage.message);
            }

        } catch (e) {
          // This catch block handles errors from fetchWithExponentialBackoff itself (e.g., if all retries failed and it throws)
          // or errors during the processing of the response.
          console.error(`[${functionName}] Pass 4 Error (Batch Update execution or response processing): ${e.message}. Stack: ${e.stack}`);
            responseMessage.success = false;
          responseMessage.message = `Error during final batch update execution: ${e.message}`;
          if (responseMessage.errors.indexOf(e.message) === -1) { // Avoid duplicate error messages
            responseMessage.errors.push(e.message);
          }
        }
    } else {
        console.log(`[${functionName}] No items in final requestsPayload to send to batchUpdate.`);
      responseMessage.success = true; // Technically successful as there was nothing to do
        responseMessage.message = "No items to create in the form.";
    }

  } catch (err) {
    // General catch for the entire function
    console.error(`[${functionName}] General script error: ${err.message}. Stack: ${err.stack}`);
    responseMessage.success = false;
    responseMessage.message = `General script error: ${err.message}`;
    if (responseMessage.errors.indexOf(err.message) === -1) { // Avoid duplicate error messages
        responseMessage.errors.push(err.message);
    }
  }

  console.log(`[${functionName}] Execution finished. Final Response:`, JSON.stringify(responseMessage));
  return responseMessage;
}

/**
 * Fetches a URL with exponential backoff.
 * Only retries on 5xx errors or network errors (caught by try-catch).
 * For 4xx errors, it fails immediately after the first attempt unless it's a 429.
 * @param {string} url The URL to fetch.
 * @param {Object} params The parameters for UrlFetchApp.fetch(). Should include muteHttpExceptions: true.
 * @param {number} maxRetries The maximum number of retries.
 * @param {number} initialDelayMs The initial delay in milliseconds.
 * @return {GoogleAppsScript.URL_Fetch.HTTPResponse} The HTTPResponse object.
 * @throws {Error} If all retries fail or a non-retryable error occurs.
 */
function fetchWithExponentialBackoff(url, params, maxRetries, initialDelayMs) {
  let delay = initialDelayMs;
  let lastResponse; // Store the last response for throwing error

  // Ensure muteHttpExceptions is true, as this function handles response codes.
  params.muteHttpExceptions = true;

  for (let i = 0; i <= maxRetries; i++) { // Changed to i <= maxRetries for i=0 to be first attempt
    try {
      console.log(`[fetchWithExponentialBackoff] Attempt ${i + 1} (retry ${i}) to fetch ${url}`);
      lastResponse = UrlFetchApp.fetch(url, params);
      const responseCode = lastResponse.getResponseCode();

      if (responseCode >= 200 && responseCode < 300) {
        console.log(`[fetchWithExponentialBackoff] Attempt ${i + 1} successful for ${url} with status ${responseCode}.`);
        return lastResponse; // Success
      }

      // Retry on 5xx errors or 429 (Too Many Requests)
      if (responseCode >= 500 || responseCode === 429) {
        console.warn(`[fetchWithExponentialBackoff] Attempt ${i + 1} failed for ${url} with retryable status ${responseCode}: ${lastResponse.getContentText().substring(0, 500)}`);
        if (i === maxRetries) { // If this was the last retry
          throw new Error(`Failed to fetch ${url} after ${maxRetries + 1} attempts. Last status: ${responseCode}. Response: ${lastResponse.getContentText().substring(0, 500)}`);
        }
        // Wait for the delay before retrying
        Utilities.sleep(delay + Math.floor(Math.random() * 1000)); // Add jitter
        delay *= 2; // Exponential backoff
      } else {
        // For other client errors (4xx other than 429) or unexpected codes, fail immediately.
        console.error(`[fetchWithExponentialBackoff] Attempt ${i + 1} failed for ${url} with non-retryable status ${responseCode}: ${lastResponse.getContentText().substring(0, 500)}`);
        throw new Error(`Failed to fetch ${url}. Status: ${responseCode}. Response: ${lastResponse.getContentText().substring(0, 500)}`);
      }
    } catch (e) { // Catches network errors or errors thrown from above
      console.warn(`[fetchWithExponentialBackoff] Attempt ${i + 1} for ${url} caught exception: ${e.message}`);
      if (i === maxRetries) { // If this was the last retry
        // If lastResponse is available, include its details in the error.
        const responseDetails = lastResponse ? `Last status: ${lastResponse.getResponseCode()}. Response: ${lastResponse.getContentText().substring(0, 500)}` : "No response received.";
        throw new Error(`Failed to fetch ${url} after ${maxRetries + 1} attempts due to exception: ${e.message}. ${responseDetails}`);
      }
      // Wait for the delay before retrying network errors
      Utilities.sleep(delay + Math.floor(Math.random() * 1000)); // Add jitter
      delay *= 2; // Exponential backoff
    }
  }
  // This line should not be reached if logic is correct.
  // If it is, it means all retries were exhausted without returning or throwing a more specific error.
  const finalErrorMsg = lastResponse ? `Last status: ${lastResponse.getResponseCode()}. Response: ${lastResponse.getContentText().substring(0, 500)}` : "No response received during retries.";
  throw new Error(`Exhausted retries (${maxRetries + 1} attempts) for ${url} without success. ${finalErrorMsg}`);
}


/**
 * Helper function to check if a string is likely a filename based on common image extensions.
 * @param {string} text The string to check.
 * @return {boolean} True if it's likely a filename, false otherwise.
 */
function isLikelyFilename(text) {
  if (!text || typeof text !== 'string') return false;
  return /\.(jpeg|jpg|gif|png|svg|bmp|webp|tif|tiff)$/i.test(text.trim());
}

/**
 * Checks if an image URL is considered secure and directly usable by Forms API.
 * It prioritizes 'googleusercontent.com' URLs which are typical for Forms' internal content URIs.
 * It also allows common image extensions from any domain if they are https.
 * @param {string} url The URL to check.
 * @return {boolean} True if secure and likely usable, false otherwise.
 */
function isSecureImageUrl(url) {
    if (!url || typeof url !== 'string') return false;
    const lcUrl = url.toLowerCase();

  // Check 1: Must be HTTPS
  if (!lcUrl.startsWith('https://')) {
    console.warn(`[isSecureImageUrl] URL "${url}" is not HTTPS.`);
    return false;
  }

  // Check 2: Google User Content / GGPHT links are generally preferred and processed by Forms
  // These often look like: https://lh3.googleusercontent.com/... or https://*.ggpht.com/...
  // The 'key=' parameter is not strictly necessary for all googleusercontent links to be valid,
  // especially for contentUris obtained from the Forms API itself.
  if (lcUrl.includes('googleusercontent.com/') || lcUrl.includes('.ggpht.com/')) {
    // Further check if it looks like a direct image link (common extensions)
    if (/\.(jpeg|jpg|gif|png|bmp|webp)$/i.test(lcUrl.split('?')[0])) { // Check path before query params
      return true;
    }
    // If it's a googleusercontent link but doesn't have a common extension, it might still be
    // a valid contentUri (e.g., from Forms API response), so we can be a bit more lenient here.
    // However, for external URLs, we are stricter below.
    // Let's assume if it's from googleusercontent it's okay for now, as Forms API should handle its own URLs.
    // This was the previous behavior: if (lcUrl.includes('googleusercontent.com/') && lcUrl.includes('key=')) return true;
    // For now, being more inclusive of googleusercontent URLs.
    console.log(`[isSecureImageUrl] URL "${url}" is from Google content domain, considered usable.`);
    return true;
  }

  // Check 3: For other domains, check for common image extensions.
  if (/\.(jpeg|jpg|gif|png|bmp|webp)$/i.test(lcUrl.split('?')[0])) { // Check path before query params
    console.log(`[isSecureImageUrl] URL "${url}" is HTTPS and has a common image extension.`);
        return true;
    }

  // Check 4: Data URIs for images are also acceptable
  if (lcUrl.startsWith('data:image/') && lcUrl.includes('base64,')) {
    if (/\/(jpeg|jpg|gif|png|bmp|webp);base64/i.test(lcUrl.substring(0, lcUrl.indexOf('base64,') + 7))) {
      console.log(`[isSecureImageUrl] URL "${url}" is a valid data URI.`);
        return true;
      }
    }

  console.warn(`[isSecureImageUrl] URL "${url}" did not pass all security/usability checks. It might not be directly embeddable or accessible by Forms API in the required way.`);
    return false;
}
