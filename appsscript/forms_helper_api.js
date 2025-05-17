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
        const itemData = request.createItem.item; // Using itemData from the loop's current request
        let imageDetailPath = null;
        let imageDetails = null;

        if (itemData.questionItem && itemData.questionItem.image && itemData.questionItem.image.sourceUri) {
          imageDetailPath = ['questionItem', 'image'];
          imageDetails = itemData.questionItem.image;
        } else if (itemData.imageItem && itemData.imageItem.image && itemData.imageItem.image.sourceUri) {
          // This path is for standalone images, which don't need the 2-step URL processing in this script's current logic
          // but if they did, the path would be identified here.
          imageDetailPath = ['imageItem', 'image'];
          imageDetails = itemData.imageItem.image;
        }

        // Only do the temporary image item dance if it's an image intended for a questionItem
        if (itemData.questionItem && itemData.questionItem.question && imageDetails && imageDetailPath && imageDetailPath[0] === 'questionItem') {
          const tempItemTitle = `TEMP_IMG_FOR_REQ_INDEX_${index}_${Utilities.getUuid().substring(0, 8)}`;
          const currentImageInfo = {
            originalRequestIndex: index,
            itemPath: imageDetailPath, // Should be ['questionItem', 'image']
            tempItemTitle: tempItemTitle,
            tempItemId: null,
            originalSourceUri: imageDetails.sourceUri,
            originalAltText: imageDetails.altText,
            processedUriFailed: false, // Flag to track if processed URI step failed
            finalSourceUriForQuestion: imageDetails.sourceUri // Default to original, update if processed one is better
          };
          imageProcessInfo.push(currentImageInfo);

          try {
            console.log(`[${functionName}] Fetching blob with backoff for temp image: ${tempItemTitle} from ${imageDetails.sourceUri}`);
            // *** INTEGRATED EXPONENTIAL BACKOFF FOR THIS FETCH ***
            const imageBlob = fetchWithExponentialBackoff(imageDetails.sourceUri, {}, 3, 1000).getBlob();

            console.log(`[${functionName}] Adding temp image: ${tempItemTitle}`);
            const tempImageItem = form.addImageItem().setImage(imageBlob).setTitle(tempItemTitle);
            currentImageInfo.tempItemId = tempImageItem.getId();
          } catch (e) {
            console.error(`[${functionName}] Error creating temp image (index ${index}, title ${tempItemTitle}): ${e.message}`);
            responseMessage.errors.push(`Temp image creation failed (index ${index}, source ${imageDetails.sourceUri}): ${e.message}`);
            currentImageInfo.processedUriFailed = true; // Mark that this step failed
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
        const formApiResponse = UrlFetchApp.fetch(formsApiUrl, params);
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
                info.finalSourceUriForQuestion = processedContentUri; // Store the processed URL
              } else {
                console.warn(`[${functionName}] URL "${processedContentUri}" for temp item "${info.tempItemTitle}" is NOT secure. Original URI will be used if secure, otherwise image removed for this question.`);
                info.processedUriFailed = true; // Mark this processed URI as unusable
                if (!isSecureImageUrl(info.originalSourceUri)){ // Check original also
                   console.warn(`[${functionName}] Original URL "${info.originalSourceUri}" for question ${info.originalRequestIndex} is also NOT secure. Image will be removed.`);
                   // The actual removal from payload happens in Pass 3 based on finalSourceUriForQuestion
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
        imagesThatHadTempItemCreated.forEach(info => info.processedUriFailed = true); // Mark all as failed if API call itself failed
      }
    }

    // --- Pass 3: Update original payload with processed URIs (or remove image if insecure) and delete temporary items ---
    console.log(`[${functionName}] Pass 3: Updating payload and deleting ${imageProcessInfo.filter(info => info.tempItemId).length} temporary image items...`);
    imageProcessInfo.forEach(info => {
      let targetImageObject = processedRequestsPayload[info.originalRequestIndex].createItem.item;
      info.itemPath.forEach(p => { if(targetImageObject) targetImageObject = targetImageObject[p]; });

      if (targetImageObject) { // targetImageObject is like { sourceUri: '...', altText: '...' }
        if (info.finalSourceUriForQuestion && isSecureImageUrl(info.finalSourceUriForQuestion)) {
          targetImageObject.sourceUri = info.finalSourceUriForQuestion;
        } else if (isSecureImageUrl(info.originalSourceUri)) {
          // If processed URI failed or was insecure, but original was secure, keep original
          targetImageObject.sourceUri = info.originalSourceUri;
           console.warn(`[${functionName}] Using original secure URI for question ${info.originalRequestIndex} as processed URI was not secure or failed.`);
        } else {
          // Both processed and original are insecure, or processing failed and original is insecure
          console.warn(`[${functionName}] Both original and processed URIs for question ${info.originalRequestIndex} are insecure or unavailable. Removing image from request.`);
          responseMessage.errors.push(`Image for item (index ${info.originalRequestIndex}, title "${processedRequestsPayload[info.originalRequestIndex].createItem.item.title}") removed due to insecure/unavailable URL.`);
          // Navigate to parent of image object to delete the 'image' property
          let parentOfImage = processedRequestsPayload[info.originalRequestIndex].createItem.item;
          for(let i = 0; i < info.itemPath.length - 1; i++) { // e.g. info.itemPath = ['questionItem', 'image']
              if(parentOfImage) parentOfImage = parentOfImage[info.itemPath[i]];
          }
          if (parentOfImage && parentOfImage[info.itemPath[info.itemPath.length - 1]]) { // if parentOfImage is questionItem, this is 'image'
              delete parentOfImage[info.itemPath[info.itemPath.length - 1]];
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
                muteHttpExceptions: true
            };

            const batchApiResponse = UrlFetchApp.fetch(batchUpdateUrl, batchUpdateOptions); // No backoff needed for this single API call
            const batchResponseCode = batchApiResponse.getResponseCode();
            const batchResponseText = batchApiResponse.getContentText();
            const batchResponseJson = JSON.parse(batchResponseText);

            console.log(`[${functionName}] BatchUpdate API Response Code: ${batchResponseCode}`);

            if (batchResponseCode >= 200 && batchResponseCode < 300) {
                responseMessage.success = true;
                if (batchResponseJson.replies) {
                    createdItemCount = batchResponseJson.replies.filter(reply => reply.createItem && reply.createItem.itemId).length;
                } else {
                    createdItemCount = processedRequestsPayload.filter(req => !(req.createItem && req.createItem.item && req.createItem.item.questionItem && req.createItem.item.questionItem.image && req.createItem.item.questionItem.image.sourceUri && info.processingFailed)).length;
                }
                responseMessage.createdItems = createdItemCount;
                responseMessage.message = `Batch update successful. Processed approximately ${createdItemCount} items.`;

                if (batchResponseJson.replies) {
                    batchResponseJson.replies.forEach((reply, i) => {
                        const reqTitle = processedRequestsPayload[i]?.createItem?.item?.title || `Request ${i + 1}`;
                        if (!(reply.createItem && reply.createItem.itemId)) {
                            console.warn(`[${functionName}] Individual item "${reqTitle}" in batch may have had issues or no specific success confirmation in reply.`);
                        }
                    });
                }
            } else {
                responseMessage.success = false;
                const errorDetails = batchResponseJson.error ? batchResponseJson.error.details || [] : [];
                let detailedErrorMessages = errorDetails.map(detail => detail.fieldViolations ? detail.fieldViolations.map(fv => `Field: ${fv.field}, Desc: ${fv.description}`).join('; ') : detail.message).join('; ');

                const errorMessage = batchResponseJson.error ? batchResponseJson.error.message : batchResponseText;
                const errorCode = batchResponseJson.error ? batchResponseJson.error.code : batchResponseCode;
                responseMessage.message = `Batch update failed with status ${errorCode}: ${errorMessage}. ${detailedErrorMessages}`;
                responseMessage.errors.push(responseMessage.message);
                console.error(`[${functionName}] BatchUpdate API Error:`, responseMessage.message);
            }

        } catch (e) {
            console.error(`[${functionName}] Pass 4 Error (Batch Update): ${e.message}. Stack: ${e.stack}`);
            responseMessage.success = false;
            responseMessage.message = `Error during final batch update: ${e.message}`;
            responseMessage.errors.push(responseMessage.message);
        }
    } else {
        console.log(`[${functionName}] No items in final requestsPayload to send to batchUpdate.`);
        responseMessage.success = true;
        responseMessage.message = "No items to create in the form.";
    }

  } catch (err) {
    console.error(`[${functionName}] General script error: ${err.message}. Stack: ${err.stack}`);
    responseMessage.success = false;
    responseMessage.message = `General script error: ${err.message}`;
    if (responseMessage.errors.indexOf(err.message) === -1) {
        responseMessage.errors.push(err.message);
    }
  }

  console.log(`[${functionName}] Execution finished. Final Response:`, JSON.stringify(responseMessage));
  return responseMessage;
}

/**
 * Fetches a URL with exponential backoff.
 * @param {string} url The URL to fetch.
 * @param {Object} params The parameters for UrlFetchApp.fetch().
 * @param {number} maxRetries The maximum number of retries.
 * @param {number} initialDelayMs The initial delay in milliseconds.
 * @return {GoogleAppsScript.URL_Fetch.HTTPResponse} The HTTPResponse object.
 * @throws {Error} If all retries fail.
 */
function fetchWithExponentialBackoff(url, params, maxRetries, initialDelayMs) {
  let delay = initialDelayMs;
  for (let i = 0; i < maxRetries; i++) {
    try {
      console.log(`[fetchWithExponentialBackoff] Attempt ${i + 1} to fetch ${url}`);
      const response = UrlFetchApp.fetch(url, params);
      if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
        return response;
      } else {
         console.warn(`[fetchWithExponentialBackoff] Attempt ${i + 1} failed for ${url} with status ${response.getResponseCode()}: ${response.getContentText().substring(0,500)}`);
         if (i === maxRetries - 1) {
            throw new Error(`Failed to fetch ${url} after ${maxRetries} attempts. Last status: ${response.getResponseCode()}`);
         }
      }
    } catch (e) {
      console.warn(`[fetchWithExponentialBackoff] Attempt ${i + 1} for ${url} caught error: ${e.message}`);
      if (i === maxRetries - 1) {
        throw e;
      }
    }
    Utilities.sleep(delay + Math.floor(Math.random() * 1000));
    delay *= 2;
  }
  // This line should ideally not be reached if errors are thrown correctly in the loop.
  throw new Error(`Exhausted retries for ${url} without success or definitive error.`);
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
 * Checks if an image URL is considered secure.
 * @param {string} url The URL to check.
 * @return {boolean} True if secure, false otherwise.
 */
function isSecureImageUrl(url) {
    if (!url || typeof url !== 'string') return false;
    const lcUrl = url.toLowerCase();

    if (/\.(jpeg|jpg|gif|png|bmp|webp)$/i.test(lcUrl)) {
        return true;
    }
    // Loosened the check for Google hosted content to be more inclusive of valid contentUris
    // A key may or may not be present, but if it's a googleusercontent or ggpht link, it's often a direct image.
    if (lcUrl.includes('googleusercontent.com/') && lcUrl.includes('key=')) {
        return true;
    }

    console.warn(`[isSecureImageUrl] URL "${url}" did not pass security checks. It may not be directly embeddable or publicly accessible in the required way.`);
    return false;
}
