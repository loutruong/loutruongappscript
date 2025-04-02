// Global variable declaration and store information to reuse
// const spreadSheetID = "1YRbL3n2iNv4hYz202pllVOzfO2Pg8V94AM5GoEX1HKY"; // Controller
// const sheetName_1 = "Link builder"; // Controller
// const sheetId_1 = 462100576; // Controller
// const ss_1 = SpreadsheetApp.openById(spreadSheetID).getSheetByName(sheetName_1); // Fixed

// Feature to get sheet id
// function checkSheetId() {
//   const ss_1 =
//     SpreadsheetApp.openById(spreadSheetID).getSheetByName(sheetName_1); // Fixed
//   try {
//     if (ss_1) {
//       console.log(`The actual ID for ${sheetName_1} is: ${ss_1.getSheetId()}`);
//     } else {
//       console.log(`The actual ID for ${sheetName_1} is not found`);
//     }
//   } catch (error) {
//     console.error("Error message:", error.message);
//   }
// }

// Feature callback onEdit
function onEdit(e) {
  try {
    if (e) {
      linkbuildId(e);
      addTimeStamp(e);
      adjustTimeStamp(e);
      webLinkBuildTool(e);
      affiliateShopeeLinkBuildTool(e); // Handles edits in Column D
      updateShopeeLinkOnSubIdEdit(e); // Handles edits in Columns K-O
      affiliateLazadaLinkBuildTool(e);
    } else {
      console.error("onEdit triggered without event object 'e'.");
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}

// Feature add timestamp
function addTimeStamp(e) {
  const sheetIdList = [462100576, 1894656836, 985243160]; // Controller
  try {
    if (
      e &&
      e.range.getRow() > 1 &&
      e.range.getColumn() === 4 &&
      sheetIdList.includes(e.source.getSheetId()) &&
      e.source.getActiveSheet().getRange(e.range.getRow(), 2).getValue() === "" // Controller
    ) {
      e.source
        .getActiveSheet()
        .getRange(e.range.getRow(), 2) // Controller
        .setValue(
          Utilities.formatDate(
            new Date(),
            "Asia/Ho_Chi_Minh",
            "yyyy-MM-dd HH:mm:ss"
          )
        );
    } else {
      console.error("addTimeStamp error");
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}

// Feature adjust timestamp
function adjustTimeStamp(e) {
  const sheetIdList = [462100576, 1894656836, 985243160]; // Controller
  try {
    if (
      e &&
      e.range.getRow() > 1 &&
      e.range.getColumn() === 4 &&
      sheetIdList.includes(e.source.getSheetId())
    ) {
      e.source
        .getActiveSheet()
        .getRange(e.range.getRow(), 3) // Controller
        .setValue(
          Utilities.formatDate(
            new Date(),
            "Asia/Ho_Chi_Minh",
            "yyyy-MM-dd HH:mm:ss"
          )
        );
    } else {
      console.error("addTimeStamp error");
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}

// Feature add id
function linkbuildId(e) {
  const sheetId = e.source.getSheetId();
  let idPrefix = "link"; // Default prefix

  // Determine the prefix based on the sheet ID
  if (sheetId === 1894656836) {
    idPrefix = "linkshp";
  } else if (sheetId === 985243160) {
    idPrefix = "linklzd";
  }

  const sheetIdList = [462100576, 1894656836, 985243160]; // Controller (still used for the initial check)

  try {
    if (
      e &&
      e.range.getRow() > 1 &&
      e.range.getColumn() === 4 &&
      sheetIdList.includes(sheetId) &&
      e.source.getActiveSheet().getRange(e.range.getRow(), 1).getValue() === ""
    ) {
      let maxIdNumber = 0;
      const valueId = e.range
        .getSheet()
        .getRange(2, 1, e.range.getSheet().getMaxRows() - 1)
        .getValues();
      for (let i = 0; i < valueId.length; i++) {
        const value = valueId[i][0]; // Get the value from the cell [row][col=0]
        if (typeof value === "string" && value.startsWith(idPrefix)) {
          const curIdNumber = parseInt(value.substring(idPrefix.length), 10);
          if (!isNaN(curIdNumber)) {
            maxIdNumber = Math.max(maxIdNumber, curIdNumber); // Update the maximum number found
          }
        }
      }
      const nextIdNumber = maxIdNumber + 1;
      const newId = idPrefix + nextIdNumber;
      e.range.getSheet().getRange(e.range.getRow(), 1).setValue(newId);
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}

// Feature website link builder
function webLinkBuildTool(e) {
  // --- Start Configuration ---
  const targetSheetId = 462100576; // Controller
  const firstDataRow = 2; // Row number where your data starts (assuming Row 1 has headers)

  // Define column numbers based on structure:
  // A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8, I=9, J=10, K=11, L=12
  const urlCol = 4; // Column D: link_original (Base URL)
  const outputCol = 5; // Column E: link_full (Generated URL Output)
  const sourceCol = 7; // Column G: utm_source
  const mediumCol = 8; // Column H: utm_medium
  const campaignCol = 9; // Column I: utm_campaign
  const utmIdCol = 10; // Column J: utm_id
  const termCol = 11; // Column K: utm_term
  const contentCol = 12; // Column L: utm_content

  // Define which columns trigger the script when edited: We trigger on changes to the base URL or any UTM parameter
  const triggerColumns = [
    urlCol,
    sourceCol,
    mediumCol,
    campaignCol,
    utmIdCol,
    termCol,
    contentCol,
  ];
  // --- End Configuration ---
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const editedRow = range.getRow();
  const editedCol = range.getColumn();

  // Exit if the edit is not on the specified sheet, is in the header row, or is not in one of the trigger columns.
  if (
    sheet.getSheetId() != targetSheetId ||
    editedRow < firstDataRow ||
    !triggerColumns.includes(editedCol)
  ) {
    return;
  }

  // Get the base URL ('link_original') from the edited row
  let baseUrl = sheet.getRange(editedRow, urlCol).getValue().toString().trim();
  // Get the cell where the output ('link_full') will be written
  const outputCell = sheet.getRange(editedRow, outputCol);

  // If the base URL ('link_original') is empty, clear the output cell and stop
  if (!baseUrl) {
    outputCell.setValue("");
    return;
  }

  // Define the parameters and their corresponding columns
  const paramsConfig = [
    { name: "utm_source", col: sourceCol },
    { name: "utm_medium", col: mediumCol },
    { name: "utm_campaign", col: campaignCol },
    { name: "utm_id", col: utmIdCol }, // Note: 'utm_id' is the standard parameter name
    { name: "utm_term", col: termCol },
    { name: "utm_content", col: contentCol },
  ];

  let queryStringParts = []; // Array to hold "key=value" strings

  // Loop through each parameter configuration
  paramsConfig.forEach((paramInfo) => {
    const value = sheet
      .getRange(editedRow, paramInfo.col)
      .getValue()
      .toString()
      .trim(); // Only add the parameter if its value is not empty

    if (value) {
      queryStringParts.push(`${paramInfo.name}=${encodeURIComponent(value)}`); // Encode the value to make it URL-safe (handles spaces, special characters, etc.)
    }
  });

  // Assemble the final URL
  let finalUrl = baseUrl;
  if (queryStringParts.length > 0) {
    // Check if the base URL already contains a query string (a '?')
    if (baseUrl.includes("?")) {
      // If yes, append parameters with '&'
      finalUrl += "&" + queryStringParts.join("&");
    } else {
      // If no, start the query string with '?'
      finalUrl += "?" + queryStringParts.join("&");
    }
  }

  // Write the generated URL ('link_full') to the output column (Column C)
  outputCell.setValue(finalUrl);
}
// Feature affilaite Shopee link builder
/**
 * Configuration section
 */
const CONFIG = {
  SHEET_ID: 1894656836, // Controller
  AFFILIATE_ID: "17371270103",
  COLUMNS: {
    link_original: 4, // D: Trigger point
    link_full: 5, // E: Final affiliate link output
    seller_id: 7, // G: Extracted seller ID output
    item_id: 8, // H: Extracted item ID output
    link_clean: 9, // I: Constructed clean link output
    link_encode: 10, // J: Encoded clean link output
    sub_id1: 11, // K: Manual input sub_id1
    sub_id2: 12, // L: Manual input sub_id2
    sub_id3: 13, // M: Manual input sub_id3
    sub_id4: 14, // N: Manual input sub_id4
    sub_id5: 15, // O: Manual input sub_id5
  },
};
// ========================================================================
// FEATURE 1: Handles FULL processing triggered by edits in Column D
// ========================================================================
function affiliateShopeeLinkBuildTool(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  // --- Pre-checks: Exit if edit is not relevant ---
  if (
    sheet.getSheetId() !== CONFIG.SHEET_ID ||
    col !== CONFIG.COLUMNS.link_original
  ) {
    return;
  }

  Logger.log(
    `affiliateShopeeLinkBuildTool: Processing edit in Sheet GID ${CONFIG.SHEET_ID}, Column ${col}, Row ${row}.`
  );

  try {
    const link_original_value = range.getValue().toString().trim();

    // --- Handle Empty Input: Clear dependent cells ---
    if (!link_original_value) {
      sheet.getRange(row, CONFIG.COLUMNS.link_full).clearContent(); // E
      sheet.getRange(row, CONFIG.COLUMNS.seller_id).clearContent(); // G
      sheet.getRange(row, CONFIG.COLUMNS.item_id).clearContent(); // H
      sheet.getRange(row, CONFIG.COLUMNS.link_clean).clearContent(); // I
      sheet.getRange(row, CONFIG.COLUMNS.link_encode).clearContent(); // J
      Logger.log(
        `Row ${row}: Input 'link_original' is empty. Cleared dependent cells (E, G, H, I, J).`
      );
      return;
    }

    // --- Initialize variables ---
    let seller_id_value = "";
    let item_id_value = "";
    let link_clean_value = "";
    let link_encode_value = "";
    let link_full_value = "";

    // --- Extract seller_id and item_id ---
    let urlPath = link_original_value;
    const queryIndex = urlPath.indexOf("?");
    if (queryIndex !== -1) {
      urlPath = urlPath.substring(0, queryIndex);
    }
    const hashIndex = urlPath.indexOf("#");
    if (hashIndex !== -1) {
      urlPath = urlPath.substring(0, hashIndex);
    }
    const lastSlashIndex = urlPath.lastIndexOf("/");
    const productPart =
      lastSlashIndex !== -1 ? urlPath.substring(lastSlashIndex + 1) : urlPath;
    const parts = productPart.split(".");

    if (parts.length >= 2) {
      const potentialItemId = parts[parts.length - 1];
      const potentialSellerId = parts[parts.length - 2];
      if (
        !isNaN(potentialItemId) &&
        potentialItemId.trim() !== "" &&
        !isNaN(potentialSellerId) &&
        potentialSellerId.trim() !== ""
      ) {
        item_id_value = potentialItemId.trim();
        seller_id_value = potentialSellerId.trim();
      } else {
        Logger.log(
          `Row ${row}: Failed numeric validation for IDs in '${productPart}'.`
        );
      }
    } else {
      Logger.log(
        `Row ${row}: URL path part '${productPart}' has unexpected structure for IDs.`
      );
    }

    // --- Populate Columns G & H ---
    sheet.getRange(row, CONFIG.COLUMNS.seller_id).setValue(seller_id_value);
    sheet.getRange(row, CONFIG.COLUMNS.item_id).setValue(item_id_value);

    // --- Process further only if IDs were extracted ---
    if (seller_id_value && item_id_value) {
      // --- Create Clean Link (Column I) ---
      link_clean_value = `https://shopee.vn/product/${seller_id_value}/${item_id_value}`;
      sheet.getRange(row, CONFIG.COLUMNS.link_clean).setValue(link_clean_value);

      // --- Encode Clean Link (Column J) ---
      link_encode_value = encodeURIComponent(link_clean_value);
      sheet
        .getRange(row, CONFIG.COLUMNS.link_encode)
        .setValue(link_encode_value);

      // --- Construct Initial Final Affiliate Link (Column E) with CORRECT default empty sub_id ---
      const sub_id_string_default = "----"; // CORRECTED Default placeholder (4 hyphens)
      link_full_value = `https://s.shopee.vn/an_redir?sub_id=${sub_id_string_default}&origin_link=${link_encode_value}&affiliate_id=${CONFIG.AFFILIATE_ID}`;
      sheet.getRange(row, CONFIG.COLUMNS.link_full).setValue(link_full_value);

      Logger.log(
        `Row ${row}: Successfully processed Col D edit. Initial Final URL generated with sub_id=${sub_id_string_default}`
      );
    } else {
      // --- Handle Failure: Clear dependent cells if IDs were not found ---
      sheet.getRange(row, CONFIG.COLUMNS.link_clean).clearContent(); // I
      sheet.getRange(row, CONFIG.COLUMNS.link_encode).clearContent(); // J
      sheet.getRange(row, CONFIG.COLUMNS.link_full).clearContent(); // E
      Logger.log(
        `Row ${row}: Failed ID extraction during Col D edit. Cleared dependent cells (E, I, J).`
      );
    }
  } catch (error) {
    Logger.log(
      `ERROR in affiliateShopeeLinkBuildTool for row ${row}: ${
        error.message
      }\nInput: ${e.range.getValue()}\nStack: ${error.stack}`
    );
    try {
      sheet
        .getRange(row, CONFIG.COLUMNS.link_full)
        .setValue(`ERROR: ${error.message}`);
    } catch (e2) {
      Logger.log(`Could not write error to sheet: ${e2}`);
    }
  }
} // --- End of affiliateShopeeLinkBuildTool function ---
// ========================================================================
// FEATURE 2: Handles partial update triggered by edits in Columns K-O
// ========================================================================
function updateShopeeLinkOnSubIdEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  // --- Pre-checks: Exit if edit is not relevant ---
  if (
    sheet.getSheetId() !== CONFIG.SHEET_ID ||
    col < CONFIG.COLUMNS.sub_id1 ||
    col > CONFIG.COLUMNS.sub_id5
  ) {
    return;
  }

  Logger.log(
    `updateShopeeLinkOnSubIdEdit: Processing edit in Sub ID Column ${col}, Row ${row}.`
  );

  try {
    // --- Get the required base encoded link from Column J ---
    const link_encode_value = sheet
      .getRange(row, CONFIG.COLUMNS.link_encode)
      .getValue()
      .toString()
      .trim();

    if (!link_encode_value) {
      Logger.log(
        `Row ${row}: Skipping sub-id update because 'link_encode' (Col J) is empty.`
      );
      return; // Cannot update sub-ids if the base encoded link isn't there
    }

    // --- Get current Sub IDs (Columns K-O) ---
    const startColSubId = CONFIG.COLUMNS.sub_id1;
    const numSubIdCols = 5; // Explicitly 5 columns
    const subIdRange = sheet.getRange(row, startColSubId, 1, numSubIdCols);
    let fetched_sub_id_values = ["", "", "", "", ""]; // Default to 5 blanks
    try {
      fetched_sub_id_values = subIdRange.getValues()[0];
    } catch (err) {
      Logger.log(
        `Row ${row}: updateShopeeLinkOnSubIdEdit - Error getting sub ID values: ${err}. Using default blanks.`
      );
    }
    // Optional log: Logger.log(`Row ${row}: updateShopeeLinkOnSubIdEdit - Fetched ${fetched_sub_id_values.length} values: [${fetched_sub_id_values.map(v => `"${v}"`).join(', ')}]`);

    // --- FIX: Force array to have exactly 5 elements ---
    const sub_id_values = fetched_sub_id_values.slice(0, numSubIdCols);
    // Optional log: Logger.log(`Row ${row}: updateShopeeLinkOnSubIdEdit - Sliced to ${sub_id_values.length} values: [${sub_id_values.map(v => `"${v}"`).join(', ')}]`);

    // --- Format the Sub ID string ---
    // CORRECTED map condition: Map empty/falsy to '' (empty string), not '-'.
    const sub_ids_array = sub_id_values.map((id) =>
      id ? id.toString().trim() : ""
    );
    const sub_id_string = sub_ids_array.join("-"); // Join will now produce '----' if all were empty.
    // Optional log: Logger.log(`Row ${row}: updateShopeeLinkOnSubIdEdit - Final sub_id_string: "${sub_id_string}"`);

    // --- Re-Construct Final Affiliate Link (Column E) ---
    const link_full_value = `https://s.shopee.vn/an_redir?sub_id=${sub_id_string}&origin_link=${link_encode_value}&affiliate_id=${CONFIG.AFFILIATE_ID}`;
    sheet.getRange(row, CONFIG.COLUMNS.link_full).setValue(link_full_value);

    Logger.log(
      `Row ${row}: Successfully updated Final URL (Col E) due to sub-id edit. New sub_id: ${sub_id_string}`
    );
  } catch (error) {
    Logger.log(
      `ERROR in updateShopeeLinkOnSubIdEdit for row ${row}: ${error.message}\nEdited Sub ID Col: ${col}\nStack: ${error.stack}`
    );
    try {
      sheet
        .getRange(row, CONFIG.COLUMNS.link_full)
        .setValue(`ERROR updating sub-ids: ${error.message}`);
    } catch (e2) {
      Logger.log(`Could not write error to sheet: ${e2}`);
    }
  }
} // --- End of updateShopeeLinkOnSubIdEdit function ---

function affiliateLazadaLinkBuildTool(e) {}
