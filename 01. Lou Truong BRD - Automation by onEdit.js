// Global variable declaration and store information to reuse
// const spreadSheetID = "1YRbL3n2iNv4hYz202pllVOzfO2Pg8V94AM5GoEX1HKY"; // Controller
// const sheetName_1 = "Aff Lzd link build"; // Controller
// const sheetId_1 = 985243160; // Controller
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
      affiliateLazadaLinkBuildTool(e); // Handles edits in Column D
      updateLazadaLinkOnSubIdEdit(e); // Handles edits in Columns J-P
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
// Feature affilaite Shopee link build
/**
 * Configuration section
 */
const CONFIG_SHOPEE = {
  SHEET_ID: 1894656836, // The GID of the Shopee sheet
  AFFILIATE_ID: "17371270103", // Your Shopee Affiliate ID

  // Column Indices (1-based index: A=1, B=2, C=3, etc.)
  COLUMNS: {
    link_original: 4, // D: Raw Shopee URL input (Trigger column for full processing)
    link_full: 5, // E: Final generated affiliate link (Updated by both functions)
    seller_id: 7, // G: Extracted seller ID
    item_id: 8, // H: Extracted item ID
    link_clean: 9, // I: Generated clean product link (shopee.vn/product/...)
    link_encode: 10, // J: URL-Encoded version of the clean link (Needed for sub-id updates)
    sub_id1: 11, // K: Tracking Sub ID 1 (Start of sub-id range)
    sub_id2: 12, // L: Tracking Sub ID 2
    sub_id3: 13, // M: Tracking Sub ID 3
    sub_id4: 14, // N: Tracking Sub ID 4
    sub_id5: 15, // O: Tracking Sub ID 5 (End of sub-id range)
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
    sheet.getSheetId() !== CONFIG_SHOPEE.SHEET_ID ||
    col !== CONFIG_SHOPEE.COLUMNS.link_original
  ) {
    return;
  }

  Logger.log(
    `SHOPEE_TOOL: affiliateShopeeLinkBuildTool - Processing edit in Sheet GID ${CONFIG_SHOPEE.SHEET_ID}, Col ${col}, Row ${row}.`
  );

  try {
    const link_original_value = range.getValue().toString().trim();

    // --- Handle Empty Input: Clear dependent cells ---
    if (!link_original_value) {
      sheet.getRange(row, CONFIG_SHOPEE.COLUMNS.link_full).clearContent(); // E
      sheet.getRange(row, CONFIG_SHOPEE.COLUMNS.seller_id).clearContent(); // G
      sheet.getRange(row, CONFIG_SHOPEE.COLUMNS.item_id).clearContent(); // H
      sheet.getRange(row, CONFIG_SHOPEE.COLUMNS.link_clean).clearContent(); // I
      sheet.getRange(row, CONFIG_SHOPEE.COLUMNS.link_encode).clearContent(); // J
      Logger.log(
        `SHOPEE_TOOL: Row ${row}: Input 'link_original' is empty. Cleared dependent cells (E, G, H, I, J).`
      );
      return;
    }

    // --- Initialize variables ---
    let seller_id_value = "";
    let item_id_value = "";
    let link_clean_value = "";
    let link_encode_value = "";
    let link_full_value = "";

    // --- Extract seller_id and item_id (Handles multiple formats) ---
    let urlPath = link_original_value;
    const queryIndex = urlPath.indexOf("?");
    if (queryIndex !== -1) {
      urlPath = urlPath.substring(0, queryIndex);
    }
    const hashIndex = urlPath.indexOf("#");
    if (hashIndex !== -1) {
      urlPath = urlPath.substring(0, hashIndex);
    }

    // Method 1: Check for '/product/seller_id/item_id' format
    const productPathMatch = urlPath.match(/\/product\/(\d+)\/(\d+)/);
    if (productPathMatch && productPathMatch.length === 3) {
      seller_id_value = productPathMatch[1];
      item_id_value = productPathMatch[2];
      // Logger.log(`SHOPEE_TOOL: Row ${row}: Extracted IDs using /product/ format: seller=${seller_id_value}, item=${item_id_value}`);
    } else {
      // Method 2: Check for 'name.seller_id.item_id' format at the end of the path
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
          // Logger.log(`SHOPEE_TOOL: Row ${row}: Extracted IDs using .seller.item format: seller=${seller_id_value}, item=${item_id_value}`);
        } else {
          Logger.log(
            `SHOPEE_TOOL: Row ${row}: Found '.' pattern, but failed numeric validation for IDs in '${productPart}'.`
          );
        }
      } else {
        Logger.log(
          `SHOPEE_TOOL: Row ${row}: URL did not match /product/seller/item or name.seller.item formats.`
        );
      }
    }
    // --- End Extraction ---

    // --- Populate Columns G & H ---
    sheet
      .getRange(row, CONFIG_SHOPEE.COLUMNS.seller_id)
      .setValue(seller_id_value);
    sheet.getRange(row, CONFIG_SHOPEE.COLUMNS.item_id).setValue(item_id_value);

    // --- Process further only if IDs were extracted ---
    if (seller_id_value && item_id_value) {
      // --- Create Clean Link (Column I) ---
      link_clean_value = `https://shopee.vn/product/${seller_id_value}/${item_id_value}`;
      sheet
        .getRange(row, CONFIG_SHOPEE.COLUMNS.link_clean)
        .setValue(link_clean_value);

      // --- Encode Clean Link (Column J) ---
      link_encode_value = encodeURIComponent(link_clean_value);
      sheet
        .getRange(row, CONFIG_SHOPEE.COLUMNS.link_encode)
        .setValue(link_encode_value);

      // --- Construct Initial Final Affiliate Link (Column E) with CORRECT default empty sub_id ---
      const sub_id_string_default = "----"; // Default placeholder (4 hyphens)
      link_full_value = `https://s.shopee.vn/an_redir?sub_id=${sub_id_string_default}&origin_link=${link_encode_value}&affiliate_id=${CONFIG_SHOPEE.AFFILIATE_ID}`;
      sheet
        .getRange(row, CONFIG_SHOPEE.COLUMNS.link_full)
        .setValue(link_full_value);

      Logger.log(
        `SHOPEE_TOOL: Row ${row}: Successfully processed Col D edit. Initial Final URL generated with sub_id=${sub_id_string_default}`
      );
    } else {
      // --- Handle Failure: Clear dependent cells if IDs were not found ---
      sheet.getRange(row, CONFIG_SHOPEE.COLUMNS.link_clean).clearContent(); // I
      sheet.getRange(row, CONFIG_SHOPEE.COLUMNS.link_encode).clearContent(); // J
      sheet.getRange(row, CONFIG_SHOPEE.COLUMNS.link_full).clearContent(); // E
      Logger.log(
        `SHOPEE_TOOL: Row ${row}: Failed ID extraction during Col D edit. Cleared dependent cells (E, I, J).`
      );
    }
  } catch (error) {
    Logger.log(
      `ERROR in SHOPEE affiliateShopeeLinkBuildTool for row ${row}: ${
        error.message
      }\nInput: ${e.range.getValue()}\nStack: ${error.stack}`
    );
    try {
      sheet
        .getRange(row, CONFIG_SHOPEE.COLUMNS.link_full)
        .setValue(`ERROR: ${error.message}`);
    } catch (e2) {
      Logger.log(`SHOPEE_TOOL: Could not write error to sheet: ${e2}`);
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

  // Define sub-id column range for Shopee
  const firstSubIdCol = CONFIG_SHOPEE.COLUMNS.sub_id1; // K = 11
  const lastSubIdCol = CONFIG_SHOPEE.COLUMNS.sub_id5; // O = 15
  const numSubIdCols = lastSubIdCol - firstSubIdCol + 1; // Should be 5

  // --- Pre-checks: Exit if edit is not relevant ---
  if (
    sheet.getSheetId() !== CONFIG_SHOPEE.SHEET_ID ||
    col < firstSubIdCol ||
    col > lastSubIdCol
  ) {
    return; // Edit was not in columns K-O of the target Shopee sheet
  }

  Logger.log(
    `SHOPEE_TOOL: updateShopeeLinkOnSubIdEdit - Processing edit in Sub ID Column ${col}, Row ${row}.`
  );

  try {
    // --- Get the required base encoded link from Column J ---
    const link_encode_value = sheet
      .getRange(row, CONFIG_SHOPEE.COLUMNS.link_encode)
      .getValue()
      .toString()
      .trim();

    if (!link_encode_value) {
      Logger.log(
        `SHOPEE_TOOL: Row ${row}: Skipping sub-id update because 'link_encode' (Col J) is empty.`
      );
      return; // Cannot update sub-ids if the base encoded link isn't there
    }

    // --- Get current Sub IDs (Columns K-O) ---
    const subIdRange = sheet.getRange(row, firstSubIdCol, 1, numSubIdCols);
    let fetched_sub_id_values = Array(numSubIdCols).fill(""); // Default to expected number of blanks
    try {
      fetched_sub_id_values = subIdRange.getValues()[0];
    } catch (err) {
      Logger.log(
        `SHOPEE_TOOL: Row ${row}: updateShopeeLinkOnSubIdEdit - Error getting sub ID values: ${err}. Using default blanks.`
      );
    }
    // Optional log: Logger.log(`SHOPEE_TOOL: Row ${row}: updateShopeeLinkOnSubIdEdit - Fetched ${fetched_sub_id_values.length} values from ${subIdRange.getA1Notation()}: [${fetched_sub_id_values.map(v => `"${v}"`).join(', ')}]`);

    // --- FIX: Force array to have exactly numSubIdCols elements ---
    const sub_id_values = fetched_sub_id_values.slice(0, numSubIdCols);
    // Optional log: Logger.log(`SHOPEE_TOOL: Row ${row}: updateShopeeLinkOnSubIdEdit - Sliced to ${sub_id_values.length} values: [${sub_id_values.map(v => `"${v}"`).join(', ')}]`);

    // --- Format the Sub ID string ---
    // Map empty/falsy to '' (empty string), not '-'.
    const sub_ids_array = sub_id_values.map((id) =>
      id ? id.toString().trim() : ""
    );
    const sub_id_string = sub_ids_array.join("-"); // Join produces '----' if all were empty.
    // Optional log: Logger.log(`SHOPEE_TOOL: Row ${row}: updateShopeeLinkOnSubIdEdit - Final sub_id_string: "${sub_id_string}"`);

    // --- Re-Construct Final Affiliate Link (Column E) ---
    const link_full_value = `https://s.shopee.vn/an_redir?sub_id=${sub_id_string}&origin_link=${link_encode_value}&affiliate_id=${CONFIG_SHOPEE.AFFILIATE_ID}`;
    sheet
      .getRange(row, CONFIG_SHOPEE.COLUMNS.link_full)
      .setValue(link_full_value);

    Logger.log(
      `SHOPEE_TOOL: Row ${row}: Successfully updated Final URL (Col E) due to sub-id edit. New sub_id: ${sub_id_string}`
    );
  } catch (error) {
    Logger.log(
      `ERROR in SHOPEE updateShopeeLinkOnSubIdEdit for row ${row}: ${error.message}\nEdited Sub ID Col: ${col}\nStack: ${error.stack}`
    );
    try {
      sheet
        .getRange(row, CONFIG_SHOPEE.COLUMNS.link_full)
        .setValue(`ERROR updating sub-ids: ${error.message}`);
    } catch (e2) {
      Logger.log(`SHOPEE_TOOL: Could not write error to sheet: ${e2}`);
    }
  }
} // --- End of updateShopeeLinkOnSubIdEdit function ---

// Feature affilaite Lazada link build
/**
 * Configuration section
 */
const CONFIG_LAZADA = {
  SHEET_ID: 985243160, // The GID of the Lazada sheet

  // Column Indices (1-based index: A=1, B=2, C=3, etc.)
  COLUMNS: {
    // Input Column:
    link_original: 4, // D: Raw Lazada URL input (Trigger for ID extraction)

    // Link Full Column (Manual Base Input, Auto Updated with Sub-IDs):
    link_full: 5, // E: Base affiliate link (manual) + Appended sub-ids (auto)

    // Output Columns (from Col D Trigger):
    product_id: 7, // G: Extracted product ID
    sku_id: 8, // H: Extracted SKU ID
    link_clean: 9, // I: Generated clean product link

    // Sub-ID Input Columns (Manual Entry - Trigger columns for Col E update):
    sub_aff_id: 10, // J -> sub_aff_id
    sub_id1: 11, // K -> sub_id1
    sub_id2: 12, // L -> sub_id2
    sub_id3: 13, // M -> sub_id3
    sub_id4: 14, // N -> sub_id4
    sub_id5: 15, // O -> sub_id5
    sub_id6: 16, // P -> sub_id6 (Assuming P is sub_id6)
  },
  // Parameter names mapping for update function
  SUB_ID_PARAMS: {
    10: "sub_aff_id", // Col J
    11: "sub_id1", // Col K
    12: "sub_id2", // Col L
    13: "sub_id3", // Col M
    14: "sub_id4", // Col N
    15: "sub_id5", // Col O
    16: "sub_id6", // Col P
  },
  // Parameter ORDER for the final URL (as specified in prompt)
  PARAM_ORDER: [
    "sub_id1",
    "sub_aff_id",
    "sub_id6",
    "sub_id3",
    "sub_id2",
    "sub_id5",
    "sub_id4",
  ],
}; // --- END OF CONFIGURATION ---
// ========================================================================
// FEATURE 1: Handles ID Extraction triggered by edits in Column D
// ========================================================================
function affiliateLazadaLinkBuildTool(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  // --- Pre-checks: Exit if edit is not relevant ---
  if (
    sheet.getSheetId() !== CONFIG_LAZADA.SHEET_ID ||
    col !== CONFIG_LAZADA.COLUMNS.link_original
  ) {
    return; // Edit was not in Column D of the target Lazada sheet
  }

  Logger.log(
    `affiliateLazadaLinkBuildTool: Processing edit in Sheet GID ${CONFIG_LAZADA.SHEET_ID}, Column ${col}, Row ${row}.`
  );

  try {
    const link_original_value = range.getValue().toString().trim();

    // --- Initialize variables ---
    let product_id_value = "";
    let sku_id_value = "";
    let link_clean_value = "";

    // --- Handle Empty Input: Clear dependent cells ---
    if (!link_original_value) {
      sheet.getRange(row, CONFIG_LAZADA.COLUMNS.product_id).clearContent(); // G
      sheet.getRange(row, CONFIG_LAZADA.COLUMNS.sku_id).clearContent(); // H
      sheet.getRange(row, CONFIG_LAZADA.COLUMNS.link_clean).clearContent(); // I
      // Do NOT clear Column E as it's manual input
      Logger.log(
        `Row ${row}: Input 'link_original' (Lazada) is empty. Cleared G, H, I.`
      );
      return;
    }

    // --- Extract product_id and sku_id using Regex ---
    // Format: ...-i{product_id}-s{sku_id}.html...
    const idMatch = link_original_value.match(/-i(\d+)-s(\d+)\.html/);
    if (idMatch && idMatch.length === 3) {
      product_id_value = idMatch[1];
      sku_id_value = idMatch[2];
      Logger.log(
        `Row ${row}: Extracted product_id=${product_id_value}, sku_id=${sku_id_value}`
      );

      // --- Construct Clean Link (Column I) ---
      link_clean_value = `https://www.lazada.vn/products/-i${product_id_value}-s${sku_id_value}.html`;
    } else {
      Logger.log(
        `Row ${row}: Could not extract product_id and sku_id from URL: ${link_original_value}`
      );
      // Clear values if extraction fails
      product_id_value = "";
      sku_id_value = "";
      link_clean_value = "";
    }

    // --- Populate Output Columns G, H, I ---
    sheet
      .getRange(row, CONFIG_LAZADA.COLUMNS.product_id)
      .setValue(product_id_value); // G
    sheet.getRange(row, CONFIG_LAZADA.COLUMNS.sku_id).setValue(sku_id_value); // H
    sheet
      .getRange(row, CONFIG_LAZADA.COLUMNS.link_clean)
      .setValue(link_clean_value); // I

    Logger.log(`Row ${row}: Finished processing Lazada Col D edit.`);
  } catch (error) {
    Logger.log(
      `ERROR in affiliateLazadaLinkBuildTool for row ${row}: ${
        error.message
      }\nInput: ${e.range.getValue()}\nStack: ${error.stack}`
    );
    // Optionally clear cells or set error status
    try {
      sheet
        .getRange(row, CONFIG_LAZADA.COLUMNS.product_id, 1, 3)
        .clearContent(); // Clear G, H, I on error
      // Maybe set status in a dedicated column if available
    } catch (e2) {
      Logger.log(`Could not clear cells on error: ${e2}`);
    }
  }
} // --- End of affiliateLazadaLinkBuildTool function ---

// ========================================================================
// FEATURE 2: Handles Sub-ID Updates triggered by edits in Columns J-P
// ========================================================================
function updateLazadaLinkOnSubIdEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const row = range.getRow();
  const col = range.getColumn();

  // --- Pre-checks: Exit if edit is not relevant ---
  // 1. Check sheet
  if (sheet.getSheetId() !== CONFIG_LAZADA.SHEET_ID) {
    return;
  }
  // 2. Check if the column is within the sub_id range (J to P)
  const firstSubIdCol = CONFIG_LAZADA.COLUMNS.sub_aff_id; // J = 10
  const lastSubIdCol = CONFIG_LAZADA.COLUMNS.sub_id6; // P = 16
  if (col < firstSubIdCol || col > lastSubIdCol) {
    return; // Edit was not in columns J-P
  }

  Logger.log(
    `updateLazadaLinkOnSubIdEdit: Processing edit in Sub ID Column ${col}, Row ${row}.`
  );

  try {
    // --- Get the manually entered base link from Column E ---
    const baseLinkRange = sheet.getRange(row, CONFIG_LAZADA.COLUMNS.link_full);
    let base_link_full = baseLinkRange.getValue().toString().trim();

    if (!base_link_full) {
      Logger.log(
        `Row ${row}: Skipping sub-id update because base link (Col E) is empty.`
      );
      return; // Cannot update if the base link isn't there
    }

    // --- Extract the base part of the URL (before the first '?') ---
    let baseUrlPart = base_link_full;
    const queryStartIndex = base_link_full.indexOf("?");
    if (queryStartIndex !== -1) {
      baseUrlPart = base_link_full.substring(0, queryStartIndex);
    }

    // --- Get current Sub ID values (Columns J-P) ---
    const subIdRange = sheet.getRange(
      row,
      firstSubIdCol,
      1,
      lastSubIdCol - firstSubIdCol + 1
    ); // J to P (7 columns)
    const subIdValues = subIdRange.getValues()[0]; // [valJ, valK, valL, valM, valN, valO, valP]

    // --- Build the new query string parameters map ---
    const params = {};
    for (let i = 0; i < subIdValues.length; i++) {
      const currentSubIdCol = firstSubIdCol + i; // Column number (10, 11, ...)
      const paramName = CONFIG_LAZADA.SUB_ID_PARAMS[currentSubIdCol]; // Get param name ('sub_aff_id', 'sub_id1', ...)
      const paramValue = subIdValues[i];

      // Only include if the parameter name exists and the value is not empty
      if (
        paramName &&
        paramValue !== null &&
        paramValue !== undefined &&
        paramValue !== ""
      ) {
        params[paramName] = encodeURIComponent(paramValue.toString().trim()); // Store URL-encoded value
      }
    }

    // --- Construct the final query string respecting PARAM_ORDER ---
    const queryStringParts = [];
    for (const key of CONFIG_LAZADA.PARAM_ORDER) {
      if (params[key]) {
        // Check if the parameter was added (i.e., had a value)
        queryStringParts.push(`${key}=${params[key]}`);
      }
    }

    let final_link_full = baseUrlPart; // Start with the base URL part
    if (queryStringParts.length > 0) {
      final_link_full += "?" + queryStringParts.join("&"); // Append new query string if any params were added
    }

    // --- Update Column E with the final reconstructed link ---
    // Avoid infinite loops by checking if the value actually changed
    if (final_link_full !== base_link_full) {
      baseLinkRange.setValue(final_link_full);
      Logger.log(
        `Row ${row}: Successfully updated Final URL (Col E) due to sub-id edit. New link: ${final_link_full}`
      );
    } else {
      Logger.log(
        `Row ${row}: Sub-id edit resulted in the same final URL. No update written to Col E.`
      );
    }
  } catch (error) {
    Logger.log(
      `ERROR in updateLazadaLinkOnSubIdEdit for row ${row}: ${error.message}\nEdited Sub ID Col: ${col}\nStack: ${error.stack}`
    );
    try {
      // Avoid writing error directly to E if possible, maybe use another column or just log
      // sheet.getRange(row, CONFIG_LAZADA.COLUMNS.link_full).setValue(`ERROR updating sub-ids: ${error.message}`);
    } catch (e2) {
      Logger.log(`Could not write error status to sheet: ${e2}`);
    }
  }
} // --- End of updateLazadaLinkOnSubIdEdit function ---
