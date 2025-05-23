// // Global variable declaration and store information to reuse
// const spreadSheetID = "1YRbL3n2iNv4hYz202pllVOzfO2Pg8V94AM5GoEX1HKY"; // Controller
// const sheetName_1 = "Aff Other link build"; // Controller
// const sheetId_1 = 1103978091; // Controller
// const ss_1 = SpreadsheetApp.openById(spreadSheetID).getSheetByName(sheetName_1); // Fixed

// // Feature to get sheet id
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
      shortLinkbuild(e);
    } else {
      console.error("onEdit triggered without event object 'e'.");
      Logger.log("onEdit triggered without event object 'e'.");
    }
  } catch (error) {
    console.error(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
    Logger.log(
      `Error in onEdit handler: ${error.message}\nStack: ${error.stack}`
    );
  }
}

// Feature add timestamp
function addTimeStamp(e) {
  const sheetIdList = [462100576, 1894656836, 985243160, 1103978091]; // Controller
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
  const sheetIdList = [462100576, 1894656836, 985243160, 1103978091]; // Controller
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

// Feature short link builder
/**
 * Configuration specific to the short link generation.
 */
const SHORT_LINK_CONFIG = {
  TARGET_SHEET_IDS: [
    462100576, // Person Link Build GID
    1894656836, // Aff Shp Link Build GID
    985243160, // Aff Lzd Link Build GID
    1103978091, // Aff Other Link Build GID
  ],
  LINK_ID_COLUMN: 1, // Column A - Source Data
  LINK_ORIGINAL_COLUMN: 4, // Column D - Also a trigger
  LINK_SHORT_COLUMN: 6, // Column F - Target Output
  BASE_URL: "https://s.loutruong.com/", // Base URL for the short link
};
function shortLinkbuild(e) {
  try {
    // --- Basic Event Object Check ---
    if (!e || !e.range) {
      return;
    }

    const range = e.range;
    const sheet = range.getSheet();
    const sheetId = sheet.getSheetId();
    const editedCol = range.getColumn();
    const editedRow = range.getRow();

    // --- Configuration ---
    const cfg = SHORT_LINK_CONFIG;

    // --- Check if the edit is relevant to this function ---
    // 1. Is it one of the target sheets?
    if (!cfg.TARGET_SHEET_IDS.includes(sheetId)) {
      return; // Not a sheet we care about
    }
    // 2. Was Column A OR Column D edited?
    if (
      editedCol !== cfg.LINK_ID_COLUMN &&
      editedCol !== cfg.LINK_ORIGINAL_COLUMN
    ) {
      return; // Edit wasn't in Col A or Col D
    }
    // 3. Ignore header row
    if (editedRow <= 1) {
      return;
    }
    // 4. Process only single cell edits for simplicity
    // (If Col D edits can be multi-cell and you need to handle that, this check needs adjustment)
    if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
      Logger.log(
        `generateShortLink: Row ${editedRow}: Ignoring multi-cell edit.`
      );
      return;
    }

    // --- Perform Action ---
    Logger.log(
      `generateShortLink: Processing edit in Col ${editedCol}, Row ${editedRow}, Sheet GID ${sheetId}.`
    );

    // ALWAYS read the Link ID from Column A of the *edited* row
    const linkIdCell = sheet.getRange(editedRow, cfg.LINK_ID_COLUMN); // Column A
    const linkIdValue = linkIdCell.getValue().toString().trim();

    // Get the target cell in Column F
    const shortLinkCell = sheet.getRange(editedRow, cfg.LINK_SHORT_COLUMN); // Column F

    if (linkIdValue) {
      // If Link ID has a value, create and set the short link
      const shortLink = cfg.BASE_URL + linkIdValue; // Assuming linkIdValue is URL-safe
      // Optional: Check if update is needed to prevent writing same value
      // if (shortLinkCell.getValue() !== shortLink) {
      shortLinkCell.setValue(shortLink);
      Logger.log(
        `generateShortLink: Row ${editedRow}: Set short link in Col F.`
      );
      // }
    } else {
      // If Link ID is empty, clear the short link cell
      // Optional: Check if already clear
      // if (shortLinkCell.getValue() !== "") {
      shortLinkCell.clearContent();
      Logger.log(
        `generateShortLink: Row ${editedRow}: Cleared short link in Col F because Link ID in Col A is empty.`
      );
      // }
    }
  } catch (error) {
    Logger.log(
      `ERROR in generateShortLink: ${error.message}\nRange: ${
        e && e.range ? e.range.getA1Notation() : "N/A"
      }\nStack: ${error.stack}`
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
  } else if (sheetId === 1103978091) {
    idPrefix = "linkother";
  }

  const sheetIdList = [462100576, 1894656836, 985243160, 1103978091]; // Controller (still used for the initial check)

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
  SHEET_ID: 985243160, // The GID (Sheet ID) of the Lazada sheet. Find in URL: gid=xxxxxxxx

  // Column Indices (1-based index: A=1, B=2, C=3, etc.)
  // Based on the structure provided in the prompt.
  COLUMNS: {
    link_id: 1, // A: (Skipped by functions)
    link_create_time: 2, // B: (Skipped by functions)
    link_modify_time: 3, // C: (Skipped by functions)
    link_original: 4, // D: Raw Lazada URL input (Trigger column for extraction)
    link_full: 5, // E: Manually entered base link / Final generated link with sub-ids (Updated by sub-id edit)
    link_short: 6, // F: (Skipped by functions)
    product_id: 7, // G: Extracted product ID
    sku_id: 8, // H: Extracted SKU ID
    link_clean: 9, // I: Generated clean product link (lazada.vn/products/-i...-s...)
    sub_aff_id: 10, // J: Tracking Sub Affiliate ID (Input for link_full update)
    sub_id1: 11, // K: Tracking Sub ID 1 (Input for link_full update)
    sub_id2: 12, // L: Tracking Sub ID 2 (Input for link_full update)
    sub_id3: 13, // M: Tracking Sub ID 3 (Input for link_full update)
    sub_id4: 14, // N: Tracking Sub ID 4 (Input for link_full update)
    sub_id5: 15, // O: Tracking Sub ID 5 (Input for link_full update)
    sub_id6: 16, // P: Tracking Sub ID 6 (Input for link_full update) - Corrected from prompt's P=sub_id5
  },
};
// ========================================================================
// FEATURE 1: Handles ID Extraction triggered by edits in Column D
// ========================================================================
function affiliateLazadaLinkBuildTool(e) {
  try {
    // --- Basic Checks ---
    if (!e || !e.range) {
      Logger.log(
        "affiliateLazadaLinkBuildTool: Event object or range missing."
      );
      return;
    }

    const sheet = e.source.getActiveSheet(); // Use e.source for trigger context
    const sheetId = sheet.getSheetId();
    const editedCol = e.range.getColumn();
    const editedRow = e.range.getRow();

    // --- Configuration Aliases ---
    const cfg = CONFIG_LAZADA; // Use the CONFIG_LAZADA defined elsewhere in your script
    const cols = cfg.COLUMNS;

    // --- Sheet and Column Check ---
    if (sheetId !== cfg.SHEET_ID || editedCol !== cols.link_original) {
      // Logger.log(`affiliateLazadaLinkBuildTool: Edit ignored. Sheet ID: ${sheetId} (Expected: ${cfg.SHEET_ID}), Col: ${editedCol} (Expected: ${cols.link_original})`);
      return;
    }

    // --- Single Cell Check ---
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
      Logger.log(
        "affiliateLazadaLinkBuildTool: Edit ignored. Only single cell edits are processed for ID extraction."
      );
      return;
    }

    // --- Get Value ---
    // Use e.value if available (often faster in onEdit), otherwise get from range
    const link_original_value =
      e.value !== undefined && e.value !== null
        ? e.value.toString().trim()
        : e.range.getValue().toString().trim();

    // --- Initialize variables ---
    let product_id_value = "";
    let sku_id_value = ""; // Will remain empty if not found
    let link_clean_value = "";

    // --- Handle Empty or Invalid Input ---
    if (!link_original_value || !link_original_value.includes("lazada.vn")) {
      // Simplified check
      Logger.log(
        `affiliateLazadaLinkBuildTool: Cell in D${editedRow} is empty or not a valid Lazada URL.`
      );
      // Clear dependent cells if the input is invalid or cleared
      sheet.getRange(editedRow, cols.product_id).clearContent(); // G
      sheet.getRange(editedRow, cols.sku_id).clearContent(); // H
      sheet.getRange(editedRow, cols.link_clean).clearContent(); // I
      return;
    }

    // --- Prepare URL for Parsing: Remove query string and fragment ---
    let urlToParse = link_original_value;
    const queryIndex = link_original_value.indexOf("?");
    if (queryIndex !== -1) {
      urlToParse = link_original_value.substring(0, queryIndex);
    }
    const hashIndex = urlToParse.indexOf("#");
    if (hashIndex !== -1) {
      urlToParse = urlToParse.substring(0, hashIndex);
    }
    Logger.log(
      `LAZADA_TOOL: Row ${editedRow}: URL Prepared for Parsing: "${urlToParse}"`
    );
    // --- End URL preparation ---

    // --- START NEW EXTRACTION LOGIC ---
    // 1. Extract Product ID (Essential)
    // Looks for 'i' + digits, followed by either '-s' or '.html'
    const productRegex = /i(\d+)(?:-s|\.html)/;
    const productMatch = urlToParse.match(productRegex);
    Logger.log(
      `LAZADA_TOOL: Row ${editedRow}: Product ID Match (i...): ${JSON.stringify(
        productMatch
      )}`
    );

    if (productMatch && productMatch[1]) {
      // Product ID Found
      product_id_value = productMatch[1];
      Logger.log(
        `LAZADA_TOOL: Row ${editedRow}: Extracted ProductID=${product_id_value}`
      );

      // 2. Extract SKU ID (Optional)
      // Looks for '-s' + digits + '.html' in the *same prepared URL*
      const skuRegex = /-s(\d+)\.html/;
      const skuMatch = urlToParse.match(skuRegex);
      Logger.log(
        `LAZADA_TOOL: Row ${editedRow}: SKU Match (-s...): ${JSON.stringify(
          skuMatch
        )}`
      );

      if (skuMatch && skuMatch[1]) {
        // SKU ID Found
        sku_id_value = skuMatch[1];
        Logger.log(
          `LAZADA_TOOL: Row ${editedRow}: Extracted SkuID=${sku_id_value}`
        );
        // Construct Clean Link WITH SKU
        link_clean_value = `https://www.lazada.vn/products/-i${product_id_value}-s${sku_id_value}.html`;
      } else {
        // SKU ID NOT Found
        Logger.log(
          `LAZADA_TOOL: Row ${editedRow}: SKU ID pattern (-s...) not found.`
        );
        // Construct Clean Link WITHOUT SKU
        link_clean_value = `https://www.lazada.vn/products/i${product_id_value}.html`;
      }
      Logger.log(
        `LAZADA_TOOL: Row ${editedRow}: Generated Clean Link: ${link_clean_value}`
      );

      // 3. Populate Columns G, H, I
      sheet.getRange(editedRow, cols.product_id).setValue(product_id_value); // G
      sheet.getRange(editedRow, cols.sku_id).setValue(sku_id_value); // H (Will be empty if SKU not found)
      sheet.getRange(editedRow, cols.link_clean).setValue(link_clean_value); // I
    } else {
      // Product ID itself wasn't found - Clear everything
      Logger.log(
        `LAZADA_TOOL: Row ${editedRow}: Essential Product ID pattern (i...) not found in "${urlToParse}"`
      );
      sheet.getRange(editedRow, cols.product_id).clearContent(); // G
      sheet.getRange(editedRow, cols.sku_id).clearContent(); // H
      sheet.getRange(editedRow, cols.link_clean).clearContent(); // I
    }
    // --- END NEW EXTRACTION LOGIC ---
  } catch (error) {
    Logger.log(
      `Error in affiliateLazadaLinkBuildTool for row ${editedRow}: ${error.message}\nStack: ${error.stack}`
    );
    try {
      // Attempt to clear output cells on error
      const sheetOnError = e.source.getActiveSheet(); // Re-get sheet just in case
      const rowOnError = e.range.getRow();
      sheetOnError
        .getRange(rowOnError, CONFIG_LAZADA.COLUMNS.product_id, 1, 3)
        .clearContent(); // Clear G, H, I
    } catch (e2) {
      Logger.log(`LAZADA_TOOL: Could not clear cells on error: ${e2}`);
    }
  }
}
// --- End of affiliateLazadaLinkBuildTool function ---

// ========================================================================
// FEATURE 2: Handles Sub-ID Updates triggered by edits in Columns J-P
// ========================================================================
function updateLazadaLinkOnSubIdEdit(e) {
  try {
    // --- Basic Checks ---
    if (!e || !e.range) {
      // Logger.log("updateLazadaLinkOnSubIdEdit: Event object or range missing.");
      return;
    }

    const sheet = e.source.getActiveSheet();
    const sheetId = sheet.getSheetId();
    const editedCol = e.range.getColumn();
    const editedRow = e.range.getRow();

    // --- Configuration Aliases ---
    const cfg = CONFIG_LAZADA; // Alias for shorter reference
    const cols = cfg.COLUMNS;

    // --- Sheet and Column Check ---
    // Check if the edit happened on the correct sheet and in the sub-id columns range (J to P)
    if (
      sheetId !== cfg.SHEET_ID ||
      editedCol < cols.sub_aff_id ||
      editedCol > cols.sub_id6
    ) {
      // Logger.log(`updateLazadaLinkOnSubIdEdit: Edit ignored. Sheet ID: ${sheetId} (Expected: ${cfg.SHEET_ID}), Col: ${editedCol} (Expected: ${cols.sub_aff_id}-${cols.sub_id6})`);
      return;
    }

    // --- Process potentially multiple rows if edited simultaneously ---
    const numRows = e.range.getNumRows();
    const startRow = editedRow;

    for (let i = 0; i < numRows; i++) {
      const currentRow = startRow + i;

      // --- Get Base URL (from 'link_full' column E) ---
      const linkFullCell = sheet.getRange(currentRow, cols.link_full);
      let baseUrl = linkFullCell.getValue(); // This is the manually entered value or previously generated one

      if (!baseUrl || typeof baseUrl !== "string") {
        // Logger.log(`updateLazadaLinkOnSubIdEdit: Skipping row ${currentRow}. Base URL in Column E ('link_full') is missing or invalid.`);
        continue; // Skip this row if base URL is missing
      }

      // Remove existing query string to rebuild it cleanly
      const questionMarkIndex = baseUrl.indexOf("?");
      if (questionMarkIndex !== -1) {
        baseUrl = baseUrl.substring(0, questionMarkIndex);
      }

      // --- Get Sub ID Values ---
      // Fetch all relevant sub-ID values for the current row (Columns J to P)
      const subValuesRange = sheet.getRange(
        currentRow,
        cols.sub_aff_id,
        1,
        cols.sub_id6 - cols.sub_aff_id + 1
      );
      const subValues = subValuesRange.getValues()[0];

      // Map values based on their position relative to the start column (sub_aff_id)
      const subAffIdValue = subValues[cols.sub_aff_id - cols.sub_aff_id]; // Index 0
      const subId1Value = subValues[cols.sub_id1 - cols.sub_aff_id]; // Index 1
      const subId2Value = subValues[cols.sub_id2 - cols.sub_aff_id]; // Index 2
      const subId3Value = subValues[cols.sub_id3 - cols.sub_aff_id]; // Index 3
      const subId4Value = subValues[cols.sub_id4 - cols.sub_aff_id]; // Index 4
      const subId5Value = subValues[cols.sub_id5 - cols.sub_aff_id]; // Index 5
      const subId6Value = subValues[cols.sub_id6 - cols.sub_aff_id]; // Index 6

      // --- Build Parameter List ---
      const params = [];
      // Use the exact parameter names specified in the initial prompt
      if (subAffIdValue)
        params.push(`sub_aff_id=${encodeURIComponent(subAffIdValue)}`);
      if (subId1Value)
        params.push(`sub_id1=${encodeURIComponent(subId1Value)}`);
      if (subId2Value)
        params.push(`sub_id2=${encodeURIComponent(subId2Value)}`);
      if (subId3Value)
        params.push(`sub_id3=${encodeURIComponent(subId3Value)}`);
      if (subId4Value)
        params.push(`sub_id4=${encodeURIComponent(subId4Value)}`);
      if (subId5Value)
        params.push(`sub_id5=${encodeURIComponent(subId5Value)}`);
      if (subId6Value)
        params.push(`sub_id6=${encodeURIComponent(subId6Value)}`);

      // --- Construct Final URL ---
      let finalFullLinkValue = baseUrl; // Variable name reflects column E
      if (params.length > 0) {
        // Reconstruct according to the format: base?param1&param2...
        // The order follows the prompt's example: K, J, P, M, L, O, N
        const orderedParams = [];
        if (subId1Value)
          orderedParams.push(`sub_id1=${encodeURIComponent(subId1Value)}`);
        if (subAffIdValue)
          orderedParams.push(`sub_aff_id=${encodeURIComponent(subAffIdValue)}`);
        if (subId6Value)
          orderedParams.push(`sub_id6=${encodeURIComponent(subId6Value)}`);
        if (subId3Value)
          orderedParams.push(`sub_id3=${encodeURIComponent(subId3Value)}`);
        if (subId2Value)
          orderedParams.push(`sub_id2=${encodeURIComponent(subId2Value)}`);
        if (subId5Value)
          orderedParams.push(`sub_id5=${encodeURIComponent(subId5Value)}`);
        if (subId4Value)
          orderedParams.push(`sub_id4=${encodeURIComponent(subId4Value)}`);

        finalFullLinkValue += "?" + orderedParams.join("&");
      }

      // --- Update 'link_full' Column E ---
      // Only update if the generated URL is different from the current one
      if (linkFullCell.getValue() !== finalFullLinkValue) {
        linkFullCell.setValue(finalFullLinkValue);
        // Logger.log(`updateLazadaLinkOnSubIdEdit: Updated Column E ('link_full') in row ${currentRow} to: ${finalFullLinkValue}`);
      } else {
        // Logger.log(`updateLazadaLinkOnSubIdEdit: No update needed for Column E ('link_full') in row ${currentRow}. URL is already correct.`);
      }
    }
  } catch (error) {
    Logger.log(
      `Error in updateLazadaLinkOnSubIdEdit: ${error.message}\nStack: ${error.stack}`
    );
    // SpreadsheetApp.getActiveSpreadsheet().toast(`Error updating Lazada link parameters: ${error.message}`);
  }
}
