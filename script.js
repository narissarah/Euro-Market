// Google Apps Script for exporting products with barcodes from Sheets to Shopify

// ======= CONFIGURATION =======
const CONFIG = {
  SHOPIFY_STORE: 'your-store.myshopify.com', // Replace with your store URL
  SHOPIFY_ACCESS_TOKEN: 'your-access-token', // Replace with your Shopify Admin API access token
  CHECK_FREQUENCY: 5, // Check for new rows every 5 minutes
  TIMESTAMP_COLUMN: 'K', // Column to mark sync status and timestamp (adjust as needed)
  BARCODE_COLUMN: 'C' // Column containing barcode data (adjust as needed)
};

// ======= CORE FUNCTIONS =======

/**
 * Initialize the script and create triggers
 */
function setup() {
  // Clear existing triggers
  clearTriggers();
  
  // Create a time-based trigger to run the checkForNewProducts function
  ScriptApp.newTrigger('checkForNewProducts')
    .timeBased()
    .everyMinutes(CONFIG.CHECK_FREQUENCY)
    .create();
  
  // Initialize tracking if needed
  initializeTracking();
  
  Logger.log('Setup complete. Script will check for new products every ' + 
             CONFIG.CHECK_FREQUENCY + ' minutes.');
}

/**
 * Initialize row tracking
 */
function initializeTracking() {
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('lastProcessedRow')) {
    // Get the current number of rows with data
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = getLastDataRow(sheet);
    properties.setProperty('lastProcessedRow', lastRow.toString());
    Logger.log('Initialized tracking at row ' + lastRow);
  }
}

/**
 * Check for new product rows and send them to Shopify
 */
function checkForNewProducts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const properties = PropertiesService.getScriptProperties();
  const lastProcessedRow = parseInt(properties.getProperty('lastProcessedRow') || '0');
  
  // Get all data including headers
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  
  // First row should be headers
  const headers = data[0];
  
  // Get the current last row with data
  const currentLastRow = getLastDataRow(sheet);
  
  // Process only new rows
  if (currentLastRow > lastProcessedRow) {
    Logger.log(`Found ${currentLastRow - lastProcessedRow} new rows to process`);
    
    // Process each new row
    for (let rowIndex = lastProcessedRow + 1; rowIndex <= currentLastRow; rowIndex++) {
      const rowData = data[rowIndex - 1]; // -1 because data array is 0-indexed
      
      // Skip completely empty rows
      if (rowData.join('').trim() === '') {
        continue;
      }
      
      try {
        // Map row data to product format
        const product = mapRowToProduct(headers, rowData);
        
        // Validate product data
        if (!isValidProduct(product)) {
          const errorMessage = 'Row ' + rowIndex + ': Invalid product data (missing required fields)';
          Logger.log(errorMessage);
          updateSyncStatus(sheet, rowIndex, 'ERROR: ' + errorMessage);
          continue;
        }
        
        // Send to Shopify
        const response = createShopifyProduct(product);
        
        // Update sync status
        if (response.success) {
          updateSyncStatus(sheet, rowIndex, 'Synced: ' + new Date().toLocaleString() + 
                           ' (ID: ' + response.id + ')');
          Logger.log('Successfully created product from row ' + rowIndex + 
                     ' (Shopify ID: ' + response.id + ')');
        } else {
          updateSyncStatus(sheet, rowIndex, 'ERROR: ' + response.error);
          Logger.log('Failed to create product from row ' + rowIndex + ': ' + response.error);
        }
      } catch (error) {
        Logger.log('Error processing row ' + rowIndex + ': ' + error);
        updateSyncStatus(sheet, rowIndex, 'ERROR: ' + error.toString());
      }
    }
    
    // Update the last processed row
    properties.setProperty('lastProcessedRow', currentLastRow.toString());
  } else {
    Logger.log('No new rows to process');
  }
}

/**
 * Maps a row of data to Shopify product format
 * @param {Array} headers - Array of column headers
 * @param {Array} rowData - Array of cell values for the row
 * @return {Object} Shopify product object
 */
function mapRowToProduct(headers, rowData) {
  const product = {
    variants: [{}] // Initialize with one variant
  };
  
  // Map each column based on its header
  headers.forEach((header, index) => {
    const value = rowData[index];
    
    // Skip empty values
    if (!value && value !== 0) return;
    
    // Convert header to lowercase for case-insensitive matching
    const headerLower = header.toString().toLowerCase().trim();
    
    switch (headerLower) {
      // Basic product fields
      case 'title':
      case 'product name':
      case 'name':
        product.title = value;
        break;
        
      case 'description':
      case 'body':
      case 'body_html':
        product.body_html = value;
        break;
        
      case 'vendor':
      case 'brand':
      case 'manufacturer':
        product.vendor = value;
        break;
        
      case 'product type':
      case 'type':
      case 'category':
        product.product_type = value;
        break;
        
      case 'tags':
      case 'keywords':
        product.tags = value;
        break;
        
      case 'published':
      case 'status':
        product.published = (value === true || 
                            value === 'true' || 
                            value === 'yes' || 
                            value === 'published');
        break;
        
      // Variant fields
      case 'sku':
      case 'product code':
        product.variants[0].sku = value.toString();
        break;
        
      case 'barcode':
      case 'upc':
      case 'ean':
      case 'isbn':
      case 'gtin':
        product.variants[0].barcode = value.toString();
        break;
        
      case 'price':
      case 'retail price':
        product.variants[0].price = value.toString();
        break;
        
      case 'compare at price':
      case 'compare price':
      case 'msrp':
        product.variants[0].compare_at_price = value.toString();
        break;
        
      case 'cost':
      case 'cost price':
        product.variants[0].cost = value.toString();
        break;
        
      case 'inventory':
      case 'quantity':
      case 'stock':
        product.variants[0].inventory_quantity = parseInt(value) || 0;
        break;
        
      case 'weight':
        product.variants[0].weight = parseFloat(value) || 0;
        break;
        
      case 'weight unit':
        product.variants[0].weight_unit = value;
        break;
        
      case 'option1':
      case 'size':
        product.variants[0].option1 = value.toString();
        if (!product.options) {
          product.options = [{ name: "Size", values: [value.toString()] }];
        }
        break;
        
      case 'option2':
      case 'color':
        product.variants[0].option2 = value.toString();
        if (!product.options) {
          product.options = [{ name: "Size", values: ["Default"] }];
        }
        if (product.options.length < 2) {
          product.options.push({ name: "Color", values: [value.toString()] });
        }
        break;
        
      case 'image':
      case 'image url':
      case 'product image':
        if (value && value.toString().trim()) {
          product.images = [{ src: value.toString().trim() }];
        }
        break;
    }
  });
  
  return product;
}

/**
 * Validates a product has the minimum required fields
 * @param {Object} product - Shopify product object
 * @return {Boolean} Whether the product is valid
 */
function isValidProduct(product) {
  // At minimum, a product needs a title
  return !!product.title;
}

/**
 * Creates a product in Shopify
 * @param {Object} product - Shopify product object
 * @return {Object} Response object with success status
 */
function createShopifyProduct(product) {
  try {
    // Create the API URL
    const apiUrl = `https://${CONFIG.SHOPIFY_STORE}/admin/api/2023-07/products.json`;
    
    // Set up the HTTP request
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'X-Shopify-Access-Token': CONFIG.SHOPIFY_ACCESS_TOKEN
      },
      payload: JSON.stringify({ product: product }),
      muteHttpExceptions: true
    };
    
    // Send the request to Shopify
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseBody = JSON.parse(response.getContentText());
    
    if (responseCode >= 200 && responseCode < 300 && responseBody.product) {
      return { 
        success: true, 
        id: responseBody.product.id,
        data: responseBody.product
      };
    } else {
      return {
        success: false,
        error: `API Error (${responseCode}): ${JSON.stringify(responseBody)}`
      };
    }
  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}

// ======= UTILITY FUNCTIONS =======

/**
 * Gets the last row with data
 * @param {Sheet} sheet - Google Sheet object
 * @return {Number} Last row number with data
 */
function getLastDataRow(sheet) {
  const lastRow = sheet.getLastRow();
  return lastRow;
}

/**
 * Updates the sync status in the specified column
 * @param {Sheet} sheet - Google Sheet object
 * @param {Number} rowNum - Row number to update
 * @param {String} status - Status message
 */
function updateSyncStatus(sheet, rowNum, status) {
  sheet.getRange(CONFIG.TIMESTAMP_COLUMN + rowNum).setValue(status);
}

/**
 * Clears all triggers for this script
 */
function clearTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
}

/**
 * Manually resets tracking to start from the current row count
 * Useful when you want to ignore existing rows
 */
function resetTracking() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = getLastDataRow(sheet);
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('lastProcessedRow', lastRow.toString());
  Logger.log('Tracking reset. Will only process rows after row ' + lastRow);
}

/**
 * Creates a menu in the Google Sheet for easy access to functions
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Shopify Sync')
    .addItem('Setup Automatic Sync', 'setup')
    .addItem('Check For New Products Now', 'checkForNewProducts')
    .addItem('Reset Tracking (Ignore Existing Rows)', 'resetTracking')
    .addToUi();
}

// Run this when the spreadsheet opens
function onOpen() {
  createShopifySyncMenu();
}

// Create the custom menu
function createShopifySyncMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Shopify Sync')
    .addItem('Set Up Automatic Sync', 'setup')
    .addItem('Sync New Products Now', 'checkForNewProducts')
    .addItem('Reset Tracking', 'resetTracking')
    .addToUi();
}
