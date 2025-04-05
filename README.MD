## Script Explanation

This script is designed to sync product data from a Google Sheet to your Shopify store. Here's how it works:

1. **Configuration**: The script starts with a CONFIG object that stores crucial settings:
   - Your Shopify store URL
   - Your Admin API access token
   - How frequently to check for new products (5 minutes)
   - Which column contains the sync status timestamps
   - Which column contains the barcode data

2. **Core Functions**:
   - `setup()`: Initializes the script, creates time-based triggers, and sets up tracking
   - `checkForNewProducts()`: The main function that checks for new rows in the spreadsheet and sends them to Shopify
   - `mapRowToProduct()`: Converts spreadsheet row data into a Shopify product object
   - `createShopifyProduct()`: Sends the product data to Shopify via their API

3. **Utility Functions**:
   - Functions for tracking which rows have been processed
   - Status updating
   - Trigger management
   - Menu creation in the Google Sheet

4. **Product Mapping**: The script intelligently maps spreadsheet columns to Shopify product fields based on column headers, handling common variations (e.g., "price" or "retail price").

## Step-by-Step Setup Guide

### 1. Create a Shopify Private App
- Log in to your Shopify admin dashboard
- Go to "Apps" → "Develop apps" → "Create an app"
- Name your app something like "Google/Shopify Integration"
- Under "API credentials," click "Configure Admin API scopes"
- Select these permissions:
  - write_products
  - read_products
- Click "Save" then "Install app"
- Copy the Admin API access token (important: use the access token, not the API key)

### 2. Set Up the Google Sheet
- Create a Google Sheet with columns for your product data
- Important columns to include:
  - Product title/name (required)
  - Barcode (in column C by default, or update the script)
  - Other product details (description, price, inventory, etc.)
- The script will use column K for sync status by default

### 3. Add the Script to Google Sheets
- Open your Google Sheet
- Click on "Extensions" → "Apps Script"
- Delete any code in the editor and paste the provided JavaScript code
- Update the CONFIG object with your specific details:
  ```javascript
  const CONFIG = {
    SHOPIFY_STORE: 'your-store.myshopify.com', // Your store URL
    SHOPIFY_ACCESS_TOKEN: 'your-access-token', // The token you copied
    CHECK_FREQUENCY: 5, // Check interval in minutes
    TIMESTAMP_COLUMN: 'K', // Column for sync status
    BARCODE_COLUMN: 'C' // Column with barcode data
  };
  ```

### 4. Check Compatibility
- Review your sheet's column headers and make sure they align with what the script expects
- The script is designed to be flexible with column naming (for example, it recognizes "price", "retail price", etc.)
- If needed, you can modify the `mapRowToProduct()` function to match your specific column structure

### 5. Save and Run the Setup
- Save the script (Ctrl+S or Cmd+S)
- Return to your Google Sheet
- Refresh the page to see the new "Shopify Sync" menu
- Click "Shopify Sync" → "Set Up Automatic Sync"
- The first time you run the script, you'll need to authorize it to access your Google Sheets

### 6. Using the Integration
- Add new product rows to your Google Sheet
- The script will automatically check for new rows every 5 minutes
- Alternatively, you can manually trigger a sync by clicking "Shopify Sync" → "Sync New Products Now"
- The script will update column K with the sync status (success or error message)

### 7. Troubleshooting
- Check the sync status column (K) for error messages
- If you want to start fresh, use "Shopify Sync" → "Reset Tracking"
- For debugging, you can view logs in the Apps Script editor under "Execution logs"

That's it! You've successfully set up an integration between Google Sheets and Shopify that automatically creates products with barcodes in your store.
