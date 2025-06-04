# Google Sheets Coupon Management System

This project is a Coupon Management System built using Google Apps Script and a Google Spreadsheet. It allows you to track coupons, their expiration dates, and manage gift certificate balances via a web app interface.

## Features

*   Add new coupons with details: barcode, expiration date.
*   Mark coupons as gift certificates and track their initial and current balances.
*   Log usage of gift certificates, automatically updating their balance.
*   View the 5 most recently added coupons.
*   View a list of all coupons, with visual indicators for:
    *   Expired coupons.
    *   Coupons nearing their expiration date (within 7 days).
*   Filter the list of all coupons by their entry date.
*   Data is stored in a Google Spreadsheet.

## Setup Instructions

1.  **Create a Google Spreadsheet:**
    *   Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet.
    *   You can name it anything you like (e.g., "My Coupons").

2.  **Get the Spreadsheet ID:**
    *   Open your newly created spreadsheet.
    *   The Spreadsheet ID is a long string of characters in the URL, between `/d/` and `/edit`.
    *   For example, if your URL is `https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_HERE/edit#gid=0`, then `SPREADSHEET_ID_HERE` is your Spreadsheet ID.
    *   Copy this ID.

3.  **Open Google Apps Script Editor:**
    *   In your spreadsheet, go to "Extensions" > "Apps Script". This will open the Apps Script editor in a new tab.

4.  **Copy Project Files:**
    *   You should have two files from this project:
        *   `Code.gs`
        *   `index.html`
    *   In the Apps Script editor:
        *   If there's a default `Code.gs` file, delete its content and paste the content of **your** `Code.gs` file.
        *   Click on the "+" icon next to "Files" in the left sidebar and choose "HTML". Name the new file `index` (it will automatically be saved as `index.html`). Delete any default content in this new file and paste the content of **your** `index.html` file.

5.  **Update Spreadsheet ID in `Code.gs`:**
    *   Go back to the `Code.gs` file in the Apps Script editor.
    *   Find the line: `var SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE";`
    *   Replace `"YOUR_SPREADSHEET_ID_HERE"` with the Spreadsheet ID you copied in Step 2.

6.  **Initial Save and Authorization (Important):**
    *   Save your project in Apps Script (File > Save, or Ctrl+S/Cmd+S).
    *   You need to run a function once to trigger the authorization process.
    *   In the Apps Script editor, select any function from the function dropdown menu (e.g., `getCoupons` or `testSaveCoupon` if available).
    *   Click the "Run" button (looks like a play icon).
    *   Google will ask you to "Review permissions". Click "Review permissions".
    *   Choose your Google account.
    *   You might see a warning "Google hasn't verified this app". Click "Advanced" and then "Go to [Your Project Name] (unsafe)".
    *   Review the permissions the script needs (it will ask for permission to manage your spreadsheets). Click "Allow".
    *   The function may fail if the sheets (`Coupons`, `UsageLog`) don't exist yet, but the authorization is the important part here. The web app will create them if they don't exist on first use.

7.  **Deploy as a Web App:**
    *   In the Apps Script editor, click on the "Deploy" button (usually in the top right).
    *   Select "New deployment".
    *   Click on the gear icon next to "Select type" and choose "Web app".
    *   In the "Description" field, you can add a short description (e.g., "Coupon Manager").
    *   For "Execute as", select **"Me (your.email@example.com)"**.
    *   For "Who has access", select:
        *   **"Only myself"** (if only you will use it).
        *   **"Anyone with Google account"** (if others in your domain/organization should use it).
        *   **"Anyone"** (if you want it to be publicly accessible - use with caution).
        *   *For most personal uses, "Only myself" is recommended.*
    *   Click "Deploy".
    *   After deployment, you will be given a **Web app URL**. Copy this URL. This is the link to your Coupon Management System.

8.  **Using the Web App:**
    *   Open the Web app URL you copied in the previous step in your browser.
    *   The "Coupons" and "UsageLog" sheets will be automatically created in your Google Spreadsheet the first time you save a coupon or log usage if they don't already exist.
    *   Start adding and managing your coupons!

## How it Works

*   **`Code.gs`**: This file contains all the server-side Google Apps Script functions. It handles:
    *   Serving the HTML page (`doGet`).
    *   Saving coupon data to the Google Sheet (`saveCoupon`).
    *   Retrieving coupon data (`getCoupons`, `getLatestCoupons`).
    *   Logging gift certificate usage and updating balances (`logGiftCertificateUsage`).
*   **`index.html`**: This file defines the structure, style, and client-side JavaScript for the web app interface. It communicates with `Code.gs` functions to send and receive data.
*   **Google Spreadsheet**: Acts as the database, storing all coupon and usage log information in two sheets: "Coupons" and "UsageLog".

## Notes

*   Ensure the `SPREADSHEET_ID` in `Code.gs` is correct.
*   When making changes to `Code.gs` or `index.html` after the initial deployment, you'll need to create a *new* deployment to see the changes live in the web app. Go to Deploy > Manage deployments, select your deployment, click the pencil icon (edit), and choose "New version" from the version dropdown, then click "Deploy".
*   The barcode recognition is manual (typing the barcode number). The file input for an image in the HTML is a placeholder for potential future enhancements and does not currently process the image to extract a barcode.
