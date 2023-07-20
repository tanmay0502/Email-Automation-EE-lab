
# Google Apps Script for Electrical Engineering Lab Component Management

![Google Apps Script](https://www.gstatic.com/images/branding/product/1x/apps_script_48dp.png)

This Google Apps Script automates the process of managing the issuance and return of components from the Electrical Engineering Lab. It includes functionalities to send reminder emails, handle partial and full returns, and notify users about component status. The script operates on a Google Sheets spreadsheet with three sheets: "Issued," "Returned," and "Template (hidden)." Additionally, it utilizes global variables to store template texts for emails.

## Table of Contents
- [Objective](#objective)
- [Spreadsheet Structure](#spreadsheet-structure)
- [Global Variables](#global-variables)
- [Functions](#functions)
- [Triggers](#triggers)
- [Functionality Flow](#functionality-flow)
- [Execution and Usage](#execution-and-usage)
- [Daily Gmail Limit](#daily-gmail-limit)
- [Date Calculation](#date-calculation)

## Objective

The objective of this Google Apps Script is to streamline the process of issuing and returning components in the Electrical Engineering Lab. It automates the tasks of sending reminder emails to users who have issued components and have not returned them. The script also handles email actions related to partial and full returns, notifying the recipients about the status of the components.

## Spreadsheet Structure

The script operates on a Google Sheets spreadsheet with the following sheets:
- "Issued": Contains the information about issued components, including user details, component names, and issuance status.
- "Returned": Stores the details of components that have been returned partially or fully.
- "Template (hidden)": A hidden sheet that stores template texts used in email notifications.

## Global Variables

The following global variables are defined at the beginning of the script:
- `ss`: Refers to the active spreadsheet.
- `ss2`: Refers to the "Returned" sheet.
- `templateText`: Stores the template text for sending reminder emails.
- `templateText2`: Stores the template text for notifying full returns.
- `templateText3`: Stores the template text for notifying partial returns.

## Functions

1. `sendEmails()`: Sends reminder emails to users who have issued components but have not returned them. It extracts relevant data from the "Issued" sheet, replaces placeholders in the template text, and sends personalized emails using `MailApp.sendEmail()`.

2. `sendCurrEmail(row, action)`: Sends different types of emails based on the action (status of the issued components). If the action is "Partially," it notifies about the partial return, and for other actions, it notifies about the full return.

3. `onOpen()`: Creates a custom menu, "Email Actions," in the spreadsheet UI, allowing users to trigger the `sendEmails()` function from the menu.

4. `newIssueEmail()`: Sends an email notification when new components are issued. It retrieves the latest row data from the "Issued" sheet and uses `MailApp.sendEmail()` to send the email.

5. `onTheEdit(e)`: Triggered on any edit in the spreadsheet. Handles edits in the "Issued" sheet when the "Status" or "Returned Date" is changed. Sends relevant emails and moves the row to the "Returned" sheet when required.

6. `copyRow(row)`: Copies the provided row data to the "Returned" sheet.

7. `sortResponses()`: Sorts the "Issued" and "Returned" sheets based on the first column (Latest First).

8. `runDaily()`: Sends daily reminder emails for components that have not been fully returned within a 30-day interval.

9. `dateAction(actionDate)`: Calculates the difference in days between the provided date and the current date. Returns `true` if the difference is a multiple of 30 and greater than or equal to 0.

## Triggers

The script uses the following triggers:
- `On Open` trigger: Runs the `onOpen()` function when the spreadsheet is opened, creating the custom menu "Email Actions."
- `On Edit` trigger: Runs the `onTheEdit(e)` function whenever any cell in the spreadsheet is edited.
- `Time-based` trigger: Runs the `runDaily()` function to send reminder emails.
- `On Form Submit` trigger: Runs the `newIssueEmail()` function whenever a new form is submitted.

## Functionality Flow

The main functionalities of the script include:
- Sending reminder emails for issued components not returned.
- Sending emails for partial and full returns.
- Notifying about new component issuances.
- Handling edits in the "Issued" sheet for status changes and returns.
- Sending daily reminder emails for overdue components.

## Execution and Usage

To use the script:
1. Open the Google Sheets spreadsheet containing the "Issued," "Returned," and "Template" sheets.
2. The script runs through custom menu options or triggers set in the script.
3. Custom menu "Email Actions" is created using the `onOpen()` trigger to execute the `sendEmails()` function.

## Daily Gmail Limit

The script considers the daily Gmail sending limit. The `MailApp.getRemainingDailyQuota()` function is used to get the remaining daily quota, which is displayed in a message box when an email is sent.

## Date Calculation

The `dateAction(actionDate)` function calculates the difference in days between the provided date (in the first column of the "Issued" sheet) and the current date. If the difference is a multiple of 30 and greater than or equal to 0, the function returns `true`, indicating that a daily reminder email should be sent.

For more information about Google Apps Script and its functionalities, you can refer to the [official documentation](https://developers.google.com/apps-script).

---
Today's date: 2023-07-20T16:49:24+05:30.
