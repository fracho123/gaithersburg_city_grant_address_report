# Gaithersburg City Grant and Montgomery County Totals Report
## About
This Excel macro workbook takes in visit data records exported as a CSV from food bank manager software such as Soxbox and produces the quarterly address listings report for Gaithersburg City Food and Rx as well as the monthly Montgomery County visit totals report.

To print this documentation, click [here](README.md) and print that page.
> [!NOTE]
> Documentation changes for v3.10 are highlighted in Notes like this one.

### Process Overview:
1) Import Food (and optionally Rx) visit data records
2) Attempt to validate addresses against the Gaithersburg Address database
3) Can additionally validate addresses against the Google Address Validation API
4) Can additionally accept user input to fix invalid addresses
5) Produce visit totals as well as the Gaithersburg City address listings report
> [!NOTE]
> 6) After adding addresses, can additionally import Rx medication records and calculate Rx totals

> [!NOTE]
> ### Overview of changes in version 3.10
> #### Sheet changes:
> * The "Interface" sheet has been renamed to "Home"
> * The "Final Report" sheet has been renamed to "NonRx Report" to differentiate it from the new Rx Final Report
> * The new Rx Final Report is in the "Rx Report" sheet
> * Adding Rx medication records and calculations for Rx medication totals are done on the new "Rx" sheet
> 
> #### Totals changes:
> * The new Food report totals are displayed first on the "Home" sheet. The old totals in earlier versions can be found by scrolling to the right on the "Home" sheet.
> * Totals now display included services/tracking methods. E.g. if you added records for tracking methods "Food", and "Food-Delivery", you will see that the report included the tracking method "Food" in the list of "Services included in total - Non-delivery: Food", and included the tracking method "Food-Delivery" in the other list of Delivery services.
> * The "Rx Asst" tracking method is now excluded from all totals except those on the "Rx" sheet.
>
> #### Process changes:
> * Confirm whether tracking methods have been sorted correctly into Non-delivery and Delivery.
> * When exporting the Non-Rx Final Report and Rx Final Report, the "Organization" column now asks for the organization name, not the initials.

# Using the XLSM file
## Downloading the XLSM
1) Ensure you have the latest version. Download the [latest release of the XLSM file](https://github.com/jimmyli97/gaithersburg_city_grant_address_report/releases). Click on the "Assets" title and then click on the XLSM to download. 
![Release assets download page](readme/1download.png)
    * If you have data in an older release version, in the new file on the "Home" sheet, click "Import Data" and select the older file. All data will be copied over to the new version.
2) The same file can be used from year to year and from quarter to quarter. The XLSM file will remember previously validated and user edited addresses. The date of the most recent imported service is stored in D1 of the "Home" sheet.
3) If this is the first quarter, name the file with the current fiscal year, "e.g. City Grant Address Listings Report v3.0 FY24.xlsm"
    * If you have a file from the last fiscal year, make a copy of it for this fiscal year and rename it. Then on the "Addresses" sheet, click "Delete All Visit Data" to delete all visit data but keep address data. This operation is quick, you'll know it's done when all of the dates are deleted below and to the right of the "Rx Totals" column, and all of the totals on the "Home" sheet are zeroed.
## Importing data
1) Log into your food bank manager and export data as a CSV. The visit data does not need to be quarterly, the XLSM file will automatically sort by quarter. The visit data can also be imported at any time, you don't have to do it all at once at the end of the quarter.
    1) For Gaithersburg HELP Soxbox, log in [here](https://ghp.soxbox.co/login). Go to Visit History Export:
       ![Soxbox Visit History Export](readme/2.1soxbox_visithistoryexport.png)
    2) Select the preset "city and county grant address v3", then select the dates you wish to export. For example, if you've already processed all the addresses up to but not including February 1, but you want to process addresses from February 1 to 29, select Visit On dates starting on February 1 and ending on February 29. Click "Export" and save the CSV file
       ![city and county grant address v3 preset and date selection](readme/2.2soxbox_preset.png)
2) Open the exported CSV file. **Filter by tracking method or service for the visit data that you want to report on.** Select all data **except for the header** (Click on B1, hold Shift and click on the last cell with data in Row B, then hold Shift and End and press Down Arrow once. Ctrl-C to copy.)
    * If you forget to import an extra tracking method, you can repeat this step but filter on only the extra tracking method.
    * If you accidentally imported an extra tracking method, you can click "Delete All Visit Data" in the "Addresses" sheet in the XLSM file to delete all tracking methods, and then reimport only the tracking methods you wish to report on.
    ![Filter exported CSV by tracking method](readme/2.3exportedcsvfilter.png)
3) Open the XLSM file. If you see the Protected View warning message, click the "Enable Editing" button. If you see that macros from the internet are disabled, close the workbook, right click on the workbook in File Explorer, go to Properties, at the bottom of the General tab check the Unlock checkbox, and open the workbook again (see [this link](https://learn.microsoft.com/en-us/deployoffice/security/internet-macros-blocked)).  If you see the Make This A Trusted Document warning message, click Yes.
4) On the "Home" sheet, click "Paste Records". This will paste records into the "Home" sheet, but will not make any changes to the database. If you need to make any edits, now is the time to do so.
5) On the "Home" sheet, click "Add Records". This will match all records against existing records in the file. All new addresses will be validated against the Gaithersburg database only. This takes about 4 minutes per 1000 records, depending on your computer and internet. Progress can be seen in the lower left corner of the screen, you cannot use Excel but you can use other programs on your computer.
    * Progress is tracked in the status bar in the bottom left. Program is finished when the status bar says "Ready" or when the mouse icon is no longer the spinning busy icon
    * If you need to stop execution, hit Esc. If you stop execution, you will need to start over from step 4.
6) Addresses matching existing addresses in the "Addresses", "Needs Autocorrect" or "Discards" sheets will be merged. Successfully validated addresses can be seen in the "Addresses" sheet and will be marked with an "In City" code of "Yes". All other addresses will be moved to the "Needs Autocorrect" sheet.
7) I recommend at least automatically validating addresses first, but you can generate a final report and county totals now (see [here](#generating-totals-and-non-rx-final-report)). Before editing addresses by hand, automatically validate addresses first.
## Automatically validating addresses
1) Google Address Validation requires a [Google Address Validation key](https://developers.google.com/maps/documentation/address-validation/get-api-key). Please be careful to avoid sharing this file with the key inside publically on the Internet (email to selected recipients is fine).
    1) Open apikeys.csv, select cell B1 (second column, first row cell), copy.
    2) Open this XLSM, paste the key into cell F1 of the "Home" sheet.
2) This XLSM file attempts to keep usage of the API within the free tier and limits you to 8,000 requests per month. To increase this limit, email me.
3) On the "Needs Autocorrect" sheet, click "Attempt validation" This will attempt to autocorrect and validate all addresses against Google Address Validation if the In City Code is "Not yet autocorrected". Addresses returned from this validation will be placed in the "Validated Address", "Validated Unit Type and No.", and "Validated Zip Code" fields. All addresses are then verified on the "Validated" address fields against the Gaithersburg database. 
    * The same restrictions apply as before while validating addresses (you cannot use excel, progress will be shown in the lower left corner, etc.)
4) After validation, addresses will show up in either the "Addresses" sheet, the "Needs Autocorrect" sheet, or the "Discards" sheet. Any addresses automatically autocorrected by the XLSM file will additionally show up in the "Autocorrected" sheet.
    * Addresses that were autocorrected and in Gaithersburg will be marked with an "In City" code of "Yes" and be placed in the "Addresses" sheet.
    * Addresses that were autocorrected and not in Gaithersburg will be marked with an "In City" code of "No" and be placed in the "Addresses" sheet.
    * Addresses that were marked as not correctable (one word addresses, etc.) will be marked with an "In City" code of "Not correctable" and be placed in the "Discards" sheet
    * Addresses that could not be autocorrected but were predicted to not be in Gaithersburg (not in zips 20877, 20878, 20879, not in Gaithersburg city bounds) will be marked with an "In City" code of "Failed autocorrection and geocoded not in city" and be placed in the "Discards" sheet
    * Addresses that could not be autocorrected but were predicted to be in Gaithersburg will be marked with an "In City" code of "Possible but failed autocorrection" and be placed in the "Needs Autocorrect" sheet
5) You can generate a Non-Rx Final Report and county totals now (see [here](#generating-totals-and-non-rx-final-report)), or continue with editing addresses by hand.
## Manually validating addresses
### Editing
1) Before manually editing addresses, automatically validate addresses first. All user editing should be done in the "Validated" fields in the "Needs Autocorrect" sheet, NOT in other sheets. This makes sure the program moves records correctly. The "Raw" fields are ignored after automatic validation and should not be edited. Go through each address and type in a valid address for the record in the corresponding "Validated" field (see [Tips for validating addresses](#Tips-for-validating-addresses)). As you do so, the "User Verified" field will be set to "True" for that record. If you accidentally edit a record, click "Toggle User Verified" to set the record back to "False".
2) Click "Attempt Validation". All "True" "User Verified" records will be validated again against the Gaithersburg database using the user input address. You don't have to correct all records before clicking "Attempt Validation".
3) Select any records which cannot be corrected and click "Discard selected" to move those records to the "Discards" sheet. Alternatively, you can click "Discard All" to discard all remaining records in the "Needs Autocorrect" sheet.
4) To fix accidentally discarded addresses, select them in the "Discards" sheet and click "Restore selected discards". They will be moved back to the "Needs Autocorrect" sheet.
5) To fix records incorrectly marked as "Yes" or "No, select them in the "Addresses" sheet and click "Move to needs autocorrect". They will be moved back to the "Needs Autocorrect" sheet.
6) If verifying Autocorrected addresses, you can toggle records as verified so you know which ones you have verified as correct
7) For additional corrections, see [Tips for editing](#Tips-for-editing)
8) See [here](#generating-totals-and-non-rx-final-report) to generate the Non-Rx Final Report and county totals
### Tips for validating addresses
1) Check if the address is similar to any Gaithersburg street names. Click "Open List of Gaithersburg Streets" in the "Needs Autocorrect" sheet to get the list. You can use Ctrl-F in the browser to look for similar street names. Discard if not in the list.
2) Check for typos by selecting record and execute the LookupInCity macro Ctrl+Shift+L to look up that record via the Gaithersburg City address search page, in a browser window. Delete some characters to find similar addresses (e.g. “3 Summit” instead of “3A S Summit St”)
	* This macro automatically uses the validated address if it exists, otherwise it will use the raw address.
    * A common error is the unit letter being in the unit number instead of in the address, e.g. 425 N Frederick Ave Unit 1C should be 425C N Frederick Ave Unit 1
    * Two streets with apostrophes exist, O’Neill and Odend’hal
    * You can click on the address in the City address search page to see a map where Gaithersburg borders are highlighted in red and house numbers are visible.
3) You can also check for typos in Google Maps, and check the apartment range by typing into the USPS lookup page with no apartment number.
    * In the USPS website, the DPV Confirmation will be Y for a valid address, and a different letter if invalid.
5) Look at the keyboard for possible typos, e.g. 419 was entered instead of 119
6) If unable to validate unit type and number but you can validate the address, the XLSM file will accept it if "User Verified" is set to "True" and the "Validated" unit type and number fields are blank before clicking "Attempt Validation"
7) The Gaithersburg [Batch Address Match tool](https://maps.gaithersburgmd.gov/batchAddressMatch/) is the standard for what is considered valid - as of 1/8/25 they do not care about valid unit type and number, only the core address matters. You can upload the Address Listings report to this tool to verify all core addresses are valid.
### Tips for editing
1) To change multiple addresses, you can filter on the addresses you want to change. You can fill on the filtered list by clicking and dragging on the bottom right corner of the cell to e.g. change multiple rows to the same value. This doesn’t affect hidden filtered rows.
    * You must disable the filters before clicking any buttons, the XLSM will check for enabled filters before allowing further changes.
2) All sheets are protected from editing except for the "Validated" fields in the "Needs Autocorrect" sheet and the rows of pasted records in the "Interface" sheet. If for some reason you need to edit something, click the "Review" tab on the menu and click "Unprotect Sheet". When the workbook is saved all sheets will be reprotected.
## Generating totals and Non-Rx Final Report ##
1) Confirm that the XLSM only has the tracking methods that you want to report on by checking the rows beneath the totals on the "Home" sheet.
    * If there is an extra tracking method, you can click "Delete All Visit Data" and then reimport all of the visits. 
> [!NOTE]
> * Confirm that the "Services included in total - Non-delivery" only includes relevant non-delivery Food tracking methods, and vice versa for delivery.
2) The total counts for addresses can be seen on the "Home" sheet. Click on the cell containing the total label and look in the formula bar to view the full description for each total. The county totals count all addresses, both valid, invalid, discarded, not yet autocorrected, etc.
3) On the "Home" sheet, click "Generate Non-Rx Final Report". This will be output to the "NonRx Report" sheet. This outputs every address per unique guest ID, sorted by street name.
> [!NOTE]
> 4) If needed, edit the "Organization" column to match your organization's name. (In previous versions this asked for initials, that is no longer the case.)
5) Right-click the "NonRx Report" sheet and select "Move or copy". Select "(new book)" and check the “Create a copy” box. 
	![Right click and copy](readme/5.1finalreport.png)
6) Save the new workbook as the final grant report to be sent, named e.g. City FYnn Qn [Service] GHELP Address Listings.xlsx, for example City FY18 Q3 Food GHELP Address Listings.xlsx.
## County Totals and Form Submission
> [!NOTE]
> * Confirm that the "Services included in total" beneath County totals only includes relevant Food tracking methods.
1) If you need to submit county totals, on the "Home" sheet, select the month you want to report County totals for. Click "Copy selected zip totals code and open county totals site". This will open the county totals form in the browser.
2) In the browser, hit F12 to open developer tools. Select the "Console" tab. Click to the right of the > symbol to set the cursor. Paste using Ctrl-V into the console and hit Enter. If you receive a warning, type in 'allow pasting' without quotes and hit Enter, then paste using Ctrl-V and hit enter. This will refresh the page and all values will be filled in. Double check the household totals and zip code total values for e.g. 20877, 20878, 20879, 20886 since this code will break if the form gets updated.
    ![Paste in developer console](readme/5.2countypasted.png)
4) Fill in the other questions.
5) **Before submission** save a copy of the form (the emailed copy does not include zip code totals). Hit Ctrl-P (Right click > Print and Menu > Print do not work). Enable "Format for printers", then select "Save as PDF" in the print dialog.
> [!NOTE]
> If you need to generate Rx Final Report and totals, see below. 
## Adding Rx medication records and generating totals and Rx Final Report
1) Records pasted in on the "Home" sheet are referred to as "address records", while records pasted in on the "Rx" sheet are referred to as "medication records".
2) Ensure you have done the following:
    1) Imported in all address records pertaining to "Rx Asst" (follow steps for [Importing Data](#importing-data), and make sure "Rx Asst" tracking methods are included when copying and pasting data). Note that it does not matter if you are under the "Gaithersburg HELP" location or the "Gaithersburg HELP Financial" location.
    2) Validated or discarded all relevant address records pertaining to "Rx Asst" tracking method, so there are no "Rx Asst" tracking methods under "Needs Autocorrect". Records which are still under "Needs Autocorrect" will not be counted in totals or added to the Rx Final Report.
3) Delete all old Rx medication records on the "Rx" sheet by clicking "Delete All Rx Records". Do this even for each quarter in the same fiscal year, otherwise duplicate records will be counted twice for Rx expenditures.
4) Import Rx medication records by following steps for [Importing Data](#importing-data) with the following changes:
	* Use the preset "city grant address rx v3"
	* Select dates from the beginning of the fiscal year to the end of the current quarter.
	* On the "Rx" sheet, click "Paste Additional Rx Records and Generate Report" instead of pasting records on the "Home" sheet.
5) Verify and fix any guest IDs in the "Discarded Guest IDs" field. After fixing the guest IDs, go back to step 2, delete all medication records, and reimport. Guest IDs can be discarded for the following reasons:
    * Address may be in "Needs Autocorrect". Validate or discard the address.
	* Address record may not have been added. Make sure all "Rx Asst" tracking methods were imported when importing address records.
	* Medication record may have no medications. Contact the Rx coordinator if you think this is an error, otherwise ignore this guest ID.
6) On the "Rx Report" sheet, check for duplicate addresses with the same initials. If they exist, it's probably because of the same name being spelled incorrectly, which the report wrongly assumes are different people. As a temporary measure, you can delete the duplicated records off of the "Rx Report" sheet. To fix this error, you can do the following:
   1) On the "Addresses" sheet, search for the duplicated address to find the associated guest ID(s). There may be multiple guest IDs associated to the same address.
   2) On the "Rx" sheet, search for the guest ID(s) to find the records which are associated with that address.
   3) Check those records for name misspellings.
   4) If names are misspelled, you can either correct the misspelling in the raw "city grant address rx v3" exported CSV, or fix it in Soxbox so it doesn't happen again.
   5) When fixed, start again with deleting old Rx medication records from step 3.
6) Totals can be seen at the top of the "Rx" sheet. To export the Rx final report, start from step 4 of [generating the Non Rx report](#generating-totals-and-non-rx-final-report) but use the "Rx Report" sheet instead.

# Software setup
This project uses Excel with the Rubberduck VBA plugin for running tests and linting. See the Rubberduck installation instructions [here](https://github.com/rubberduck-vba/Rubberduck/wiki/Installing)