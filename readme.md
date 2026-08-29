# Customer Data Cleaner
This tool can be used to help a business quickly clean a customer or vendor list they have.

![before and after.PNG](images/before%20and%20after.PNG)

# Cleaning
- removes unnecessary white spaces throughout
- gives appropriate capitalisation to names and addresses
- formats phone numbers to a standard format

# Removes bad data 
- provides a dialog panel with buttons to filter to: 
  - show duplicate entries
  - show email addresses that have incorrect syntax
# Useful features
- This project was built using a configuration file so users can change the following if they ever need to:
  - the folder location of the spreadsheet
  - column names used in the spreadsheet
  - worksheet name used in the spreadsheet
- Interface buttons allow user to navigate between seeing duplicates/bad syntax emails/final clean data

![customer list cleaner GUI Panel.PNG](images/customer%20list%20cleaner%20GUI%20Panel.PNG)
- User has control - duplicates entries and bad syntax emails can be manually reviewed before deletion.

![example output from duplicates and invalid emails.PNG](images/example%20output%20from%20duplicates%20and%20invalid%20emails.PNG)
- Error handling is built in. For example, if the list gets moved to another folder, the tool will show an error message.
- Duplicates entries have multiple checkpoints to ensure all duplicates are caught.
  - if the following match across two rows it shows as a duplicate:
      - name and email address 
      - name and phone number
      - name and address
      - name and postcode
- Rather than doing a seperate report of duplicates and bad syntax emails this automation filters to
them on the click of a button meaning the user can update immediately rather than switching
between screens.

# Limitations

* This tool is currently designed for Windows users with Microsoft Excel installed. The application uses Windows-specific Excel automation, meaning it is not currently compatible with macOS, Linux or other spreadsheet applications.
* The tool is currently designed primarily for Irish users. For example, Eircode formatting is hardcoded to insert a space after the third character, although this could be adapted to support postcode formats used in other countries.
* Duplicate detection does not currently use fuzzy matching, so customers with spelling mistakes or variations in their details may not always be identified. However, the tool uses multiple customer fields to identify potential duplicates rather than relying on a single piece of information. It checks customer names and cross-checks them against email addresses, phone numbers, street addresses and postcodes, making it possible to identify duplicate records even when the same customer details appear across different fields.
* The email checker does not verify whether an email mailbox exists or is actively monitored. However, it performs email syntax validation, allowing the tool to automatically identify incorrectly formatted email addresses and common data entry errors that could otherwise affect the quality and usability of customer data.

# Future Improvements

* Add email domain checks. This would check the part of the email address after the `@` symbol by querying the Domain Name System (DNS) to confirm that the domain exists and can receive emails.
* Add statistics showing how many records were cleaned and how many updates were made.
* The current email syntax check does not confirm whether an email address actually exists; it only checks whether the syntax is valid. Users who need to verify whether an email address exists could use an email verification service, although these are typically paid services once free usage limits have been reached.
* Add fuzzy matching to help identify potential duplicate records containing small differences or typos.
* Make the application cross-platform. The current version uses Windows-specific automation to interact directly with Microsoft Excel. A future version could instead use a file-based workflow, where users provide an `.xlsx` or `.csv` file, the application processes the data, and then saves a cleaned version. This would make the core functionality usable on Windows, macOS and Linux, with the resulting files compatible with applications such as Microsoft Excel, LibreOffice and Google Sheets.


# What I Learned

One of the main things I learned while building this project was the importance of considering platform compatibility when choosing technologies and designing an application's architecture.

The application was originally designed to automate Microsoft Excel directly using Windows-specific tools. This was suitable for the original goal of building a desktop tool for Windows and Excel users, but it also meant that the spreadsheet automation layer became dependent on a specific operating system and application.

At the same time, the core functionality of the project — including data cleaning, validation and duplicate detection — is largely platform-independent because it works with Pandas DataFrames rather than directly with Excel.

This highlighted the importance of separating core business logic from external integrations. If I revisit this project in the future, I would consider moving towards a file-based workflow where users import an `.xlsx` or `.csv` file and export a cleaned version. This would make it possible to support users on Windows, macOS and Linux without needing to automate a specific spreadsheet application.
