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

# Future Improvements
- I could also include email domain checks (this is the part of the email after the @ symbol).
This queries the Domain Name System (DNS) to confirm that the domain part of the email exists
and is able to receive emails.
- I could also include statistics to show how many records have been cleaned/updates done.
- The email syntax check doesn't confirm if an email actually exists. It just checks the
syntax is valid. If the user wants to check if an email exists they could use
an email verification service. These are usually paid services after free limits are used.   

# Limitations
- This tool is currently only built for Windows users using Excel.
- This tool is built for Irish users - for example, the post codes are hardcoded to input a
space after the third character (though this could be easily updated if and when needed for other countries)
- No fuzzy matching - currently catches exact matches only would be useful to also catch
entries that have typos - would be a good future update. 