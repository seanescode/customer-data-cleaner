import pandas
import win32com.client as win32
from dialogs import (
    show_cleanup_panel,
    create_dialog_root,
    show_temporary_message
)
from cleanup import (
    clean_column,
    clean_name,
    clean_email,
    clean_phone_number,
    clean_address,
    clean_city,
    clean_postcode,
    filter_to_invalid_emails,
    update_excel_from_dataframe,
    remove_filters,
    get_duplicate_customers,
    filter_duplicate_customers
)

def main():

    # set up dataframe
    file_loc = r"C:\Users\seane\Documents\customer_data_clean_up\client_data_cleanup_project.xlsx"
    df = pandas.read_excel(file_loc)

    #clean data
    clean_column(df, "name", clean_name)
    clean_column(df, "email", clean_email)
    clean_column(df, "phone", clean_phone_number)
    clean_column(df, "address", clean_address)
    clean_column(df, "city", clean_city)
    clean_column(df, "postcode", clean_postcode)

    all_duplicates = get_duplicate_customers(df)
    # Sort by Name so the matching duplicates sit right next to each other
    sorted_duplicates = all_duplicates.sort_values(by='name')
    # get the customer_id for each duplicate value so can use this to filter data to show duplicates
    customer_ids = sorted_duplicates["customer_id"].astype(str).tolist()

    # Open spreadsheet
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True
    wb = excel.Workbooks.Open(r"C:\Users\seane\Documents\customer_data_clean_up\client_data_cleanup_project.xlsx")
    ws = wb.Worksheets("Customer Records")

    update_excel_from_dataframe(df, ws)
    root = create_dialog_root()
    show_cleanup_panel(
        root,
        excel.Hwnd,
        lambda: filter_duplicate_customers(
            file_loc,
            ws,
            lambda text: show_temporary_message(root, text)
        ),
        lambda: filter_to_invalid_emails(
            file_loc,
            ws,
            "email",
            lambda text: show_temporary_message(root, text)
        ),
        lambda: remove_filters(ws)
    )
    root.mainloop()

if __name__ == '__main__':
    main()
