import pandas
from cleanup import cleanup_customer_data
import win32com.client as win32
from dialogs import show_email_cleanup_panel, create_dialog_root

from cleanup import (
    clean_column,
    clean_name,
    clean_email,
    clean_phone_number,
    clean_address,
    clean_city,
    clean_postcode,
    filter_to_invalid_emails,
    update_excel_from_dataframe
)

def main():
    file_loc = r"C:\Users\seane\Documents\customer_data_clean_up\client_data_cleanup_project.xlsx"
    df = pandas.read_excel(file_loc)

    clean_column(df, "name", clean_name)
    clean_column(df, "email", clean_email)
    clean_column(df, "phone", clean_phone_number)
    clean_column(df, "address", clean_address)
    clean_column(df, "city", clean_city)
    clean_column(df, "postcode", clean_postcode)

    # 1. Create flags for both of your duplicate conditions
    dup_email = df.duplicated(subset=['name', 'email'], keep=False)
    dup_phone = df.duplicated(subset=['name', 'phone'], keep=False)
    dup_address_first_line = df.duplicated(subset=['name', 'address'], keep=False)
    dup_postcode = df.duplicated(subset=['name', 'postcode'], keep=False)

    # 2. Combine the rules using "OR" logic (|)
    all_duplicates = df[dup_email | dup_phone | dup_address_first_line | dup_postcode]

    # 3. Sort by Name so the matching duplicates sit right next to each other
    preview_df = all_duplicates.sort_values(by='name')
    customer_ids = preview_df["customer_id"].astype(str).tolist()

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Open(r"C:\Users\seane\Documents\customer_data_clean_up\client_data_cleanup_project.xlsx")
    ws = wb.Worksheets("Customer Records")

    update_excel_from_dataframe(df, ws)

    root = create_dialog_root()
    # show_cleanup_dialog(
    #     root,
    #     lambda: cleanup_customer_data(customer_ids, ws),
    #     excel.Hwnd
    # )

    show_email_cleanup_panel(
        root,
        excel.Hwnd,
        lambda: filter_to_invalid_emails(pandas.read_excel(file_loc), "email", ws)
    )

    root.mainloop()


if __name__ == '__main__':
    main()