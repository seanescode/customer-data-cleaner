import pandas
from cleanup import cleanup_customer_data
import win32com.client as win32
from dialogs import show_cleanup_dialog

from cleanup import (
    clean_column,
    clean_name,
    clean_email,
    clean_phone_number,
    clean_address,
    clean_city,
    clean_postcode,
    is_valid_email,
    update_excel_from_dataframe
)

def main():
    df = pandas.read_excel(r"C:\Users\seane\Documents\customer_data_clean_up\client_data_cleanup_project.xlsx")

    clean_column(df, "name", clean_name)
    clean_column(df, "email", clean_email)
    clean_column(df, "phone", clean_phone_number)
    clean_column(df, "address", clean_address)
    clean_column(df, "city", clean_city)
    clean_column(df, "postcode", clean_postcode)


    # 1. Create flags for both of your duplicate conditions
    # (keep=False marks ALL copies as True so you can see the original and the duplicate together)
    dup_email = df.duplicated(subset=['name', 'email'], keep=False)
    dup_phone = df.duplicated(subset=['name', 'phone'], keep=False)

    # 2. Combine the rules using "OR" logic (|)
    # This finds rows that violate either the Email rule OR the Phone rule
    all_duplicates = df[dup_email | dup_phone]

    # 3. Sort by Name so the matching duplicates sit right next to each other
    preview_df = all_duplicates.sort_values(by='name')
    customer_ids = preview_df["customer_id"].astype(str).tolist()

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True

    wb = excel.Workbooks.Open(r"C:\Users\seane\Documents\customer_data_clean_up\client_data_cleanup_project.xlsx")
    ws = wb.Worksheets("Customer Records")

    update_excel_from_dataframe(df, ws)
    show_cleanup_dialog(
        lambda: cleanup_customer_data(customer_ids, ws),
        excel.Hwnd
    )


    # my_list = []
    # validation_message = ""
    # for email in df["email"]:
    #     if is_valid_email(email):
    #         my_list.append(email)
    #         validation_message = validation_message + email + " is valid.\n"
    #
    #
    # print(my_list)
    #
    #
    # show_error_dialog("emails that are invalid", validation_message)


if __name__ == '__main__':
    main()


