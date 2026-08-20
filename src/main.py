import pandas
from dialogs import show_error_dialog

from cleanup import (
    clean_column,
    clean_name,
    clean_email,
    clean_phone_number,
    clean_address,
    clean_city,
    clean_postcode,
    is_valid_email
)

def main():
    df = pandas.read_excel(r"C:\Users\seane\Documents\customer_data_clean_up\client_data_cleanup_project.xlsx")

    clean_column(df, "name", clean_name)
    clean_column(df, "email", clean_email)
    clean_column(df, "phone", clean_phone_number)
    clean_column(df, "address", clean_address)
    clean_column(df, "city", clean_city)
    clean_column(df, "postcode", clean_postcode)



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


