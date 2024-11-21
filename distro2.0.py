import pandas as pd
import win32com.client as win32
from datetime import datetime
import os


def load_excel_data(file_path):
    """Load data from an Excel file."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    return pd.read_excel(file_path)


def format_load_data(load_df):
    """Format load data into a compact array of strings."""
    load_data_array = []
    for _, row in load_df.iterrows():
        try:
            load_data = (
                f"Trip: {row['Trip']} | {row['Orig City/St']} → {row['Dest City/St']}\n"
                f"Pickup: {row['Date'].strftime('%m/%d/%y')} {row['Time'].strftime('%H:%M')} | "
                f"Delivery: {row['Date_2'].strftime('%m/%d/%y')} {row['Time_3'].strftime('%H:%M')}\n"
                f"Weight: {row['Weight']} Lbs | Temp: {row['Temp']} | Stops: {row['P/S']}\n"
                f"-------------------------\n"
            )
            load_data_array.append(load_data)
        except Exception as e:
            print(f"Error processing row: {row}\n{e}")
    return load_data_array


def construct_email_body(load_data_array):
    """Construct a minimal email body."""
    header = "Dear Carrier,\n\nHere are the available loads:\n\n"
    footer = "\nThank you,\nTrue Blue SCM Team"
    return header + "\n".join(load_data_array) + footer


def send_emails(email_df, email_body, subject):
    """Send emails with Outlook."""
    outlook = win32.Dispatch("Outlook.Application")
    count = 0
    for _, row in email_df.iterrows():
        try:
            recipient_email = row["Email"]
            mail = outlook.CreateItem(0)
            mail.To = recipient_email
            mail.Subject = subject
            mail.Body = email_body
            mail.Send()
            count += 1
            print(f"{count} email(s) sent to {recipient_email}")
        except Exception as e:
            print(f"Error sending email to {row['Email']}: {e}")


def main():
    # File paths
    carrier_list = "distro_folder/carrier_list.xlsx"
    load_excel_file = "distro_folder/load_list.xlsx"

    # Load data from Excel files
    try:
        load_df = load_excel_data(load_excel_file)
        email_df = load_excel_data(carrier_list)
    except FileNotFoundError as e:
        print(e)
        return

    # Format load data
    load_data_array = format_load_data(load_df)
    email_body = construct_email_body(load_data_array)

    # Prepare email subject
    current_datetime = datetime.now().strftime("%m/%d/%y %H:%M")
    subject = f"Available Loads True Blue SCM - {current_datetime}"

    # Send emails
    send_emails(email_df, email_body, subject)

    print("All emails sent successfully.")


if __name__ == "__main__":
    main()
