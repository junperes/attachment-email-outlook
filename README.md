# Attachment Email via Outlook

This script automates the process of sending emails with attachments via Outlook. It reads data from an Excel file containing email addresses and merges it with another Excel file to extract relevant information. Then, it sends emails with attached PDF files based on the merged data.

## Requirements

- Python 3.x
- `win32com` library
- `pandas` library

## Usage

1. Clone this repository:

    ```
    git clone <repository_url>
    ```

2. Install the required libraries:

    ```
    pip install pandas pywin32
    ```

3. Prepare your Excel files:
   
   - Ensure you have an Excel file containing email addresses (e.g., `Base_nome_cnpj_email.xlsx`).
   - Ensure you have another Excel file containing the necessary data for the email content and file attachments (e.g., `base.xlsx`). Make sure the column names match those expected by the script (`Número`, `Nome`, `Data de vencimento`, `EMAIL`).

4. Run the script:

    ```
    python attachment_email_outlook.py
    ```

5. Follow the prompts to input the email subject and the name of the base Excel file.

6. The script will send emails with attachments to the specified email addresses via Outlook.

## Notes

- This script assumes you have Microsoft Outlook installed and configured on your system.
- Make sure to customize the script according to your specific requirements, such as adjusting column names, file paths, and email content.
- Ensure that the Excel files and PDF attachments are located in the same directory as the script.
- The script logs the sent emails in an Excel file named `emails_enviados.xlsx`.

If you encounter any issues or have suggestions for improvements, feel free to open an issue or submit a pull request!
