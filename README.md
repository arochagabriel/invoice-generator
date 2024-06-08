# Invoice Generator

This Python script generates invoices based on data from a CSV file and a Word document template. The script reads configuration settings from a `config.json` file and outputs the generated invoices to a specified directory.

## Requirements

- Python 3.x
- pandas
- python-docx

## Installation

1. Clone the repository or download the script files.
2. Install the required Python packages using pip:

    ```sh
    pip install pandas python-docx
    ```

3. Ensure you have a `config.json` file in the same directory as the script. The `config.json` file should have the following structure:

    ```json
    {
        "csv_path": "path/to/your/csvfile.csv",
        "template_file": "path/to/your/template.docx",
        "output_dir": "path/to/output/directory",
        "day_rate_per_week": 100.0,
        "client": "Client Name",
        "account_holder": "Account Holder Name",
        "company_name": "Your Company Name",
        "company_address": "Company adress",
        "company_email": "Your Company Email",
        "routing_number": "Your Routing Number",
        "swift_bic": "Your SWIFT/BIC",
        "account_number": "Your Account Number",
        "wise_address": "Your Wise Address"
    }
    ```

## Usage

1. Ensure your CSV file is formatted correctly and located at the path specified in the `config.json` file.
2. Ensure your Word document template is located at the path specified in the `config.json` file.
3. Run the script:

    ```sh
    python invoice_generator.py
    ```

4. The generated invoices will be saved in the output directory specified in the `config.json` file.

## Script Details

- `load_dataframe(csv_path)`: Loads the CSV file into a pandas DataFrame.
- `format_date(date)`: Formats a date to 'Month Day'.
- `get_weekday_ranges_for_current_month()`: Gets weekday ranges for the current month.
- `count_workdays_in_week(df, start_date, end_date)`: Counts workdays in a given week.
- `create_data_dictionary(invoice_number, formatted_date, ACCOUNT_HOLDER, COMPANY_NAME, COMPANY_ADDRESS, COMPANY_EMAIL)`: Creates a dictionary with invoice data.
- `fill_template(template_file, output_path, data)`: Fills the Word document template with data and saves the output.
- `generate_invoice(df, invoice_number, formatted_date)`: Generates an invoice based on the DataFrame and other details.
- `main()`: Main function that loads the DataFrame and generates the invoice.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
