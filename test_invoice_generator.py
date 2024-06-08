import unittest
from unittest.mock import patch, mock_open, MagicMock
import pandas as pd
from datetime import datetime
import invoice_generator  
import json
import logging

class TestInvoiceGenerator(unittest.TestCase):

    @patch('builtins.open', new_callable=mock_open, read_data='Day\tHours\n1\t8\n2\t8\n3\t8\n4\t8\n5\t8')
    def test_format_date(self,mock_open):
        date = datetime(2023, 10, 5)
        formatted_date = invoice_generator.format_date(date)
        self.assertEqual(formatted_date, "October 5")

    @patch('builtins.open', new_callable=mock_open, read_data='Day\tHours\n1\t8\n2\t8\n3\t8\n4\t8\n5\t8')
    def test_count_workdays_in_week(self,mock_open):
        df = pd.DataFrame({'1': [8], '2': [8], '3': [8], '4': [8], '5': [8]})
        start_date = datetime(2023, 10, 1)
        end_date = datetime(2023, 10, 5)
        workdays = invoice_generator.count_workdays_in_week(df, start_date, end_date)
        self.assertEqual(workdays, 5)

    @patch('invoice_generator.config', {
        'routing_number': '123456789',
        'swift_bic': 'ABCDEF12',
        'account_number': '987654321',
        'wise_address': '123 Wise St.',
        'company_name': 'Test Company',
        'company_address': '123 Test St.',
        'company_email': 'test@example.com'
    })
    def test_create_data_dictionary(self):
        invoice_number = '202310'
        formatted_date = '05 Oct 2023'
        account_holder = 'Test User'
        company_name = 'Test Company'
        company_address = '123 Test St.'
        company_email = 'test@example.com'

        data_dict = invoice_generator.create_data_dictionary(invoice_number, formatted_date, account_holder, company_name, company_address, company_email)
        self.assertEqual(data_dict['{{INVOICE_NUMBER}}'], invoice_number)
        self.assertEqual(data_dict['{{DATE}}'], formatted_date)
        self.assertEqual(data_dict['{{ACCOUNT_HOLDER}}'], account_holder)
        self.assertEqual(data_dict['{{COMPANY_NAME}}'], company_name)
        self.assertEqual(data_dict['{{COMPANY_ADDRESS}}'], company_address)
        self.assertEqual(data_dict['{{COMPANY_EMAIL}}'], company_email)

    def test_total_due_calculation(self):
        invoice_number = '202310'
        formatted_date = '05 Oct 2023'
        
        # Mocking the load_dataframe function
        with patch('invoice_generator.load_dataframe') as mock_load_dataframe:
            # Mocking the return value of load_dataframe
            mock_load_dataframe.return_value = pd.DataFrame({
                '1': [8, 8, 8, 8, 8],
                '2': [8, 8, 8, 8, 8],
                '3': [8, 8, 8, 8, 8],
                '4': [8, 8, 8, 8, 8],
                '5': [0, 0, 0, 0, 0]
            })
        
            # Mocking the get_weekday_ranges_for_csv_month function
            with patch('invoice_generator.get_weekday_ranges_for_csv_month') as mock_get_weekday_ranges:
                mock_get_weekday_ranges.return_value = ([], [])
            
            # Mocking the fill_template function
            with patch('invoice_generator.fill_template') as mock_fill_template:
                # Mocking the main function to generate invoice
                with patch('invoice_generator.main') as mock_main:
                    mock_main.side_effect = lambda: invoice_generator.fill_template(
                        invoice_generator.TEMPLATE_FILE,
                        f'{invoice_generator.OUTPUT_DIR}/invoice_Test User_{invoice_number}.docx',
                        {
                            '{{INVOICE_NUMBER}}': invoice_number,
                            '{{DATE}}': formatted_date,
                            '{{ACCOUNT_HOLDER}}': 'Test User',
                            '{{ROUTING_NUMBER}}': '123456789',
                            '{{SWIFT_BIC}}': 'ABCDEF12',
                            '{{ACCOUNT_NUMBER}}': '987654321',
                            '{{WISE_ADDRESS}}': '123 Wise St.',
                            '{{COMPANY_NAME}}': 'Test Company',
                            '{{COMPANY_ADDRESS}}': '123 Test St.',
                            '{{COMPANY_EMAIL}}': 'test@example.com',
                            '{{TOTAL_DUE}}': '$160.00'
                        }
                    )
                    invoice_generator.main()
                    # Asserting that fill_template was called with the correct total due value
                    mock_fill_template.assert_called_with(
                        invoice_generator.TEMPLATE_FILE,
                        f'{invoice_generator.OUTPUT_DIR}/invoice_Test User_{invoice_number}.docx',
                        {
                            '{{INVOICE_NUMBER}}': invoice_number,
                            '{{DATE}}': formatted_date,
                            '{{ACCOUNT_HOLDER}}': 'Test User',
                            '{{ROUTING_NUMBER}}': '123456789',
                            '{{SWIFT_BIC}}': 'ABCDEF12',
                            '{{ACCOUNT_NUMBER}}': '987654321',
                            '{{WISE_ADDRESS}}': '123 Wise St.',
                            '{{COMPANY_NAME}}': 'Test Company',
                            '{{COMPANY_ADDRESS}}': '123 Test St.',
                            '{{COMPANY_EMAIL}}': 'test@example.com',
                            '{{TOTAL_DUE}}': '$160.00'
                        }
                    )
    @patch("builtins.open", new_callable=mock_open, read_data="{}")
    @patch("json.load")
    def test_handle_empty_or_malformed_config_json(self, mock_json_load, mock_open):
        # Mocking json.load to raise a JSONDecodeError
        mock_json_load.side_effect = json.JSONDecodeError("Expecting value", "", 0)
        
        with self.assertLogs(level="ERROR") as cm:
            invoice_generator.main()
            self.assertIn("Error parsing configuration JSON", cm.output[0])
            
    @patch('invoice_generator.config', {
        'template_file': 'Invoice_Template.docx',
        'csv_path': 'dummy_path.csv',
        'output_dir': 'dummy_output_dir',
        'day_rate_per_week': 100,
        'client': 'Test Client',
        'account_holder': 'Test User',
        'company_name': 'Test Company',
        'company_address': '123 Test St.',
        'routing_number': '123456789',
        'swift_bic': 'ABCDEF12',
        'account_number': '123456789',
        'wise_address': 'wiseaddress',
        'company_email': 'test@example.com'
    })
    def test_handle_non_existent_template_file(self):
        with patch("os.path.exists") as mock_exists:
            mock_exists.return_value = False
            with patch("logging.error") as mock_logging_error:
                invoice_generator.main()
                mock_logging_error.assert_called_with("The file Invoice_Template.docx was not found.")

if __name__ == '__main__':
    unittest.main()