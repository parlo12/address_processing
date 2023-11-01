import pandas as pd

class ExcelProcessor:
    def split_address(self, mailing_address):
        parts = mailing_address.split()
        if len(parts) >= 3:
            state = parts[-1]
            city = parts[-2]
            address = ' '.join(parts[:-2])
            return address, city, state
        #return address, None, None
        else:
            print('Warning: could not split address: {}'.format(mailing_address))
            return mailing_address, None, None

    def process_excel_file(self, input_file, output_file, callback):
        try:
            # Load the excel file
            df = pd.read_excel(input_file)
            # Check if "Mailing address" Column exists
            if 'Mailing Address' in df.columns:
                # Split the Mailing Address into Address, City and State
                df['Address'], df['City'], df['State'] = zip(*df['Mailing Address'].apply(self.split_address))
                df.to_excel(output_file, index=False)
                callback('process completed. The file is saved as {}'.format(output_file))
            else:
                callback('Mailing Address column does not exist in the excel file')
                print('Mailing Address column does not exist in the excel file')
        except Exception as e:
            callback('Error: {}'.format(e))
            print('Error: {}'.format(e))