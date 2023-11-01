import pandas as pd

def split_address(mailing_address):
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

def process_excel_file(input_file, output_file):

    # Load the excel file

    df = pd.read_excel(input_file)

    # Check if "Mailing address" Column exists
    if 'Mailing Address' in df.columns:
        # Split the Mailing Address into Address, City and State
        df['Address'], df['City'], df['State'] = zip(*df['Mailing Address'].apply(split_address))
    else:
        print('Mailing Address column does not exist in the excel file')
        return
    # Save the new Excel file
    df.to_excel(output_file, index=False)
    print('process completed. The file is saved as {}'.format(output_file))

if __name__ == '__main__':
    input_excel_file = 'Private Lenders-OHIO-Mahoning.xlsx'
    output_excel_file = 'output.xlsx'
    process_excel_file(input_excel_file, output_excel_file)
