import os
import requests
import json

# Import dot env for sensitive information
from dotenv import load_dotenv
load_dotenv()

def convert_excelsheet_to_pdf(excel_sheet):
    '''Converts an excelsheet to a pdf.'''
    # Check if file exists
    if os.path.exists(excel_sheet)==True:
        api_key = os.environ.get('API_KEY')
        instructions = {
        'parts': [
            {
            'file': 'document'
            }
        ]
        }

        response = requests.request(
        'POST',
        'https://api.pspdfkit.com/build',
        headers = {
            'Authorization': f'Bearer {api_key}'
        },
        files = {
            'document': open(excel_sheet, 'rb')
        },
        data = {
            'instructions': json.dumps(instructions)
        },
        stream = True
        )

        # If response is successful write a pdf file
        if response.ok:
            with open('result.pdf', 'wb') as fd:
                for chunk in response.iter_content(chunk_size=8096):
                    fd.write(chunk)

        # If response is unsuccessful print response message and exit with a code of -1
        else:
            print(response.text)
            exit(-1)

    # If file does not exist print error and exit with code of -1
    else:
        print("Enter a file that exists!")
        exit(-1)