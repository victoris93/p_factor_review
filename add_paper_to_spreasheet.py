import openpyxl
import requests
import sys
import datetime

def fetch_article_metadata(doi):
    # Fetch metadata from a DOI, for example, using CrossRef API
    # if the doi link is given, strip it of the doi.org part
    if doi.startswith('https://doi.org/'):
        doi = doi[16:]
    url = f"https://api.crossref.org/works/{doi}"
    response = requests.get(url)
    if response.status_code == 200:
        return response.json()['message']
    else:
        return None

def fill_excel_with_metadata(excel_path, sheet_name, metadata):
    # Open an Excel file and select a specific tab
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook[sheet_name]

    empty_row = sheet.max_row + 1

    # Assuming specific columns for specific metadata fields
    sheet[f'B{empty_row}'] = metadata.get('title', '')[0]  # Title
    sheet[f'C{empty_row}'] = ', '.join([f'{i["family"]} {i["given"]}' for i in metadata.get('author', '') if "family" in i and "given" in i])  # Author
    try:
        sheet[f'D{empty_row}'] = metadata.get('published-print', '')['date-parts'][0][0]  # Year
    except:
        sheet[f'D{empty_row}'] = metadata.get('published-online', '')['date-parts'][0][0]
    sheet[f'F{empty_row}'] = datetime.datetime.now().strftime('%d.%m.%Y')  # Current date
    sheet[f'G{empty_row}'] = f'https://doi.org/{metadata.get("DOI", "")}'  # DOI link

    # Save the changes
    workbook.save(excel_path)

def main():
    for doi in sys.argv[1:]:
        try:
            metadata = fetch_article_metadata(doi)
            if metadata:
                excel_path = 'metalab_decisions.xlsx'
                sheet_name = 'Studies'
                fill_excel_with_metadata(excel_path, sheet_name, metadata)
                print(f"Metadata for DOI {doi} has been added to the Excel file.")
            else:
                print(f"Metadata could not be fetched for the DOI {doi}.")
        except Exception as e:
            print(f"An error occurred while processing DOI {doi}: {e}")

if __name__ == "__main__":
    main()
