from pypdf import PdfReader
import pandas as pd
import csv
from pathlib import Path


PUNCTUATION = '''!()-[]{};:'"\,<>./?@#$%^&*_~'''

def excel_column_to_int(column):
    n = len(column)
    column_num = 0
    for i in range(n):
        digit = ord(column[n - i - 1]) - ord('A') + 1
        column_num += digit * (26 ** i)
    return column_num -1


def remove_punctuation(text):
    """Removes punctuation from a string."""
    res = ""
    for char in text:
        if char not in PUNCTUATION:
            res += char
    return res


def read_pdf_text(pdf_filepath):
    """Reads text from a PDF file and returns a set of words in the text."""
    reader = PdfReader(pdf_filepath)
    pages = reader.pages
    page_texts = [page.extract_text() for page in pages]
    all_text = ' '.join(page_texts)
    all_text_no_punctuation = remove_punctuation(all_text)
    pdf_words = set(all_text_no_punctuation.lower().split())
    return pdf_words


def xlsx_to_csv(xlsx_filepath):
    """Converts an xlsx file to a temporary csv file."""
    read_file = pd.read_excel(xlsx_filepath)
    read_file.to_csv("temp.csv", index=None, header=False)
    
    
def delete_temp_csv():
    """Deletes the temporary csv file created by xlsx_to_csv."""
    Path.unlink(Path("temp.csv"))


def read_csv_to_dict(csv_filepath, key_col=0, val_col=1):
    """Reads a csv file and returns a dictionary with the 
       first column as keys and the second column as values."""
    csv_dict = {}
    with open(csv_filepath, 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            key, value = row[key_col], row[val_col]
            if type(key) == str:
                key = key.lower()
            csv_dict[key] = value
    return csv_dict
            
            
def read_xlsx_to_dict(xlsx_filepath, key_col, val_col):
    """Reads an xlsx file and returns a dictionary with the
       first column as keys and the second column as values."""
    xlsx_to_csv(xlsx_filepath)
    csv_dict = read_csv_to_dict('temp.csv', key_col, val_col)
    delete_temp_csv()
    return csv_dict


def find_values_in_pdf(pdf_filepath, excel_filepath, excel_key_col, excel_val_col):
    """Finds keys from an excel file in a pdf file and returns a list of pairs."""
    pdf_filepath = Path(pdf_filepath)
    excel_filepath = Path(excel_filepath)
    
    key_col = excel_column_to_int(excel_key_col)
    val_col = excel_column_to_int(excel_val_col)
    
    pdf_words = read_pdf_text(pdf_filepath)
    if excel_filepath.suffix == '.xlsx':
        csv_dict = read_xlsx_to_dict(excel_filepath, key_col, val_col)
    else:
        csv_dict = read_csv_to_dict(excel_filepath, key_col, val_col)
        
    keys = set(csv_dict.keys())
    keys_in_text = []
    
    for key in keys:
        if key.lower() in pdf_words:
            keys_in_text.append(key)
            
    pairs = [(key, csv_dict[key]) for key in keys_in_text]
    return pairs


def save_pairs(destination, pairs, original_filename):
    """Saves a list of key value pairs to a text file."""
    header = f'Keys found in {original_filename}:\n'
    pair_strings = [f"{key}: {value}\n" for key, value in pairs]
    with open(destination, 'w') as f:
        f.write(header)
        f.writelines(pair_strings)
        
        
def find_and_save_values_in_pdf(pdf_filepath, excel_filepath, output_directory, output_filename, excel_key_col, excel_val_col):
    """Finds keys from an excel file in a pdf file and saves the pairs to a text file."""
    pairs = find_values_in_pdf(pdf_filepath, excel_filepath, excel_key_col, excel_val_col)
    path = Path(output_directory)
    destination = path / (output_filename + '.txt')
    save_pairs(destination, pairs, excel_filepath.split('/')[-1])    


# def xlsx_test():
#     excel = "testing/excel_test.xlsx"
#     pdf = "testing/test.pdf"
#     destination = "testing/new_test/result.txt"
#     pairs = find_values_in_pdf(pdf, excel)
#     save_pairs(destination, pairs, excel)


# def main():
#     xlsx_test()
    

# if __name__ == '__main__':
#     main()