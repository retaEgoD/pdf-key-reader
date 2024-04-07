from pypdf import PdfReader
import pandas as pd
import csv
import os


PUNCTUATION = '''!()-[]{};:'"\,<>./?@#$%^&*_~'''


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
    os.remove("temp.csv")


def read_csv_to_dict(csv_filepath):
    """Reads a csv file and returns a dictionary with the 
       first column as keys and the second column as values."""
    csv_dict = {}
    with open(csv_filepath, 'r') as f:
        reader = csv.reader(f)
        for key, value in reader:
            if type(key) == str:
                key = key.lower()
            csv_dict[key] = value
    return csv_dict
            
            
def read_xlsx_to_dict(xlsx_filepath):
    """Reads an xlsx file and returns a dictionary with the
       first column as keys and the second column as values."""
    xlsx_to_csv(xlsx_filepath)
    csv_dict = read_csv_to_dict('temp.csv')
    delete_temp_csv()
    return csv_dict


def find_values_in_pdf(pdf_filepath, excel_filepath):
    """Finds keys from an excel file in a pdf file and returns a list of pairs."""
    pdf_words = read_pdf_text(pdf_filepath) 
    
    excel_extension = excel_filepath.split('.')[-1]
    if excel_extension == 'xlsx':
        csv_dict = read_xlsx_to_dict(excel_filepath)
    else:
        csv_dict = read_csv_to_dict(excel_filepath)
        
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
        
        
def find_and_save_values_in_pdf(pdf_filepath, excel_filepath, output_directory, output_filename):
    """Finds keys from an excel file in a pdf file and saves the pairs to a text file."""
    # TODO: Use OS for filepaths.
    pairs = find_values_in_pdf(pdf_filepath, excel_filepath)
    destination = os.path.join(output_directory, output_filename)
    save_pairs(destination, pairs, excel_filepath)    


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