#! /usr/bin/envoc python
import argparse
import os
import openpyxl
import pytesseract

# if the inputs files are not provided,all files matching one of the allowed patterns will be used
# unless they match one of the ignored patterns
ALLOWED_FILE_EXTENSIONS = ['.png', '.PNG']
IGNORE_LIST = []

# change this string to the location of the tesseract install if it is not in the PATH
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'


def get_args():  # all the argparse flags and help page setup
    parser = argparse.ArgumentParser(description='Use OCR to get Excel from webex meeting participant list screenshot. '
                                                 'It is recommended to take screenshots as cropped as possible in order'
                                                 ' to avoid interference from other UI elements.',
                                     epilog=f'Written by Daniel Karpelevitch')
    parser.add_argument('inputs', type=str, nargs='*',
                        help='files and directories to get data from - leave blank to use all .png files in current '
                             'directory\nif directories are specified, all .png files in those directories that match '
                             'will be used')
    parser.add_argument('-o', '--output', default='output.xlsx',
                        help='the name of the output file (remember to add ".xlsx")\ndefault is "output.xlsx"')
    return parser.parse_args()


def get_files(inputs):
    if not inputs:
        return [file for file in os.listdir(os.curdir)
                if any(ext in file for ext in ALLOWED_FILE_EXTENSIONS)
                and not any(ign in file for ign in IGNORE_LIST)]
    files = []
    for item in inputs:
        if os.path.isdir(item):
            files.extend([os.path.join(item, file) for file in os.listdir(item)
                          if any(ext in file for ext in ALLOWED_FILE_EXTENSIONS)
                          and not any(ign in file for ign in IGNORE_LIST)])
        else:
            if any([ext in item for ext in ALLOWED_FILE_EXTENSIONS]) and not any([ign in item for ign in IGNORE_LIST]):
                files.append(item)
    return files


def get_data(files):  # does ocr on each file and adds it to one big string, then splits to a list on each newline and
    # doing a bit of simple filtering for simple rows that are definitely not a name (such as a blank line, etc)
    data = ''
    for file in files:
        print(file)
        data += pytesseract.image_to_string(file, lang='eng')

    return [x for x in data.split('\n') if x not in ['', ' ', '\n', 'Cohost', 'Host', 'Me', chr(12), 'x']]


def parse_rows(data):  # given a list of strings, parses each string to normalize all the data and filter out edge cases
    # returns a sorted list of (name, "Cisco" or "Guest" or "") tuples
    rows = []
    for x in data:
        # each row looks something like this: "John Doe (Cisco)", so we split the string on the "(" to get the two parts
        name, _, cisco_or_guest = x.partition('(')
        # strip out any leading or trailing whitespace
        name = name.strip()
        # split string into words, filter out 1 and 2 letter words which are usually errors unless
        # they are all caps like JW. Capitalize each word, then combine them back into a string separated by spaces
        name = parse_words(name)
        # look at the cisco_or_guest string, which might look something like "Guest)  " so we search for "Guest" and
        # "Cisco" within the string, and manually set it to the exact string; if it can't be found, then make it blank
        if 'Guest' in cisco_or_guest:
            cisco_or_guest = 'Guest'
        elif 'Cisco' in cisco_or_guest:
            cisco_or_guest = 'Cisco'
        else:
            cisco_or_guest = ''
        # techx appears as Guest but we want it to be counted as Cisco
        if 'techx' in name.lower():
            cisco_or_guest = 'Cisco'
        # sanity check - is the total name longer than 3 characters?
        if len(name) > 3:
            rows.append((name, cisco_or_guest))

    # filter out duplicates
    rows = list(set(rows))
    # sort rows by name first
    rows.sort(key=lambda item: item[0])
    # then sort by the 2nd column ("Cisco" or "Guest")
    rows.sort(key=lambda item: item[1], reverse=True)
    # return the sorted row
    return rows


def parse_words(name):
    words = name.split()
    words = [word for word in words if len(word) > 2 or word.isupper()]
    for word in words:
        word.capitalize()
    return ' '.join(words)


def write_excel(data, output):  # write the list of tuples to excel, and add formulas for the totals
    # make a new workbook and get the current sheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # make the headers
    sheet["A1"] = "Name"
    sheet["B1"] = "Type"
    # for each item in our list, write a row in our sheet
    for row, (name, attendee_type) in enumerate(data, start=2):
        sheet[f'A{row}'] = name
        sheet[f'B{row}'] = attendee_type
    # make the headers for the totals
    sheet['C1'] = 'Total:'
    sheet['C2'] = 'Cisco:'
    sheet['C3'] = 'Guest:'
    # make the formulas for counting the total people, and how many are from cisco vs. guest
    # formulas are used rather than doing the math in python so that it updates dynamically if you make manual tweaks
    sheet['D1'] = '=COUNTA(B2:B1000)'
    sheet['D2'] = '=COUNTIF(B:B, "Cisco")'
    sheet['D3'] = '=COUNTIF(B:B, "Guest")'

    # save the file with the filename provided
    workbook.save(filename=output)


def print_result(data):  # print all the data to the console for quick debugging
    print(*[f'{x[0].ljust(30)}{x[1]}' for x in data], sep='\n')
    print('Total:', len(data))


def main():  # the main function :) drives all the others
    args = get_args()
    files = get_files(args.inputs)
    data = get_data(files)
    data = parse_rows(data)

    write_excel(data, args.output)
    print_result(data)


if __name__ == '__main__':
    main()
