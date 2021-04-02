#! /usr/bin/python
import argparse
import os
import openpyxl
import pytesseract

# all files matching one of the allowed patterns will be used
# unless they match one of the ignored patterns
ALLOWED_FILE_EXTENSIONS = ['.png', '.PNG', '.jpg']
IGNORE_LIST = []

# change this string to the location of the tesseract install if it is not in the PATH
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract'


def get_args():  # all the argparse flags and help page setup
    parser = argparse.ArgumentParser(description='Use OCR to get Excel from webex meeting participant list screenshot. '
                                                 'It is recommended to take screenshots as cropped as possible in order'
                                                 ' to avoid interference from other UI elements. Do not include the'
                                                 ' icons on the right or the thumbnails on the left of the participant'
                                                 ' list, and expand the participant list so no words are cut off.',
                                     epilog=f'Written by Daniel Karpelevitch (dkarpele@cisco.com)')
    parser.add_argument('inputs', type=str, nargs='*',
                        help='files and directories to get data from - leave blank to use all .png files in current '
                             'directory\nif directories are specified, all .png files in those directories that match '
                             'will be used')
    parser.add_argument('-o', '--output', default='output.xlsx',
                        help='the name of the output file (remember to add ".xlsx")\ndefault is "output.xlsx"')
    return parser.parse_args()


def get_files(inputs):  # gather the files based on provided input or if no input, all matching files in current dir
    if not inputs:  # inputs is a list of file names and directories
        return [file for file in os.listdir(os.curdir)
                if any(ext in file for ext in ALLOWED_FILE_EXTENSIONS)
                and not any(ign in file for ign in IGNORE_LIST)]
    files = []
    for item in inputs:  # iterate through each item in the inputs list
        if os.path.isdir(item):  # if the item is a directory, add all matching files in that directory to files list
            files.extend([os.path.join(item, file) for file in os.listdir(item)
                          if any(ext in file for ext in ALLOWED_FILE_EXTENSIONS)
                          and not any(ign in file for ign in IGNORE_LIST)])
        else:  # otherwise, the item is a filename. Add it to the files list
            files.append(item)
    return files  # return the files list containing all the files to be used


def get_data(files):  # does ocr on each file and adds it to one big string, then splits to a list on each newline and
    # doing a bit of simple filtering for simple rows that are definitely not a name (such as a blank line, etc)
    data = ''
    for file in files:
        print(file)
        data += pytesseract.image_to_string(file, lang='eng')

    return [x for x in data.split('\n') if x not in ['', ' ', '\n', 'Cohost', 'Host', 'Me', chr(12), 'x', 'Q_ Search']]


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


def parse_words(name):  # the whole name is passed in and each word processed separately
    filter_out = ['Guest', 'Desk', 'Pro', 'DX80', 'Participants', ]  # add strings here and they will be filtered out
    words = name.split()  # turn the name string into list of words
    new_words = []
    num_words = len(words)  # things are handled differently depending on how many words there are in the name

    for word in words:
        if '@' not in word or '.' not in word:  # check if it looks like an email
            word = ''.join([char for char in word if char.isalnum() or char in '-_'])  # strip out non-alphanum or -_
        else:  # if it looks like an email, don't process it
            new_words.append(word)
            continue
        if len(word) > 2:  # if a word and does not contain any of the filtered words, capitalize and add back to list
            if not (any(f in word for f in filter_out)):
                new_words.append(word.capitalize())
        elif len(word) == 2 and not(word.islower()) and num_words <= 2:
            new_words.append(word)  # handle 2-letter initials, but only if there are 2 or less words, otherwise ignore
            # this is because an initial in a 3-word name is usually caused by the initials in the thumbnail being read
        # a 1-letter "word" is ignored

    return ' '.join(new_words)  # put all the words back together


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


def print_result(data):  # print all the data gathered to the console for quick debugging
    print(*[f'{x[0].ljust(30)}{x[1]}' for x in data], sep='\n')
    print('Total:', len(data))


def main():  # the main function :) drives all the others
    args = get_args()
    files = get_files(args.inputs)
    data = get_data(files)
    parsed_data = parse_rows(data)

    write_excel(parsed_data, args.output)
    print_result(parsed_data)


if __name__ == '__main__':
    main()
