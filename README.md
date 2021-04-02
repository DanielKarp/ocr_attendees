# ocr_attendees
usage: ocr_attendees.py [-h] [-o OUTPUT] [inputs [inputs ...]]

Use OCR to get Excel from webex meeting participant list screenshot. It is
recommended to take screenshots as cropped as possible in order to avoid
interference from other UI elements. Do not include the icons on the right or
the thumbnails on the left of the participant list, and expand the participant
list so no words are cut off.

positional arguments:
  inputs                files and directories to get data from - leave blank
                        to use all .png files in current directory if
                        directories are specified, all .png files in those
                        directories that match will be used

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        the name of the output file (remember to add ".xlsx")
                        default is "output.xlsx"

Written by Daniel Karpelevitch (dkarpele@cisco.com)
