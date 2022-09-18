#!/usr/bin/env python3
import os
import re
import csv
import pandas as pd
import PySimpleGUI as sg # Requires 'pysimplegui' package
from docx import Document # Requires 'python-docx' package
from pdfminer.high_level import extract_text # Requires 'pdfminer.six' package

def main():
    
    # Base 64-encoded icon for GUI
    icon = b'iVBORw0KGgoAAAANSUhEUgAAAQsAAADxCAMAAADr9WKrAAAACXBIWXMAAC4jAAAuIwF4pT92AAAAk1BMVEX/////MmT/MmT/MmT/MmT/MmT/MmT/MmT/MmT/MmT/MmT/MmT/MmT/MmT/MmT/MmT/MmT/Pm3/P27/S3f/THf/WIH/WYH/ZYr/ZYv/cpT/cpX/fp7/f57/i6f/jKj/mLH/mbL/pbv/prz/ssX/s8X/ucr/v87/v8//zNj/zNn/2OL/2eL/5ez/5uz/8vX/8/b////kxu4TAAAAEHRSTlMAECAwQFBgcICQoLDA0ODwVOCoyAAACdxJREFUeNrtXVt7ojoUBURFRQyiw1i00dJx0FrL//9156FzWhHQXHaSHep6ln6wmmTtexwHEv4gCEJNmARBv+fgRH80JfoRDtHx0RvPiClEQxcRE35IzGLsPZi4YAPD2nDHBAVmQ/OLYkawIDS8NAKCCDP/sT++MTBHxZSQBxlYqTBGxoSQBxmfGBGcMHCADpBSQWbapdWbYeWChLq5CAleDB87xNAucSPMXJDxjzW9G6AxvuMhp0Ln8TnBzgXpa/PT0VNBIl3H5xQ/FyR46Om3rmqJgLozG7jQo6sjYgf8h55+YfqzHRHNkQzfGirUuyWRPVyo1tWA2ATvoadfmKjkYkzsgkJd7VlGBYkeeqoh3Ne3jgpluupG9nFBRg89Vayr3sxKLsKHnioN9/mWUqFCV0NbuYB3SwbWUgGuq+7MXi6gw30BsRmgbolnNRWwujqxmwvIcJ9vORWQabSp7VzA6erQeirA0mhW6ylwuG9EugD/oadfmP5sRwRcV/sdoQLCLYm6woW8rgakO/AeegrlloxJlyClq71OUSEX7gu7xYVMGm3QMSokdBUgURanGaX7oigKSmm6vPPr+fZQsuPj8DznfZ+RIT1NNvvT9fsXu7T9gfVHyYmPNe9LCVbNSyXKks2x5f3PeQsd61IAT3p0VUJPl/nNDzhlTRvkQ4SLD95tIpRGEw/sJfndTzitak89l0J41qGrwoE9yvQNRXL12EGMi78a3BJRPU2OjB9xXoFwcVAf7hN1RNIz+2dQI1zwp9EE9TTj+o788tG/Yly8KndLPB1UVMlYi3HxW7m/KuaIpNxfspPdJAeR9xyo19Plmf9bLiyN+Tv/429zkRflSqMJOSLxUeD/er7wUeav3IfFXGwvc+iqWKJs1/S2x5yu0jRN05TuisYfXP6JxfOWA+uF+nCfmJ4uG6zLTdWgilf7+o82qNNoYomygsHOJiSp/ewcIw73iQX2VjWJaPnG7HzT5MKVRgtBlkXGLDdmFgaTroolylJmKupkGDkxmMJ9YoG9/IZ1fc88PRmJfY5UOSLx1dfdWfZXcrI0QoanyD/N2HcIIYQk7ZY4ojSaYGBvz7nocwSb5J6uigb2zlzLomaYJRjTaIKJsqtPYxDJarZgZWZhBCoCe9XjYs/tvFAzXNzSVeFE2Y7bXqiaqYUZLm6F+4QTZVWjM+UW4aMhLtrTaOKJsmrkgumRkv8RnboqXgFf9S8ElpIpLtrSaBIV8AKbHwkXLeG+KRAXbJt/j4OLZl2VKTw5839YxqvCqnTVA67YK/ityPjMqTwadVWqAj4XsCJXpn2zVrdErgKecgQvvsk4o6CirqtyFXupkNsZ02NZnvMlMYwBbEdZyW14IsKVWyJbsVeIbBI0CEA7yjYY4hEg4T75CviroJ0Ke+HXdrtdKOJiAloBf1R9YrxKJZOZdRWioyxTnP95+b+G8VkJF1PIxoj4Kv9zhCVj8f2X33+rIGMA2Wh4Xcl4BD0/KwVMh4WyhQHTL3O9MGo1i1LYVv/2y1zRiQHURlUvcS1SVVwIFMUzuWhgTSINJUrHzVINF2X5Bn1suJCzLZYtxc4FpZRSukqvT5DlnwMrmsra/sIeG324LVI3Phurt7JLQuK8lMHHdg67SSDb65g+7UiTVrOEF++Ax0YEO9yCtapxn9aDGIJ4hXt7D7YDlbnEs/g6U5eSZMCtDB+47TIumKuf49tHLnPlL6DnDt2kzXwcHpcwZwYgF+AGXMa66L8MU4qDixCeC5LseQvjCxR7RAUXhKQFHxmJxPn5hJwLQlLGtbFktdLa8ELQc3GrWbcp5nMS3CCQTok6Lggh8WpX3Fv9/5LyKyEbHNZZVcrFP/shTdOMUtrCy78j4w+/b7YFDmJMtM4GWm5OpVjLRN1nhw9uBbrHfaSnUqRl4poLFUHPof4Jc1RkYVS5UBMM9w3MPskEmqzWiqOdhBBiZNzepuSuabrICRx+qXmrqZnBagV/y8T/m+T9SdVLjcyMpEwE+oq2n1E9dS/VMzRDKxcoglyst+u5uleKIKcPpnlxLHZs2bIljqrfqnUBNs433nMVX50Q9FhdRzuh5g9eBDpz/k2SmadiDDeXcsd5FG4wtI/UlwV4WQ5LGV+KoPeuaVkApFQz3u2/RNFK84XLQmjZwfCUu0AJFxcBYHmn5VxUGycmkFxQ27joQw5E5z4KUZ0XIeh8ypT301DpSA92bimvTY3JvhgDz7PlbUNGZHc2NexKBT5zztr4E56+ggD6Ip6ML4SJyE+NwOeiJ3yx7RyPjPjw832rCcNzAsmcTj2F0NVNyfGvFoh3qvZPIe/XiEv24UEbDKODPjFScu9KzkyGSH5Eo54C6GpSK1yMWXwXU6PGPjFUdK9ErWjt3GRCpUehfKoaTFXd3xXXSwpOu+qx2FSfYnJZ+MruoWksNDoXO0rTNN1Q2lx/UTktFmtN8zsJIfcHV8okCETK8S7FV99c1zt6CqGr/OV4l0bZ/E3XvN92RwTOLeGv7b70yvTNgW4I7MHfy8NLxqXSPGmbD/4JlgGeUuE+PjIqovsqxoVo0wTbwPRAExlX9oe2+wRY9BSmu525a+J05ZLp5YL1ggXJNNqGaWnURiNr5YL94g3JNBrDvTQNraua7qVh1VOwe7fvsNHYxKvpvqIbgT1V97HHWVtp/KmtdkfLPVaf4LnAyYPovIpXtRsEjvnmRvRPx/1mHHoKfm1qkq4+25ZpdvfaO7J44TLD314EnTO+C9/c7lynK66nnb0Wki2w9zOuC2UL7P2Ia2TF9LSb1wvzOiIdvnZaVE+7eB05R2Cv69fUizgicGm0zuhpd3V1IEqFif4rtZg64hh1jAtfgouO6erEkcGwUwenJ8VFp3Q1kKNCPtyHyBFxJbmAG1FnsZ52TldDRx5dcUt8AC46oqtjBwKdCPfNXBAupLvRuqCnHdLVyIGC/eG+PhgXnu3HZ+jAwXZd7QFyYXkabeRAwupwH5SediHcN4SlwuY0WuRAw15d9cG5sNYtmTjwsFVXPQVcWBruC1RQYadbAq2nNuvqQA0VNob7po4q2JdG85VxYZ2ujh11sCyNNvMUcmFZuC9QSYVduhq5SrmwKo02UEuFTboaOqphj1vSU86FNeG+saMelqTRZq4GLixxSwIdVNihq5GjBzaE+3xNXFiQRgsdXcCvq542LtCH+0b6qMCeRtOjp3bo6lAnFbjTaFNHLzAfnz3NXCCO6gS6qXB6WI2M0NEPpFoSuQa4wOmwznqO8yDDKBUIyTBHBToypgapcJweJs9k4jpG4aJJEsyGjnH0cSyN0HMwIDBvdoW+gwTuMHowcREQHptaHFHgOejgBxPdy2M6HiAk4ktkfX2ApuE/fZzcNYqAivkAAAAASUVORK5CYII='
    
    def search(src_dir, out_dir, target_string, chars, file_types):
        '''
        Principal function (void)
        Takes source directory (str), out directory (str), target string (str), number of subsequent characters (int), and file types (dict)
        Iterates through valid files in the source directory, searching for the target string
        Writes to a .tsv file containing four columns: file path, filename, target string, and following x number of characters after the target
        '''

        paths = [] # Create list to contain all file paths within the source directory
        output = [] # Create list to contain extracted data

        # Iterate through source directory and append paths to list
        for subdir, dirs, files in os.walk(src_dir):
            for file in files:
                if file[0] != '~': # Ignore temporary files
                    paths.append(os.path.join(subdir, file))

        def pdf(path):
            '''
            Takes .pdf file path
            Extracts text, collapses excess white space, and searches for target string (case-insensitive)
            Returns EITHER (if found) x number of characters after target string OR (if not found) empty string
            '''

            text = extract_text(path) # Use pdfminer package to extract text from pdf file
            content = text.replace('\n', ' ') # Remove new-line characters
            
            # Substitute multiple consecutive spaces with single space characters 
            content = re.sub(' +', ' ', content)

            # Find index at which the target string ends within the the content string
            index = content.lower().find(target_string.lower()) + len(target_string)

            # Return x characters after target string (if found)
            if index != (-1 + len(target_string)):
                return content[index:index + chars]
            else:
                return ''

        def word(path):
            '''
            Takes .docx file path
            Performs same process as pdf function
            '''

            raw = Document(path) # Instantiate Document class using docx path
            content = ''

            # Iterate through Document paragraphs and append to content string (separated using single space)
            for paragraph in raw.paragraphs:
                content += paragraph.text.replace('\n', ' ') + ' '

            # Substitute multiple consecutive spaces with single space characters 
            content = re.sub(' +', ' ', content)

            # Find index at which the target string ends within the the content string
            index = content.lower().find(target_string.lower()) + len(target_string)

            # Return x characters after target string (if found)
            if index != (-1 + len(target_string)):
                return content[index:index + chars]
            
            return ''

        def excel(path):
            '''
            Takes .xls or .xlsx file path
            Iterates through each cell and performs same function as pdf and word functions
            Note: if target string spans multiple cells, it will not be found
            '''

            # Use pandas to read excel file and convert to array
            data = pd.read_excel(path)
            matrix = data.values

            extracted = []
            
            # Iterate through row, then cells, and append content to 'extracted' list
            for row in matrix:
                output = []

                for cell in row:
                    if not isinstance(cell, float): # Ignores blank cells
                        output.append(re.sub(' +', ' ', str(cell).replace('\n', ' ')))

                extracted.append(output)

            # Iterate through extracted cells and search for target string
            for row in extracted:
                for cell in row:
                    index = cell.lower().find(target_string.lower()) + len(target_string)

                    if index != (-1 + len(target_string)):
                        return cell[index:index + chars]

            # Return empty string if no matching string is found
            return ''

        #Iterate through paths, differentiate between file types, extract text, and write to .tsv file
        for path in paths:

            if os.name == 'nt':
                filename = path.split('\\')[-1]
                
                if path.replace(src_dir, '').replace('\\' + filename, '')[1:]:
                    folder = src_dir.split('/')[-1] + '\\' + path.replace(src_dir, '').replace('\\' + filename, '')[1:]
                else:
                    folder = src_dir.split('/')[-1]
            
            else:
                filename = path.split('/')[-1]

                if path.replace(src_dir, '').replace('/' + filename, '')[1:]:
                    folder = src_dir.split('/')[-1] + '/' + path.replace(src_dir, '').replace('/' + filename, '')[1:]
                else:
                    folder = src_dir.split('/')[-1]

            result = ''

            if file_types['pdf'] and filename[-3:].lower() == 'pdf':
                result = pdf(path)

            if file_types['word'] and filename[-4:].lower() == 'docx':
                result = word(path)

            if file_types['excel'] and (filename[-3].lower() == 'xls' or filename [-4:].lower() == 'xlsx'):
                result = excel(path)

            if result:
                output.append([folder, filename, target_string, result])

        if out_dir != 'tsv':
            out_dir = out_dir.split('.')[0] + '.tsv'

        with open(out_dir, 'w', newline = '') as out:
            
            writer = csv.writer(out, delimiter = '\t')
            writer.writerow(['folder', 'file', 'target string', 'following characters'])
            
            for row in output:
                writer.writerow(row)

    def notAnInteger():
        '''If value entered in 'characters' text input is not an int, display error dialog box (void)'''

        popup = sg.Window(
            title = 'Error', 
            layout = [
                [sg.Text('Characters: Please enter an integer.', pad = (5, 7.5))], 
                [sg.Button('   OK   ', key = 'OK', pad = (5, 7.5))]
            ], 
            margins = (20, 20),
            icon = icon,
            element_justification = 'c',
        )

        while True:
            popup_event, popup_values = popup.read()

            if popup_event == "OK" or popup_event == sg.WIN_CLOSED:
                break

        popup.close()

    def incompleteFields():
        '''If any text input fields are empty, display error dialog box (void)'''

        popup = sg.Window(
            title = 'Error', 
            layout = [
                [sg.Text('Please complete every field.', pad = (5, 7.5))], 
                [sg.Button('   OK   ', key = 'OK', pad = (5, 7.5))]
            ], 
            margins = (20, 20),
            icon = icon,
            element_justification = 'c',
        )

        while True:
            popup_event, popup_values = popup.read()

            if popup_event == "OK" or popup_event == sg.WIN_CLOSED:
                break

        popup.close()

    def complete():
        '''Display dialog box once process is complete (void)'''

        popup = sg.Window(
            title = 'Complete', 
            layout = [
                [sg.Text('Process Complete', pad = (5, 7.5))], 
                [sg.Button('   OK   ', key = 'OK', pad = (5, 7.5))]
            ], 
            margins = (20, 20),
            icon = icon,
            element_justification = 'c',
        )

        while True:
            popup_event, popup_values = popup.read()

            if popup_event == "OK" or popup_event == sg.WIN_CLOSED:
                break

        popup.close()

    def fatalError():
        '''If unknown error occurs, display fatal error dialog box'''

        popup = sg.Window(
            title = 'Error', 
            layout = [
                [sg.Text('Fatal Error', pad = (5, 7.5))], 
                [sg.Text('Please try again.', pad = (5, 7.5))], 
                [sg.Button('   OK   ', key = 'OK', pad = (5, 7.5))]
            ], 
            margins = (20, 20),
            icon = icon,
            element_justification = 'c',
        )

        while True:
            popup_event, popup_values = popup.read()

            if popup_event == "OK" or popup_event == sg.WIN_CLOSED:
                break

        popup.close()

    # Specify layout for GUI
    layout = [
        [
            sg.Text("Select your source folder:", pad = (0, 0)), 
            sg.Input(size = (43, 1), pad = (5, 15), key = 'SOURCE'), 
            sg.FolderBrowse('  Select...  ')
        ],
        [sg.HSeparator()],
        [
            sg.Text("Target string: ", pad = (0, 0)), 
            sg.InputText(size = (41, 1), key = 'TARGET'), 
            sg.VSeparator(pad = (15, 12)), 
            sg.Text("Characters: ", pad = (5, 15)), 
            sg.InputText(size = (5, 1), key = 'CHARS', justification = 'c')
        ],
        [
            sg.Text("Files: ", pad = (5, 15)),
            sg.Checkbox("Adobe PDF (.pdf)", key = 'PDF'),
            sg.Checkbox("Excel Workbook (.xls .xlsx)", key = 'EXCEL'), 
            sg.Checkbox("Word Document (.docx)", key = 'WORD'), 

        ],
        [sg.HSeparator()],
        [
            sg.Text("Select your output destination:", pad = (0, 0)), 
            sg.Input(size = (39, 1), pad = (5, 15), key = 'OUT'), 
            sg.FileSaveAs('  Select...  ', file_types = [('Tab-separated values', '.tsv')])
        ],
        [sg.Button('     Run     ', pad = (5, 15), key = 'RUN')]
    ]

    # Construct principal GUI window
    window = sg.Window(
        title = 'String Extractor', 
        layout = layout, 
        margins = (30, 30),
        icon = icon,
        element_justification = 'c',
    )

    # Create event loop for GUI
    while True:
        event, values = window.read()

        if event == sg.WIN_CLOSED:
            break
        
        if event == 'RUN': 
            if values['SOURCE'] and values['TARGET'] and values['CHARS'] and values['OUT']:
                try:
                    search(
                        values['SOURCE'],
                        values['OUT'],
                        values['TARGET'], 
                        int(values['CHARS']),
                        {'pdf': values['PDF'], 'excel': values['EXCEL'], 'word': values['WORD']}
                    )

                    complete()

                except ValueError:
                    notAnInteger()

                except:
                    fatalError()
            
            else:
                incompleteFields()
                
    window.close()

if __name__ == "__main__":
    main() 
