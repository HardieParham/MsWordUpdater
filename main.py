# Standard Imports
import logging
import os

# External Imports
import docx

# Local Imports
from data.update_list import data



def match_loop(i, data):
    match = False

    # Cycle through all key,value pairs of data
    for key, value in data.items():

        # Setting the text in the document to a variable
        textline = doc.paragraphs[int(i)].text
        print(textline)

        # Searching that variable for the given key, if match returns the index where it occurs
        start_index = textline.find(key)
        end_index = start_index + len(key)

        # If there is a key match
        if start_index != -1:

            # If the match is at the start of the textline
            if start_index == 0:
                logging.warning(f'Match at paragraph {i} for {key} at the start')
                match = True

                splits = textline.split(key)
                new_line = value + splits[1]
                doc.paragraphs[i].text = new_line

            # If the match is at the end of the textline
            elif end_index == len(textline):
                logging.warning(f'Match at paragraph {i} for {key} at the end')
                match = True

                splits = textline.split(key)
                new_line = splits[0] + value
                doc.paragraphs[i].text = new_line

            # If the match is somewhere in the middle
            else:
                logging.warning(f'Match at paragraph {i} for {key} in the middle')
                match = True

                splits = textline.split(key)
                new_line = splits[0] + value + splits[1]
                doc.paragraphs[i].text = new_line

    # If no match for the given paragraph
    if match == False:
        print('\n # NO MATCH # \n')




class Text():
    def __init__(self):
        # Future TODO, setup the ability for text formatting memorization.
        # As of now, any updated text reformats text to document default
        self.alignment = ''
        self.font = ''
        self.font_size = ''
        self.content = ''


    def text_check(phrase, dic):
        # If phrase is blank, skip the check
        if phrase == None or phrase == '':
            pass

        else:
            # Cycle through all key,value pairs of data
            for key, value in dic.items():

                # Searching the phrase for the given key, returns the index where the key starts and ends. If no match, start_index is returned as -1
                start_index = phrase.find(key)
                end_index = start_index + len(key)

                # If a match was found, update key, value pair
                if start_index != -1:
                    print(f'MATCH! Changing {key} to {value}')

                    # If the match is at the start of the textline
                    if start_index == 0:
                        splits = phrase.split(key)
                        phrase = value + splits[1]

                    # If the match is at the end of the textline
                    elif end_index == len(phrase):
                        splits = phrase.split(key)
                        phrase = splits[0] + value

                    # If the match is somewhere in the middle
                    else:
                        splits = phrase.split(key)
                        phrase = splits[0] + value + splits[1]
                    
            # Returns back new updated phrase
            return phrase









class Word():
    def __init__(self, file):
        self.old_name = file
        self.new_name = Text.text_check(phrase = self.old_name, dic = data) # Updating file name, if needed
        self.data = data
        self.doc = self.load_doc()


    # Method to load the word document to extract its information
    def load_doc(self):
        self.log_change(content=f'starting transition from {self.old_name} to {self.new_name}')
        return docx.Document(f'input/{self.old_name}')


    # Method to save word document after changes have been made
    def save_doc(self):
        return self.doc.save(f'output/{self.new_name}')


    # Method to update log.txt if any changes were made (useful if errors after script was run)
    def log_change(self, content):
        with open('data/log.txt', 'a') as f:
            f.write(f'\n {content}')


    # Loop through the document's Headers
    def header_loop(self):
        #self.doc.sections[0].header.tables[0]._cells
        i = 0
        # Loop through the tables on the header
        while i < len(self.doc.sections[0].header.tables):
            print(f'\n Checking Header Table {i} \n')
            j = 0
            # Loop through the cells of the table
            while j < len(self.doc.sections[0].header.tables[i]._cells):
                text_line = self.doc.sections[0].header.tables[i]._cells[j].text
                if text_line == '' or text_line == None:
                    pass
                else:
                    new_textline = Text.text_check(phrase=text_line, dic=self.data)
                    if new_textline != text_line:
                        self.doc.sections[0].header.tables[i]._cells[j].text = new_textline
                        self.log_change(content=f'{new_textline} /// on header table {i}, cell {j}')
                j += 1
            i += 1


    # Loop through the document's Paragraphs
    def paragraph_loop(self):
        i = 0
        while i < len(self.doc.paragraphs):
            print(f'\n Checking Paragraph {i} \n')
            text_line = self.doc.paragraphs[int(i)].text
            new_textline = Text.text_check(phrase=text_line, dic=self.data)
            if new_textline != text_line:
                self.doc.paragraphs[i].text = new_textline
                self.log_change(content=f'{new_textline} on paragraph {i}')
            i += 1


    # Loop through the document's Tables
    def table_loop(self):
        i = 0
        # Loop through the tables on the doc
        while i < len(self.doc.tables):
            #if i == 6:
                #test = 'test'
                #print('test')
            print(f'\n Checking Table {i} \n')
            j = 0
            # Loop through the cells of the table
            while j < len(self.doc.tables[i]._cells):
                text_line = self.doc.tables[i]._cells[j].text
                print(text_line)
                if text_line == '' or text_line == None:
                    pass
                else:
                    new_textline = Text.text_check(phrase=text_line, dic=self.data)
                    print(new_textline)
                    if new_textline != text_line:
                        try:
                            self.doc.tables[i]._cells[j].text = new_textline
                            self.log_change(content=f'{new_textline} on body table {i}, cell {j}')
                        except:
                            logging.error('Table could not be looped thru!')
                            self.log_change(content=f' E R R O R! on body table {i}, cell {j}')
                j += 1
            i += 1


    # Loop through the document's Footers
    def footer_loop(self):
        i = 0
        # Loop through the tables on the header
        while i < len(self.doc.sections[0].footer.tables):
            print(f'\n Checking Footer Table {i} \n')
            j = 0
            # Loop through the cells of the table
            while j < len(self.doc.sections[0].footer.tables[i]._cells):
                text_line = self.doc.sections[0].footer.tables[i]._cells[j].text
                if text_line == '' or text_line == None:
                    pass
                else:
                    new_textline = Text.text_check(phrase=text_line, dic=self.data)
                    if new_textline != text_line:
                        self.doc.sections[0].footer.tables[i]._cells[j].text = new_textline
                        self.log_change(content=f'{new_textline} on footer table {i}, cell {j}')
                j += 1
            i += 1


    def loop_thru_document(self):
        self.header_loop()
        self.paragraph_loop()
        self.table_loop()
        self.footer_loop()




class App():
    def __init__(self):
        path = './input'
        self.name_list = [f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))]
        print(self.name_list)


    def delete_old_log(self):
        if os.path.exists('log.txt'):
            os.remove('log.txt')
        else:
            print('Log file not found')


    def main_loop(self):
        self.delete_old_log()
        for file in self.name_list:
            print(f'\n Loading {file} \n')
            word = Word(file)
            word.loop_thru_document()
            #test = word.doc.sections[0].header.tables
            #print(test)
            word.save_doc()




if __name__ == '__main__':
    app = App()
    app.main_loop()
    #for name in app.name_list:
        #print(name)