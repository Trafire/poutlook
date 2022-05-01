import os

import win32com.client


def ensure_dir(file_path):

    if not os.path.exists(file_path):
        os.makedirs(file_path)


def save_attachments(base_directory, search_function, filetypes, savebydate=True, overwriteexistig=False):
    if type(search_function == str):
        search_function = search_text(search_function)
    if type(filetypes) == str:
        filetypes = [filetypes]
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # "6" is the number for your inbox
    messages = inbox.Items
    messages.Sort("[ReceivedTime]", True)
    directory = base_directory
    for m in list(filter(search_function, messages)):
        if savebydate:
            directory = os.path.join(base_directory, str(m.ReceivedTime.year), m.ReceivedTime.strftime("%B"),
                                     str(m.ReceivedTime.day))
        for att in m.Attachments:
            for f in filetypes:
                if att.Filename.endswith(f):
                    filename = os.path.join(directory, att.Filename)
                    ensure_dir(directory)
                    att.SaveAsFile(filename)


def search_text(text):
    return lambda x: text.lower() in x.Subject.lower()

import datetime
if __name__ == '__main__':
    start = datetime.datetime.now()
    directory = 'D:/reports'
    savebydate = False
    overwriteexisting = False
    filetypes = ['.pdf']
    search_phrase = 'Freedom'

    save_attachments(directory, search_phrase, filetypes, savebydate=True, overwriteexistig=False)
    end = datetime.datetime.now()
    print(end-start)
