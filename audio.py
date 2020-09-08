import os
import PyPDF2
import docx2txt
import pyttsx3
import tkinter as tk
from tkinter import filedialog

root=tk.Tk()
root.withdraw()

file_name=filedialog.askopenfilename()

if os.path.splitext(file_name)[1] == '.docx':
    text=docx2txt.process(file_name)
    speaker = pyttsx3.init()
    voices = speaker.getProperty('voices')
    speaker.setProperty('voice',voices[1].id)
    speaker.setProperty('rate',150)
    speaker.say(text)
    speaker.runAndWait()
elif os.path.splitext(file_name)[1] == '.txt' :
    fh=open(file_name)
    readtext= fh.read()
    speaker = pyttsx3.init()
    voices = speaker.getProperty('voices')
    speaker.setProperty('voice',voices[1].id)
    speaker.setProperty('rate',150)
    speaker.say(readtext)
    speaker.runAndWait()
elif os.path.splitext(file_name)[1] == '.pdf' :
    pdfreader=PyPDF2.PdfFileReader(file_name)
    pages=pdfreader.numPages

    for num in range(0, pages):
        page=pdfreader.getPage(num)
        text = page.extractText()
        speaker = pyttsx3.init()
        voices = speaker.getProperty('voices')
        speaker.setProperty('voice',voices[1].id)
        speaker.setProperty('rate',175)
        speaker.say(text)
        speaker.runAndWait()
else:
    print("error in file selection")
