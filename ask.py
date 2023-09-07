import pandas as pd
import random
import webbrowser
from tkinter import filedialog
import os

ask_index = None
question_asked_times = None
link=None
answer= None
filePath=None
file = None

def path(p):
    global file, filePath
    filePath = p
    file = pd.read_excel(filePath)

def askq(focusRange=[0,-1]):
    global file,link,question_asked_times,answer
    question_list = file["Question"].tolist()[focusRange[0]:focusRange[1]]
    question_asked_times = file["Asked_Times"].tolist()[focusRange[0]:focusRange[1]]
    link = file["Link"].tolist()[focusRange[0]:focusRange[1]]
    answer = file["Answer"].tolist()[focusRange[0]:focusRange[1]]

    total_asked = sum(question_asked_times)

    if total_asked==0:
        least_asked = [1]*len(question_list)
    else:
        least_asked = [total_asked - i for i in question_asked_times]

    ask = random.choices(question_list, weights=least_asked, k=1)
    ask_index = question_list.index(ask[0])

    return ask[0], ask_index

def ans(index):
    global answer
    return answer[index]

def updateAsked(rowNo):
    global file,question_asked_times
    file.iat[rowNo, -1] = question_asked_times[rowNo]+1
    file.to_excel(filePath, index=False)

def openLink(ask_index):
    webbrowser.open(link[ask_index])

def reset():
    file.iloc[:,-1 ] = 0
    file.to_excel(filePath, index=False) 