import tkinter as tk
import ask
import os

question, index = None,None
filePath= None

def browse():
    global filePath
    try:
        filePath = tk.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        ask.path(filePath)
    except Exception as e:
        print("No Files Given. Loading From Default.")
        try:
            if os.path.exists("D:\Programming\Projects\Question_Assistance\Default.xlsx"):
                filePath="D:\Programming\Projects\Question_Assistance\Default.xlsx"
                print("Default File Loaded.")
                ask.path(filePath)
            else:
                raise FileNotFoundError(f"The file does not exist.")
        except Exception as e:
            print("Fail Loading File")
    file_text.config(text=f"Asking from File: {fileN()}")
        

def fileN():
    if filePath:
            file_name = os.path.basename(filePath)  
            return file_name

def text(show):
    show_text.config(text=show)

def show_ans():
    if filePath:
        text(ask.ans(index))

def reset():
    if filePath:
        reset_value = Reset_var.get()
        if reset_value == "RESET":
            ask.reset()
            text("Reset Completed")
        else:
            text("Invalid Text")


def ques():
    global question,index
    if filePath:
        start_value = int(start_var.get())
        end_value = int(end_var.get())
        update_value = update.get()
        question,index = ask.askq(focusRange=[start_value,end_value])
        
        if update_value:
            ask.updateAsked(index)
            
        custom_text.config(text=question)
        text("Question's Up!")

        TA_text.config(text=f"Total Asked Till Date: {sum(ask.question_asked_times)}")

root = tk.Tk()
root.title("Custom Text and Options")
root.geometry("{0}x{1}+0+0".format(root.winfo_screenwidth(), root.winfo_screenheight()))

text_frame = tk.Frame(root)
custom_text = tk.Label(text_frame, text="Question Region", font=("Helvetica", 16))
custom_text.pack(padx=10, pady=10)
text_frame.pack(side=tk.LEFT, padx=20, pady=20)

options_frame = tk.Frame(root)
options_label = tk.Label(options_frame, text="Options:", font=("Helvetica", 14))
options_label.pack()

file_text = tk.Label(options_frame, text="<< Select File >>", font=("Helvetica", 10))
file_text.pack(padx=10, pady=10)

start_label = tk.Label(options_frame, text="Focus start index:")
start_label.pack()
start_var = tk.StringVar(value=0)
start_entry = tk.Entry(options_frame, textvariable=start_var)
start_entry.pack()

end_label = tk.Label(options_frame, text="Focus end index:")
end_label.pack()
end_var = tk.StringVar(value=-1)
end_entry = tk.Entry(options_frame, textvariable=end_var)
end_entry.pack()

update = tk.BooleanVar(value=True)
update_checkbox = tk.Checkbutton(options_frame, text="Update", variable=update)
update_checkbox.pack()

ask_button = tk.Button(options_frame, text="Ask", command=ques)
ask_button.pack(pady=10)

if filePath:
    openLink_button = tk.Button(options_frame, text="Open Link", command=lambda: ask.openLink(index))
    openLink_button.pack(pady=10)

Ans_button = tk.Button(options_frame, text="Answer", command=show_ans)
Ans_button.pack(pady=10)

Reset_label = tk.Label(options_frame, text="Reset")
Reset_label.pack()
Reset_var = tk.StringVar(value="Type RESET")
Reset_entry = tk.Entry(options_frame, textvariable=Reset_var)
Reset_entry.pack()

reset_button = tk.Button(options_frame, text="Reset", command=reset)
reset_button.pack(pady=10)

show_text = tk.Label(options_frame, text="Console", font=("Helvetica", 10))
show_text.pack(padx=10, pady=10)

if fileN()=="Lib.xlsx":
    text("File Not Selected. Loaded From Default.")
if fileN()== None:
    text("File Not Found")

TA_text = tk.Label(options_frame, font=("Helvetica", 10))
TA_text.pack(padx=10, pady=10)

reset_button = tk.Button(options_frame, text="Open File", command=browse)
reset_button.pack(pady=10)

options_frame.pack(side=tk.RIGHT, padx=20, pady=20)


root.mainloop()
