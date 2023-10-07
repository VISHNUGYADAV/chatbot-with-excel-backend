from tkinter import *
import openpyxl
import urllib.request
from PIL import Image

# GUI
root = Tk()
root.title("Chatbot")

BG_GRAY = "#ABB2B9"
BG_COLOR = "#17202A"
TEXT_COLOR = "#EAECEE"

FONT = "Helvetica 14"
FONT_BOLD = "Helvetica 13 bold"

path = "GymExercisesDataset.xlsx"
wb_obj = openpyxl.load_workbook(path)
sheet = wb_obj.active


muscles = ["abdominals", "abductors", "adductors", "biceps", "calves", "chest", "forearms", "glutes", "hamstrings",
           "lats", "lower back", "middle back", "neck", "quadriceps", "quads", "shoulders", "traps", "triceps"]

# Send function


def send():
    send = "You -> " + e.get()
    txt.insert(END, "\n" + send)

    user = e.get().lower().split()

    muscle = ""

    i = 0
    while (i < len(user)):
        if user[i] == "ab" or user[i] == "abs":
            muscle = "abdominals"
        if user[i] == "abdominals":
            muscle = "abdominals"
        if user[i] == "abductor":
            muscle = "abductors"
        if user[i] == "abductors":
            muscle = "abductors"
        if user[i] == "adductor":
            muscle = "adductors"
        if user[i] == "adductors":
            muscle = "adductors"
        if user[i] == "calve":
            muscle = "calves"
        if user[i] == "calves":
            muscle = "calves"
        if user[i] == "chest":
            muscle = "chest"
        if user[i] == "lower":
            if user[i+1] == "back":
                muscle = "lower back"
                i = i+1
        if user[i] == "middle":
            if user[i+1] == "back":
                muscle = "middle back"
                i = i+1
        if user[i] == "forearm":
            muscle = "forearms"
        if user[i] == "forearms":
            muscle = "forearms"
        if user[i] == "glute":
            muscle = "glutes"
        if user[i] == "glutes":
            muscle = "glutes"
        if user[i] == "hamstring":
            muscle = "hamstrings"
        if user[i] == "hamstrings":
            muscle = "hamstrings"
        if user[i] == "lat":
            muscle = "lats"
        if user[i] == "lats":
            muscle = "lats"
        if user[i] == "quad" or user[i] == "quads" or user[i] == "quadricep":
            muscle = "quadriceps"
        if user[i] == "quadriceps":
            muscle = "quadriceps"
        if user[i] == "shoulder":
            muscle = "shoulders"
        if user[i] == "shoulders":
            muscle = "shoulders"
        if user[i] == "trap":
            muscle = "traps"
        if user[i] == "traps":
            muscle = "traps"
        if user[i] == "tricep":
            muscle = "triceps"
        if user[i] == "triceps":
            muscle = "triceps"
        if user[i] == "biceps":
            muscle = "biceps"   
        if user[i] == "bicep":
            muscle = "biceps"   

        i = i+1

    # txt.insert(END, "\n" + "Bot ->" + muscle)

    for i in muscles:
        if i == muscle:
            # txt.insert(END, "\n" + "Bot -> Present")
            for row in sheet:
                if (row[5].value.lower() == i):
                    txt.insert(END, "\n" + "Bot -> " + str(row[0].value))
                    if (row[2].value != ""):
                        urllib.request.urlretrieve(
                            str(row[2].value), "exercise.jpg")
                        img = Image.open("exercise.jpg")
                        img.show()
                    break

    # for i in muscles:
    # 	if i in user:
    # 		# txt.insert(END, "\n" + "Bot -> Present")
    # 		n=1
    # for row in sheet:
    # 	# if (row[5].value == i):
    # 	if (row[5].value.lower()=="forearms"):
    # 		txt.insert(END, "\n" + "Bot ->" + str(row[0].value))
    # 		break

    # if (user == "hello"):
    # 	txt.insert(END, "\n" + "Bot -> Hi there, how can I help?")

    # elif (user == "hi" or user == "hii" or user == "hiiii"):
    # 	txt.insert(END, "\n" + "Bot -> Hi there, what can I do for you?")

    # elif (user == "how are you"):
    # 	txt.insert(END, "\n" + "Bot -> fine! and you")

    # elif (user == "fine" or user == "i am good" or user == "i am doing good"):
    # 	txt.insert(END, "\n" + "Bot -> Great! how can I help you.")

    # elif (user == "thanks" or user == "thank you" or user == "now its my time"):
    # 	txt.insert(END, "\n" + "Bot -> My pleasure !")

    # elif (user == "what do you sell" or user == "what kinds of items are there" or user == "have you something"):
    # 	txt.insert(END, "\n" + "Bot -> We have coffee and tea")

    # elif (user == "tell me a joke" or user == "tell me something funny" or user == "crack a funny line"):
    # 	txt.insert(
    # 		END, "\n" + "Bot -> What did the buffalo say when his son left for college? Bison.! ")

    # elif (user == "goodbye" or user == "see you later" or user == "see yaa"):
    # 	txt.insert(END, "\n" + "Bot -> Have a nice day!")

    # else:
    # 	txt.insert(END, "\n" + "Bot -> Sorry! I didn't understand that")

    e.delete(0, END)


lable1 = Label(root, bg=BG_COLOR, fg=TEXT_COLOR, text="Welcome", font=FONT_BOLD, pady=10, width=20, height=1).grid(
    row=0)

txt = Text(root, bg=BG_COLOR, fg=TEXT_COLOR, font=FONT, width=60)
txt.grid(row=1, column=0, columnspan=2)

scrollbar = Scrollbar(txt)
scrollbar.place(relheight=1, relx=0.974)

e = Entry(root, bg="#2C3E50", fg=TEXT_COLOR, font=FONT, width=55)
e.grid(row=2, column=0)

send = Button(root, text="Send", font=FONT_BOLD, bg=BG_GRAY,
              command=send).grid(row=2, column=1)

# root.bind('<Return>', send)

root.mainloop()
