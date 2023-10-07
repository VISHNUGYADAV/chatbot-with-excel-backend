import openpyxl
import nltk

nltk.download('punkt')
nltk.download('stopwords')
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import string

# Load Excel workbook
workbook = openpyxl.load_workbook('D:\josephs java lav\RAW Fitness\chatbot in excel\GymExercisesDataset.xlsx')
sheet = workbook.active

# Preprocess user input
def preprocess_input(user_input):
    user_input = user_input.lower()
    tokens = word_tokenize(user_input)
    tokens = [token for token in tokens if token not in stopwords.words('english') and token not in string.punctuation]
    return tokens

# Process user command
def process_command(tokens):
    if "read" in tokens and "cell" in tokens:
        row = int(tokens[tokens.index("row") + 1])
        col = tokens[tokens.index("column") + 1].upper()
        cell_value = sheet[f"{col}{row}"].value
        return f"The value in cell {col}{row} is: {cell_value}"
    elif "write" in tokens and "cell" in tokens:
        row = int(tokens[tokens.index("row") + 1])
        col = tokens[tokens.index("column") + 1].upper()
        value = ' '.join(tokens[tokens.index("as") + 1:])
        sheet[f"{col}{row}"].value = value
        workbook.save('example.xlsx')
        return f"Value '{value}' written to cell {col}{row}"
    else:
        return "I'm sorry, I don't understand the command."

# Main chat loop
def chat():
    print("Chatbot: Hello! I'm here to help you with Excel tasks. What would you like to do?")
    while True:
        user_input = input("You: ")
        if user_input.lower() == "exit":
            print("Chatbot: Goodbye!")
            break
        else:
            tokens = preprocess_input(user_input)
            response = process_command(tokens)
            print("Chatbot:", response)

if __name__ == "__main__":
    chat()
