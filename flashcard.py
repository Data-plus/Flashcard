import win32com.client
import random

# Reads data from excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('C:\\Users\\username\\Desktop\\test.xlsx') # Location for the dataset  
ws = wb.ActiveSheet

correct = 0
wrong = 0

class Flash_card:
    def __init__(self, tot_question):
        self.tot_question = tot_question

    def print_question(self, tot_question):
        number = random.randrange(2,tot_question+1)
        q_number = ws.Cells(number, 1)  # returns Q number
        question = ws.Cells(number, 2).Value  # Question1, column 2 = Question
        answer = ws.Cells(number, 3).Value  # Answer1, column 3 = Answer
        print(int(q_number), question)
        user_answer = input("Your Answer: ")

        global correct
        global wrong

        if user_answer == answer:
            print("That's Correct!")
            correct += 1 # for the correct input, adds count +1 to correct
            print("Total Correct answers:", correct,'\n')

        elif user_answer == "q": # Quits the program when user inputs "q"
            print("Bye Bye")
            print("Final Percentage: {:.0%}".format(correct / (correct+wrong))) # Calculates / Prints final percentage
            exit()

        else:
            print("That's not correct.")
            wrong += 1 # Adds count to Wrong answers
            print("Total wrong answers:", wrong,'\n')

    def print_menu(self):
        print()
        print("---------------------------------")
        print("Welcome! choose your option below")
        print("1. Start")
        print("2. Finish")
        print("---------------------------------")
        print()
        menu = input("Your option: ")
        return int(menu)

    def run(self, tot_question):
        self.tot_question = tot_question
        menu = Flash_card.print_menu(self)
        if menu == 1:
            while True:
                print("Type 'q' to exit")
                user_answer = Flash_card.print_question(self, tot_question)
                if user_answer == "q":
                    break
                else:
                    Flash_card.print_question(self, tot_question)
        elif menu == 2:
            print("Exit")


if __name__ == '__main__':
    flash = Flash_card(10) # 10 numbers of qs
    flash.run(10) # 10 numbers of qs
