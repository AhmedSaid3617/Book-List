from Books import Book
import openpyxl

filename = "Records.xlsx"
wb = openpyxl.load_workbook(filename)
sheet = wb.active

def record(data):
    sheet[f'A{data.id}'] = data.title
    sheet[f'B{data.id}'] = data.author
    sheet[f'C{data.id}'] = data.year


def create_book():
    max_row = sheet.max_row
    title = input("Enter book title:")
    author = input("Enter book author:")
    year = input("Enter book year:")
    new_book = Book(max_row+1, title, author, year)
    record(new_book)
    print("Book recorded successfully.")


while True:
    userInput = input()

    if userInput == "new":
            create_book()
            wb.save(filename)

    elif userInput == "close":
        break

    elif userInput.startswith("delete"):
        subject = userInput.replace("delete ", "")
        if subject.isnumeric():
            id = int(subject)
            sheet.delete_rows(id)
            wb.save(filename)
            print("Deleted successfully")
    else:
        print("Unkown command.")



print("done")