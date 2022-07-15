from Books import Book
import openpyxl

filename = "Records.xlsx"
wb = openpyxl.load_workbook(filename)
books_sheet = wb.active


"""def record(data):
    books_sheet[f'A{data.id}'] = data.title
    books_sheet[f'B{data.id}'] = data.author
    books_sheet[f'C{data.id}'] = data.year"""


def create_book():
    max_row = books_sheet.max_row
    title = input("Enter book title:")
    author = input("Enter book author:")
    year = input("Enter book year:")
    new_book = Book(max_row+1, title, author, year)
    new_book.record(books_sheet)
    print("Book recorded successfully.")


def rundown():
    count = 1
    for row in books_sheet.rows:
        print(f"{count}  Book: {row[0].value}, Author: {row[1].value}, Year: {row[2].value}")
        count += 1


rundown()

print('\n')
print("Commands:")
print("\"new\": to record a new book.")
print("\"delete {number of book}\": to delete a new book.")
print("\"rundown\": to displays a list of all books.")
print("\"close\": to close application.")

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
            books_sheet.delete_rows(id)
            wb.save(filename)
            print("Deleted successfully")

    elif userInput == "rundown":
        rundown()

    else:
        print("Unkown command.")