from Books import Book
import openpyxl

filename = "Records.xlsx"
wb = openpyxl.load_workbook(filename)
sheet = wb.active
maxRow = sheet.max_row

def Record(data):
    sheet[f'A{data.id}'] = data.title
    sheet[f'B{data.id}'] = data.author
    sheet[f'C{data.id}'] = data.year

another = 0
while another != "no":
    title = input("Enter book title:")
    author = input("Enter book author:")
    year = input("Enter book year:")
    newBook = Book(maxRow+1, title, author, year)
    Record(newBook)
    another = input("Type \"no\" to stop.")


wb.save(filename)
