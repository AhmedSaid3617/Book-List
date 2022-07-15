class Book:
    def __init__(self, order, title, author, year):
        self.order = order
        self.title = title
        self.author = author
        self.year = year

    def record(self, sheet):
        sheet[f'A{self.order}'] = self.title
        sheet[f'B{self.order}'] = self.author
        sheet[f'C{self.order}'] = self.year
