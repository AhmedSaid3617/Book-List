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

    def rewrite(self, sheet, edit_order, new_title, new_author, new_year):
        sheet[f'A{edit_order}'] = new_title
        sheet[f'B{edit_order}'] = new_author
        sheet[f'C{edit_order}'] = new_year

    def edit_title(self, sheet, edit_order, new_title):
        sheet[f'A{edit_order}'] = new_title

    def edit_author(self, sheet, edit_order, new_author):
        sheet[f'B{edit_order}'] = new_author

    def edit_year(self, sheet, edit_order, new_year):
        sheet[f'C{edit_order}'] = new_year


