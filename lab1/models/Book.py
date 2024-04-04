# Class for view books
class Book(object):
    def __init__(self, category, title, authors, year, price):
        self.category = category
        self.title = title
        self.authors = authors
        self.year = year
        self.price = price

    def to_list(self):
        return [self.category, self.title, ', '.join(self.authors), self.year, self.price]
