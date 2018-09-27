from openpyxl import load_workbook

class BirthDate():
        def __init__(self, year, month=0, number=0):
                self._year = year
                self._month = moth
                self._number = number

        @property
        def year(self):
                return self._year

        @property
        def month(self):
                return self._month


        @property
        def number(self):
                return self._number

class DistStrategy():
        def __init__(self):
                #TODO FIGURE OUT
                pass

        def __call__(self, cls, row):
                print("SEPARATOR")
                cls.id
                cls.name
                cls.surname
                cls.middle
                cls.patronymic
                number
                month
                year
                cls.date = BirthDate(year, month, number)
                gender
                for cell in row:
                        print(cell.value)


class Person():
        def __init__(self, row, strategy = None):
                if strategy:
                        strategy(self, row);

                self._id = row[0].value
                self._name = row[1].value
                self._surname = row[2].value
                self._school = None if len(row) < 4 else row[3].value

        def __eq__(self, lhc):
                if self.name == lhc.name and self.surname == lhc.surname:
                        self._school = lhc.school
                        return True
                else:
                        return False

        def __str__(self):
                return("{} {} from {}".format(self._name, self._surname, self._school))

        @property
        def name(self):
                return self._name

        @property
        def surname(self):
                return self._surname

        @property
        def school(self):
                   return self._school

        @school.setter
        def school(self, value):
                self._school = value if not self._school else self._school

        @property
        def dump(self):
                return [self._id, self._name, self._surname, self._school]



class Students(list):
        def __init__(self, book, strategy=None):
                self._book = book
                for row in book.rows:
                        self.append(Person(row, strategy))

        def sheet(self, wb, ws='dummy'):
                internal_sheet = wb.create_sheet(ws, 0)
                for row in self:
                        internal_sheet.append(row.dump)
                return internal_sheet


def entry_point():
        wb = load_workbook('./Exel/Syxiv_2000.xlsx')
        sheets = wb.sheetnames
        town_by_year = []
        for sh in sheets:
                town_by_year.append(Students(wb[sh],DistStrategy()))


if __name__ == '__main__':
        entry_point()
