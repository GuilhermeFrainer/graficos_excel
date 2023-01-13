import datetime
from sys import exit


# Class to handle caged data and ensure a value is connected to its date
class Data:
    def __init__(self, date, value):
        self.date = date
        self.value = value

    @property
    def date(self):
        return self._date

    # Expects either date already
    @date.setter
    def date(self, date):
        
        if isinstance(date, datetime.datetime):
            self._date = date

        else:
            months = {'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8, 'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12}
            
            month, year = date.split("/")
            month = months[month]
            date = datetime.date.fromisoformat(f'{year}-{month:02d}-01')
            self._date = date

    @property
    def value(self):
        return self._value

    # Not checking anything for now cause it doesn't seem necessary
    @value.setter
    def value(self, value):
        self._value = value