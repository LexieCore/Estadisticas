import sys
from termcolor import colored
usage = "You must type : python %s filename.xlsx"
if len(sys.argv) != 2:
    print >> sys.stderr, \
    colored(usage % sys.argv[0], "yellow")
    sys.exit(1)

import pyexcel as excel
import pyexcel.ext.xls
import pickle
from collections import OrderedDict
from pyexcel_xlsx import save_data




class Stats:
    """Produce stats per user """
    def __init__(self):
        self.restore = pickle.load(open('results.p', 'rb'))

    def get_users(self):
        '''Abre un archivo .xlsx y obtiene las dos columnas (usuario y fecha) obtiene los datos de before during y after del evento'''
        data = excel.get_records(file_name=sys.argv[1])
        users = {}
        for record in data:
            event = record['Date']
            usname = record['Screenname']
            if not users.has_key(usname):
                users.update({usname:event})
        return users


    def create_table(self,users,record):
        '''Crea un archivos .xlsx que contienen estadisticas (mediana y promedio) por usuario'''
        data = OrderedDict()
        spread = [["User", "Before", "During", "After"]]
        for user in users:
            try:
                b = self.restore[user]["%s_before"%record]
                d = self.restore[user]["%s_during"%record]
                a = self.restore[user]["%s_after"%record]
                spread.append([user, b, d,a])
            except KeyError:
                pass
        spread.append(["Average", "=AVERAGE(B2:B297)", "=AVERAGE(C2:C297)","=AVERAGE(D2:D297)"])
        spread.append(["Median", "=MEDIAN(B2:B297)", "=MEDIAN(C2:C297)","=MEDIAN(D2:D297)"])
        data.update({"Sheet 1": spread})

        save_data("results/%s.xlsx"%record, data)

    def print_data(self,user):
        '''Imprime los datos por cada usuario'''
        print colored(user,"cyan")

        for record in records:
            try:
                print colored(record,"magenta")
                print self.restore[user]["%s_before"%record]
                print self.restore[user]["%s_during"%record]
                print self.restore[user]["%s_after"%record]
            except KeyError:
                pass

if __name__ == '__main__':
    records = ["hashtags","mentions","replies","retweets","status"]
    s = Stats()
    for record in records:
        users = s.get_users()
        s.create_table(users,record)
