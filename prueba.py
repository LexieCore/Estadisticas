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


restore = pickle.load(open('results.p', 'rb'))




def get_users():
    '''Abre un archivo .xlsx y obtiene las dos columnas (usuario y fecha) obtiene los datos de before during y after del evento'''
    data = excel.get_records(file_name=sys.argv[1])
    users = {}
    for record in data:
        event = record['Date']
        usname = record['Screenname']
        if not users.has_key(usname):
            users.update({usname:event})
    return users


def create_table(users,record):
    data = OrderedDict()
    spread = [["User", "Before", "During", "After"]]
    for user in users:
        try:
            b = restore[user]["%s_before"%record]
            d = restore[user]["%s_during"%record]
            a = restore[user]["%s_after"%record]
            spread.append([user, b, d,a])
        except KeyError:
            pass
    data.update({"Sheet 1": spread})

    save_data("results/%s.xlsx"%record, data)

def print_data(user):
    print colored(user,"cyan")

    for record in records:
        try:
            print colored(record,"magenta")
            print restore[user]["%s_before"%record]
            print restore[user]["%s_during"%record]
            print restore[user]["%s_after"%record]
        except KeyError:
            pass


records = ["hashtags","mentions","replies","retweets","status"]
for record in records:
    users = get_users()
    create_table(users,record)
