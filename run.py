import os
from interpreter import Interpreter
#from upload import upload

#Variables
file_dir = os.getcwd() + "\\raw\\"
cal_dir = os.getcwd() + "\\cal\\"


name = 'Alexander'

file_path = "Test"

# Run that shit!
def cal_name(f,n):
    base, ext = os.path.splitext(f)
    return 'cal_%s_%s.ics' % (name, base)


def have_cal(f):
    return os.path.exists(os.path.join(cal_dir, cal_name(f, name)))

#Scan all calendar files
for f in os.listdir(file_dir):
    if have_cal(f):
        print('%s is not needed. Passing.' % f)
        continue
    print("Scanning file %s..." % str(f))
    i = Interpreter(f)
    print("Done")

    #get log
    i.print_log()

    p = i.person_by_name(name)
    #for j in p.jobs:
     #   print(j)
    try:
        calfile = p.write_calendar(cal_name(f, name))
    except Exception as e:
        print(e)
