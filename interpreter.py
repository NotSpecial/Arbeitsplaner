from xlrd import open_workbook
import os
import re
import datetime
import icalendar
import sys
from icalendar.prop import vText
import pytz

# Variables
file_dir = os.getcwd() + "\\raw\\"
cal_dir = os.getcwd() + "\\cal\\"


def remove_non_ascii(s):
    return "".join(i for i in s if ord(i) < 256)

# contant minutes if there is no assigned starting time
class Jobinfo(object):
    """Containing important info on jobs"""
    def __init__(self, name, t_einf, t_start, t_duration):
        self.name = name
        self.t_einf = t_einf
        self.t_start = t_start
        self.t_duration = t_duration

# Joblist, ordered by importance
job_list =  [
            Jobinfo("VEF"      , 45,  0, 210),
            Jobinfo("V"        ,  0, 75, 210),
            Jobinfo("PEF"      , 45,  0, 210),
            Jobinfo("P"        ,  0, 75, 210),
            Jobinfo("NEF"      , 35,  0, 180),
            Jobinfo("N"        ,  0, 65, 180),
            Jobinfo("B/N"      ,  0, 65, 180),
            Jobinfo("B"        ,  0, 60,  60),
            Jobinfo("E+"       ,  0, 45, 120),
            Jobinfo("E"        ,  0, 45,  60),
            Jobinfo("Sonstiges",  1,  1,  60),
            ]


class Timeinfo(object):
    """All Time information for an event
        strings for:
        - time of introduction t_einf
        - stating time
        - ending time
        - additional info (to be parsed somewhere else"""

    def __init__(self, t_einf, t_start, t_end, additional):
        self.t_einf = t_einf
        self.t_start = t_start
        self.t_end = t_end
        self.additional = additional

    def __str__(self):
        temp = ""
        for t in [self.t_einf, self.t_start, self.t_end]:
            if t == "":
                temp += "\t"
                # temp += "nux\t"
            else:
                temp += t + "\t"
                # temp += "wot\t"
        return temp


# Define Class for events
class Event(object):
    """Veranstaltung, enthaelt Name, Zeit, Ort, Spalte in der Tabelle"""
    def __init__(self, title, file, col, date, place, time):
        # Titel
        self.title = title
        # Ueberpruefen, ob Premiere oder Derniere
        if re.search('prem', time.additional, re.IGNORECASE):
            self.title += " (Premiere)"
        if re.search('dern', time.additional, re.IGNORECASE):
            self.title += " (Derniere)"

        self.file = file
        self.col = col
        self.date = date
        if time.t_end == '00:00':
            self.date_end = self.date + datetime.timedelta(days=1)
        else:
            self.date_end = self.date

        self.place = place
        self.time = time
        self.jobs = []

    # Generate description
    def get_description(self):
        temp = ""
        # Check time-class additional info for
        # # Fester Arbeitsbeginn
        if re.search('arb.*', self.time.additional, re.IGNORECASE):
            temp += "Fester Arbeitsbeginn\n\n"

        # # Publikumsgespraech
        if re.search('publ', self.time.additional, re.IGNORECASE):
            temp += "Publikumsgespraech im Anschluss\n\n"

        # Use general joblist to order
        for master in job_list:
            for slave in self.jobs:
                if slave.t == master.name:
                    temp += slave.str_with_person() + "\n"

        return temp

    def __str__(self):
        temp = str(self.date) + "\t" + str(self.time) + self.place + "\t"
        if self.place == "Pfauen":
            temp += "\t"

        if self.place == "":
            temp += "\t\t"

        temp += self.title
        return temp


# Class for jobs, with link to person (NOT YET IMPLEMENTED) and event
class Job(object):
    """Dienst, enthaelt Typ und Veranstaltung"""
    def __init__(self, eventtype, person, event):
        self.t = eventtype
        # remeber events
        self.event = event
        self.person = person
        # hook into objects
        event.jobs.append(self)
        person.jobs.append(self)

        # Get Title
        self.title = self.t + " " + str(self.event.title)

    def get_t_start(self):
        return

    def str_with_person(self):
        return self.t + "\t" + str(self.person)

    def str_with_event(self):
        return self.t + "\t" + str(self.event)

    def __str__(self):
        return self.t + "\t" + str(self.person) + "\t" + str(self.event)


# Class for person, contain name, position and jobs
class Person(object):
    """Person, enthaelt Name, Zeile in der Tabelle und Liste an Diensten"""
    def __init__(self, name, row):
        self.name = name
        self.row = row
        self.jobs = []

    def get_jobs(self):
        temp = ""
        for j in self.jobs:
            temp += j.str_with_event() + "\n"
        return temp

    def __str__(self):
        return self.name

    # Needs a string of format xx:yy and turns it into time object
    def format_time(self, timestring):
        tz = pytz.timezone('CET')
        if re.match('\d\d:\d\d$', timestring):
            return datetime.time(hour=int(timestring[0:2]),
                                 minute=int(timestring[3:5]), tzinfo=tz)
        else:
            # Crash
            sys.exit("Wanted to format incompatible timestring:\'"
                     + timestring + "\'")

    # Use the Info to bring it to correct format. all time formatting is here
    def get_calendar(self):
        # Create calendar object
        cal = icalendar.Calendar()

        for j in self.jobs:
            e = icalendar.Event()

            # Add title and description
            e.add('summary', j.title.encode('utf-8'))
            e.add('description', j.event.get_description().encode('utf-8'))

            # Add place
            e.add('location', vText(j.event.place).encode('utf-8'))

            # Now, get start and end-time
            for info in job_list:
                if info.name == j.t:
                    start_date = j.event.date
                    # Look for fixed start
                    if re.search('arb.*', j.event.time.additional, re.IGNORECASE):
                        start = self.format_time(j.event.time.t_start)
                        start_delta = 0
                    # Has introduction
                    elif info.t_einf and j.event.time.t_einf:
                        start = self.format_time(j.event.time.t_einf)
                        start_delta = - info.t_einf
                    # Has regular start
                    elif (info.t_start and j.event.time.t_start):
                        start = self.format_time(j.event.time.t_start)
                        start_delta = - info.t_start
                    # Has no start, no intro, but an end time! -> comp start time
                    elif j.event.time.t_end:
                        start_date = j.event.date_end  # Use different date!
                        start = self.format_time(j.event.time.t_end)
                        start_delta = - info.t_duration
                    else:
                        print('time data insufficient for:\n%s!' % str(j))
                        exit

                    e.add('dtstart', datetime.datetime.combine(start_date,
                          start) + datetime.timedelta(minutes=start_delta))

                    # Compute ending time
                    # # Case 1: there is a fixed ending
                    if not j.event.time.t_end == "":
                        end = self.format_time(j.event.time.t_end)
                        end_delta = 0
                    # # Case 2: no fixed end
                    else:
                        end = self.format_time(j.event.time.t_start)
                        end_delta = start_delta + info.t_duration

                    e.add('dtend', datetime.datetime.combine(j.event.date_end,
                          end) + datetime.timedelta(minutes=end_delta))

            # timestamp
            e.add('dtstamp', datetime.datetime.now())

            # Add to calendar
            cal.add_component(e)

        # give calendar
        return cal

    def write_calendar(self, filename):
        """Write calendar to file, returns filename
        """
        cal = self.get_calendar()
        path = cal_dir + filename
        calfile = open(path, "wb")
        calfile.write(cal.to_ical())
        calfile.close()

        # Write backup
        # path = cal_dir + filename + '_backup_' + str(datetime.date.today())
        # calfile = open(path , "wb")
        # calfile.write(cal.to_ical())
        # calfile.close()
        return path


# Main Class, will contain everything. opens file when created
class Interpreter(object):
    """docstring for ClassName"""
    def __init__(self, filename):
        self.logstring = "Started Log...\n"
        self.alertstring = "Started Alerts...\n"
        self.persons = []
        self.events = []
        self.jobs = []
        self.filename = filename
        self.scan_file()

    def log(self, text):
        # print(remove_non_ascii(text))
        self.logstring += remove_non_ascii(text)  # To excape unicode hell
        self.logstring += "\n"

    def alert(self, text):
        self.alertstring += text
        self.alertstring += "\n"
        # print(text)

    def print_log(self):
        logfile = open("%s_log.txt" %self.filename, "w")
        logfile.write(self.logstring)
        logfile.close()

    def print_alerts(self):
        alertfile = open("%s_alerts.txt" % self.filename, "w")
        alertfile.write(self.alertstring)
        alertfile.close()

    # Functions
    # # access functions
    def event_by_index_and_file(self, index, file):
        for e in self.events:
            if (e.col == index) and (e.file == file):
                return e

    def event_by_day(self, day):
        for e in self.events:
            if e.day == day:
                return e

    def person_by_name(self, name):
        for p in self.persons:
            if not (p.name.find(name) == -1):
                return p

    def not_working(self, event):
        w = []
        for j in event.jobs:
            w.append(j.person)

        temp = []

        for p in self.persons:
            if p not in w:
                temp.append(p)

        return temp

    # # cell functions
    def has_left_border(self, row, col):
        # check self
        c = self.sheet.cell(row, col)
        if not (self.book.xf_list[c.xf_index].border.left_line_style == 0):
            return True

        # No left border -> check Cell left for right border
        c = self.sheet.cell(row, col - 1)
        return not (self.book.xf_list[c.xf_index].border.right_line_style == 0)

    def has_right_border(self, row, col):
        # check self
        c = self.sheet.cell(row, col)
        if not (self.book.xf_list[c.xf_index].border.right_line_style == 0):
            return True

        # No right border -> check Cell left for left border
        c = self.sheet.cell(row, col + 1)
        return not (self.book.xf_list[c.xf_index].border.left_line_style == 0)

    def get_digits(self, s):
        return ''.join(i for i in s if i.isdigit())

    def get_date(self, cell):
        v = cell.value

        # If we already have a number, just convert it
        if isinstance(v, float):
            return str(int(v))

        return self.get_digits(v)

    # Returns date as a string (as parsed from file)
    def scan_date(self, col):
        self.log("Looking for date, starting with column " + str(col) + "...")
        # Dates are below events
        dates_row = self.events_row + 1

        # check current cell
        c = self.sheet.cell(dates_row, col)
        s = self.get_date(c)
        if not(s == ""):
            # Found date!
            self.log("Found date in current position: " + s + " ")
            return s

        # check left cell if current cell empty until you hit a border
        self.log("No date found, looking to the left...")
        newcol = col
        while not self.has_left_border(dates_row, newcol):
            newcol -= 1
            c = self.sheet.cell(dates_row, newcol)
            s = self.get_date(c)
            if not(s == ""):
                self.log("Found date in column " + str(newcol) + ": "
                         + s + " ")
                return s

        # check right cells if nothing found left, again until you hit a border
        self.log("No date found, looking to the right...")
        newcol = col
        while not self.has_right_border(dates_row, newcol):
            newcol += 1
            c = self.sheet.cell(dates_row, newcol)
            s = self.get_date(c)
            if not (s == ""):
                self.log("Found date in column " + str(newcol) + ": "
                         + s + " ")
                return s

        self.log("No date found. Warning issued and returned empty string")
        self.alert("Found no date starting from column " + str(col)
                   + ", empty string returned")
        return ""

    # Wrapper for scan_date to turn output into datetime
    def find_date(self, col):
        d = self.scan_date(col)

        # Convert
        if d == "":
            d = 0
        else:
            d = int(d)
        return datetime.date(self.year, self.month, d)

    # Find the Place, maybe parse already (NOT YET IMPLEMENTED)
    def find_place(self, col):
        # Places are two rows below events
        place_row = self.events_row + 2

        t = str(self.sheet.cell(place_row, col).value)
        if t == "PF":
            return "Pfauen"
        elif t == "SB":
            return "Schiffbau"
        else:
            return t

    # Find the time
    def find_time(self, col):
        self.log("Checking for time...")
        # Times are from events_row + 3 to events_row + 7
        time_row_first = self.events_row + 3
        time_row_last = self.events_row + 7

        # Get time string
        timecol = self.sheet.col_slice(col, time_row_first, time_row_last)
        timestring = ""
        for field in timecol:
            timestring += " " + str(field.value)

        self.log("Interpreting string: " + timestring)

        # replace . and , with : for uniform formatting
        # add zero in front to bring all time stamps to the format xx:xx
        timestring = re.sub('[, .]', ':', timestring)
        timestring = re.sub('(?<=\D)(\d:\d\d)', '0\g<1>', timestring)
        timestring = re.sub('(\d\d:\d)\D', '\g<1>0', timestring)

        # Substitute "bis" with 9, since the dates will never start with 9
        timestring = re.sub('bis', '9', timestring)

        # Building the ultimate regex
        m = re.match('\D*((?P<t_einf>\d\d:\d\d)*\D*(?P<t_start>\d\d:\d\d))*\D*(9\D*?(?P<t_end>\d\d:\d\d))*\D*$', timestring, re.IGNORECASE)

        # Check for error
        if not m:
            self.log("Could not find time for event!")
            return

        # introduction time
        if bool(m.group('t_einf')):
            self.log("Found introduction at: " + str(m.group('t_einf')))
            t_einf = str(m.group('t_einf'))
        else:
            t_einf = ""
        # end time
        if bool(m.group('t_end')):
            self.log("Found end at: " + str(m.group('t_end')))
            t_end = str(m.group('t_end'))
        else:
            t_end = ""
            # starting time
        if bool(m.group('t_start')):
            self.log("Found start at: " + str(m.group('t_start')))
            t_start = str(m.group('t_start'))
        else:
            t_start = ""

            # t_start = t_end - t_duration_no_start

        # Return the whole timestring as rest so it can be searched for keywords
        return Timeinfo(t_einf, t_start, t_end, timestring)

    # Checks the job string if the job description is valid. Issues warning and makes default entry if job is unknown
    def check_job(self, job, person):
        # Convert to uppercase first.
        temp = job.upper()
        self.log("Checking " + job)
        for j in job_list:
            if temp == j.name:
                self.log("Valid!")
                return temp
        # If not here, it is not valid!
        self.alert(person.name + ": Job " + job + " not found! Creating as default.")
        self.log("Warning issued: Job not found, entry: default job")
        return "Sonstiges"

    def scan_file(self):
        # Check if File exists
        if not (os.path.isfile(file_dir + self.filename)):
            # Error if file is missing
            self.log("File %s does not exist.\nExiting." %
                     (file_dir + self.filename))
            return

        if not re.match('\d{6}.xls$',self.filename):
            self.log("File name not in format: MMYYYY.xls")
            self.log("Exiting")
            return

        # Workbook oeffnen und sheet waehlen
        self.log("Loading file: " + file_dir + self.filename + "...")
        self.book = open_workbook(file_dir + self.filename,
                                  formatting_info=True)
        self.sheet = self.book.sheet_by_index(0)
        self.month = int(self.filename[0:2])
        self.year = int(self.filename[2:6])
        self.log("Set month to %s" % self.month)
        self.log("Set year to %s" % self.year)
        self.log("Done.")

        # Find first row with events
        self.log("Checking which row contains events.")
        temp = self.sheet.col(0)
        self.events_row = 0
        while temp[self.events_row].value == "":
            self.events_row += 1
        self.log("Found events in row %i" % self.events_row)

        # Check events
        self.log("Looking for events...")

        eventrow = self.sheet.row(self.events_row)

        index = 0
        for cell in eventrow:
            # Check if empty or contains numbers

            if cell.value == "" or not isinstance(cell.value, str):
                self.log("Skipped column " + str(index))
            else:
                self.log("Found event '" + cell.value + "' in column " + str(index))
                self.events.append(Event(
                    cell.value,
                    self.filename,
                    index,
                    self.find_date(index),
                    self.find_place(index),
                    self.find_time(index)
                    ))
            index += 1

        self.log("Done.")

        # Check persons
        self.log("Looking for persons...")

        # THIS IS ALSO NOT FINAL
        perscol_start = 13
        perscol = self.sheet.col_slice(0, perscol_start)

        index = perscol_start
        for cell in perscol:
            # Check if empty or contains numbers
            if (cell.value == "" or re.search('(T|t)otal',cell.value)
                or not isinstance(cell.value, str)):
                self.log("Stopping in row " + str(index))
                break
            else:
                self.log("Found person '" + cell.value + "' in row " + str(index))
                # check if person is already in list
                p = self.person_by_name(cell.value)
                if p:
                    self.log("Exists already, updating")
                    p.row = index
                else:
                    self.log("Creating new")
                    self.persons.append(Person(cell.value, index))
            index += 1

        self.log("Done.")

        # Find jobs
        self.log("Looking for jobs...")

        # Going through all events
        first_col = self.events[1].col
        last_col = self.events[-1].col + 1

        # for every person, find jobs
        for p in self.persons:
            self.log("Looking for jobs for %s in row %i..." % (p.name, p.row))
            # using the row of the person and the event boundaries, get cells
            j = self.sheet.row_slice(p.row, first_col, last_col)

            index = first_col
            for cell in j:

                # Check if empty, or not str
                if cell.value == "" or not isinstance(cell.value, str):
                    self.log("Cell in column " + str(index) + " empty. Skipping.")
                # Continue with alert if value appears to be name (too long)
                elif (len(cell.value) > 4):
                    self.log("Cell in column " + str(index) + " appears to be no job, alert issued. Skipping.")
                    self.alert("Discarded " + cell.value + " as job because the string is too long. See log for more information.")
                else:
                    # get event by index
                    ev = self.event_by_index_and_file(index, self.filename)
                    job = self.check_job(cell.value, p)
                    self.log("Found: %s in column %i:\n%s" % (job,index,ev))
                    # Create job -> will hook into everybody
                    self.log("Adding job...")
                    self.jobs.append(Job(job, p, ev))
                    self.log("Done!")
                index += 1

        self.log("Done.")
