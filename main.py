#!/usr/bin/env python
import sys
import os
import subprocess
import shutil
import tempfile
import calendar
import locale
from datetime import datetime, timedelta
from dateutil import parser

import click
import jinja2

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials


LATEX_JINJA_ENV = jinja2.Environment(
    block_start_string='\BLOCK{',
    block_end_string='}',
    variable_start_string='\VAR{',
    variable_end_string='}',
    comment_start_string='\#{',
    comment_end_string='}',
    line_statement_prefix='%%',
    line_comment_prefix='%#',
    trim_blocks=True,
    autoescape=False,
)

MAIN_TEMPLATE = LATEX_JINJA_ENV.from_string(r'''
\documentclass{article}% neither 10pt nor headsepline are doing anything whatsoever as far as I can tell - certainly the class doesn't recognise them
\usepackage{geometry}
\geometry{%
  paperheight=594pt,
  paperwidth=369pt,
  layoutsize={364.5pt,585pt},
  layoutoffset={9pt,9pt},
  top=-43pt,
  bottom=10pt,
  right=10pt,
  left=30pt,
  twoside,
  showcrop
}

% Load needed packages
\usepackage{xcolor}
\usepackage{lmodern}
\usepackage{tikz}
    \usetikzlibrary{positioning}
\usepackage{afterpage}

\newcommand\blankpage{%
    \null
    \thispagestyle{empty}%
    \addtocounter{page}{-1}%
    \newpage}

\newcommand\blackpage{%
    \null
    \thispagestyle{empty}%
    \addtocounter{page}{-1}%
    \pagecolor{black}%
    \newpage}

\begin{document}
    \pagecolor{black}
    \shipout\null
    \nopagecolor
    \begin{titlepage}
        \vspace*{\fill}
        \begin{center}
          \huge{\VAR{year}}\\[1cm]
          \large{\VAR{author}}\\[0.5cm]
          \footnotesize{\VAR{author_mail}}\\[0.1cm]
          \footnotesize{\VAR{author_phone}}
        \end{center}
        \vspace*{\fill}
    \end{titlepage}
    \VAR{weeks}
    \afterpage{\blankpage}
    \afterpage{\blackpage}
\end{document}
''')

WEEK_TEMPLATE = LATEX_JINJA_ENV.from_string(r'''
\begin{tikzpicture}[%
        inner sep=3 pt,
        dayname/.style={%
            node font=\footnotesize, 
        },
        tiny/.style={%
            node font=\tiny\scshape, 
        },
        daynumber/.style={%
            anchor=north east,
            node font=\normalsize\bfseries, 
        },
        xscale = 6,
        yscale=-2.6,
    ]

    \node (day_number) at (0,7) [daynumber] {\VAR{sunday}};
    \node (day)[base right = 1em of day_number, anchor=base west] [dayname] {Sonntag};
    \VAR{events_sunday}

    \node (day_number) at (0,6) [daynumber] {\VAR{saturday}};
    \node (day)[base right = 1em of day_number, anchor=base west] [dayname] {Samstag};
    \VAR{events_saturday}

    \node (day_number) at (0,5) [daynumber] {\VAR{friday}};
    \node (day)[base right = 1em of day_number, anchor=base west] [dayname] {Freitag};
    \VAR{events_friday}

    \node (day_number) at (0,4) [daynumber] {\VAR{thursday}};
    \node (day)[base right = 1em of day_number, anchor=base west] [dayname] {Donnerstag};
    \VAR{events_thursday}

    \node (day_number) at (0,3) [daynumber] {\VAR{wednesday}};
    \node (day)[base right = 1em of day_number, anchor=base west] [dayname] {Mittwoch};
    \VAR{events_wednesday}

    \node (day_number) at (0,2) [daynumber] {\VAR{tuesday}};
    \node (day)[base right = 1em of day_number, anchor=base west] [dayname] {Dienstag};
    \VAR{events_tuesday}

    \node (day_number) at (0,1) [daynumber] {\VAR{monday}};
    \node (day)[base right = 1em of day_number, anchor=base west] [dayname] {Montag}; 
    \VAR{events_monday}

    \node (year_number) at (0,1) [anchor = south east, minimum height = 2em] {\VAR{year}};
    \node [base right = 1em of year_number, anchor=base west, node font=\Large] {\VAR{month}};
    \node [below = 53em of year_number.east, node font=\footnotesize] {KW \VAR{weeknr}};

    \foreach \i in {0,1,...,8} {%
        \draw (-0.25,\i) -- (2,\i);
    }

    %\node (SW-corner) at (0,7) {};
\end{tikzpicture}
''')

EVENTS_TEMPLATE = LATEX_JINJA_ENV.from_string(r'''
    \node (a)[below=1em of day.west,anchor=west] [tiny] {\VAR{am8}};
    \node (b)[below=2em of day.west,anchor=west] [tiny] {\VAR{am9}};
    \node (c)[below=3em of day.west,anchor=west] [tiny] {\VAR{am10}};
    \node (d)[below=4em of day.west,anchor=west] [tiny] {\VAR{am11}};
    \node (e)[below=5em of day.west,anchor=west] [tiny] {\VAR{am12}};
    \node (f)[below=6em of day.west,anchor=west] [tiny] {\VAR{pm1}};
    \node [right= 13.5em of day.west,anchor=west] [tiny] {\VAR{pm2}};
    \node [right= 13.5em of a.west,anchor=west] [tiny] {\VAR{pm3}};
    \node [right= 13.5em of b.west,anchor=west] [tiny] {\VAR{pm4}};
    \node [right= 13.5em of c.west,anchor=west] [tiny] {\VAR{pm5}};
    \node [right= 13.5em of d.west,anchor=west] [tiny] {\VAR{pm6}};
    \node [right= 13.5em of e.west,anchor=west] [tiny] {\VAR{pm7}};
    \node [right= 13.5em of f.west,anchor=west] [tiny] {\VAR{pm8}};
    \node (full)[below=6em of day_number,rotate=90,anchor=west][tiny] {\VAR{full1}};
    \node [left=1em of full,anchor=west,rotate=90][tiny] {\VAR{full2}};
    \node [left=2em of full,anchor=west,rotate=90][tiny] {\VAR{full3}};
''')


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']


# Connects to the Google API.
# An application with API credentials must be set up and a credentials.json file present in the local dir.
def google_auth():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('calendar', 'v3', credentials=creds)
    return service


def get_events(date, service):
    start = datetime.combine(date, datetime.min.time())
    end = start + timedelta(days=1) + timedelta(microseconds=-1)
    start = start.isoformat() + 'Z'  # 'Z' indicates UTC time
    end = end.isoformat() + "Z"
    calendars_result = service.calendarList().list(minAccessRole="owner").execute()
    calendars = calendars_result.get('items', [])
    events = []
    for cal in calendars:
        # get all events for current day and add to day-event-list
        events_result = service.events().list(calendarId=cal['id'],
                                              timeMin=start,
                                              timeMax=end,
                                              maxResults=10, singleEvents=True,
                                              orderBy='startTime').execute()
        eventlist = events_result.get('items', [])
        for i, event in enumerate(eventlist):
            # Don't include the calendar-label for the primary calendar
            if "primary" in cal:
                continue
            eventlist[i]['cal'] = cal['summary']
        events += eventlist
    return events


def get_weeknumber(week):
    """Week number for given week"""
    weeknumber = week[1][1].isocalendar()[1]
    return weeknumber


def get_weekdays(week):
    """Week number for given week"""
    out_tuples = []
    for day in enumerate(week[1]):
        out_tuples.append("%d" % (day[1].day))
    return out_tuples


def render_events_for_day(day, service):
    events = get_events(day, service)
    if len(events) == 0:
        return ""
    slots = [""] * 13
    full = [""] * 3
    for event in enumerate(events):
        start = dict(event[1]['start'])
        event_summary = event[1]['summary'].replace('&', 'und')
        if 'date' in start:
            if full[0] == "":
                full[0] = event_summary
            elif full[1] == "":
                full[1] = event_summary
            elif full[2] == "":
                full[2] = event_summary
            else:
                print("Warning: No space left. I will not display this full day event.")
            continue
        date = parser.parse(start['dateTime'])
        if "cal" in event[1]:
            event_string = date.strftime('%H:%M') + " " + event[1]['cal'] + ": " + event_summary + " "
        else:
            event_string = date.strftime('%H:%M') + " " + event_summary + " "
        if date.hour < 8:
            slots[0] += event_string
        elif date.hour > 20:
            slots[12] += event_string
        else:
            slots[date.hour-8] += event_string
    return EVENTS_TEMPLATE.render(dict(
        am8=slots[0],
        am9=slots[1],
        am10=slots[2],
        am11=slots[3],
        am12=slots[4],
        pm1=slots[5],
        pm2=slots[6],
        pm3=slots[7],
        pm4=slots[8],
        pm5=slots[9],
        pm6=slots[10],
        pm7=slots[11],
        pm8=slots[12],
        full1=full[0],
        full2=full[1],
        full3=full[2],
    ))


@click.command()
@click.help_option('--help', '-h')
@click.argument('year')
@click.argument('outfile')
@click.argument('author')
@click.argument('author_mail')
@click.argument('author_phone')
def main(year, outfile, author, author_mail, author_phone):
    """Generates a weekly calendar for one YEAR that fetches events from a Google calendar.

    A credentials.json file pointing to a google calender API must be present in the main directory.
    AUTHOR name, AUTHOR_MAIL and AUTHOR_PHONE must be provided as parameters for the title page.
    The OUTFILE must either have a tex or a pdf extension. If tex, the resulting file should be
    compiled with lualatex: `lualatex OUTFILE`. If pdf, the compilation with lualatex will be
    done in the background and the resulting pdf will be stored in OUTFILE.
    """
    # TODO: create parameter for locale
    locale.setlocale(locale.LC_ALL, "de_CH")
    service = google_auth()

    cal = calendar.Calendar()
    months = cal.yeardatescalendar(int(year), 1)
    renderedweeks = ""
    weeknumbers = []
    for i, month in enumerate(months):
        weeks = month[0]
        monthname = calendar.month_name[i + 1]
        for week in enumerate(weeks):
            # TODO: Make month label prettier for overlapping months (eg. "Februar/März")
            weeknr = get_weeknumber(week)
            if weeknr in weeknumbers:
                continue
            weeknumbers.append(weeknr)
            print("Rendering week " + str(weeknr))
            # Uncomment to fetch & render one week only
            #if weeknr > 10 or weeknr < 10:
            #    continue
            weekdays = get_weekdays(week)
            renderedweeks += WEEK_TEMPLATE.render(dict(
                year=year,
                month=monthname,
                weeknr=weeknr,
                monday=weekdays[0],
                events_monday=render_events_for_day(week[1][0], service),
                tuesday=weekdays[1],
                events_tuesday=render_events_for_day(week[1][1], service),
                wednesday=weekdays[2],
                events_wednesday=render_events_for_day(week[1][2], service),
                thursday=weekdays[3],
                events_thursday=render_events_for_day(week[1][3], service),
                friday=weekdays[4],
                events_friday=render_events_for_day(week[1][4], service),
                saturday=weekdays[5],
                events_saturday=render_events_for_day(week[1][5], service),
                sunday=weekdays[6],
                events_sunday=render_events_for_day(week[1][6], service),
            ))

    # Render and save output file
    if outfile.endswith('.tex'):
        with open(outfile, "w") as out_fh:
            out_fh.write(MAIN_TEMPLATE.render(dict(weeks=renderedweeks)))
    elif outfile.endswith('.pdf'):
        try:
            dir = tempfile.mkdtemp()
            with open(os.path.join(dir, 'calendar.tex'), "w") as out_fh:
                out_fh.write(MAIN_TEMPLATE.render(dict(
                    weeks=renderedweeks,
                    year=year,
                    author=author,
                    author_mail=author_mail,
                    author_phone=author_phone
                )))
            cmd = ['lualatex', '--halt-on-error', '--interaction=nonstopmode',
                   'calendar.tex']
            proc = subprocess.run(cmd, cwd=dir, stdout=subprocess.PIPE)
            if proc.returncode != 0:
                click.echo(proc.stdout)
            proc.check_returncode()  # raises SubprocessError
            shutil.copy(os.path.join(dir, 'calendar.pdf'), outfile)
        except (OSError, subprocess.SubprocessError) as exc_info:
            click.echo(str(exc_info))
            sys.exit(1)
        finally:
            shutil.rmtree(dir, ignore_errors=True)
    else:
        click.echo("OUTFILE must have either a .pdf or a .tex extension")
        sys.exit(1)


if __name__ == "__main__":
    main()
