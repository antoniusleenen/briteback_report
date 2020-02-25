"""Import modules"""
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import Table
from reportlab.platypus.tables import TableStyle

import datetime as dt
import numpy as np
import pandas as pd
import os

# Define Input Briteback Datafile
input_file = "KNBSB_CTO_Weekly.xlsx"

# Define Briteback Datafile Path
datafile = os.path.abspath(os.path.join("data/" + input_file))

# Load Briteback Datafile
dataset = pd.read_excel(datafile, usecols=[0, 3, 4, 5, 6, 7],
                        names=["Names", "Time", "Recurrence", "Strength", "Tennis", "Others"])

# Briteback Questionnaire Week
resp_datetime = dataset.iloc[0, 1]

# Select Year, Month and Day from Datetime
resp_year = int(resp_datetime[0:4])
resp_month = int(resp_datetime[5:7])
resp_day = int(resp_datetime[8:10])

# Convert Date to Week
resp_week = dt.date(resp_year, resp_month, resp_day).isocalendar()[1]

# Number of Respondents
num_respondents = (dataset.iloc[:, 2] == dataset.iloc[0, 2]).sum()

# Respondents Names
resp_names = dataset.iloc[0:num_respondents, 0]

# Hours Training of Respondents
training_1 = dataset.iloc[0:num_respondents, 4]
training_2 = dataset.iloc[0:num_respondents, 3]
training_3 = dataset.iloc[0:num_respondents, 5]

# DESIGN OF THE PDF-REPORT

# Register 'Calibri' Fonts
pdfmetrics.registerFont(TTFont('calibri_bold', 'fonts/calibri_bold.ttf'))
pdfmetrics.registerFont(TTFont('calibri_light', 'fonts/calibri_light.ttf'))

# Create CANVAS PDF Document
c = canvas.Canvas('briteback_report.pdf', pagesize=A4)

# Author
c.setAuthor('Ton (AJR) Leenen')

# Header
c.setFont('calibri_bold', 28, leading=None)
c.drawCentredString(300, 800, 'WEEKLY REPORT')
c.line(205, 794, 395, 794)
c.setFont('calibri_bold', 12.25, leading=None)
c.drawCentredString(300, 780, 'BREAKING THE HIGH LOAD IN TENNIS')
c.line(30, 760, 565, 760)

# Respondents and Week Number in Subheader
c.setFont('calibri_bold', 12, leading=None)
c.drawString(30, 745, 'Week: ' + str(resp_week))
c.drawString(30, 730, 'Respondents: ' + str(num_respondents))
c.line(30, 720, 565, 720)

# Respondents Names
c.setFont('calibri_bold', 12, leading=None)
c.drawString(30, 705, 'Respondents')

# Set Position and Size
x = 30
y = 693
size = 12
for line in resp_names:
    c.setFont('calibri_light', 12, leading=None)
    c.drawString(x, y, line)
    y = y-size

# Hours Training of Respondents

# Tennis Training
c.setFont('calibri_bold', 12, leading=None)
c.drawString(175, 705, 'Tennis Training (Hours)')

# Set Position and Size
x = 225
y = 693
size = 12
for line in training_1:
    c.setFont('calibri_light', 12, leading=None)
    c.drawString(x, y, str(line))
    y = y-size

# Strength Training
c.setFont('calibri_bold', 12, leading=None)
c.drawString(305, 705, 'Strength Training (Hours)')

# Set Position and Size
x = 365
y = 693
size = 12
for line in training_2:
    c.setFont('calibri_light', 12, leading=None)
    c.drawString(x, y, str(line))
    y = y-size

# Other Training
c.setFont('calibri_bold', 12, leading=None)
c.drawString(450, 705, 'Other Training (Hours)')

# Set Position and Size
x = 500
y = 693
size = 12
for line in training_3:
    c.setFont('calibri_light', 12, leading=None)
    c.drawString(x, y, str(line))
    y = y-size
c.line(30, 410, 565, 410)

# Logos Institutes
c.drawImage('images/KNLTB_Logo.jpg', 105, 27, 0.50 * 6 * cm, 0.50 * 2 * cm)
c.drawImage('images/VU_Logo.jpg', 250, 30, 0.50 * 8 * cm, 0.50 * 2 * cm)
c.drawImage('images/TU_Logo.jpg', 415, 25, 0.50 * 6 * cm, 0.50 * 3 * cm)
c.line(30, 80, 565, 80)

c.showPage()
c.save()
