%matplotlib inline

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import numpy as np
import random
import matplotlib.pyplot as plt
import mpld3

SCOPES     = ['https://www.googleapis.com/auth/spreadsheets.readonly']
data       = {"x":[],"y":[],"color":[],"size":[],"label":[]}
size_code  = {'Small': 500, 'Medium': 1000, 'Large': 3000, 'Very Large': 5000}

def main(ID, RANGE):
    '''Establishes connection with Google Sheets API and collects data'''
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server()

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds,token)

    service = build('sheets', 'v4', credentials=creds)

    sheet  = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ID, range=RANGE).execute()
    values = result.get('values', [])

    org_data = []

    n = 0
    while n < len(values[0]):
        mid = []
        for row in values:
            mid.append(row[n])
        org_data.append(mid)
        n = n + 1
    return(org_data)

def browser_plot(name, xname, yname):
    '''creates plot displayed in browser'''
    fig, ax = plt.subplots(figsize=(10,8))
    x1 = np.linspace(0,6,100)
    y1 = 0*x1 + 3
    plt.plot(x1, y1, 'k', linewidth=2)
    plt.plot(y1, x1, 'k', linewidth=2)
    plt.title('Organization Collaboration Matrix: {}'.format(name), fontsize=20)
    plt.xlabel(xname, fontsize=15)
    plt.ylabel(yname, fontsize=15)
    scatter = ax.scatter(data["x"], data["y"], marker = 'o', c = data["color"],
                          s = data["size"], alpha = 0.6)
    ax.grid(True)
    ax.set_xlim([0,6])
    ax.set_ylim([0,6])

    tooltip = mpld3.plugins.PointLabelTooltip(scatter, labels=data["label"])
    mpld3.plugins.connect(fig, tooltip)

    mpld3.show()

def save_plot(name, xname, yname, fh):
    '''Generates parameters for all plots'''
    fig = plt.figure(figsize=(10,8))
    plt.title('Data Visualization Matrix: {}'.format(name), fontsize=20)
    plt.xlabel(xname, fontsize=15)
    plt.ylabel(yname, fontsize=15)
    ax = plt.gca()
    scatter = ax.scatter(data["x"], data["y"], marker = 'o', c = data["color"],
                          s = data["size"], alpha = 0.6)

    ax.spines['right'].set_position('center')
    ax.spines['top'].set_position('center')
    ax.xaxis.set_ticks_position('bottom')
    ax.spines['bottom'].set_color('none')
    ax.yaxis.set_ticks_position('left')
    ax.spines['left'].set_color('none')
    ax.grid(True)

    axes = plt.gca()
    axes.set_xlim([0,6])
    axes.set_ylim([0,6])

    plt.savefig('Beta Pictures/{}'.format(fh))

def matrix(ID, RANGE, x, xname, y, yname, size, color_row, color, spacing=FALSE):
    '''Fills in measurement data for plots'''
    raw_data = main(ID, RANGE)
    if data["x"] == []:
        for item in raw_data:
            if "spacing" == TRUE: #slightly spreads out observations with repeat values
                data["x"].append(float(item[x]) + random.uniform(-0.35,0.35))
                data["y"].append(float(item[y]) + random.uniform(-0.35,0.35))
            else:
                data["x"].append(float(item[x])
                data["y"].append(float(item[y])
            sizzle = size_code[item[size]]
            data["size"].append(sizzle)
            data["label"].append(item[0])
    else:
        data["color"] = []

    for item in raw_data:
        if item[color_row] == color:
                data["color"].append('r')
        else:
            data["color"].append('none')
    #print(data)
    browser_plot(color, xname, yname)
    #save_plot(color, xname, yname, '{}_Matrix.png'.format(color))

#Meant to automate helping friend conducting similar project
#Switched to google form because friend didn't feel comfortable with python
def interface():
    '''collects data from user'''
    link = raw_input('Please enter Google Sheets Shareable Link:').split("/")
    ID = link[5]
    RANGE = raw_input('Enter Range for Sheet:')
    horizontal = raw_input('Input row number for X axis:')
    x = int(horizontal) - 1
    vertical = raw_input('Input row number for Y axis:')
    y = int(vertical) - 1
    big = raw_input('Which row should determine size')
    color_row = raw_input('Which row ')
