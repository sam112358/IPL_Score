import pygsheets
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('Scoresheet.json', scope)
client = pygsheets.authorize(service_file='client_secret.json')

sheet = client.open('ipl_scores')

raw_data = sheet.worksheet_by_title('raw_data')
analysis = sheet.worksheet_by_title('analysis_sheet')

for i in range(1, 791):
#    A = 'A' + str(i)
#    B = 'B' + i
#    C = 'C' + i
#    D = 'D' + i
#    E = 'E' + i
#    F = 'F' + i
#    G = 'G' + i
#    H = 'H' + i
#    I = 'I' + i
    analysis.update_cell(i+1,1,raw_data.cell(i+1,1))

val = analysis.cell(1,1).value
analysis.update_cell(2,2, "1")
