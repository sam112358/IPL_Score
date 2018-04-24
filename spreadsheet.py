import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

sheet = client.open('ipl_scores')

raw_data = sheet.worksheet('raw_data')
analysis = sheet.worksheet('analysis_sheet')

#updating the analysis sheet
#    A = 'A' + str(i)
#    B = 'B' + i
#    C = 'C' + i
#    D = 'D' + i
#    E = 'E' + i
#    F = 'F' + i
#    G = 'G' + i
#    H = 'H' + i
#    I = 'I' + i
a = 2
i = 1
k = 0
while i in range(1, 10):
    #Checking for team names
    check_term = raw_data.cell(i, 2).value
    check_team = check_term.split()
    if 'INNINGS' in check_team:
        if k == 0:
            team1 = check_term
            k = 1
        elif k == 1:
            team2 = check_term
            k = 2
        else:
            k = 0
        i += 1
    elif raw_data.cell(i,2).value == 'Batsmen':
        category = 'Batsmen'
        i += 1
    elif raw_data.cell(i,2).value == 'Bowler': 
        category = 'Bowler'
        i += 1
    elif raw_data.cell(i,2).value == None: 
        analysis.update_cell(a, 1, analysis.cell(i, 1).value) #matchNo
        analysis.update_cell(a, 2, team1)
        analysis.update_cell(a, 3, team2)
        analysis.update_cell(a, 4, k) #inning
        analysis.update_cell(a, 6, raw_data.cell(i, 3).value) #playerName
        analysis.update_cell(a, 7, category) #category
        
        if category == 'Batsmen':
            distype = (raw_data.cell(i, 4).value).split()
            if 'c' in distype:
                analysis.update_cell(a, 8, 'Catch') #dismissalType
                analysis.update_cell(a, 10, distype[1,2]) #fielder
            elif 'b' in distype:
                analysis.update_cell(a, 8, 'Bowled') #dismissalType
                analysis.update_cell(a, 9, distype[1,2]) #bowler
            elif 'lbw' in distype:
                analysis.update_cell(a, 8, 'lbw') #dismissalType
                analysis.update_cell(a, 9, distype[1,2]) #bowler
            
#        elif category == 'Bowler':
    
        a += 1
        i += 1
