#importing libraries
import xlrd
import xlwt

#loading file
workbook = xlrd.open_workbook('ipl_scores.xlsx')
raw_data = workbook.sheet_by_index(0)

#creating new file to work on
wb = xlwt.Workbook()
analysis = wb.add_sheet('analysis', cell_overwrite_ok=True)

#working on new file
team_set = i = k = name_check = a = t = num_matches = 0
team1 = team2 = ' '

mno = raw_data.cell_value(num_matches, 0)
while mno:
    mno = raw_data.cell_value(num_matches-1, 0)    
    num_matches += 1

while i in range(0,num_matches-1):
    mno = raw_data.cell_value(i, 0)
    name = raw_data.cell_value(i, 2)
    check_term = raw_data.cell_value(i, 1)
    check_team = check_term.split() 
    if 'INNINGS' in check_team:
        if 'MUMBAI' in check_team:
            check_term = 'MI'
        if 'KINGS' in check_team:
            check_term = 'KXI'
        if 'CHENNAI' in check_team:
            check_term = 'CSK'
        if 'KOLKATA' in check_team:
            check_term = 'KKR'
        if 'HYDERABAD' in check_team:
            check_term = 'SRH'
        if 'BANGALORE' in check_team:
            check_term = 'RCB'
        if 'RAJASTHAN' in check_team:
            check_term = 'RR'
        if 'DELHI' in check_team:
            check_term = 'DD'
        if k == 0 or k == 2:
            team1 = check_term
            k = 1
        elif k == 1:
            team2 = check_term
            k = 2
    elif check_term == 'Batsmen':
        team_set = 0
        category = 'Batsman'
    elif check_term == 'Bowler':
        team_set = 1
        category = 'Bowler'
    elif check_term == '': 
        analysis.write(a, 0, mno) #matchNo
        analysis.write(a, 1, team1)
        analysis.write(a, 2, team2)
        analysis.write(a, 3, k) #inning
        analysis.write(a, 5, name) #playerName
        analysis.write(a, 6, category) #category
        if category == 'Batsman':
            distype = (raw_data.cell_value(i, 3)).split()
            if 'c' in distype:
                analysis.write(a, 7, 'Catch') #dismissalType
                if 'b' in distype:
                    analysis.write(a, 9, distype[3] + ' ' + distype[4]) #caught and bowled
                analysis.write(a, 9, distype[1] + ' ' + distype[2])
            elif 'b' in distype:
                if 'st' in distype:
                    analysis.write(a, 7, 'Stump') #dismissalType
                    analysis.write(a, 8, distype[4] + ' ' + distype[5]) #bowler
                    analysis.write(a, 9, distype[1] + ' ' + distype[2]) #fielder
                else:
                    analysis.write(a, 7, 'Bowled') #dismissalType
                    analysis.write(a, 8, distype[1] + ' ' + distype[2]) #bowler
                
                
            elif 'lbw' in distype:
                analysis.write(a, 7, 'lbw') #dismissalType
                analysis.write(a, 8, distype[1] + ' ' + distype[2]) #bowler
            elif 'NOT' in distype:
                analysis.write(a, 7, 'Not Out') #dismissalType
            runs_scored = raw_data.cell_value(i, 4)
            balls_played = raw_data.cell_value(i, 5)
            strike_rate = raw_data.cell_value(i, 6)
            num_4 = raw_data.cell_value(i, 7)
            num_6 = raw_data.cell_value(i, 8)
            analysis.write(a, 10, runs_scored)
            analysis.write(a, 11, balls_played)
            analysis.write(a, 12, strike_rate)
            analysis.write(a, 13, num_4)
            analysis.write(a, 14, num_6)
    
        if category == 'Bowler':
            overs_bowled = raw_data.cell_value(i, 3)
            runs_against = raw_data.cell_value(i, 4)
            wickets_taken = raw_data.cell_value(i, 5)
            economy = raw_data.cell_value(i, 6)
            dot_balls = raw_data.cell_value(i, 7)
            analysis.write(a, 15, overs_bowled)
            analysis.write(a, 16, runs_against)
            analysis.write(a, 17, wickets_taken)
            analysis.write(a, 18, economy)
            analysis.write(a, 19, dot_balls)
            
        a += 1
    i += 1
        
wb.save('analysis_test.xlsx') #saving the file

#setting teams
sheet = xlrd.open_workbook('analysis_test.xlsx')
update = sheet.sheet_by_index(0)
    
i = a - 1
while i in range(0, a):
    mno = update.cell_value(i, 0)
    team1 = update.cell_value(i, 1)
    team2 = update.cell_value(i, 2)
    while mno == update.cell_value(i, 0):
        analysis.write(i, 1, (team1))
        analysis.write(i, 2, (team2))
        i -= 1
wb.save('analysis_test.xlsx') #saving the file

#concatenating the player with bowling and batting
for i in range(0, a-1):
    if update.cell_value(i, 5) == '':
        for k in range(0, 20):
                analysis.write(i, k) == ''
    mno = update.cell_value(i, 0)
    player_name = update.cell_value(i, 5)
    innings = update.cell_value(i, 3)
    category = update.cell_value(i, 6)
    num_catch = num_stump = num_runout = 0
    for j in range(i, a-1):
        if mno == update.cell_value(j, 0) and player_name == update.cell_value(j, 5) and innings != update.cell_value(j, 3):
            if category == 'Batsman':
                analysis.write(i, 15, int(update.cell_value(j, 15)))
                analysis.write(i, 16, int(update.cell_value(j, 16)))
                analysis.write(i, 17, int(update.cell_value(j, 17)))
                analysis.write(i, 18, int(update.cell_value(j, 18)))
                analysis.write(i, 19, int(update.cell_value(j, 19)))
            elif category == 'Bowler':
                analysis.write(i, 10, int(update.cell_value(j, 10)))
                analysis.write(i, 11, int(update.cell_value(j, 11)))
                analysis.write(i, 12, int(update.cell_value(j, 12)))
                analysis.write(i, 13, int(update.cell_value(j, 13)))
                analysis.write(i, 14, int(update.cell_value(j, 14)))
            for k in range(0, 20):
                analysis.write(j, k) == ''
        if mno == update.cell_value(j, 0):
            if player_name == update.cell_value(j, 8):
                if update.cell_value(k, 6) == 'Catch':
                    num_catch += 1
                if update.cell_value(k, 6) == 'Stump':
                    num_stump += 1
    analysis.write(i, 20, num_catch)
    analysis.write(i, 20, num_stump)

wb.save('analysis_test.xlsx') #saving the file