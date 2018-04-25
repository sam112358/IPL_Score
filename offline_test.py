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

i = 0
while i in range(0,num_matches-1):
    mno = raw_data.cell_value(i, 0)
    name = raw_data.cell_value(i, 2)
    check_term = raw_data.cell_value(i, 1)
    check_team = check_term.split() 
    if 'INNINGS' in check_team:
        if 'MUMBAI' in check_team:
            check_term = 'MI'
        if 'KINGS' in check_team:
            check_term = 'KXIP'
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
            dismiss  = raw_data.cell_value(i, 3)
            distype = dismiss.split()
            if 'c' in distype: #caught
                analysis.write(a, 7, 'Catch') #dismissalType
                if '&' in distype: #caught and bowled
                    start = dismiss.index("c & b") + 4
                    analysis.write(a, 9, (dismiss[start:])[:-1])
                else:
                    start = dismiss.index("c ") + 2
                    end = dismiss.index("b ")
                    analysis.write(a, 9, (dismiss[start:end])[:-1])
            elif 'b' in distype:
                if 'st' in distype:
                    start = dismiss.index("st ") + 2
                    end = dismiss.index("b ")
                    analysis.write(a, 7, 'Stump') #dismissalType
                    analysis.write(a, 8, (dismiss[end+1:])[:-1]) #bowler
                    analysis.write(a, 9, (dismiss[start:end])[:-1]) #fielderr
                else:
                    start = dismiss.index("b ") + 1
                    analysis.write(a, 7, 'Bowled') #dismissalType
                    analysis.write(a, 8, (dismiss[start:])[:-1]) #bowler  
            elif 'lbw' in distype:
                start = dismiss.index("lbw ") + 3
                analysis.write(a, 7, 'LBW') #dismissalType
                analysis.write(a, 8, (dismiss[start:])[:-1]) #bowler
            elif 'NOT' in distype:
                analysis.write(a, 7, 'Not Out') #dismissalType
            elif 'run' in distype:
                start = dismiss.index("(")
                end = dismiss.index(")")
#                slash_posi = run_out_second_name.index("/")
                analysis.write(a, 7, 'Run Out') #dismissalType
                analysis.write(a, 9, (dismiss[start+1:end])[:-1])
#                run_out_second_name = run_out_second_name[:slash_posi]
#                analysis.write(a, 8, run_out_first_name + run_out_second_name)
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
sheet = xlrd.open_workbook('analysis_test.xlsx')
update = sheet.sheet_by_index(0)

for i in range(0, a):
    #removing the extra columns that were added
    if update.cell_value(i, 5) == '': 
        for k in range(0, 22):
                analysis.write(i, k) == ''
     
    #concatenating the stats of 1 match for the same player           
    mno = update.cell_value(i, 0)
    player_name = update.cell_value(i, 5)
    innings = update.cell_value(i, 3)
    category = update.cell_value(i, 6)
    num_catch = num_stump = num_runout = 0
    for j in range(i, a-1):
        if mno == update.cell_value(j, 0) and player_name == update.cell_value(j, 5) and innings != update.cell_value(j, 3):
            if category == 'Batsman':
                analysis.write(i, 15, (update.cell_value(j, 15)))
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
                analysis.write(j, k) == '' #removing redundant stats
                
        #adding the number of fielding stats for each player for each mach
        if mno == update.cell_value(j, 0):
            if player_name == update.cell_value(j, 9):
                if update.cell_value(j, 7) == 'Catch':
                    num_catch += 1
                if update.cell_value(j, 7) == 'Stump':
                    num_stump += 1
                if update.cell_value(j, 7) == 'Run Out':
                    num_runout += 1
    analysis.write(i, 20, num_catch)
    analysis.write(i, 21, num_stump)
    analysis.write(i, 22, num_runout)

wb.save('analysis_test.xlsx') #saving the file


#POINTS CALCULATOR
#playerdb = workbook.sheet_by_index(2)
sheet = xlrd.open_workbook('analysis_test.xlsx')
update = sheet.sheet_by_index(0)

for i in range(0, a-2):
    strike_rate = update.cell_value(i, 12)
    base_pts = update.cell_value(i, 10)
    wickets = update.cell_value(i, 17)
    num_6s = update.cell_value(i, 14)
    runs_scored = base_pts
    num_stump = num_runout = num_catch = fielding_points = bowling_impact = bowling_milestone = economy_pts = impact_pts = base_points = pace_bonus_pts = bowling_base = milestone_pts = 0
    wickets_taken = update.cell_value(i, 17)
    economy_rate = update.cell_value(i, 18)
    
    try:
        num_catch = update.cell_value(i, 20)
    except:
        analysis.write(i, 20, 0)
            
    try:
        num_stump = update.cell_value(i, 21)
    except:
        analysis.write(i, 21, 0)

    try:
        num_runout = update.cell_value(i, 22)
    except:
        analysis.write(i, 22, 0)
    try:
        if base_pts > 10: #calculating batsman base
            if strike_rate < 75:
                pace_bonus_pts = -15
            elif strike_rate < 100:
                pace_bonus_pts = -10
            elif strike_rate < 150:
                pace_bonus_pts = 5
            elif strike_rate < 200:
                pace_bonus_pts = 10
            else:
                pace_bonus_pts = 15
    except:
        analysis.write(i, 10, 0)
        
    try:
        if runs_scored > 25: #calculating batsman milestone
            milestone_pts = int(runs_scored/25)
    except:
        analysis.write(i, 10, 0)
        
    impact_pts = num_6s * 2 #calculating batsman impact
    if runs_scored == 0:
        impact_pts -= 5
        
    if wickets_taken: #calculating bowling_base
        bowling_base = wickets_taken * 20
    try:    
        if economy_rate <= 5: #calculating economy points
            economy_pts = 15
        elif economy_rate <= 8:
            economy_pts = 10
        elif economy_rate <= 10:
            economy_pts = 5
        elif economy_rate <= 12:
            economy_pts = -10
        else:
            economy_pts = -15
    except:
        analysis.write(i, 18, 0)
        
    try:
        if wickets_taken == 2: #calculating bowling milestone
            bowling_miestone = 10
        elif wickets_taken >= 2:
            bowling_milestone = (wickets_taken - 2) * 10
    except:
        analysis.write(i, 17, 0)
        
    bowling_impact = update.cell_value(i, 19) #calculating bowling impact points
    
    fielding_points = (num_catch * 10) + (num_stump * 15) + (num_runout * 10)
    
    analysis.write(i, 24, base_pts)
    analysis.write(i, 25, pace_bonus_pts)
    analysis.write(i, 26, milestone_pts)
    analysis.write(i, 27, impact_pts)
    analysis.write(i, 28, economy_pts)
    analysis.write(i, 29, bowling_milestone)
    analysis.write(i, 30, bowling_impact)
    
for i in range(0, a):
    if update.cell_value(i, 0) == '':
        for j in range(0, 31):
            analysis.write(i, j, '')
wb.save('analysis_test.xlsx') #saving the file

