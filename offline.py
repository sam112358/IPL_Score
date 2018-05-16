#importing libraries
import xlrd #reading files
import xlwt #creating files

#loading file
workbook = xlrd.open_workbook('ipl_scores.xlsx') 
wb = xlwt.Workbook()

#creating new file to work on
raw_data = workbook.sheet_by_index(1)
analysis = wb.add_sheet('analysis_test', cell_overwrite_ok=True)

#working on new file
team_set = i = k = name_check = a = t = num_matches = 0
team1 = team2 = ' '

mno = raw_data.cell_value(num_matches, 0)
try:
    while mno:
        mno = raw_data.cell_value(num_matches-1, 0)   
        num_matches += 1
except:
    i = 0





#run till here first
i = 0
while i in range(0,num_matches - 1):        
    mno = raw_data.cell_value(i, 0)
    name = raw_data.cell_value(i, 2)
    check_term = raw_data.cell_value(i, 1)
    check_team = check_term.split() 
    if 'INNINGS' in check_team: #setting teams
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
        if 'PUNE' in check_team:
            check_term = 'PWI'
        if 'DECCAN' in check_team:
            check_term = 'DC'
        if k == 0 or k == 2: #sets innings
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
        analysis.write(a, 1, team1) #team1
        analysis.write(a, 2, team2) #team2
        analysis.write(a, 3, k) #inning
        analysis.write(a, 5, name) #playerName
        analysis.write(a, 6, category) #category
        
        if category == 'Batsman':
            dismiss  = raw_data.cell_value(i, 3)
            distype = dismiss.split()
            if 'c' in distype: #caught
                analysis.write(a, 7, 'Catch') #dismissalType
                if '&' in distype: #caught and bowled
                    start = dismiss.index("c & b") + 6
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
                    analysis.write(a, 8, (dismiss[end+2:])) #bowler
                    analysis.write(a, 9, (dismiss[start+1:end])[:-1]) #fielder
                else:
                    start = dismiss.index("b ") + 1
                    analysis.write(a, 7, 'Bowled') #dismissalType
                    analysis.write(a, 8, (dismiss[start+1:])) #bowler  
            elif 'lbw' in distype:
                start = dismiss.index("lbw ") + 3
                analysis.write(a, 7, 'LBW') #dismissalType
                analysis.write(a, 8, (dismiss[start+1:])) #bowler
            elif 'NOT' in distype:
                analysis.write(a, 7, 'Not Out') #dismissalType
            elif 'run' in distype:
                start = dismiss.index("(")
                end = dismiss.index(")")
                analysis.write(a, 7, 'Run Out') #dismissalType
                check = dismiss[start+1:end]
                if '/' in check:
                    end = check.index('/')
                    analysis.write(a, 9, (check[:end]))
                
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
            #setting 0 for values that don't apply here
            analysis.write(a, 15, 0)
            analysis.write(a, 16, 0)
            analysis.write(a, 17, 0)
            analysis.write(a, 18, 0)
            analysis.write(a, 19, 0)    

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
            #setting 0 for values that don't apply here
            analysis.write(a, 10, 0)
            analysis.write(a, 11, 0)
            analysis.write(a, 12, 0)
            analysis.write(a, 13, 0)
            analysis.write(a, 14, 0)
        a += 1
        
    elif check_term == 'DID NOT BAT:':
        m = i+1
        while(raw_data.cell_value(m, 1) != 'Bowler'):
            if raw_data.cell_value(m+1, 1) == 'Bowler':
                analysis.write(a, 5, (raw_data.cell_value(m, 1)))
            else : 
                analysis.write(a, 5, (raw_data.cell_value(m, 1))[:-2])
            analysis.write(a, 1, team1)
            analysis.write(a, 2, team2)
            analysis.write(a, 0, mno)
            analysis.write(a, 10, 0)
            analysis.write(a, 11, 0)
            analysis.write(a, 12, 0)
            analysis.write(a, 13, 0)
            analysis.write(a, 14, 0)
            analysis.write(a, 15, 0)
            analysis.write(a, 16, 0)
            analysis.write(a, 17, 0)
            analysis.write(a, 18, 0)
            analysis.write(a, 19, 0)
            m += 1
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

for i in range(0, a-1):
    #removing the extra columns that were added
    if update.cell_value(i, 5) == '': 
        for k in range(0, 22):
                analysis.write(i, k, '')

for i in range(0, a-1):
    #concatenating the stats of 1 match for the same player           
    mno = update.cell_value(i, 0)
    player_name = update.cell_value(i, 5)
    innings = update.cell_value(i, 3)
    category = update.cell_value(i, 6)
    
    for j in range(i, a):
        if update.cell_value(i, 0) != '':
            if mno == update.cell_value(j, 0) and player_name == update.cell_value(j, 5) and innings != update.cell_value(j, 3):
                analysis.write(i, 7, update.cell_value(j, 7))
                analysis.write(i, 8, update.cell_value(j, 8))
                analysis.write(i, 9, update.cell_value(j, 9))
                analysis.write(i, 10, int(update.cell_value(j, 10) + update.cell_value(i, 10)))
                analysis.write(i, 11, int(update.cell_value(j, 11) + update.cell_value(i, 11)))
                analysis.write(i, 12, int(update.cell_value(j, 12) + update.cell_value(i, 12)))
                analysis.write(i, 13, int(update.cell_value(j, 13) + update.cell_value(i, 13)))
                analysis.write(i, 14, int(update.cell_value(j, 14) + update.cell_value(i, 14)))
                analysis.write(i, 15, (update.cell_value(j, 15) + update.cell_value(i, 15)))
                analysis.write(i, 16, int(update.cell_value(j, 16) + update.cell_value(i, 16)))
                analysis.write(i, 17, int(update.cell_value(j, 17) + update.cell_value(i, 17)))
                analysis.write(i, 18, update.cell_value(j, 18) + update.cell_value(i, 18))
                analysis.write(i, 19, int(update.cell_value(j, 19) + update.cell_value(i, 19)))
                
                for k in range(0, 34):
                    analysis.write(j, k) == '' #removing redundant stats
wb.save('analysis_test.xlsx') #saving the file






#Adding fielding stats
for i in range(0, a-1):
    mno = update.cell_value(i, 0)
    player_name = update.cell_value(i, 5)
    num_catch = num_stump = num_runout = 0
    for j in range(0, a-1):
        #adding the number of fielding stats for each player for each mach
        if mno == update.cell_value(j, 0):
            if player_name == update.cell_value(j, 9):
                if update.cell_value(j, 7) == 'Catch':
                    num_catch += 1
                elif update.cell_value(j, 7) == 'Stump':
                    num_stump += 1
                elif update.cell_value(j, 7) == 'Run Out':
                    num_runout += 1
    analysis.write(i, 20, num_catch)
    analysis.write(i, 21, num_stump)
    analysis.write(i, 22, num_runout)
wb.save('analysis_test.xlsx') #saving the file





#adding the player roles and teams from playerDB sheets 
sheet = xlrd.open_workbook('analysis_test.xlsx')
update = sheet.sheet_by_index(0)

player_db = workbook.sheet_by_index(0)
for i in range(0, a-1):
    player_name = update.cell_value(i, 5)
    for j in range(1, 189):
        if player_name == player_db.cell_value(j, 1):
            analysis.write(i, 4, player_db.cell_value(j, 5)) #write player team by crosschecking with playerdb
            analysis.write(i, 6, player_db.cell_value(j, 2)) #write player role by crosschecking with playerdb
            
    mno = update.cell_value(i, 0)
    for j in range(4, 55):
        if mno == player_db.cell_value(j, 22):
            if player_db.cell_value(j, 24) == update.cell_value(i, 4): #added isWinner
                analysis.write(i, 35, 1)
                
            if player_db.cell_value(j, 25) == update.cell_value(i, 5): #added isMVP
                analysis.write(i, 36, 1)

wb.save('analysis_test.xlsx') #saving the file






#Calculate fantasy points using the stats
sheet = xlrd.open_workbook('analysis_test.xlsx')
update = sheet.sheet_by_index(0)

for i in range(0, a-2):
    if update.cell_value(i, 0):
        num_stump = num_runout = num_catch = bowling_impact_pts = bowling_milestone_pts = economy_pts = impact_pts = 0
        extra_bonus_pts = pace_bonus_pts = bowling_base_pts = milestone_pts = 0
        
        base_pts = update.cell_value(i, 10)
        strike_rate = update.cell_value(i, 12)
        num_6s = update.cell_value(i, 14)
        wickets_taken = update.cell_value(i, 17)
        economy_rate = update.cell_value(i, 18)
        
        runs_scored = base_pts
        
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
                milestone_pts = int(runs_scored/25) * 10
        except:
            analysis.write(i, 10, 0)
            
        impact_pts = num_6s * 2 #calculating batsman impact
        if runs_scored == 0 and update.cell_value(i, 11) != 0 and update.cell_value(i, 6) != "Bowler":
            impact_pts -= 5
            
        if wickets_taken: #calculating bowling_base
            bowling_base_pts = wickets_taken * 20
        try:    
            if update.cell_value(i, 15) != 0:
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
                bowling_milestone_pts = 10
            if wickets_taken > 2:
                bowling_milestone_pts += ((wickets_taken - 2) * 10) + 10
        except:
            analysis.write(i, 17, 0)
            
        bowling_impact_pts = update.cell_value(i, 19) #calculating bowling impact points 
        
        fielding_pts = (num_catch * 10) + (num_stump * 15) + (num_runout * 10) #calculating total fielding points      
        total_batting_pts = base_pts + pace_bonus_pts + milestone_pts + impact_pts #calculating total batting points
        total_bowling_pts = economy_pts + bowling_milestone_pts + bowling_impact_pts + bowling_base_pts #calculating total bowling points
        
        if update.cell_value(i, 35) == 1:
            extra_bonus_pts += 5
        if update.cell_value(i, 36) == 1:
                extra_bonus_pts += 25
                
        grand_total_pts = total_batting_pts + total_bowling_pts + fielding_pts
                
        #writing the calculated points on the sheet
        analysis.write(i, 24, base_pts)
        analysis.write(i, 25, pace_bonus_pts)
        analysis.write(i, 26, milestone_pts)
        analysis.write(i, 27, impact_pts)
        analysis.write(i, 28, bowling_base_pts)
        analysis.write(i, 29, economy_pts)
        analysis.write(i, 30, bowling_milestone_pts)
        analysis.write(i, 31, bowling_impact_pts)
        analysis.write(i, 32, fielding_pts)
        analysis.write(i, 33, total_batting_pts)
        analysis.write(i, 34, total_bowling_pts)
        analysis.write(i, 37, extra_bonus_pts)
        analysis.write(i, 38, grand_total_pts)
    
wb.save('analysis_test.xlsx') #saving the file


j = 2
k = 0
for i in range(0, a-1):       
    if update.cell_value(i, 0) != '':
        if k % 11 == 0:
            if j == 1:
                j = 2
            elif j == 2:
                j = 1
        analysis.write(i, 3, j)
        k += 1
                
wb.save('analysis_test.xlsx') #saving the file

#A0 - MATCH NUMBER
#B1 - TEAM1
#C2 - TEAM2
#D3 - INNINGS
#E4 - TEAM OF THE PLAYER 
#F5 - PLAYER NAME
#G6 - PLAYER ROLE
#H7 - DISMISSAL TYPE
#I8 - BOWLER
#J9 - FIELDER
#K10 - RUNS
#L11 - BALLS PLAYED
#M12 - STRIKE RATE
#N13 - NUMBER OF 4S
#O14 - NUMBER OF 6S
#P15 - OVER BOWLED
#Q16 - RUNS AGAINST
#R17 - WICKETS TAKEN 
#S18 - ECONOMY
#T19 - DOT BALLS
#U20 - NUMBER OF CATCHES
#V21 - NUMBER OF STUMPS
#W22 - NUMBER OF RUNOUT
#X23 - 
#Y24 - BATTING BASE POINTS
#Z25 - BATTING PACE BONUS POINTS
#AA26 - BATTING MILESTONE POINTS
#AB27 - BATTING IMPACT POINTS
#AC28 - BOWLING BASE POINTS
#AD29 - BOWLING ECONOMY POINTS
#AE30 - BOWLING MILESTONE POINTS
#AF31 - BOWLING IMPACT POINTS
#AG32 - TOTAL FIELDING POINTS
#AH33 - TOTAL BATTING POINTS
#AI34 - TOTAL BOWLIMG POINTS
#AJ35 - WINNER (YES = 1)        
#AK36 - MVP (YES = 1)
#AL37 - EXTRA BONUS (= WINNER + MVP BONUS)
#AM3 - EXTRA BONUS POINTS(5 FOR WIN, 25 FOR MVP)
#AN38 - GRAND TOTAL POINTS