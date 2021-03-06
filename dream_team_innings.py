#importing libraries
import xlrd #reading files
import xlwt #creating files

#loading file
workbook = xlrd.open_workbook('ipl_scores.xlsx') 
wb = xlwt.Workbook()

raw_data = workbook.sheet_by_index(2)
analysis = wb.add_sheet('dream_team', cell_overwrite_ok = True)

#working on new file
mno = 1
rows = 1
try:
    while mno:
        mno = raw_data.cell_value(rows-1, 0)   
        rows += 1
except:
    i = 0





a = 0
n = 1
inn = 0
i = 1

#try:
    for i in range(1,61):
        player_name = []
        player_score = []
        player_role = []
        
        while raw_data.cell_value(n, 0) == i and raw_data.cell_value(n, 3) == 1:
            name = raw_data.cell_value(n, 5)
            score = raw_data.cell_value(n, 38)
            role = raw_data.cell_value(n, 6)
            player_name.append(name)
            player_score.append(score)
            player_role.append(role)
            
            n += 1
            inn = 1
                    
        if inn == 1:
            for k in range(0, 11):        
                for l in range(0, 10):
                    if player_score[l] < player_score[l + 1]:
                        temp = player_name[l + 1]
                        player_name[l + 1] = player_name[l]
                        player_name[l] = temp
                        
                        temp = player_score[l + 1]
                        player_score[l + 1] = player_score[l]
                        player_score[l] = temp
                        
                        temp = player_role[l + 1]
                        player_role[l + 1] = player_role[l]
                        player_role[l] = temp
                            
            no_batsmen = 0
            no_bowlers = 0
            no_all_rounders = 0
            no_wicket_keepers = 0
            dream_team = []
            total_dream_players = 0
            
            for j in range(0, 11):
                if player_role[j] == 'Batsman' and no_batsmen < 2:
                    dream_team.append(player_name[j])
                    analysis.write(a, 0, i)
                    analysis.write(a, 1, player_name[j])
                    analysis.write(a, 2, player_role[j])
                    analysis.write(a, 3, player_score[j])
                    analysis.write(a, 4, 1)
                    no_batsmen += 1
                    a += 1  
    #            j += 1    
                        
            for j in range(0, 11):
                if player_role[j] == 'Bowler' and no_bowlers < 2:
                    dream_team.append(player_name[j])
                    analysis.write(a, 0, i)
                    analysis.write(a, 1, player_name[j])
                    analysis.write(a, 2, player_role[j])
                    analysis.write(a, 3, player_score[j])
                    analysis.write(a, 4, 1)
                    no_bowlers += 1
                    a += 1
    #            j += 1
                    
            total_dream_players = no_batsmen + no_bowlers + no_wicket_keepers + no_all_rounders
                        
            for j in range(0,11):
                if total_dream_players < 6:
                    if player_role[j] == 'Batsman' and no_batsmen < 4 and player_name[j] not in dream_team:
                        dream_team.append(player_name[j])
                        analysis.write(a, 0, i)
                        analysis.write(a, 1, player_name[j])
                        analysis.write(a, 2, player_role[j])
                        analysis.write(a, 3, player_score[j])
                        analysis.write(a, 4, 1)
                        no_batsmen += 1
                        total_dream_players += 1
                        a += 1        
                    
                    if player_role[j] == 'Bowler' and no_bowlers < 4 and player_name[j] not in dream_team:
                        dream_team.append(player_name[j])
                        analysis.write(a, 0, i)
                        analysis.write(a, 1, player_name[j])
                        analysis.write(a, 2, player_role[j])
                        analysis.write(a, 3, player_score[j])
                        analysis.write(a, 4, 1)
                        no_bowlers += 1
                        total_dream_players += 1
                        a += 1
                        
                    if player_role[j] == 'All Rounder' and no_all_rounders < 1 and player_name[j] not in dream_team:
                        dream_team.append(player_name[j])
                        analysis.write(a, 0, i)
                        analysis.write(a, 1, player_name[j])
                        analysis.write(a, 2, player_role[j])
                        analysis.write(a, 3, player_score[j])
                        analysis.write(a, 4, 1)
                        no_all_rounders += 1
                        total_dream_players += 1
                        a += 1
                    
                    if player_role[j] == 'Wicket Keeper' and no_wicket_keepers < 1 and player_name[j] not in dream_team:
                        analysis.write(a, 0, i)
                        analysis.write(a, 1, player_name[j])
                        analysis.write(a, 2, player_role[j])
                        analysis.write(a, 3, player_score[j])
                        analysis.write(a, 4, 1)
                        no_wicket_keepers = 1
                        total_dream_players += 1
                        a += 1
        
        wb.save('dream_team_innings.xlsx')
                                                
        player_name = []
        player_score = []
        player_role = []
        
        try:
            while raw_data.cell_value(n, 0) == i and raw_data.cell_value(n, 3) == 2:        
                name = raw_data.cell_value(n, 5)
                score = raw_data.cell_value(n, 38)
                role = raw_data.cell_value(n, 6)
                player_name.append(name)
                player_score.append(score)
                player_role.append(role)
                
                n += 1
                inn = 2
        except:
            rand = 0
    
        if inn == 2:
            for k in range(0, 11):        
                for l in range(0, 10):
                    if player_score[l] < player_score[l + 1]:
                        temp = player_name[l + 1]
                        player_name[l + 1] = player_name[l]
                        player_name[l] = temp
                        
                        temp = player_score[l + 1]
                        player_score[l + 1] = player_score[l]
                        player_score[l] = temp
                        
                        temp = player_role[l + 1]
                        player_role[l + 1] = player_role[l]
                        player_role[l] = temp
                            
            no_batsmen = 0
            no_bowlers = 0
            no_all_rounders = 0
            no_wicket_keepers = 0
            dream_team = []
            total_dream_players = 0
            
            for j in range(0, 11):
                if player_role[j] == 'Batsman' and no_batsmen < 2:
                    dream_team.append(player_name[j])
                    analysis.write(a, 0, i)
                    analysis.write(a, 1, player_name[j])
                    analysis.write(a, 2, player_role[j])
                    analysis.write(a, 3, player_score[j])
                    analysis.write(a, 4, 2)
                    no_batsmen += 1
                    a += 1  
    #            j += 1    
                        
            for j in range(0, 11):
                if player_role[j] == 'Bowler' and no_bowlers < 2:
                    dream_team.append(player_name[j])
                    analysis.write(a, 0, i)
                    analysis.write(a, 1, player_name[j])
                    analysis.write(a, 2, player_role[j])
                    analysis.write(a, 3, player_score[j])
                    analysis.write(a, 4, 2)
                    no_bowlers += 1
                    a += 1
    #            j += 1
                    
            total_dream_players = no_batsmen + no_bowlers + no_wicket_keepers + no_all_rounders
                        
            for j in range(0,11):
                if total_dream_players < 6:
                    if player_role[j] == 'Batsman' and no_batsmen < 4 and player_name[j] not in dream_team:
                        dream_team.append(player_name[j])
                        analysis.write(a, 0, i)
                        analysis.write(a, 1, player_name[j])
                        analysis.write(a, 2, player_role[j])
                        analysis.write(a, 3, player_score[j])
                        analysis.write(a, 4, 2)
                        no_batsmen += 1
                        total_dream_players += 1
                        a += 1        
                    
                    if player_role[j] == 'Bowler' and no_bowlers < 4 and player_name[j] not in dream_team:
                        dream_team.append(player_name[j])
                        analysis.write(a, 0, i)
                        analysis.write(a, 1, player_name[j])
                        analysis.write(a, 2, player_role[j])
                        analysis.write(a, 3, player_score[j])
                        analysis.write(a, 4, 2)
                        no_bowlers += 1
                        total_dream_players += 1
                        a += 1
                        
                    if player_role[j] == 'All Rounder' and no_all_rounders < 1 and player_name[j] not in dream_team:
                        dream_team.append(player_name[j])
                        analysis.write(a, 0, i)
                        analysis.write(a, 1, player_name[j])
                        analysis.write(a, 2, player_role[j])
                        analysis.write(a, 3, player_score[j])
                        analysis.write(a, 4, 2)
                        no_all_rounders += 1
                        total_dream_players += 1
                        a += 1
                    
                    if player_role[j] == 'Wicket Keeper' and no_wicket_keepers < 1 and player_name[j] not in dream_team:
                        analysis.write(a, 0, i)
                        analysis.write(a, 1, player_name[j])
                        analysis.write(a, 2, player_role[j])
                        analysis.write(a, 3, player_score[j])
                        analysis.write(a, 4, 2)
                        no_wicket_keepers = 1
                        total_dream_players += 1
                        a += 1        
#except:
#    print("Error")
#    
        
wb.save('dream_team_innings.xlsx')

            
            
    
#Batsmen = 2 - 4
#All rounder = 0 - 1
#Wicket keeper = 0 - 1
#Bowlers = 2 - 4