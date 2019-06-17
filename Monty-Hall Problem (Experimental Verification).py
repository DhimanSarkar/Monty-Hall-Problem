import xlsxwriter
import random
import os
import time

i=1
dataset = input("Number of dataset :: ")
dataset = int(dataset)
os.system("cls")

# Excel Export #
workbook = xlsxwriter.Workbook('Monty-Hall Problem (Experimental Data).xlsx')
worksheet = workbook.add_worksheet('Monty-Hall Problem')
worksheet.write(0, 0, 'Observation Count')
worksheet.write(0, 1, 'Prize Location')
worksheet.write(0, 2, "Player's First Choice")
worksheet.write(0, 3, 'Host Revealed')
worksheet.write(0, 4, "Player's Second Choice")
worksheet.write(0, 5, "Switched")
worksheet.write(0, 6, "Won/Loss")

while (i <= dataset):
    os.system("color F5")
    print ("\n")
    print (" ----------------------------------------------- ")
    print ("  Monty Hall Problem (Experimental Verification)")
    print (" ----------------------------------------------- ")

    player_options = ['A', 'B', 'C']                # Player choice of doors > A, B, C

    prize = random.choice(player_options)      # Computer calculates random between A,B,C to store the prize

    host_options = ['A', 'B', 'C']
    host_options.remove(prize)                      # Host's options to reveal doors

    print ("\n")
################################################################################################################
    #player_choice1 = input("Choose your door :: ").upper()  # Player's first Choice as uppercase
    player_choice1 = random.choice(player_options)  # Randomized Player's first Choice as uppercase
################################################################################################################

    if player_choice1 == host_options[0]:
        host_choice = host_options[1]               # Host reveals Door
    elif player_choice1 == host_options[1]:
        host_choice = host_options[0]               # Host reveals Door
    else:
        host_choice = random.choice(host_options)   # Host reveals Door

    print ("\n\n")
    print ("Host revealed the DOOR %s and there's NOTHING beghind DOOR %s" %(host_choice, host_choice))
    print ("\n\n")
        
    player_options2 = ['A', 'B', 'C']
    player_options2.remove(host_choice) # Players second choice -- for random value generation

#############################################################################################################################
    #player_choice2 = input("New choice of door :: ").upper()  # Player Swithching Choice
    player_choice2 = random.choice(player_options2)  # Randomized Player Swithching Choice
#############################################################################################################################

    while True: # loop for illegal choice

        if player_choice2 == player_choice1:
            switching_status = "Did NOT Switched"
            break
            
        if player_choice2 == host_choice:
            print ("Illegal Choice! Try again")
            player_choice2 = input("New choice of door :: ").upper()
            pass
            
        else:
            switching_status = "Switched"
            break
            

    if player_choice2 == prize:
        player_status = "WON"
    else:
        player_status = "LOST"

    # Game Status
    print ("\n\n")
    print ("Options given to the Player :: A, B ,C")
    print ("Prize is behind :: %s " %prize)
    print ("Player Choose :: %s" %player_choice1)
    print ("Game Host reveals an empty door at :: %s" %host_choice)
    print ("\n\n")
    print ("Player %s" %(switching_status))
    print ("Player %s" %(player_status))
    print ("\n\n")

    
    # Excel Export -  adding data #
    worksheet.write(i, 0, i)
    worksheet.write(i, 1, prize)
    worksheet.write(i, 2, player_choice1)
    worksheet.write(i, 3, host_choice)
    worksheet.write(i, 4, player_choice2)
    if switching_status == "Switched":
        switching_status = True
    elif switching_status == "Did NOT Switched":
        switching_status = False
    else:
        pass
    worksheet.write(i, 5, switching_status)
    worksheet.write(i, 6, player_status)


    i=i+1 # next game
    #time.sleep(1)
    os.system("cls")

workbook.close()