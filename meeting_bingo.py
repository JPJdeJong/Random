# -*- coding: utf-8 -*-
"""
Created on Fri Jan 22 21:59:02 2021

@author: JP
"""

import random
import numpy as np
import pandas as pd
import os
import xlsxwriter

cwd = os.getcwd()
os.chdir(cwd)

words  = ['Nice', 'Fantastic', 'Corona', 'Covid', 'Home Office', 'Resilient', 'Daycare',\
         'Trainees', 'Network', 'Plan', 'Horizon', 'Pushing', 'Employer', 'Professional', 'Consultant',\
         'We', 'Employee', 'Board', '2021', '2020', '2022', 'Goals', 'Target', 'Connect', 'Zoom', 'Internet',\
         'Talent', 'Mission', 'Vision','Remark', 'Tech', 'Personal', 'Ambition', 'IT','Agile', 'Epic', \
         'Milestone', 'Bad connection']
words = sorted(words)
names = ['Rick', 'Roll', 'Never', 'Gonna', 'Give', 'You', 'Up']

print ('Amount contenders: '  + str(len(names)))
print("Amount words: " + str(len(words)))
print(cwd)

card_list = []
#Get size of the card by input from user. (cols and rows)
while True:
    try:
        size_row = int(input("How many rows should the bingo card have? "))
        size_col  = int(input("How many rows should the bingo card have?"))
        break
    except ValueError:
        #error handling if input is not a number, ask again.
        print("Please enter an integer\n")

#make size of card.
card_words_update = int(size_col * size_row) #necessary for random sample
print((size_col, size_row), "Card Size = ", card_words_update)
#make a card for every participant
for i in names:
    try:
        bingo_card = random.sample(words, k=card_words_update) #create card with words
        card_final=pd.DataFrame(np.array(bingo_card).reshape(size_col,size_row)) #reshape to array of col x row
        card_list.append(card_final) #add to list of df's 
    except ValueError:
        print("\nAmount of words on card exceeds possible words. Please add more words to the list! \n")
        break

#open an excel to save the cards.
writer = pd.ExcelWriter('Meeting_presentation_bingo.xlsx', engine='xlsxwriter')

counter = 0

for df in card_list:
    df.to_excel(writer, index = 0, sheet_name = names[counter])
    worksheet = writer.sheets[names[counter]]  # pull worksheet object
    #update column width for card
    for idx, col in enumerate(df):  # loop through all columns
        series = df[col]
        max_len = max((
            series.astype(str).map(len).max(),  # len of largest item
            len(str(series.name))  # len of column name/header
            )) + 1  # adding a little extra space
        worksheet.set_column(idx, idx, max_len)  # set column width
    counter+=1
writer.save()

print("\nCards are saved in the following folder:  \n\n"
      +str(cwd)
     )