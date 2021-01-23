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
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def make_bingo_card(rows, cols, card_words, names = participants, words = words):
    #make a card for every participant
    card_list=[]
    for i in names:
        try:
            bingo_card = random.sample(words, k=card_words) #create card with words
            card_final=pd.DataFrame(np.array(bingo_card).reshape(cols,rows)) #reshape to array of col x row
            card_list.append(card_final) #add to list of df's 
        except ValueError:
            print("\nAmount of words on card exceeds possible words. Please add more words to the list! \n")
            break
    return card_list

def cards_to_excel(list_of_cards, name_list = participants):
    #open an excel to save the cards.
    writer = pd.ExcelWriter('Meeting_presentation_bingo.xlsx', engine='xlsxwriter')
    card_list = []
    counter = 0
    for df in list_of_cards:
        df.to_excel(writer, index = 0, sheet_name = name_list[counter])
        worksheet = writer.sheets[name_list[counter]]  # pull worksheet object
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
    return writer

def cards_to_pdf(list_of_cards, name_list):
    counter=0
    pdf = PdfPages("Meeting_Bingo.pdf")    
    for df in list_of_cards:
        plt.ioff()
        fig, ax =plt.subplots(figsize=(12,4))
        ax.axis('tight')
        ax.axis('off')
        the_table = ax.table(cellText=df.values,colLabels=df.columns,loc='center')
        the_title = plt.title(name_list[counter])
        pdf.savefig(fig, bbox_inches='tight')
        counter+=1
    pdf.close()
    return pdf

if __name__ == '__main__':
    cwd = os.getcwd()
    os.chdir(cwd)
    words  = sorted(['Nice', 'Fantastic', 'Corona', 'Covid', 'Home Office', 'Resilient', 'Daycare',\
             'Trainees', 'Network', 'Plan', 'Horizon', 'Pushing', 'Employer', 'Professional', 'Consultant',\
             'We', 'Employee', 'Board', '2021', '2020', '2022', 'Goals', 'Target', 'Connect', 'Zoom', 'Internet',\
             'Talent', 'Mission', 'Vision','Remark', 'Tech', 'Personal', 'Ambition', 'IT','Agile', 'Epic', \
             'Milestone', 'Bad connection'])
    participants = ['Rick', 'Roll', 'Never', 'Gonna', 'Give', 'You', 'Up']
    
    print ('Amount participants: '  + str(len(participants)))
    print("Amount words: " + str(len(words)))
    print(cwd)
        #Get size of the card by input from user. (cols and rows)
    while True:
        try:
            size_row = int(input("How many rows should the bingo card have? "))
            size_col  = int(input("How many rows should the bingo card have? "))
            break
        except ValueError:
            #error handling if input is not a number, ask again.
            print("Please enter an integer\n")
    #make size of card.
    card_words = int(size_col * size_row) #necessary for random sample
    print("Card size: " ,(size_col, size_row), "Words on card = ", card_words)
    
    list_bingo_cards = make_bingo_card(size_row, size_col, card_words, participants, words)
    cards_to_excel = cards_to_excel(list_bingo_cards, participants)
    cards_to_pdf = cards_to_pdf(list_bingo_cards, participants)
    
    print("\nCards (PDF or Excel) for your meeting bingo are saved in the following folder:  \n\n" +str(cwd))
