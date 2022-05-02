# -*- coding: utf-8 -*-
"""
Created on Tue Mar 30 15:52:10 2021

@author: asandhaa
"""


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os


"--------------------------------------------"
"""MANUAL SETTINGS"""
"--------------------------------------------"
#Enter WZ_Code of industry type. See "All_data_industry_types.xlsx" or upload dataframe as "all_info_wz" in GENERAL DATA
industry_type = 10 
#Enter local path, where Excel files and Code is stored
path = "P:\INES\INES\Projekte\GaIN\\5 - Projektarbeit\AP1_Kategorisierung der Nachfrage\\6 - Publication\AP 1\\03_Python_Code"
"--------------------------------------------"


"--------------------------------------------"
"""GENERAL DATA"""
"--------------------------------------------"
"""Uploads the synthetical load profiles for the enduser and converts them into dataframe"""
load_prof_enduser = pd.read_excel(path+'\\Load_profiles_enduser.xlsx', usecols=('A:J'), index_col=(0))
load_prof_enduser = load_prof_enduser.round(2)
load_prof_enduser.drop('Sum', axis=0 ,inplace=True)
load_prof_enduser.dropna(axis=0,inplace=True)
load_prof_enduser.index = pd.to_datetime(load_prof_enduser.index, format ='%H:%M:%S')
load_prof_enduser.index = load_prof_enduser.index.time
load_prof_enduser =load_prof_enduser.iloc[1:]
"""Uploads the energy distribution for enduser and industry types"""
all_info_wz = pd.read_excel(path+'\\All_data_industry_types.xlsx'.format())
all_info_wz.dropna(how='all',axis=0, inplace=True)
all_info_wz.dropna(how='all',axis=1, inplace=True)
all_info_wz.fillna(0, inplace=True)
"--------------------------------------------"

"--------------------------------------------"
"""GENERATION OF SYNTHETIC LOAD PROFILES"""
"--------------------------------------------"
"""Step 2: Generating a basic load profile """
data_industry_type = all_info_wz[all_info_wz.WZ_Code.eq(industry_type)] #filters the row with specific industry_wz
energy_enduser_industry_type = data_industry_type[['Space heating', 'Hot water', 'Process heat', 'Space cooling',
       'Process cooling', 'Lighting', 'ICT', 'Mechanical drives']].values #extracts columns with specific industry_wz
energy_enduser_industry_type = energy_enduser_industry_type.astype(np.float)

"""Step 3: Classifying mechanical drive processes"""
y=pd.DataFrame()
if data_industry_type['Percentage of discontinuous mechanical drive'].item() != 0:
    mechanical_1 = float(data_industry_type['Percentage of discontinuous mechanical drive'])/100*energy_enduser_industry_type.item(7) #float mechanical1
    mechanical_2 = energy_enduser_industry_type.item(7) - mechanical_1  #float mechanical2
    energy_enduser_industry_type2 = np.delete(energy_enduser_industry_type, [7]) #remove 'Mechanische Antriebe'
    energy_enduser_industry_type2 = np.append(energy_enduser_industry_type2, [mechanical_2, mechanical_1]) #add 'stetig# und 'unstetig'
    
    y = load_prof_enduser.mul(energy_enduser_industry_type2, axis=1) #Multiplying daily profile with industry type data
    y=y.round(2)

elif data_industry_type['Percentage of discontinuous mechanical drive'].item() == 0:
    load_prof_enduser2 = load_prof_enduser.drop(columns= 'Discontinuous mechanical drives')
    load_prof_enduser2 = load_prof_enduser2.rename(columns={'Continuous mechanical drives': 'Mechanical drives'})
    
    y = load_prof_enduser2.mul(energy_enduser_industry_type, axis=1) #Multiplying daily profile with industry type data
    y=y.round(2)
    
"""Step 4: Applying a fluctuation rate"""

if data_industry_type['Percentage of discontinuous mechanical drive'].item() !=0:
    fluc_industry_type = float(data_industry_type['Fluctuations'])
    s = np.random.normal(0, fluc_industry_type, len(y.index)).round(2) #random numbers generated and saved as s. Always refresh this, to gain new numbers
    y['Discontinuous mechanical drives'] = y['Discontinuous mechanical drives'] + s
      
    sec_row = [{'Space heating' : 'in kW', 'Hot water' : 'in kW',	
               'Process heat' : 'in kW',	'Space cooling' : 'in kW',	
               'Process cooling' : 'in kW',	'Lighting' : 'in kW',	
               'ICT' : 'in kW',	'Continuous mech. drive' : 'in kW',	
               'Discontinuous mech. drive' : 'in kW'}]    
    
elif data_industry_type['Percentage of discontinuous mechanical drive'].item()  == 0:
    fluc_industry_type = float(data_industry_type['Fluctuations'])
    s = np.random.normal(0, fluc_industry_type, len(y.index)).round(2) #random numbers generated and saved as s. Always refresh this, to gain new numbers
    y['Mechanical drives'] = y['Mechanical drives'] + s
      
    sec_row = [{'Space heating' : 'in kW', 'Hot water' : 'in kW',	
               'Process heat' : 'in kW',	'Space cooling' : 'in kW',	
               'Process cooling' : 'in kW',	'Lighting' : 'in kW',	
               'ICT' : 'in kW',	'Mechanical drive' : 'in kW'}]
"--------------------------------------------"


"--------------------------------------------"
'''Plotting and saving diagram'''
if data_industry_type['Percentage of discontinuous mechanical drive'].item() !=0:
    
    def diagram(y, industry_type):
        
        industry_name = str(data_industry_type.iloc[0]['Industry_name']) #filters the row with specific industry_wz
        
        new_colors = [ (254/255, 188/255, 195/255),     #Farbe Raumwärme RGB/255
         (255/255, 89/255, 105/255),                    #Farbe Warmwasser 
         (172/255, 0/255, 16/255),                      #Farbe Prozesswärme
         (82/255, 203/255, 190/255),                    #Farbe Klimakälte
         (49/255, 164/255, 151/255),                    #Farbe Prozesskälte
         (254/255, 198/255, 48/255),                    #Farbe Beleuchtung
         (146/255, 208/255, 80/255),                    #Farbe IKT
         (93/255, 115/255, 115/255),                    #Farbe Mechanische Antriebe
         (200/255, 200/255, 200/255)]                   #Farbe unstetige mechanische Antriebe
        
        
        labels = ['Space heating', 'Hot water', 'Process heat', 'Space cooling',
           'Process cooling', 'Lighting', 'ICT', 'Continuous mechanical drives',
           'Discontinuous mechanical drives']
          
        x_plot = y.index
        x_plot=[x.strftime('%H:%M') for x in x_plot]
        y_plot = np.vstack([y['Space heating'].tolist(), 
                    y['Hot water'].tolist(),
                    y['Process heat'].tolist(),
                    y['Space cooling'].tolist(),
                    y['Process cooling'].tolist(),
                    y['Lighting'].tolist(),
                    y['ICT'].tolist(),
                    y['Continuous mechanical drives'].tolist(),
                    y['Discontinuous mechanical drives'].tolist()])
        
        fig, ax = plt.subplots(figsize= (15,10))              
        ax.stackplot(x_plot, y_plot, labels=labels, colors = new_colors, )
        ax.set_xlabel('Time',fontsize=18)
        ax.set_ylabel('Normalized Power in kW',fontsize=18)
        plt.grid()
        plt.xticks(x_plot[::16], fontsize=18)
        plt.yticks(fontsize=18)
        plt.legend(reversed(plt.legend().legendHandles), reversed(labels), loc='upper right',fontsize=16)
        plt.xlim(left=0, xmax=max(x_plot))
        plt.ylim(bottom =0, ymax=160)
        #plt.title('Synthetical load profile of WZ08 '+ str(industry_type) + ' '+ industry_name, fontsize=20)
        plt.savefig(path +'\\Results\\'+ str(industry_type) +' Diagram.png')
        return plt.show()
    
    diagram(y, industry_type)
    sec_row = pd.DataFrame(sec_row)
    data=pd.concat([sec_row, y])
    if not os.path.exists(path +'\\Results\\'):    
        os.mkdir(path +'\\Results')    
    data.to_excel(path +'\\Results\\' +str(industry_type)+ ' Dataframe.xlsx')



"--------------------------------------------"

   
    


