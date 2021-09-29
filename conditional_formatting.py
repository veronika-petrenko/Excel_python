#already created df, save to excel, for the sheet df3 add conditional formatting with this rules:
# green when the value is between -5% and 5%, yellow when it is between -10% and - 5% and 5% and 10% and red when lower than -10% or higher than 10%
#convert results to %

from datetime import datetime
import pandas as pd
import gspread
import numpy as np
import xlsxwriter
from datetime import datetime


game = 'Some_game'
TODAY = datetime.today().strftime('%Y-%m-%d')

writer = pd.ExcelWriter(f'{game}_Prices_GooglePlay_{TODAY}.xlsx', engine='xlsxwriter')
df1.to_excel(writer, sheet_name= 'Prices_local')
df2.to_excel(writer, sheet_name = 'Prices_USD')
df3.to_excel(writer, sheet_name = 'Difference (local and USD)')

workbook  = writer.book
worksheet = writer.sheets['Difference (local and USD)']

format_red = workbook.add_format({'bg_color': '#f27b77'})
format_green = workbook.add_format({'bg_color': '#5ed16a'})
format_yellow = workbook.add_format({'bg_color': '#edde34'})

worksheet.conditional_format(1,1,df3.shape[0],df3.shape[1], {'type':     'cell',
                                        'criteria': 'between',
                                          'minimum': -0.05,
                                          'maximum':0.05,
                                        'format':   format_green})

worksheet.conditional_format(1,1,df3.shape[0],df3.shape[1], {'type':     'cell',
                                        'criteria': 'between',
                                          'minimum': -0.1,
                                          'maximum':-0.05,
                                        'format':   format_yellow})


worksheet.conditional_format(1,1,df3.shape[0],df3.shape[1], {'type':     'cell',
                                        'criteria': 'between',
                                          'minimum': 0.05,
                                          'maximum': 0.1,
                                        'format':   format_yellow})


worksheet.conditional_format(1,1,df3.shape[0],df3.shape[1], {'type':     'cell',
                                        'criteria': '<',
                                          'value': -0.1,
                                        'format':   format_red})

worksheet.conditional_format(1,1,df3.shape[0],df3.shape[1], {'type':     'cell',
                                        'criteria': '>',
                                          'value': 0.1,
                                        'format':   format_red})

perc_fmt = workbook.add_format({'num_format': '0.00%'})


worksheet.conditional_format(1,1,df3.shape[0],df3.shape[1],{
                                        'type': 'cell',
                                        'criteria': 'between',
                                        'minimum': -10000,
                                        'maximum': 10000,
                                        'format': perc_fmt
                                    } )
writer.save() 
