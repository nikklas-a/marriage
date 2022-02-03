import matplotlib
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.ticker as ticker
from matplotlib.patches import Ellipse
from matplotlib.text import OffsetFrom
from matplotlib import pyplot as plt
import numpy as np
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl import Workbook
import os, shutil
import os
from os import listdir
from os.path import isfile, join
from time import time
from datetime import datetime
import matplotlib.pyplot as plt

#Data-files
data_divorce = pd.ExcelFile('divorce.xlsx')
data_marriage = pd.ExcelFile('marrige.xlsx')
data_nr_divorces = pd.ExcelFile('divorces.xlsx')

#Read data-files
df_divorce = pd.read_excel(data_divorce, sheet_name="1", skiprows=1)
df_marriage = pd.read_excel(data_marriage, sheet_name="1", skiprows=1)
df_divorces = pd.read_excel(data_nr_divorces, sheet_name="1", skiprows=1)


filtered_df_divorce = df_divorce
filtered_df_marriage = df_marriage
filtered_df_divorces = df_divorces

#Actual data to plot
filtered_df_marriage_data_year = df_marriage['år'].tail(21)
filtered_df_marriage_data_women = df_marriage['kvinnor'].tail(21)
filtered_df_divorce_year = df_divorce['år'].tail(21)
filtered_df_divorce_length = df_divorce['längd'].tail(21)
filtered_df_divorces_year = df_divorces['år']
filtered_df_divorces_nr = df_divorces['antal, tusen']

#Start year
start_year = filtered_df_divorce.iloc[2][0].astype(int)

#Last year
end_year = filtered_df_marriage.iloc[149][0].astype(int)

#Complete timeline
timeline = filtered_df_marriage['år'].head(150)

#Graph layout
#------------------------------------------------

fig = plt.figure(figsize = (12,7))

ax  = fig.add_subplot(111) #Bars
ax2 = ax.twinx()
#ax3 = ax.twinx()

#Fonts for graph
heading_font = fm.FontProperties(fname='./static/fonts/playfair-display/PlayfairDisplay-Regular.ttf', size=22)
subtitle_font = fm.FontProperties(fname='./static/fonts/Roboto/Roboto-Regular.ttf', size=12)

#Color Themes
color_bg = '#FEF1E5'
lighter_highlight = '#FAE6E1'
darker_highlight = '#FBEADC'
bar_color ='#437BBB'
line_color ='#F8483C'
line_color2 ='#000000'
legend_bg = '#FFFFFF'

#Bar width
barwidth = 0.7

#Background color
fig.patch.set_facecolor(color_bg)
ax.set_facecolor(darker_highlight)

#Remove/Show borders
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['bottom'].set_visible(True)
ax.spines['left'].set_visible(False)

#Annotations
fig.text(x=0.05, y=1.01, s=f'Marriage & Divorce from {start_year} to {end_year}',fontproperties=heading_font, horizontalalignment='left',color='#524939')    
fig.text(x=0.8, y=1.01, s='FamilStat 0.1',fontproperties=subtitle_font, horizontalalignment='left',color='#524939')    

#Format x-axel
plt.setp(ax.get_xticklabels(), rotation=30, ha="right") #rotates the x-axis labels 30 degrees

#Axis labels
#ax.set(xlabel='', ylabel='Number of divorces, thousands')
ax2.set(xlabel='Year', ylabel='Years / Age / # of divorces, thousands')

#Axis limits
ax.set_ylim(0,40) #Number of divorces
ax2.set_ylim(0,40)  #Marriage Age?

#Legend(s)
legendlabel_1 = "Average age for first marriage (swedish women)"
legendlabel_2 = "Average duration of marrige before divorce, years"
legendlabel_3 = "Divorces in Sweden, thousands"

#acutal bars and lines
plt.plot(filtered_df_marriage_data_year, filtered_df_marriage_data_women, scalex=True, label=legendlabel_1, color=line_color, lw=3)
plt.plot(filtered_df_divorce_year, filtered_df_divorce_length, label=legendlabel_2, scalex=True, color=line_color2, lw=3, ls='--',)
ax.bar(filtered_df_divorces_year, filtered_df_divorces_nr, barwidth, align = 'center', label=legendlabel_3, color=bar_color)

#X-axis ticks
plt.xticks(np.arange(min(filtered_df_marriage_data_year), max(filtered_df_marriage_data_year)+1, 1.0))

#legend
fig.legend(loc="upper left", shadow=False, bbox_to_anchor=(0.05,0.98), bbox_transform=ax.transAxes, facecolor=legend_bg, borderpad=1.2)

#y ticker
ax.tick_spacing = 2
ax2.tick_spacing = 2
ax.yaxis.set_major_locator(ticker.MultipleLocator(ax.tick_spacing))
ax2.yaxis.set_major_locator(ticker.MultipleLocator(ax2.tick_spacing))


ax.yaxis.grid()  # grid lines
ax.set_axisbelow(True)  # grid lines are behind the rest
ax.set_xlim(start_year-1)

#Draw the chart and print to file
plt.tight_layout()
plt.savefig(f'marriageanddivorce.png', facecolor=fig.get_facecolor(), edgecolor='none', bbox_inches="tight")
plt.show()