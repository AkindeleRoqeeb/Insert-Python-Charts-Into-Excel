
How to easily insert Python charts into Excel and generate professional data visualizations for your work.<br>
With xlwings, you can run python code in excel without any difficulty at all.<br>
<br>
Check out the xlwings documentation here: https://docs.xlwings.org/en/stable/matplotlib.html#<br>
<br>
Import/Install Dependencies &amp; Load Dataset<br>
<br>
In [1]:<br>
<br>
import xlwings as xw # pip install xlwings --upgrade import pandas as pd # pip install pandas <br>
<br>
NOTE: Ensure to use the latest xlwings version to have access the newest features.<br>
<br>
C:\Users\YOUR_USERNAME&gt; pip install xlwings --upgrade <br>
<br>
In [2]:<br>
<br>
# Create an empty workbook &amp; rename sheet wb = xw.Book() sht = wb.sheets[0] sht.name = &quot;Python Charts&quot; <br>
<br>
In [3]:<br>
<br>
# Helper function to insert &apos;Headings&apos; into Excel cells def insert_heading(rng, text): rng.value = text rng.font.bold = True rng.font.size = 24 rng.font.color = (0, 0, 139) <br>
<br>
In [4]:<br>
<br>
# Load seaborn &apos;tips&apos; dataset df = pd.read_csv(&quot;https://raw.githubusercontent.com/mwaskom/seaborn-data/master/tips.csv&quot;) df <br>
<br>
Out[4]:<br>
<br>
total_billtipsexsmokerdaytimesize016.991.01FemaleNoSunDinner2110.341.66MaleNoSunDinner3221.013.50MaleNoSunDinner3323.683.31MaleNoSunDinner2424.593.61FemaleNoSunDinner4........................23929.035.92MaleNoSatDinner324027.182.00FemaleYesSatDinner224122.672.00MaleYesSatDinner224217.821.75MaleNoSatDinner224318.783.00FemaleNoThurDinner2<br>
<br>
244 rows × 7 columns<br>
<br>
Matplotlib Chart<br>
<br>
In [5]:<br>
<br>
insert_heading(sht.range(&quot;A2&quot;), &quot;Matplotlib Chart&quot;) <br>
<br>
Generate Chart<br>
<br>
In [6]:<br>
<br>
import matplotlib.pyplot as plt # pip install matplotlib fig = plt.figure() x = df[&quot;day&quot;] y = df[&quot;total_bill&quot;] plt.bar(x, y) plt.grid(False) plt.ylabel(&quot;in USD&quot;) plt.title(&quot;Total Bill Amount By Day&quot;) <br>
<br>
Out[6]:<br>
<br>
Text(0.5, 1.0, &apos;Total Bill Amount By Day&apos;)<br>
<br>
<img src="content://com.samsung.android.memo/file/acf52803-0752-a9d1-0000-018440295693" orientation="0" altText="null"   width="608" /><br>
<br>
Insert Chart Into Excel<br>
<br>
In [7]:<br>
<br>
sht.pictures.add( fig, name=&quot;Matplotlib&quot;, update=True, left=sht.range(&quot;A4&quot;).left, top=sht.range(&quot;A4&quot;).top, height=200, width=300, ) <br>
<br>
Out[7]:<br>
<br>
&gt;<br>
<br>
Pandas Chart<br>
<br>
In [8]:<br>
<br>
insert_heading(sht.range(&quot;A19&quot;), &quot;Pandas Chart&quot;) <br>
<br>
Generate Chart<br>
<br>
In [9]:<br>
<br>
df_grouped = df.groupby(by=&quot;day&quot;, as_index=False).sum() ax = df_grouped.plot(kind=&quot;bar&quot;, x=&quot;day&quot;, y=&quot;tip&quot;, color=&quot;#50C878&quot;, grid=False) <br>
<br>
<img src="content://com.samsung.android.memo/file/acf52803-0752-a9d1-0000-0184402956ec" orientation="0" altText="null"   width="608" /><br>
<br>
Insert Chart Into Excel<br>
<br>
In [10]:<br>
<br>
fig = ax.get_figure() sht.pictures.add( fig, name=&quot;Pandas&quot;, update=True, left=sht.range(&quot;A21&quot;).left, top=sht.range(&quot;A21&quot;).top, height=200, width=300, ) <br>
<br>
Out[10]:<br>
<br>
&gt;<br>
<br>
Seaborn Chart<br>
<br>
In [11]:<br>
<br>
insert_heading(sht.range(&quot;A35&quot;), &quot;Seaborn Chart&quot;) <br>
<br>
Generate Chart<br>
<br>
In [12]:<br>
<br>
import seaborn as sns # pip install seaborn fig = plt.figure() sns.set_style({&apos;axes.grid&apos; : False}) sns.barplot(data=df, x=&quot;day&quot;, y=&quot;total_bill&quot;, hue=&quot;sex&quot;, ci=None) <br>
<br>
Out[12]:<br>
<br>
<img src="content://com.samsung.android.memo/file/acf52803-0752-a9d1-0000-01844029572c" orientation="0" altText="null"   width="608" /><br>
<br>
Insert Chart Into Excel<br>
<br>
In [13]:<br>
<br>
sht.pictures.add( fig, name=&quot;Seaborn1&quot;, update=True, left=sht.range(&quot;A37&quot;).left, top=sht.range(&quot;A37&quot;).top, height=200, width=300, ) <br>
<br>
Out[13]:<br>
<br>
&gt;<br>
<br>
Generate Chart<br>
<br>
In [14]:<br>
<br>
fig = plt.figure() sns.scatterplot(data=df, x=&quot;total_bill&quot;, y=&quot;tip&quot;, hue=&quot;day&quot;, style=&quot;time&quot;) <br>
<br>
Out[14]:<br>
<br>
<img src="content://com.samsung.android.memo/file/acf52803-0752-a9d1-0000-018440295777" orientation="0" altText="null"   width="608" /><br>
<br>
Insert Chart Into Excel<br>
<br>
In [15]:<br>
<br>
sht.pictures.add( fig, name=&quot;Seaborn2&quot;, update=True, left=sht.range(&quot;A51&quot;).left, top=sht.range(&quot;A51&quot;).top, height=200, width=300, ) <br>
<br>
Out[15]:<br>
<br>
&gt;<br>
<br>
Not all Seaborn Charts seems to be supported 😯<br>
<br>
NOTE: Not all seaborn charts seems to be supported by xlwings. For instance, I had issues inserting a seaborn &apos;pairplot&apos; into Excel.<br>
<br>
Plotly Chart<br>
<br>
🔥 Plotly Charts 🔥<br>
Since v0.24.0 (Jun 25, 2021), support for Plotly images was moved from xlwings PRO to the Open Source version. 🎉<br>
<br>
In addition to plotly, you will need kaleido, psutil, and requests. The easiest way to get it is via pip:<br>
<br>
C:\Users\YOUR_USERNAME&gt; pip install kaleido psutil requests <br>
<br>
In [16]:<br>
<br>
insert_heading(sht.range(&quot;A66&quot;), &quot;Plotly Chart&quot;) <br>
<br>
Generate Chart<br>
<br>
In [17]:<br>
<br>
import plotly.express as px # pip install plotly-express fig = px.histogram(df, x=&quot;day&quot;, y=&quot;total_bill&quot;, color=&quot;sex&quot;) fig <br>
<br>
Insert Chart Into Excel<br>
<br>
In [18]:<br>
<br>
sht.pictures.add( fig, name=&quot;Plotly1&quot;, update=True, left=sht.range(&quot;A68&quot;).left, top=sht.range(&quot;A68&quot;).top, height=200, width=300, ) <br>
<br>
Out[18]:<br>
<br>
&gt;<br>
<br>
Generate Chart<br>
<br>
In [19]
