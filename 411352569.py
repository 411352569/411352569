import pandas as pd
import matplotlib.pyplot as plot

df_XLSX=pd.read_excel("https://drive.google.com/uc?id=1R98vTvkpIxPSEBM9ZYqvxi3Vq50zakpM&export=download",header=0,sheet_name="Data")
df_XLSX.to_excel('imdb1000資料集.xlsx')
print(df_XLSX)

df_XLSX=df_XLSX.rename(columns={'Series_Title': '電影片名','Runtime': '影片長度'})
df_XLSX.to_excel('imdb1000資料集.xlsx')
print(df_XLSX)
df_XLSX.info()

#方法一
df_XLSX['Released_Year'].to_string()
df_XLSX['Released_Year']=df_XLSX['Released_Year'].astype(object)
df_XLSX['Gross'] = df_XLSX['Gross'].astype(str)
df_XLSX['Gross'] = df_XLSX['Gross'].str.replace(',', '')
df_XLSX['Gross'] = df_XLSX['Gross'].replace('nan', 0)
df_XLSX['Gross'] = df_XLSX['Gross'].astype(int)
print(df_XLSX)
df_XLSX.info()

#方法二
df_XLSX['Released_Year'].to_string()
df_XLSX['Released_Year']=df_XLSX['Released_Year'].astype(object)
df_XLSX['Gross'] = df_XLSX['Gross'].astype(str)
df_XLSX['Gross'] = df_XLSX['Gross'].str.replace(',', '')
df_XLSX['Gross'] = pd.to_numeric(df_XLSX['Gross'], errors='coerce')
df_XLSX=df_XLSX.dropna(subset=['Gross'])
df_XLSX['Gross'] = df_XLSX['Gross'].astype(int)
print(df_XLSX)
df_XLSX.info()

imdb_drop=df_XLSX.dropna()
imdb_drop.to_excel('imdb_drop.xlsx')
print(imdb_drop)

#方法一
IMDB_Rating=list()
for i in imdb_drop["IMDB_Rating"]:
    if i >=9.0:
         IMDB_Rating.append("A")
    elif i <9.0 and i >=8.5:
         IMDB_Rating.append("B")
    elif i <=8.5 and i >=8.0:
         IMDB_Rating.append("C")
    elif i <8.0:
         IMDB_Rating.append("D")
imdb_drop["評等"]=IMDB_Rating
print(imdb_drop['評等'])

#方法二
imdb_drop['評等'] = imdb_drop['IMDB_Rating'].apply(lambda x: 'A' if x >= 9.0 else 'B' if x >= 8.5 else 'C' if x >= 8.0 else 'D')
print(imdb_drop['評等'])

imdb_drop = imdb_drop.set_index('電影片名')
imdb_drop.to_excel('imdb_drop.xlsx')

imdb_numeric = imdb_drop[['影片長度', 'IMDB_Rating', 'Meta_score', 'No_of_Votes', 'Gross']]
imdb_numeric.to_excel('imdb_numeric.xlsx')

imdb_numeric.mean()
imdb_numeric.std()
imdb_numeric.max()
imdb_numeric.min()
imdb_numeric.cov()
imdb_numeric.corr()

imdb_rate = imdb_drop[['評等']]
imdb_rate.to_excel('imdb_rate.xlsx')

imdb_merge = pd.merge(imdb_rate, imdb_numeric, left_index=True, right_index=True)
imdb_merge.to_excel('imdb_merge.xlsx')

imdb_merge.groupby(by=['評等'])[['影片長度','IMDB_Rating','Meta_score','No_of_Votes','Gross']].mean()

imdb_merge.groupby(by=['評等'])[['影片長度','IMDB_Rating','Meta_score','No_of_Votes','Gross']].count()

#方法一
plot.figure(figsize=(18,12))
tab=imdb_drop['Certificate'].value_counts()
tab.plot.pie(title="認證項目圓餅圖",legend=True)
tab.plot(title='Certificate',fontsize='small')

#方法二
plot.figure(figsize=(18,12))
imdb_drop['Certificate'].value_counts().plot(kind='pie',title="認證項目圓餅圖",legend=True)
plot.legend(title='Certificate',fontsize='small')

imdb_U = imdb_drop[imdb_drop['Certificate'] == 'U'][['Certificate', 'Meta_score']]
imdb_U.to_excel('imdb_U.xlsx')

import sys
!{sys.executable} -m pip install stemgraphic

import stemgraphic
fig, ax = stemgraphic.stem_graphic(imdb_U["Meta_score"],scale=10,asc=False)
ax.set_title("Stem-and_Leaf display")

imdb_drop.boxplot(by='評等',column='No_of_Votes',figsize=(10,10))
lower_limit = 0
upper_limit = 2500000
plot.ylim(lower_limit,upper_limit)

df_Dat1=pd.read_csv(r"https://mopsfin.twse.com.tw/opendata/t187ap46_O_1.csv",header=0)
df_Dat2=pd.read_csv(r"https://mopsfin.twse.com.tw/opendata/t187ap46_L_1.csv",header=0)
df_merge_data= pd.concat([df_Dat1,df_Dat2], axis=0)
print("上下合併：\n",df_merge_data)
