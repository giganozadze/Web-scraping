import pandas as pd
import os
import re
import numpy as np
os.chdir('C:/Users/gigan/OneDrive/Desktop/Tenders')

# df_begin1 = pd.read_csv('data_tenders 034 1-630.csv', header=None)
# df_begin2 = pd.read_csv('data_tenders 034 630-1532.csv', header=None)
# df_conc = pd.concat([df_begin1, df_begin2])
# directory = 'C:\\Users\\user\Desktop\TendersF'
# os.listdir(directory)
# for filename in os.listdir(directory):
#     dfx = pd.read_csv(filename, header=None)
#     frames = [df_begin, dfx]
#     df_begin = pd.concat(frames)

df_conc = pd.read_excel('Procurement_data.xlsx', header=None)
df_conc.reset_index(inplace=True, drop=True)
df_conc.to_excel('test1.xlsx')
# df_conc.rename(columns={3: 'შესყიდვის_ტიპი'}, inplace=True)
df_conc.rename(columns={1: 'შესყიდვის_ტიპი', 2: 'განცხადების_ნომერი',3:'შესყიდვის_სტატუსი',4:'შემსყიდველი',5:'შესყიდვის_გამოცხადების_თარიღი',6:'წინადადებების_მიღება_იწყება',7:'წინადადებების მიღება მთავრდება',8:'შესყიდვის_სავარაუდო_ღირებულება',9:'წინადადება_წარმოდგენილი_უნდა_იყოს',10:'შესყიდვის_კატეგორია',11:'კლასიფიკატორის_კოდები',12:'მოწოდების_ვადა',13:'დამატებითი_ინფორმაცია',14:'შესყიდვის_რაოდენობა_ან_მოცულობა',15:'შეთავაზების_ფასის_კლების_ბიჯი',16:'გარანტიის_ოდენობა',17:'გარანტიის_მოქმედების_ვადა',18:'ქრონოლოგია',19:'შეთავაზებები'},inplace=True)

'''CPV code
22100000-ნაბეჭდი წიგნები, ბროშურები და საინფორმაციო ფურცლები
22200000-გაზეთები, სამეცნიერო ჟურნალები, პერიოდიკა და ჟურნალები
22300000-ღია ბარათები, მისალოცი ბარათები და სხვა ნაბეჭდი მასალა
22400000-მარკები, ჩეკების წიგნაკები.....
22500000-საბ
22600000-
22800000-
22900000-
'''
'''We need
Top companies: Their value won over years, products they won,
Top Institutions: Their purchasing, top products'''

df_conc.columns
df_conc.შესყიდვის_სტატუსი.unique()
df_conc.groupby('შესყიდვის_სტატუსი').nunique()
won = df_conc[df_conc['შესყიდვის_სტატუსი']=='ხელშეკრულება დადებულია']
offers = won['შეთავაზებები']
won.reset_index(drop=True, inplace=True)


won = df_conc
won['გამარჯვებული'] = np.nan
won['გამარჯვებული_ფასი'] = np.nan

text = won.iloc[i, 19]
print(list(text))
#Price
for i in range(1, len(won)):
    text = won.iloc[i, 19]
    pattern = re.compile(r'ნახვა(\n.+\n)')
    price = re.findall(pattern, text)
    if not price:
        try:
            pattern = re.compile(r'.+\n')
            price = re.findall(pattern, text)
            price=price[0]
            price = price.replace('`', '')
            price = re.findall(r'\d+\.\d\d?', price)[0]
            price = float(price)
            won.iloc[i, 21] = price
        except Exception as e:
            won.iloc[i, 21] = 0
    else:
        try:
            price=price[len(price)-1]
            price=price.replace('`','')
            price=re.findall(r'\d+\.\d\d?',price)[0]
            price=float(price)
            won.iloc[i,21]=price
        except Exception as e:
            won.iloc[i, 21] = 0

    #Winner name
    text=won.iloc[i,19]
    pattern=re.compile(r'ნახვა(\n.+\n)')
    winner=re.findall(pattern,text)
    if not winner:
        try:
            pattern = re.compile(r'.+\n')
            winner = re.findall(pattern, text)
            winner=winner[0]
            winner = re.findall(r'(.+`|.+\d\d\d)', winner)[0]
            winner=re.sub('(\s\d\d?\d?`|\d\d\d)','',winner)
            won.iloc[i,20] = winner
        except Exception as e:
            won.iloc[i, 20] = e

    else:
        try:
            winner=winner[len(winner)-1]
            winner=winner.replace('`','')
            winner=re.sub(r'\d+.\d?','',winner)
            winner=winner.replace('\n','')
            winner=re.sub('\s$','',winner)
            won.iloc[i, 20] = winner
        except Exception as e:
            won.iloc[i, 20] = e

#correcting for VAT
won['გამარჯვებული_ფასი_Corr'] = np.nan

for i in range(1, len(won)):
    if won.iloc[i,9]=='დღგ-ს გათვალისწინებით':
        won.iloc[i,22]=won.iloc[i,21]
    else:
        won.iloc[i,22]=won.iloc[i,21]/0.82

# Add year
won['year'] = np.nan
for i in range(1, len(won)):
    text = won.iloc[i, 5]
    pattern = re.compile(r'\d\d\d\d')
    year = re.findall(pattern, text)
    year = int(year[0])
    won.iloc[i, 23] = year

won['month']=np.nan
for i in range(1, len(won)):
    text = won.iloc[i,5]
    pattern = re.compile(r'\.\d\d\.')
    month = re.findall(pattern, text)
    month = re.sub('\.0','', str(month))
    month = month.replace('[', '')
    month = month.replace(']', '')
    month = month.replace('.', '')
    month = month.replace("'", '')
    # month = int(month[0])
    # print(month[0])
    won.iloc[i, 24] = month

won['გამარჯვებული'] = won['გამარჯვებული'].astype('str')

won.to_excel('Tenders_summary.xlsx')

#2011-2020
top=won.groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum')
top.to_excel('top.xlsx')

topP=won.groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
topP.to_excel('topP.xlsx')

yearly=won[won['year']==2011].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2011.xlsx')
yearly=won[won['year']==2012].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2012.xlsx')
yearly=won[won['year']==2013].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2013.xlsx')
yearly=won[won['year']==2014].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2014.xlsx')
yearly=won[won['year']==2015].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2015.xlsx')
yearly=won[won['year']==2016].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2016.xlsx')
yearly=won[won['year']==2017].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2017.xlsx')
yearly=won[won['year']==2018].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2018.xlsx')
yearly=won[won['year']==2019].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2019.xlsx')
yearly=won[won['year']==2020].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2020.xlsx')
yearly=won[won['year']==2021].groupby('შესყიდვის_კატეგორია')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('year2021.xlsx')

topI=won.groupby('შემსყიდველი')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean']).nlargest(200,'sum').reset_index()
topI.to_excel('topI.xlsx')

year=won.groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
year.to_excel('year.xlsx')


strange=won[(won['შესყიდვის_კატეგორია']=='79000000-ბიზნეს მომსახურებები: იურისპრუდენცია, მარკეტინგი, კონსულტირება, რეკრუტირება, ბეჭდვა და უსაფრთხოება')]

#2016 & 2018 one-off
top2014=won[won['year']==2014].groupby('შემსყიდველი')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
top2015=won[won['year']==2015].groupby('შემსყიდველი')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
top2016=won[won['year']==2016].groupby('შემსყიდველი')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
top2017=won[won['year']==2017].groupby('შემსყიდველი')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
top2018=won[won['year']==2018].groupby('შემსყიდველი')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
top2019=won[won['year']==2019].groupby('შემსყიდველი')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
det=won[(won['year']==2016)&(won['შემსყიდველი']==top2016.iloc[0,0])]
det18=won[(won['year']==2018)&(won['შემსყიდველი']==top2018.iloc[0,0])]


#By product analysis
#Yearly
#22400000
yearly=won[won['შესყიდვის_კატეგორია']=='22400000-მარკები, ჩეკების წიგნაკები, ბანკნოტები, აქციები, სარეკლამო მასალა, კატალოგები და სახელმძღვანელოები'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index().sort_values('year',ascending=True)
yearly.to_excel('year224.xlsx')

top2012=won[won['year']==2012].groupby('შემსყიდველი')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
det2012=won[(won['year']==2012)&(won['შემსყიდველი']==top2012.iloc[0,0])]
det18=won[(won['year']==2018)&(won['შემსყიდველი']==top2018.iloc[0,0])]

#228000000
yearly=won[won['შესყიდვის_კატეგორია']=='22800000-ქაღალდის ან მუყაოს სარეგისტრაციო ჟურნალები/წიგნები, საბუღალტრო წიგნები, ფორმები და სხვა ნაბეჭდი საკანცელარიო ნივთები'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index().sort_values('year',ascending=True)
yearly.to_excel('year228.xlsx')

det2013=won[(won['year']==2013)&(won['შესყიდვის_კატეგორია']=='22800000-ქაღალდის ან მუყაოს სარეგისტრაციო ჟურნალები/წიგნები, საბუღალტრო წიგნები, ფორმები და სხვა ნაბეჭდი საკანცელარიო ნივთები')]
top20_13=det2013.nlargest(20,'გამარჯვებული_ფასი_Corr')

det2012=won[(won['year']==2014)&(won['შესყიდვის_კატეგორია']=='22800000-ქაღალდის ან მუყაოს სარეგისტრაციო ჟურნალები/წიგნები, საბუღალტრო წიგნები, ფორმები და სხვა ნაბეჭდი საკანცელარიო ნივთები')]
top20_14=det2012.nlargest(20,'გამარჯვებული_ფასი_Corr')

det2015=won[(won['year']==2015)&(won['შესყიდვის_კატეგორია']=='22800000-ქაღალდის ან მუყაოს სარეგისტრაციო ჟურნალები/წიგნები, საბუღალტრო წიგნები, ფორმები და სხვა ნაბეჭდი საკანცელარიო ნივთები')]
top20_15=det2015.nlargest(20,'გამარჯვებული_ფასი_Corr')

det2016=won[(won['year']==2016)&(won['შესყიდვის_კატეგორია']=='22800000-ქაღალდის ან მუყაოს სარეგისტრაციო ჟურნალები/წიგნები, საბუღალტრო წიგნები, ფორმები და სხვა ნაბეჭდი საკანცელარიო ნივთები')]
top20_16=det2016.nlargest(20,'გამარჯვებული_ფასი_Corr')

det2017=won[(won['year']==2017)&(won['შესყიდვის_კატეგორია']=='22800000-ქაღალდის ან მუყაოს სარეგისტრაციო ჟურნალები/წიგნები, საბუღალტრო წიგნები, ფორმები და სხვა ნაბეჭდი საკანცელარიო ნივთები')]
top20_17=det2017.nlargest(20,'გამარჯვებული_ფასი_Corr')
det18=won[(won['year']==2018)&(won['შემსყიდველი']==top2018.iloc[0,0])]

#221000000
yearly=won[won['შესყიდვის_კატეგორია']=='22100000-ნაბეჭდი წიგნები, ბროშურები და საინფორმაციო ფურცლები'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index().sort_values('year',ascending=True)
yearly.to_excel('year221.xlsx')

det2016=won[(won['year']==2016)&(won['შესყიდვის_კატეგორია']=='22100000-ნაბეჭდი წიგნები, ბროშურები და საინფორმაციო ფურცლები')]
top20_16=det2016.nlargest(20,'გამარჯვებული_ფასი_Corr')

det2018=won[(won['year']==2018)&(won['შესყიდვის_კატეგორია']=='22100000-ნაბეჭდი წიგნები, ბროშურები და საინფორმაციო ფურცლები')]
top20_18=det2018.nlargest(20,'გამარჯვებული_ფასი_Corr')

#229000000
yearly=won[won['შესყიდვის_კატეგორია']=='22900000-სხვადასხვა ნაბეჭდი მასალა'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index().sort_values('year',ascending=True)
yearly.to_excel('year229.xlsx')

#222000000
yearly=won[won['შესყიდვის_კატეგორია']=='22200000-გაზეთები, სამეცნიერო ჟურნალები, პერიოდიკა და ჟურნალები'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index().sort_values('year',ascending=True)
yearly.to_excel('year222.xlsx')

#220000000
yearly=won[won['შესყიდვის_კატეგორია']=='22000000-საბეჭდი მასალა და მონათესავე პროდუქცია'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index().sort_values('year',ascending=True)
yearly.to_excel('year220.xlsx')

#225000000
yearly=won[won['შესყიდვის_კატეგორია']=='22500000-საბეჭდი ფორმები ან ცილინდრები ან ბეჭდვისას გამოსაყენებელი სხვა საშუალებები'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index().sort_values('year',ascending=True)
yearly.to_excel('year225.xlsx')

#223000000
yearly=won[won['შესყიდვის_კატეგორია']=='22300000-ღია ბარათები, მისალოცი ბარათები და სხვა ნაბეჭდი მასალა'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index().sort_values('year',ascending=True)
yearly.to_excel('year223.xlsx')

#226000000
yearly=won[won['შესყიდვის_კატეგორია']=='22600000-საღებავი'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index().sort_values('year',ascending=True)
yearly.to_excel('year226.xlsx')

#798000000
yearly=won[won['შესყიდვის_კატეგორია']=='79800000-ბეჭდვა და მასთან დაკავშირებული მომსახურებები'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index().sort_values('year',ascending=True)
yearly.to_excel('year79.xlsx')

#ტოპ კომპანიები პროდუქტების მიხედვით
#ტოპ კომპანიები პროდუქტების მიხედვით
yearly=won[won['შესყიდვის_კატეგორია']=='79800000-ბეჭდვა და მასთან დაკავშირებული მომსახურებები'].groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
yearly.to_excel('year79C.xlsx')



yearly=won[won['შესყიდვის_კატეგორია']=='22400000-მარკები, ჩეკების წიგნაკები, ბანკნოტები, აქციები, სარეკლამო მასალა, კატალოგები და სახელმძღვანელოები'].groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
yearly.to_excel('year224C.xlsx')

yearly=won[won['შესყიდვის_კატეგორია']=='22800000-ქაღალდის ან მუყაოს სარეგისტრაციო ჟურნალები/წიგნები, საბუღალტრო წიგნები, ფორმები და სხვა ნაბეჭდი საკანცელარიო ნივთები'].groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
yearly.to_excel('year228C.xlsx')

yearly=won[won['შესყიდვის_კატეგორია']=='22100000-ნაბეჭდი წიგნები, ბროშურები და საინფორმაციო ფურცლები'].groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
yearly.to_excel('year221C.xlsx')

yearly=won[won['შესყიდვის_კატეგორია']=='22900000-სხვადასხვა ნაბეჭდი მასალა'].groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
yearly.to_excel('year229C.xlsx')

yearly=won[won['შესყიდვის_კატეგორია']=='22200000-გაზეთები, სამეცნიერო ჟურნალები, პერიოდიკა და ჟურნალები'].groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
yearly.to_excel('year222C.xlsx')

yearly=won[won['შესყიდვის_კატეგორია']=='22000000-საბეჭდი მასალა და მონათესავე პროდუქცია'].groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
yearly.to_excel('year220C.xlsx')

yearly=won[won['შესყიდვის_კატეგორია']=='22500000-საბეჭდი ფორმები ან ცილინდრები ან ბეჭდვისას გამოსაყენებელი სხვა საშუალებები'].groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
yearly.to_excel('year225C.xlsx')

yearly=won[won['შესყიდვის_კატეგორია']=='22300000-ღია ბარათები, მისალოცი ბარათები და სხვა ნაბეჭდი მასალა'].groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
yearly.to_excel('year223C.xlsx')

yearly=won[won['შესყიდვის_კატეგორია']=='22600000-საღებავი'].groupby('გამარჯვებული')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).nlargest(200,'sum').reset_index()
yearly.to_excel('year226C.xlsx')




un2017=won[won['year']==2017]['გამარჯვებული'].unique()
un2017=pd.DataFrame(data=un2017.flatten())
un2018=won[won['year']==2018]['გამარჯვებული'].unique()
un2018=pd.DataFrame(data=un2018.flatten())
un2019=won[won['year']==2019]['გამარჯვებული'].unique()
un2019=pd.DataFrame(data=un2019.flatten())

len(won[won['year']==2011]['გამარჯვებული'].unique())


yearly=won[won['შემსყიდველი']=='სსიპ საქართველოს შინაგან საქმეთა სამინისტროს მომსახურების სააგენტო'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('topI1.xlsx')
yearly=won[won['შემსყიდველი']=='საგანმანათლებლო და სამეცნიერო ინფრასტრუქტურის განვითარების სააგენტო'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('topI2.xlsx')
yearly=won[won['შემსყიდველი']=='შემოსავლების სამსახური'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('topI3.xlsx')
yearly=won[won['შემსყიდველი']=='ივანე ჯავახიშვილის სახელობის თბილისის სახელმწიფო უნივერსიტეტი'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('topI4.xlsx')
yearly=won[won['შემსყიდველი']=='საზღვაო ტრანსპორტის სააგენტო'].groupby('year')['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
yearly.to_excel('topI5.xlsx')

test=won[won['შესყიდვის_კატეგორია']=='79800000-ბეჭდვა და მასთან დაკავშირებული მომსახურებები'].groupby(['შემსყიდველი','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
test.to_excel('test.xlsx')
test2=won.groupby(['გამარჯვებული','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
test2.to_excel('test2.xlsx')




test=won.groupby(['year','month'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
test.to_excel('month.xlsx')



comp=won[won['გამარჯვებული']=='სეზანი'].groupby(['შემსყიდველი','შესყიდვის_კატეგორია','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
comp.to_excel('cezan.xlsx')


comp=won[won['გამარჯვებული']=='ფავორიტი სტილი'].groupby(['შემსყიდველი','შესყიდვის_კატეგორია','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
comp.to_excel('ფავორიტი სტილი.xlsx')

comp=won[won['გამარჯვებული']=='ელვა.ჯი'].groupby(['შემსყიდველი','შესყიდვის_კატეგორია','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
comp.to_excel('ელვა.ჯი.xlsx')


comp=won[won['გამარჯვებული']=='შპს "კაბადონი +"'].groupby(['შემსყიდველი','შესყიდვის_კატეგორია','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
comp.to_excel('kabadoni.xlsx')

comp=won[won['გამარჯვებული']=='Garsu pasaulis UAB`'].groupby(['შემსყიდველი','შესყიდვის_კატეგორია','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
comp.to_excel('garsu.xlsx')

comp=won[won['გამარჯვებული']=='ქეჩერა'].groupby(['შემსყიდველი','შესყიდვის_კატეგორია','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
comp.to_excel('ქეჩერა.xlsx')

comp=won[won['გამარჯვებული']=='შპს ჰორიზონტი'].groupby(['შემსყიდველი','შესყიდვის_კატეგორია','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
comp.to_excel('შპს ჰორიზონტი.xlsx')


comp=won[won['გამარჯვებული']=='პოლიგრაფისტი'].groupby(['შემსყიდველი','შესყიდვის_კატეგორია','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
comp.to_excel('პოლიგრაფისტი.xlsx')

comp=won[won['გამარჯვებული']=='გამომცემლობა კოლორი'].groupby(['შემსყიდველი','შესყიდვის_კატეგორია','year'])['გამარჯვებული_ფასი_Corr'].agg(['sum','count','mean','max']).reset_index()
comp.to_excel('გამომცემლობა კოლორი.xlsx')



un2017=won['გამარჯვებული'].unique()
un2017 = pd.DataFrame(data=un2017.flatten())
un2017.to_excel('uniq.xlsx')

won.to_excel('Data_Tenders_Summary.xlsx')
