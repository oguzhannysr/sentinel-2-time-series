import pandas as pd
import os
import geopandas as gpd
import datetime
import numpy as np

rasters_path_dir = []
parcels = gpd.read_file("D:\Projects\BarleyWheat\Parcels\Doktar_Yozgat_2021_polygon_selection_edit_EPSG32636.shp")  



for files in os.walk("D:\Projects\BarleyWheat\Output"):
    rasters_path_dir.append(files[0])
rasters_path_dir = sorted(rasters_path_dir,key=lambda s: s[42:50])
rasters_path_dir.remove("D:\Projects\BarleyWheat\Output")
dates = []
for i in range(0,len(rasters_path_dir)):
    split = rasters_path_dir[i].split('\\')
    dates.append(split[4][11:19])
    
dates.sort()

for i in range(0,len(rasters_path_dir)):
    read_zonal_stats=pd.read_excel(rasters_path_dir[i] + "/zonal_stats.xlsx")
    globals()['d_' + '%s' % dates[i]] = read_zonal_stats 
#%%    
dates_df = pd.to_datetime(dates)
dates_range = pd.date_range(start=min(dates_df),end=max(dates_df),freq="5D")
diff = dates_range.difference(dates_df)
diff = diff.to_native_types()

for i in range(0,len(diff)):
    diff[i] = diff[i].replace('-','')
    read_zonal_stats=pd.DataFrame(columns=read_zonal_stats.columns,index=read_zonal_stats.index)
    globals()['d_' + '%s' % diff[i]] = read_zonal_stats
    globals()['d_' + '%s' % diff[i]]['_ROWID_'] = d_20201005['_ROWID_']  #d_20201005 bu değişken en baştaki tarih bunu değiştirmen gerekebilir
    dates.append(diff[i])
dates.sort()


#%%    
    
a=0
b=parcels.shape[0] 
c = 0   
bands_vgind_name = d_20211030['Unnamed: 0'].unique()
bands_vgind_name = [bands_vgind_name for bands_vgind_name in bands_vgind_name if str(bands_vgind_name) != 'nan']

for i in range(0,len(bands_vgind_name)):
    bands_vgind_name[i] = bands_vgind_name[i][:-4]
    
while b<=parcels.shape[0]*len(bands_vgind_name): #19 band count
    
    for x in range(0,len(bands_vgind_name)):  
        split = bands_vgind_name[x].split('_')
        globals()['%s' % split[2]] = pd.DataFrame()
        
        for i in range(0,len(dates)):
            globals()['%s' % split[2]] = pd.concat([globals()['%s' % split[2]],globals()['d_' + '%s' % dates[i]].iloc[a:b]])
            globals()['%s' % split[2]] = globals()['%s' % split[2]].reset_index(drop=True)
            del globals()['%s' % split[2]]['Unnamed: 1']
            


            if i == (len(dates)-1):
       
                i = 0
                a = b
                b = b + parcels.shape[0]
a = 0
b = parcels.shape[0] 

j = 0       
while j < len(bands_vgind_name):             
    for i in range(0,len(dates)):
        split = bands_vgind_name[j].split('_')
        globals()['%s' % split[2]]['Unnamed: 0'][a:b] = int(dates[i])
        a = b
        b = b + parcels.shape[0]      
        if i == (len(dates)-1):
            i = 0
            j = j + 1
            a = 0
            b = parcels.shape[0] 
            globals()['%s' % split[2]].rename(columns = {'Unnamed: 0':'Date'}, inplace = True) 
            globals()['%s' % split[2]] = globals()['%s' % split[2]].sort_values(by=['_ROWID_','Date'],ascending=[True,True])
            globals()['%s' % split[2]]['Date'] = pd.to_datetime(globals()['%s' % split[2]]['Date'],format='%Y%m%d',errors='coerce')
            globals()['%s' % split[2]]['Date'] = globals()['%s' % split[2]]['Date'].dt.strftime('%Y-%m-%d')
            globals()['%s' % split[2]].loc[globals()['%s' % split[2]]['nodata'] > 2, 'mean'] = np.nan
            globals()['%s' % split[2]] = globals()['%s' % split[2]].set_index('Date')
            globals()['%s' % split[2]].index = pd.to_datetime(globals()['%s' % split[2]].index)
            globals()['%s' % split[2]]['mean'] = globals()['%s' % split[2]]['mean'].astype('float64') 
            q=0
            w= len(dates)
            while q < globals()['%s' % split[2]].shape[0]:
                globals()['%s' % split[2]]['mean'][q:w] = globals()['%s' % split[2]][q:w]['mean'].interpolate(method='linear') #akima,cubicspline,pchip
                q = q+ len(dates)
                w = w + len(dates)
            #globals()['%s' % split[2]]["mean"] = globals()['%s' % split[2]]["mean"].ewm(span=5).mean()#0.25 0.5 0.75
                        
#%%
with pd.ExcelWriter("C:/Users/Lenovo/Desktop/cloudper25_TimeSeries_{}.xlsx".format(datetime.datetime.today().date())) as writer:
    for z in range(0,len(bands_vgind_name)):
        split = bands_vgind_name[z].split('_')     
        globals()['%s' % split[2]].to_excel(writer, sheet_name=split[2])     
        workbook=writer.book
        worksheet = writer.sheets[split[2]]
        chart = workbook.add_chart({'type': 'line',
                                    'name': 'VegInd'})

        for i in range(2,(parcels.shape[0]*len(dates)),len(dates)):
                
            chart.add_series({'values': '={}!C{}:C{}'.format(split[2],i,i+len(dates)-1),
                              'categories': '={}!A2:A{}'.format(split[2],len(dates)+1),
                              'name': '={}!B{}'.format(split[2],i)})
    
            
            worksheet.insert_chart('M2', chart)
            
#%%
'''ValueError: method must be one of ['linear', 'time', 'index', 'values', 'nearest', 'zero', 'slinear', 'quadratic', 'cubic', 'barycentric', 'krogh', 'spline', 'polynomial', 'from_derivatives', 'piecewise_polynomial', 'pchip', 'akima', 'cubicspline']. Got 'nearest-up' instead.'''

