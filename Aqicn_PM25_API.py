print("Running : requests")

import requests
import os
from datetime import datetime, date,  timedelta
import pandas as pd
from pandas import json_normalize
import geopandas as gpd
import seaborn as sns
import matplotlib.pyplot as plt
from SendEmail import Send_Email
from Credentials import *

# set seaborn "whitegrid" theme
sns.set_style("darkgrid")

todayStr=date.today().strftime('%Y-%m-%d')
nowStr=datetime.today().strftime('%Y-%m-%d %H:%M:%S')

NextDay_Date = datetime.today() + timedelta(days=1)
NextDay_Date_Formatted = NextDay_Date.strftime ('%Y-%m-%d') # format the date to ddmmyyyy

print(todayStr, '--- date ---',NextDay_Date_Formatted)

##################################################################
data_path='C:\\Users\\70018928\\Documents\\Project2021\\Experiment\\'
#################################################################


#### https://aqicn.org/data-platform/     https://aqicn.org/api/


#### template AQI
aqi_file='Template_AQI.xlsx'
aqi_template=pd.read_excel(data_path+'\\data\\'+aqi_file, sheet_name='สำหรับ Power BI')


##### parameters
output_name='thailand_aqi_waqi_info_'+str(todayStr)+'.xlsx'
#######
##### Employee location
file_name='OriginLocation_Tambon_20220406.csv'
cvt={'EID':str}
empDf=pd.read_csv(data_path+'\\data\\'+file_name,converters=cvt)
empDf['province_district']=empDf['p_name_t']+'_'+empDf['a_name_t']
print(len(empDf),' --- emp read in --- ',empDf.head(3),' :: ',empDf.columns)

#### find reference location : province center
### read provinces
path=data_path+'\\data\\th_boundary_v2\\'
p_boundary = gpd.read_file(path+'th_province_boundary_v2.shp')
province=p_boundary[['p_name_t','geometry']].copy().reset_index(drop=True)
print(province.crs)
province.set_crs(epsg=32647, allow_override=True, inplace=True)
province.to_crs(epsg=4326, inplace=True)
print(province.crs)
province['center']=province['geometry'].apply(lambda x: x.representative_point())
province['LONGITUDE']=province['center'].x
province['LATITUDE']=province['center'].y
province=province[['p_name_t','LATITUDE','LONGITUDE']].copy()
print(province)
# province.to_csv(data_path+'\\'+'check_latlng.csv',index=False)

#### only bangkok, breaks to district
p_boundary = gpd.read_file(path+'th_district_boundary_v2.shp')
p_boundary=p_boundary[p_boundary['p_name_t']=='กรุงเทพมหานคร'].copy().reset_index(drop=True)
p_boundary['p_a_name_t']=p_boundary['p_name_t']+'_'+p_boundary['a_name_t']
p_district=p_boundary[['p_a_name_t','geometry']].copy().reset_index(drop=True)
print(p_district.crs)
p_district.set_crs(epsg=32647, allow_override=True, inplace=True)
p_district.to_crs(epsg=4326, inplace=True)
print(p_district.crs)
p_district['center']=p_district['geometry'].apply(lambda x: x.representative_point())
p_district['LONGITUDE']=p_district['center'].x
p_district['LATITUDE']=p_district['center'].y
p_district=p_district[['p_a_name_t','LATITUDE','LONGITUDE']].copy()
print(p_district)

provinceList=list(province['p_name_t'].unique())
p_district_List=list(p_district['p_a_name_t'].unique())

#### write in sheets in same spreadsheet
############## Write data
writer = pd.ExcelWriter(output_name, engine='openpyxl')
#############

mainDf= pd.DataFrame()
for provinceId in provinceList:
    print(' ===> ',provinceId)
    dummyEmp=empDf[empDf['p_name_t']==provinceId].copy().reset_index(drop=True)
    numEmp=len(dummyEmp)
    del dummyEmp

    dummyDf=province[province['p_name_t']==provinceId].copy().reset_index(drop=True)
    LATITUDE=list(dummyDf['LATITUDE'].values)[0]
    LONGITUDE=list(dummyDf['LONGITUDE'].values)[0]
    print(provinceId,' :: ',LATITUDE,' :: ',LONGITUDE)

    url="https://api.waqi.info/feed/geo:"+str(LATITUDE)+";"+str(LONGITUDE)+"/?token="+str(token)
    response = requests.get(url)

    print("Running : response.json")
    data = response.json()
    print("Running : dict")
    dict_1=data
    print("Running : pandas")
    aqi=dict_1['data']['aqi']    

    pm25_forecast=json_normalize(dict_1['data']['forecast']['daily']['pm25'])
    pm25_dummy=pm25_forecast[pm25_forecast['day']==NextDay_Date_Formatted].copy().reset_index(drop=True)
    pm25_fc_nextday=list(pm25_dummy['avg'].values)[0]
    # print(' ====> ',pm25_forecast,' :: ',pm25_fc_nextday)
    del pm25_forecast, pm25_dummy

    mainDf=mainDf.append({"province":provinceId,"Latitude":LATITUDE,"Longitude":LONGITUDE,"aqi":aqi,"Employee":numEmp,"Tomorrow_aqi":pm25_fc_nextday},ignore_index=True)

includeList=['province','Latitude','Longitude','aqi', 'Employee','Tomorrow_aqi']
mainDf=mainDf[includeList].copy()

mainDf.rename(columns={'province':'Province','Latitude':'Lat','Longitude':'Long','aqi':'AQI','Tomorrow_aqi':'Tomorrow_AQI'},inplace=True)

mainDf['Updated']=nowStr
mainDf=aqi_template.merge(mainDf,on='Province',how='left')
print(' --- mainDf : ',mainDf)
# mainDf.to_excel(data_path+'\\'+'thailand_aqi_waqi_info.xlsx',index=False)
##### write to sheet
mainDf.to_excel(writer, sheet_name='Thailand_'+str(todayStr),index=False) 
####  continue

mainDf= pd.DataFrame()
for p_district_Id in p_district_List:
    print(' ===> ',p_district_Id)
    dummyEmp=empDf[empDf['province_district']==p_district_Id].copy().reset_index(drop=True)
    numEmp=len(dummyEmp)
    del dummyEmp

    dummyDf=p_district[p_district['p_a_name_t']==p_district_Id].copy().reset_index(drop=True)
    LATITUDE=list(dummyDf['LATITUDE'].values)[0]
    LONGITUDE=list(dummyDf['LONGITUDE'].values)[0]
    print(p_district_Id,' :: ',LATITUDE,' :: ',LONGITUDE)

    url="https://api.waqi.info/feed/geo:"+str(LATITUDE)+";"+str(LONGITUDE)+"/?token="+str(token)
    response = requests.get(url)

    print("Running : response.json")
    data = response.json()
    print("Running : dict")
    dict_1=data
    print("Running : pandas")
    aqi=dict_1['data']['aqi']    

    pm25_forecast=json_normalize(dict_1['data']['forecast']['daily']['pm25'])
    pm25_dummy=pm25_forecast[pm25_forecast['day']==NextDay_Date_Formatted].copy().reset_index(drop=True)
    pm25_fc_nextday=list(pm25_dummy['avg'].values)[0]
    # print(' ====> ',pm25_forecast,' :: ',pm25_fc_nextday)
    del pm25_forecast, pm25_dummy

    mainDf=mainDf.append({"province_district":p_district_Id,"Latitude":LATITUDE,"Longitude":LONGITUDE,"aqi":aqi,"Employee":numEmp,"Tomorrow_aqi":pm25_fc_nextday},ignore_index=True)

mainDf['province']=mainDf['province_district'].apply(lambda x: x.split('_')[0])
mainDf['district']=mainDf['province_district'].apply(lambda x: x.split('_')[1])

includeList=['province_district','province','district','Latitude','Longitude','aqi',"Employee","Tomorrow_aqi"]
mainDf=mainDf[includeList].copy()

mainDf.rename(columns={'province':'Province','district':'District','Latitude':'Lat','Longitude':'Long','aqi':'AQI','Tomorrow_aqi':'Tomorrow_AQI'},inplace=True)
mainDf['Updated']=nowStr
print(' --- mainDf : ',mainDf)
# mainDf.to_excel(data_path+'\\'+'thailand_bangkok_aqi_waqi_info.xlsx',index=False)
####### write to sheet
mainDf.to_excel(writer, sheet_name='Bangkok_'+str(todayStr), index=False)
writer.save()
################  end Write data

del writer
del mainDf, empDf

##################################



filepath=data_path+'\\'+'thailand_aqi_waqi_info_'+str(todayStr)+'.xlsx'
print(' **************** Sending Email *********************')
Send_Email(receiverList,filepath,todayStr,nowStr)
print(' **********************************************')


print(' ****************************************************** ')
print(' ****************************************************** ')
print(' *************  C o M p L e T e D ********************* ')
print(' ****************************************************** ')
print(' ****************************************************** ')


######################################################################################################
# #### test scatter bubble plot
# dfIn=pd.read_excel(data_path+'\\'+'thailand_aqi_waqi_info_'+str(todayStr)+'.xlsx')
# print(len(dfIn),' ---- read in ---- ',dfIn.head(3),' :: ',dfIn.columns)

# # sns.scatterplot(data=dfIn, x="province", y="aqi", size="Employee", hue="continent", palette="viridis", edgecolors="black", alpha=0.5, sizes=(10, 1000))
# sns.scatterplot(data=dfIn, x="province", y="aqi", size="Employee", palette="viridis", edgecolors="black", alpha=0.5, sizes=(10, 1000))

# # Add titles (main and on axis)
# plt.xlabel("province")
# plt.ylabel("aqi")


# # Locate the legend outside of the plot
# # plt.legend(bbox_to_anchor=(1, 1), loc='upper left', fontsize=17)

# # show the graph
# plt.show()

# del dfIn