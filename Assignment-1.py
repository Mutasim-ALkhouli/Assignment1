import pandas as pd

initFilesList=[]
finalDF=pd.DataFrame()

for i in range(10):
    df=pd.read_excel(f"XLSXs\\{i+1}.xlsx")
    initFilesList.append({1:df,2:f"{i+1}"})

allFilesData=[]

for i in initFilesList:
    allFilesData.append(i[1][['name','Email','Guest']] )
    
#print(allFilesData)

basicData=pd.concat(allFilesData)
basicData.drop_duplicates(subset="name",inplace=True)
 
for n,e,g in zip(basicData["name"],basicData["Email"],basicData["Guest"]) :
    sum=0
    absent=0
    dates={}
    for ifl in initFilesList:
        dataRow=ifl[1][ifl[1]["name"] == n]
        dataRow=dataRow.squeeze()
        if(dataRow.empty==False):
            sum+=dataRow["Time"]
            dates[ifl[2]]=dataRow["Time"]

        else:
            absent+=1
            dates[ifl[2]]=0

        
    rowData = {'Full Name': n, 'Email': e, 'Guest': g, **dates ,"Attendees":sum,"Absence":absent}
    row = pd.Series(data=rowData)
    row=row.to_frame().T
    finalDF= pd.concat([finalDF,row],axis=0)

print(finalDF)
finalDF.to_excel("final.xlsx",index=False)
