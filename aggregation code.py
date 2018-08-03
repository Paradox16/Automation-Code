Test2

import pandas as pd
import os
import openpyxl as o

 

Dir = r'D:\Master'

filelist = os.listdir(Dir)

 

output = pd.DataFrame([])

for file in filelist:

    wb = o.load_workbook(Dir + '\\' + file, read_only = True)

    worksheetlist = wb.sheetnames

   

    for ws in worksheetlist:

        df = pd.read_excel(Dir + '\\' + file , ws)

        output = pd.concat([output,df])

        print(file + " : " + ws)


output.to_csv(r'D:\D:\Master\complete1.csv', index = False)

---------------------------------------------------------

import pandas as pd
import os
import openpyxl as o

 

Dir = r'D:\Master'

filelist = os.listdir(Dir)

 

output = pd.DataFrame([])

for file in filelist:

    df = pd.read_csv(Dir + '\\' + file, usecols = list(range(0,6)))

    output = pd.concat([output,df])

    print(file + " : Done")

 


output.Period[output.Period=='P1 2016'] = '2016_P1'
output.Period[output.Period=='P2 2016'] = '2016_P2'
output.Period[output.Period=='P3 2016'] = '2016_P3'
output.Period[output.Period=='P4 2016'] = '2016_P4'
output.Period[output.Period=='P5 2016'] = '2016_P5'
output.Period[output.Period=='P6 2016'] = '2016_P6'

output.Period[output.Period=='P1 2017'] = '2017_P1'
output.Period[output.Period=='P2 2017'] = '2017_P2'
output.Period[output.Period=='P3 2017'] = '2017_P3'
output.Period[output.Period=='P4 2017'] = '2017_P4'
output.Period[output.Period=='P5 2017'] = '2017_P5'
output.Period[output.Period=='P6 2017'] = '2017_P6'


output.drop(columns=['BaseUnitSales','IncrementalUnits','PPGID'],inplace = True)
output.columns = ['PPG','Period','Units','GEOG']
unitsdf = output

unitsdf.drop_duplicates(subset=None, keep='first', inplace=True)
unitsdf.to_csv(r'D:\Master\complete2.csv', index = False)









