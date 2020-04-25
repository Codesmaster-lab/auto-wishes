import pandas as pd
import xlrd
from datetime import datetime
data = pd.read_excel (r'DATA.xlsx')
df = pd.DataFrame(data, columns= ['EMAIL','DOB'])
mylist = df.values.tolist()
for da in mylist:
    print(datetime.date(da[1]))