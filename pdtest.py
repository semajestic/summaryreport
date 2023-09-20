import pandas as pd
import numpy as np
technologies = {
    'Courses':["Spark","PySpark","Hadoop","Python","pandas","Oracle","Java"],
    'Fee' :[20000,25000,26000,22000,24000,21000,22000],
    'Duration':['30day','40days','35days','40days','60days','50days','55days'],
    'Discount':[1000,2300,1500,1200,2500,2100,2000]
              }
df = pd.DataFrame(technologies)
print(df)
df.index = np.arange(1, len(df) + 1)
df=df.reset_index()

df = df.rename({'index':' '},axis=1)
print(df)
