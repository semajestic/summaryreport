import pandas as pd
  
# define data
data = {'A': [1, 2, 3], 'B': [4, 5, 6], 'Name': [7, 8, 9]}
  
# create dataframe
df = pd.DataFrame(data)
  
print("Original DataFrame:")
print(df)
  
# shift column 'C' to first position
col = df.pop('Name')
  
# insert column using insert(position,column_name,first_column) function
df.insert(0, col.name, col)
  
print()
print("Final DataFrame")
print(df)