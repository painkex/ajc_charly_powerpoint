import pandas as pd

df = pd.read_excel('c.xlsx')
df_copy = df.copy()

new_copy = df_copy.drop(columns=['POEI','CHARGES DE RECRUTEMENT'],axis=1)

print(new_copy.columns)

