from prettytable import PrettyTable
import pandas as pd

df = pd.read_excel("C:\\Users\\rdeb\\OneDrive - Novo Nordisk\\DataVerseProject\\Copy of Latest SLA Prod Data 08 june (003).xlsx")
print(df.columns)
print(type(df.columns))


# tabular_fields = df.columns
# tabular_table = PrettyTable()
# tabular_table.field_names = tabular_fields 
# # tabular_table.add_row(["Jill","Smith", 50])
# # tabular_table.add_row(["Eve","Jackson", 94])
# # tabular_table.add_row(["John", "Doe", 80])
# for row in df.iterrows():
#     tabular_table.add_row(row)

# print(tabular_table)

html = df.to_html()
print(html)


