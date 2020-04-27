import pandas as pd

colsToUse = [1, 2, 3, 4, 12, 14, 15]

firstSheet = pd.read_excel("lekiRefundowane.xlsx", sheet_name= 0, skiprows = 2, usecols = colsToUse)
restOfSheets = pd.read_excel("lekiRefundowane.xlsx", sheet_name= [1, 2], skiprows = 1, usecols = colsToUse)
book = pd.concat(restOfSheets.values(), ignore_index=True)
book = pd.concat([firstSheet, book], ignore_index=True)
book.columns = ["substancja", "nazwaPostacDawka", "zawartosc", "EAN", "wskazania", "poziom", "doplata"]

book[["nazwa", "postacDawka"]] = book.nazwaPostacDawka.str.split(pat = ", ", n = 1, expand = True)
book = book.drop("nazwaPostacDawka", axis = 1)
book[["postac", "dawka"]] = book.postacDawka.str.split(pat = r', \d', n = 1, expand = True)
book = book.drop("postacDawka", axis = 1)
book = book[["nazwa", "substancja", "postac", "dawka", "zawartosc", "EAN", "poziom", "wskazania", "doplata"]]

book.sort_values(by = book.columns.values.tolist(), ignore_index = True, inplace = True)

# we create a table
f = open("init.sql", "w", encoding = 'utf-8')
f.write("DROP TABLE IF EXISTS Medicine;\n\n")
f.write("CREATE TABLE Medicine (\n")
f.write("\tId INTEGER NOT NULL,\n")
f.write("\tName TEXT NOT NULL,\n")
f.write("\tSubstance TEXT NOT NULL,\n")
f.write("\tForm TEXT NOT NULL,\n")
f.write("\tDose TEXT NOT NULL,\n")
f.write("\tContent TEXT NOT NULL,\n")
f.write("\tEAN INTEGER NOT NULL,\n")
f.write("\tRefundation TEXT NOT NULL,\n")
f.write("\tScope TEXT NOT NULL,\n")
f.write("\tPrice TEXT NOT NULL,\n")
f.write("\tPRIMARY KEY (Id)\n);\n\n")

for i in range(len(book.index)):
    f.write("INSERT INTO Medicine(")
    f.write("Name, Substance, Form, Dose, Content, EAN, Refundation, Scope, Price) ")
    f.write("VALUES(")
    for j in range(len(book.columns)):
        if j > 0:
            f.write(", ")
        f.write("\"")
        try: 
            f.write(book.iloc[i,j])
        except TypeError:
            f.write(str(book.iloc[i,j]))
        f.write("\"")
    f.write(");\n")
f.close()
























