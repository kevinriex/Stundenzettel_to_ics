import terminaltables as tt

f = open("data.csv")
data = f.readlines()
f.close()
data.pop(0)

tt_data = [["Datum","Start","Ende","Pause"]]
for date in data:
    #print(type(date))
    if ";;;" not in date:
        a = date[0:10]
        b = date[11:16]
        c = date[17:22]
        d = date[23:-1]
        tt_data.append([a,b,c,d])


#print(data)
table = tt.AsciiTable(tt_data,"Stundenzettel").table
print(table)
d = open("out.txt", "w")
d.write(table)
d.close()