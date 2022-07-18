import terminaltables as tt

def readfile():
    f = open("data.csv")
    data = f.readlines()
    f.close()
    data.pop(0)
    return data

def generate_table_data(data):
    tt_data = [["Datum","Start","Ende","Pause"]]
    for date in data:
        #print(type(date))
        if ";;;" not in date:
            a = date[0:10]
            b = date[11:16]
            c = date[17:22]
            d = date[23:-1]
            tt_data.append([a,b,c,d])
    return tt_data

def write_data_to_txt(table):
    d = open("out.txt", "w")
    d.write(table)
    d.close()

def write_data_to_ics(table):
    table.pop(0)
    i = open("test1.ics", "w")

    i.write("BEGIN:VCALENDAR\n")
    i.write("VERSION:2.0\n")
    i.write("CALSCALE:GREGORIAN\n")
    for x in table:
        i.write("BEGIN:VEVENT\n")
        i.write("SUMMARY:KomMITT Arbeitstag\n")

        startdata = str(x[0][6:10]) + str(x[0][3:5]) + str(x[0])[0:2] + "T" + str(x[1])[0:2] + str(x[1])[3:5]
        enddata = str(x[0][6:10]) + str(x[0][3:5]) + str(x[0])[0:2] + "T" + str(x[2])[0:2] + str(x[2])[3:5]
        start = f"DTSTART;TZID=Europe/Berlin:{startdata}00\n"
        end = f"DTEND;TZID=Europe/Berlin:{enddata}00\n"
        i.write(start)
        i.write(end)
        
        i.write("LOCATION:Kaiserswerther Str. 85, 40878 Ratingen\n")
        
        pausedata = x[3] if x[3] != "" else "///"
        pause= f"DESCRIPTION:Ein normaler Arbeitstag bei der KomMITT. Pause {pausedata}\n"
        i.write(pause)

        i.write("STATUS:CONFIRMED\n")
        i.write("SEQUENCE:0\n")
        i.write("END:VEVENT\n")
    i.write("END:VCALENDAR\n")

    i.close()

tt_data = generate_table_data(readfile())
table = tt.AsciiTable(tt_data,"Stundenzettel")
# print(tt_data)
# print(table.table)
# write_data_to_txt(table.table)
write_data_to_ics(tt_data)

