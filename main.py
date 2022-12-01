import terminaltables as tt

def readfile(fileN):
    f = open(fileN)
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
        else:
            tt_data.append(["///","///","///","///"])
    return tt_data

def write_data_to_txt(table, name):
    d = open(name, "w")
    d.write(table)
    d.close()

def write_data_to_ics(table):
    table.pop(0)
    i = open("out.ics", "w")

    i.write("BEGIN:VCALENDAR\n")
    i.write("VERSION:2.0\n")
    i.write("CALSCALE:GREGORIAN\n")
    i.write("BEGIN:VTIMEZONE\n")
    i.write("TZID:Europe/Berlin\n")
    i.write("X-LIC-LOCATION:Europe/Berlin\n")
    i.write("BEGIN:DAYLIGHT\n")
    i.write("TZOFFSETFROM:+0100\n")
    i.write("TZOFFSETTO:+0200\n")
    i.write("TZNAME:CEST\n")
    i.write("DTSTART:19700329T020000\n")
    i.write("RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=3\n")
    i.write("END:DAYLIGHT\n")
    i.write("BEGIN:STANDARD\n")
    i.write("TZOFFSETFROM:+0200\n")
    i.write("TZOFFSETTO:+0100\n")
    i.write("TZNAME:CET\n")
    i.write("DTSTART:19701025T030000\n")
    i.write("RRULE:FREQ=YEARLY;INTERVAL=1;BYDAY=-1SU;BYMONTH=10\n")
    i.write("END:STANDARD\n")
    i.write("END:VTIMEZONE\n")
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

def compare_old_new(new, old):
    for x in range(1,len(new)):
        if len(new) >= len(old):
            if new[x] == old[x]:
                old[x] = ["///","///","///","///"]
                new[x] = ["///","///","///","///"]
    data = [item for item in new if item != ["///","///","///","///"]]
    return data

data_new = readfile("data.csv")
data_old = readfile("data_old.csv")

tt_data_new = generate_table_data(data_new)
tt_data_old = generate_table_data(data_old)

table_new = tt.AsciiTable(tt_data_new,"Stundenzettel")
table_old = tt.AsciiTable(tt_data_old,"Stundenzettel")

write_data_to_txt(table_new.table,"new.txt")
write_data_to_txt(table_old.table,"old.txt")



data = compare_old_new(tt_data_new,tt_data_old)

table = tt.AsciiTable(data,"Stundenzettel")
# print(tt_data)
print(table.table)
write_data_to_txt(table.table,"out-new.txt")
write_data_to_ics(data)

