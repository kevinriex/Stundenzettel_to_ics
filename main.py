import terminaltables as tt
from sys import argv

def readfile(file_name):
    f = open(file_name)
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

def write_data_to_txt(table, file_name):
    d = open(file_name, "w")
    d.write(table)
    d.close()

def write_data_to_ics(table, file_name, eventname, location, description):
    table.pop(0)
    i = open(file_name, "w")

    i.write("BEGIN:VCALENDAR\n")
    i.write("VERSION:2.0\n")
    i.write("CALSCALE:GREGORIAN\n")
    for x in table:
        i.write("BEGIN:VEVENT\n")
        i.write(f"SUMMARY:{eventname}\n")

        startdata = str(x[0][6:10]) + str(x[0][3:5]) + str(x[0])[0:2] + "T" + str(x[1])[0:2] + str(x[1])[3:5]
        enddata = str(x[0][6:10]) + str(x[0][3:5]) + str(x[0])[0:2] + "T" + str(x[2])[0:2] + str(x[2])[3:5]
        start = f"DTSTART;TZID=Europe/Berlin:{startdata}00\n"
        end = f"DTEND;TZID=Europe/Berlin:{enddata}00\n"
        i.write(start)
        i.write(end)
        
        i.write(f"LOCATION:{location}\n")
        
        pausedata = x[3] if x[3] != "" else "///"
        pause= f"DESCRIPTION:{description} Pause {pausedata}\n"
        i.write(pause)

        i.write("STATUS:CONFIRMED\n")
        i.write("SEQUENCE:0\n")
        i.write("END:VEVENT\n")
    i.write("END:VCALENDAR\n")

    i.close()

def main():
    userdata = start()
    if userdata != None:
        tt_data = generate_table_data(readfile(userdata[0]))
        table = tt.AsciiTable(tt_data,"Stundenzettel")
        # print(tt_data)
        print(table.table)
        write_data_to_txt(table.table,userdata[2])
        write_data_to_ics(tt_data,userdata[1],userdata[3],userdata[4],userdata[5])
    else:
        man()

def start():
    if len(argv) != 1:
        if argv[1] == "/?" or "help":
            man()
        else:
            if len(argv) < 6:
                print("You have to provide all arguments...")
            else:
                incsv = argv[1]
                outics = argv[2]
                outtxt = argv[3]
                eventname = argv[4]
                location = argv[5]
                description = argv[6]
                return [incsv,outics,outtxt,eventname,location,description]
    else:
        print("You have to provide arguments...")
        return None

def man():
    print("Manual")

main()