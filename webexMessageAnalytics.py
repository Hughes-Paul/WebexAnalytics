from webexteamssdk import WebexTeamsAPI # pip install webexteamssdk
import pandas as pd # pip install pandas, pip install xlsxwriter
import time
from datetime import datetime
from datetime import date
from sys import platform
import re
import difflib
import os


#####################
# Getting api token #
#####################

#Checks if there is a cached token, and uses it if there is
if os.path.exists("tokencache.txt"):
    f = open("tokencache.txt", "r")
    personalToken = f.read()
else:
    personalToken = ""

#If there was a cached token, tries to use it for the api call
if personalToken:
    api = WebexTeamsAPI(access_token=personalToken)

#Repeatedly asks the user for an api token until a valid token is given
while True:
    try:
        api.people.me()
        print("Api token looks good!\n")
        f = open("tokencache.txt", "w")
        f.write(personalToken)
        f.close()
        break

    except:
        print("Your token was either invalid, expired, or non-existent")
        personalToken = input("Enter api token (found at https://developer.webex.com/docs/getting-started): ")
        if personalToken:
            api = WebexTeamsAPI(access_token=personalToken)


####################
# Selecting a room #
####################

#Gets user input on which room is wanted, and also formats that input
userRoom = re.sub('\W+','', re.split('\[|\(|http', input("What room would you like to gather analytics on?: ").lower() )[0] )

print("Please wait... this could take a minute or so.")

#Gets a list of rooms from the api
rooms = api.rooms.list(type="group", max=1000)

#Lists
roomTitles = []
roomTitlesUnformatted = []
roomIds = []

#Formats all of the rooms names and puts them into a list, as well as filling an unformatted list, and an ID list (this is for efficiency)
for room in rooms:
    roomTitles.append(re.sub('\W+','', re.split('\[|\(|http', room.title )[0] ).lower())
    roomTitlesUnformatted.append(room.title)
    roomIds.append(room.id)

#Finds the room with the name most similar to the one typed by the user
roomsOrdered = difflib.get_close_matches(userRoom, roomTitles, n=7, cutoff=0.35)

#Prints a list of the close matches, and lets the user select which is closest to what they want
print("Cloest matches: ")
for i in range(len(roomsOrdered)):
    for j, room in enumerate(roomTitles):
        if room == roomsOrdered[i]:
            print(str(i+1)+": "+roomTitlesUnformatted[j])
targetRoom = roomsOrdered[int(input("Select a room (1-"+str(len(roomsOrdered))+"): "))-1]


#Finds the ID of the room
for i, room in enumerate(roomTitles):
    if room == targetRoom:
        roomId = roomIds[i]
        roomName = roomTitlesUnformatted[i]


#######################
# Collecting the data #
#######################

#Name for the csv output file
csv_file = targetRoom+" "+datetime.now().strftime("%Y-%m-%d %H-%M-%S")+".csv"

#Amount of days the program will search (based on user input)
dateRange = input("Enter days back to search (e.g. 30) or a date range (e.g. mm/dd/yy-mm/dd/yy): ")
#Fixes some potential issues (forgetting a year, using 4 didgit year instead of 2 digit) and returns the dates in unix format
def fixDays(date1, date2):
    if len(date1) == 5:
        print("Year ommited, using " + date.today().year)
        date1 = date1+"/"+str(date.today().year)[2:4]
    if len(date2) == 5:
        print("Year ommited, using " + date.today().year)
        date2 = date2+"/"+str(date.today().year)[2:4]
    if date1[-4:-2] == "20":
        date1 = date1[:-4]+date1[-2:]
    if date2[-4:-2] == "20":
        date2 = date2[:-4]+date2[-2:]
    firstDate = int(time.mktime(time.strptime(str(date1), "%m/%d/%y")))
    lastDate = int(time.mktime(time.strptime(str(date2), "%m/%d/%y")))
    return((firstDate, lastDate))
try:
    date1, date2 = dateRange.split("-")
    firstDate, lastDate = fixDays(date1, date2)
except:
    try:
        date1, date2 = dateRange.split(" ") #In case the user tries to use a space instead of a dash
        firstDate, lastDate = fixDays(date1, date2)
    except:
        try:
            float(dateRange)
            firstDate = int(round(time.time() - (float(dateRange) * 86400)))
            lastDate = int(round(time.time()))
        except:
            print("Input unclear. Follow the mm/dd/yy-mm/dd/yy format, or enter an number above 0.")
            print("Exiting.")
            exit()

print("Thanks, running search on "+ datetime.fromtimestamp(firstDate).strftime("%m/%d/%y") + '-' + datetime.fromtimestamp(lastDate).strftime("%m/%d/%y") +" now. Large groups can take a while.")
search_start = time.time() #marks the start of the search so that the run time can be calculated later

messages = api.messages.list(roomId = roomId) #returns all messages from the room

messageCount = 0 #initializes a message count variable for later

people = {}

#Returns a list of all users in the room
fullRoom = api.memberships.list(roomId = roomId)
fullRoomIds = [] #Will later be a list of Ids from all people in the room
sentMsg = [] #Will later be a list of all people who sent a message 

for message in messages:

    #Checks to make sure the message isn't older than the oldest allowed message. If it is, the loop stops running
    if int(time.mktime(time.strptime(str(message.created), "%Y-%m-%dT%H:%M:%S.%fZ"))) < firstDate:
        break
    if int(time.mktime(time.strptime(str(message.created), "%Y-%m-%dT%H:%M:%S.%fZ"))) > lastDate:
        continue
    personId =  message.personId #id of the person who sent the message

    #bool based on whether or not the question is a child message
    isChild = "parentId" in str(message)

    #If the message sender has previously sent a message that was counted, simply adds one to the message count
    if personId in people:
        people[personId]["Total Messages"] = people[personId]["Total Messages"] + 1
        if isChild:
            people[personId]["Replies"] = people[personId]["Replies"] + 1
        else:
            people[personId]["Posts"] = people[personId]["Posts"] + 1

    #If the message sender hasn't previously sent a message that was counted, adds the person to a dictionary of all people in the space who have sent a message
    else:
        people[personId] = {}
        try:
            person = api.people.get(personId) #all info the api returns on the person
        except:
            person = "Null" #If the api fails to return info on the person, uses null instead
        try:
            people[personId]["Name"] = person.displayName
        except:
            people[personId]["Name"] = message.personEmail #If the display name fails to return, uses the email instead
        people[personId]["Email"] = message.personEmail
        people[personId]["Total Messages"] = 1
        if isChild: #Adds to either posts or answers, initializes the other
            people[personId]["Replies"] = 1
            people[personId]["Posts"] = 0
        else:
            people[personId]["Replies"] = 0
            people[personId]["Posts"] = 1
        
        people[personId]["InSpace"] = "yes"
        sentMsg.append(personId)
    
    #Increments the total message count by one
    messageCount = messageCount + 1


#Checks to see if a user has 0 messages in the specified time frame
noMessages = 0
for person in fullRoom:
    personId = person.personId
    fullRoomIds.append(personId) #Creates a list of all people in the room, this is used later to check for users who left
    if not personId in people:
        people[personId] = {}
        try:
            people[personId]["Name"] = person.personDisplayName
        except:
            people[personId]["Name"] = person.personEmail #If the display name fails to return, uses the email instead
        people[personId]["Email"] = person.personEmail
        people[personId]["Total Messages"] = 0
        people[personId]["Replies"] = 0
        people[personId]["Posts"] = 0
        people[personId]["InSpace"] = "yes"
        noMessages = noMessages+1

for id in sentMsg:
    if not id in fullRoomIds:
        people[id]["InSpace"] = "no"


##########
# Output #
##########

print("\nSearch complete, run time: " + str(round(time.time() - search_start, 2)) + " seconds.\n") #Prints the search time

print("Total message count: " + str(messageCount)) #More info for the user
print("Total user count: " + str(len(people))+"\n")

#Makes a list of keys (which are user ids) sorted by message count and by name alphebetically. The order in the csv file is based off of this
sortedKeys = sorted(people, key=lambda x: (people[x]["Total Messages"]), reverse=True)


#Writes each person from the dictionary into a csv file
with open(csv_file, "w", encoding="utf-8") as f:
    f.write("Email,Display name,Total Messages,Replies,Posts,User in Space\n")
    for key in sortedKeys:
        try:
            f.write(people[key]["Email"]+',"'+str(people[key]["Name"]).replace(",","")+'",'+str(people[key]["Total Messages"])+','+str(people[key]["Replies"])+','+str(people[key]["Posts"])+','+people[key]["InSpace"]+'\n')
        except:
            print("Error writing info for: "+people[key]["Email"])

#Write csv data to an excel spread sheet using pandas
data = pd.read_csv(csv_file)
df = pd.DataFrame(data)
with pd.ExcelWriter(csv_file[:-4]+'.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Report')
    messagesList = df["Total Messages"].tolist()

    #Accesses the xlsxwriter workbook and worksheet objects from the dataframe
    workbook  = writer.book
    worksheet1 = writer.sheets['Report']
    worksheet2 = workbook.add_worksheet('More info')

    #Creates a chart object
    chart = workbook.add_chart({'type': 'bar', 'subtype': 'stacked'})

    #Hex colors for the chart.
    colors = ['#930404', '#4200AD']

    #Configures the series of the chart from the dataframe data
    for col_num in range(4, 6):
        chart.add_series({
            'name':       ['Report', 0, col_num], #Legend
            'categories': ['Report', len(sentMsg), 2, 1, 2], #Name of user
            'values':     ['Report', len(sentMsg), col_num, 1, col_num], #Values
            'fill':       {'color':  colors[col_num - 4]},
            'gap':        14,
        })

    #Finds longest name to determine width of chart
    max_len = 0
    for msgSender in sentMsg:
        if len(msgSender) > max_len:
            max_len = len(msgSender)

    #Configures the chart
    chart.set_y_axis({'major_gridlines': {'visible': True}, 'num_font': {'size': 11}, 'reverse': True})
    chart.set_x_axis({'major_gridlines': {'visible': True}, 'label_position': 'high'})
    chart.set_size({'width': 460+round(max_len*8), 'height': 160+len(sentMsg)*21}) #Sets width based on longest user, sets height of chart based on user count
    chart.set_legend({'position': 'top', 'font': {'size': 14, 'bold': True}})
    chart.set_title({'name': roomName+'\n'+datetime.fromtimestamp(firstDate).strftime("%m/%d/%Y") + ' - ' + datetime.fromtimestamp(lastDate).strftime("%m/%d/%Y"), 'overlay': False})

    #Inserts the chart into the worksheet, also inserts and formats all the data on the 2nd page
    alignRight = workbook.add_format()
    alignRight.set_align('right')
    worksheet1.insert_chart(0, 7, chart)
    worksheet1.set_column(0, 0, 0)
    worksheet1.set_column(1, 2, 22)
    worksheet1.set_column(1, 2, 22)
    worksheet1.set_column(3, 3, 14)
    worksheet1.set_column(4, 5, 8)
    worksheet1.set_column(6, 6, 11.9, alignRight)
    worksheet2.set_column(1, 1, 22)
    worksheet2.set_column(2, 2, 22, alignRight)
    worksheet2.write_string(1, 1, 'Report date range:')
    worksheet2.write_string(2, 1, 'Report run date:')
    worksheet2.write_string(3, 1, 'Total messages:')
    worksheet2.write_string(4, 1, 'Users in space:')
    worksheet2.write_string(5, 1, '0 message users:')
    worksheet2.write_string(6, 1, 'Messages per week:')
    worksheet2.write_string(7, 1, 'Messages per user:')
    worksheet2.write_string(1, 2, datetime.fromtimestamp(firstDate).strftime("%m/%d/%Y") + ' - ' + datetime.fromtimestamp(lastDate).strftime("%m/%d/%Y")) #Report range
    worksheet2.write_string(2, 2, datetime.now().strftime("%m/%d/%Y"), alignRight) #Report run date
    worksheet2.write_number(3, 2, messageCount) #Message count
    worksheet2.write_number(4, 2, len(people)) #User count
    worksheet2.write_number(5, 2, noMessages) #Users who have sent 0 messages
    worksheet2.write_number(6, 2, round(messageCount/((lastDate-firstDate)/604800), 2)) #Messages per week
    worksheet2.write_number(7, 2, round(messageCount/len(people), 2)) #Messages per user
    worksheet2.write_string(9, 1, "*dates are in mm/dd/yyyy format") #info for user

#Determines if the user is on mac or windows, and then tells them the file location
if str(os.path.dirname(os.path.abspath(__file__))).startswith("/"):
    print("Csv output written to "+str(os.path.dirname(os.path.abspath(__file__)))+"/"+csv_file)
    print("Excel output written to "+str(os.path.dirname(os.path.abspath(__file__)))+"/"+csv_file[:-4]+".xlsx\n")
else:
    print("Csv output written to "+str(os.path.dirname(os.path.abspath(__file__)))+"\\"+csv_file)
    print("Excel output written to "+str(os.path.dirname(os.path.abspath(__file__)))+"\\"+csv_file[:-4]+".xlsx\n")

if platform == "darwin" or platform == "win32": 
    openFile = input("Would you like to open the excel file? yes/no: ").lower()
    if openFile == "yes" or openFile == "y":
        if platform == "darwin": 
            os.system('open  "'+csv_file[:-4]+'.xlsx" -a  "Microsoft Excel.app"')
        else:
            os.system('start "EXCEL.EXE" "'+csv_file[:-4]+'.xlsx"')