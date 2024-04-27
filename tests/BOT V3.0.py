import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
from datetime import datetime, timedelta
import time
from typing import List


def postOnFacebook(confession, epochTime, fbPageid, fbToken):
    response = requests.post("https://graph.facebook.com/" + fbPageid + "/feed" +
                    "?published=false" +
                    "&message=" + confession +
                    "&scheduled_publish_time=" + str(int(epochTime)) +
                    "&access_token=" + fbToken)
    return response


def get_time(sheet_time:str, timeList: list)-> str:
    """
    Get the closest time based on the current time and a provided target time.
    
    Parameters:
    - sheet_time (str): The target time in the format "dd-mm-yyyy hh:mm[am/pm]".

    Returns:
    - str: The closest time to the current time. 
        If the current time is later than the target time, it returns the closest upcoming hour from the target times list.
        Otherwise, it returns the target time.
    """

    # Convert the string to datetime
    target_time = datetime.strptime(sheet_time, "%d-%m-%Y %I:%M%p")

    # Add 1 hour to the  now time
    time_now = datetime.now() + timedelta(hours=1)

    current_time = max(time_now, target_time)
        
    # Find the closest upcoming time to current_time from the target_times
    future_time = datetime.strptime(current_time.strftime('%I:%M%p'), "%I:%M%p").time()

    closest_time = min([time for time in timeList if datetime.strptime(time, "%I:%M%p").time() > future_time], 
            key=lambda x: abs(future_time.hour - datetime.strptime(x, "%I:%M%p").time().hour))

    # Convert closest_time to full datetime
    closest_datetime = f'{current_time.strftime("%d-%m-%Y")} {closest_time}'

    postTime = closest_datetime

    return postTime


def get_confessions(sheet: object ,startPoint: int, endPoint:int)-> List[List]:
    if endPoint == 0:
        # get index of last confession
        endPoint = len(sheet.get_all_values())
    else:
        #get starting index and add number of confessions posted to it 
        #Note: -1 beacuse it includes endpoint row with it, ie row 5 to 8 is 4 elements not 3
        endPoint = startPoint + endPoint-1

    # get all confessions one call as a list
    confessions = sheet.get(f'A{startPoint}:Z{endPoint}')
    return confessions

def connect_sheets(creds):
    client = gspread.authorize(creds)
    sheet1 = client.open("NU Confessions (Responses)").sheet1  # Open the spreadsheet
    sheet2 = client.open("NU Confessions (Responses)").worksheet("Sheet2")

    return sheet1, sheet2

# Tokens and Credentials
#------------------------------------------------------------------------------------------------------------------------------#
f = open("Facebook token.txt", "r+")
fbPageid = f.readline().strip("\n")
fbToken = f.readline().strip("\n")
f.close()

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
        "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
#------------------------------------------------------------------------------------------------------------------------------#

print("Loading spreadsheet...")

sheet, sheet2 = connect_sheets(creds)
confNum = int(sheet2.acell("A1").value)
startPoint = int(sheet2.acell("B1").value)
print("Number of new confessions: " + str(len(sheet.get_all_values()) - startPoint + 1))
print()

timeList = []
for i in range(1, 100):
    value = sheet2.acell("F" + str(i)).value
    if value == None:
        break
    else:
        timeList.append(value)

try:
    postTime = get_time(sheet2.acell("C1").value, timeList)
    timeIndex = timeList.index(postTime.split(" ")[1].lower())  # Get index of current time to post
except:
    print("Error within last scheduled time in sheet2!")
    startPoint = 9999999

print(postTime)


day = int(postTime.split(" ")[0].split("-")[0])
month = int(postTime.split(" ")[0].split("-")[1])
year = int(postTime.split(" ")[0].split("-")[2])

endPoint = int(input("Enter number of posts - 0 to post all: "))
confessions = get_confessions(sheet,startPoint, endPoint)

c = 1

conf_index = startPoint
for confession_data in confessions:
    # delay for Gspread api write request quota
    if c % 5 == 0:
        time.sleep(30)

    confession = "%23" + str(confNum) + "\n" + confession_data[2] +', '+ confession_data[3] +', '+ confession_data[4] + ': \n\n' + confession_data[1]
    confession = confession.replace(", Prefer not to say", "").replace("Prefer not to say", "").replace("#", "%23")

    print(confession)
    print()
    confNum+=1

    is_skipped = confession_data[-1]
    if is_skipped == '1':
        print("Skipped")
        sheet.format("B" + str(conf_index), {"backgroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}})
        
    else:
        timeIndex += 1
        if timeIndex >= len(timeList):
            print("Schedule time overflow!")
            date = datetime(year, month, day)
            date += timedelta(days=1)
            day = date.day
            month = date.month
            year = date.year
            timeIndex = 0

        postTime = timeList[timeIndex].split(":")
        hour = int(postTime[0])
        minute = int(postTime[1][:-2])
        if (postTime[1][-2] + postTime[1][-1]).lower() == "pm":
            hour += 12
        date = datetime(year, month, day, hour, minute)
        epochTime = date.timestamp()
        print("Post Time: ", date)
        # Facebook scheduling request
        response = postOnFacebook(confession, epochTime, fbPageid, fbToken)
        print(str(response.json()))
        if str(response.json())[2:7] == "error":
            break

        sheet.update(range_name="G" + str(conf_index), values=[[str(confNum)]])
        sheet2.update(range_name="C1", values=[[str(day) + "-" + str(month) + "-" + str(year) + " " + postTime[0] + ":" + postTime[1]]])
        sheet.update(range_name="F" + str(conf_index), values=[[str(day) + "-" + str(month) + "-" + str(year) + " " + postTime[0] + ":" + postTime[1]]])
        sheet.format("B" + str(conf_index), {"backgroundColor": {"red": 0.0, "green": 1.0, "blue": 0.0}})
        print("Posted successfully!")
        confNum += 1

    sheet2.update(range_name="A1", values=[[str(confNum)]])
    sheet2.update(range_name="B1", values=[[str(conf_index + 1)]])

    conf_index+=1
    c += 1

    print("--------------------------------------------------\n")

print("Program End!")
input("")
