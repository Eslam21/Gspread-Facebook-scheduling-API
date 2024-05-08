import gspread
import requests
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime, timedelta
import time
from typing import List


def post_on_facebook(confession: str, epoch_time: int, fb_page_id: str, fb_token: str) -> requests.Response:
    """
    Post a confession on Facebook.

    Parameters:
    - confession (str): The content of the confession.
    - epoch_time (int): The scheduled publish time in epoch format.
    - fb_page_id (str): The Facebook page ID where the confession will be posted.
    - fb_token (str): The Facebook access token.

    Returns:
    - requests.Response: Response from the Facebook API.
    """
    response = requests.post(
        f"https://graph.facebook.com/{fb_page_id}/feed"
        f"?published=false"
        f"&message={confession}"
        f"&scheduled_publish_time={epoch_time}"
        f"&access_token={fb_token}"
    )
    return response


def get_time(sheet_time: str, time_list: List[str]) -> str:
    """
    Get the closest time based on the current time and a provided target time.

    Parameters:
    - sheet_time (str): The target time in the format "dd-mm-yyyy hh:mm[am/pm]".
    - time_list (List[str]): List of target times.

    Returns:
    - str: The closest time to the current time.
    """
    target_time = datetime.strptime(sheet_time, "%d-%m-%Y %I:%M%p")
    time_now = datetime.now() + timedelta(hours=1)
    current_time = max(time_now, target_time)
    future_time = datetime.strptime(current_time.strftime('%I:%M%p'), "%I:%M%p").time()
    closest_time = min(
        [time for time in time_list if datetime.strptime(time, "%I:%M%p").time() > future_time],
        key=lambda x: abs(future_time.hour - datetime.strptime(x, "%I:%M%p").time().hour)
    )
    closest_datetime = f'{current_time.strftime("%d-%m-%Y")} {closest_time}'
    return closest_datetime


def get_confessions(sheet: gspread.worksheet, start_point: int, end_point: int) -> List[List[str]]:
    """
    Retrieve confessions from a Google Sheets document.

    Parameters:
    - sheet (gspread.models.Worksheet): The Google Sheets document.
    - start_point (int): Starting index of confessions.
    - end_point (int): Ending index of confessions.

    Returns:
    - List[List[str]]: List of confessions.
    """
    if end_point == 0:
        end_point = len(sheet.get_all_values())
    else:
        end_point = start_point + end_point - 1
    confessions = sheet.get(f'A{start_point}:Z{end_point}')
    return confessions


def connect_sheets(creds: str) -> tuple:
    """
    Connect to Google Sheets.

    Parameters:
    - creds (str): Google Sheets credentials.

    Returns:
    - tuple: Tuple containing connected sheets.
    """
    client = gspread.authorize(creds)
    sheet1 = client.open("NU Confessions (Responses)").sheet1
    sheet2 = client.open("NU Confessions (Responses)").worksheet("Sheet2")
    return sheet1, sheet2


# Tokens and Credentials
#------------------------------------------------------------------------------------------------------------------------------#
f = open("Facebook token.txt", "r+")
fb_page_id = f.readline().strip("\n")
fb_token = f.readline().strip("\n")
f.close()

scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
        "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
#------------------------------------------------------------------------------------------------------------------------------#

print("Loading spreadsheet...")

sheet, sheet2 = connect_sheets(creds)
conf_num = int(sheet2.acell("A1").value)
start_point = int(sheet2.acell("B1").value)
print("Number of new confessions: " + str(len(sheet.get_all_values()) - start_point + 1))
print()

time_list = []
for i in range(1, 100):
    value = sheet2.acell("F" + str(i)).value
    if value is None:
        break
    else:
        time_list.append(value)

try:
    post_time = get_time(sheet2.acell("C1").value, time_list)
    time_index = time_list.index(post_time.split(" ")[1].lower())  # Get index of current time to post
except:
    print("Error within last scheduled time in sheet2!")
    start_point = 9999999

print(post_time)

day = int(post_time.split(" ")[0].split("-")[0])
month = int(post_time.split(" ")[0].split("-")[1])
year = int(post_time.split(" ")[0].split("-")[2])

end_point = int(input("Enter number of posts - 0 to post all: "))
confessions = get_confessions(sheet, start_point, end_point)

c = 1

conf_index = start_point
for confession_data in confessions:
    if c % 5 == 0:
        time.sleep(30)

    confession = "%23" + str(conf_num) + "\n" + confession_data[2] + ', ' + confession_data[3] + ', ' + confession_data[4] + ': \n\n' + confession_data[1]
    confession = confession.replace(", Prefer not to say", "").replace("Prefer not to say", "").replace("#", "%23")

    print(confession)
    print()

    is_skipped = confession_data[-1]
    if is_skipped == '1':
        print("Skipped")
        sheet.format("B" + str(conf_index), {"backgroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0}})
    else:
        time_index += 1
        if time_index >= len(time_list):
            print("Schedule time overflow!")
            date = datetime(year, month, day)
            date += timedelta(days=1)
            day = date.day
            month = date.month
            year = date.year
            time_index = 0

        post_time = time_list[time_index].split(":")
        hour = int(post_time[0])
        minute = int(post_time[1][:-2])
        if (post_time[1][-2] + post_time[1][-1]).lower() == "pm":
            hour += 12
        date = datetime(year, month, day, hour, minute)
        epoch_time = date.timestamp()
        print("Post Time: ", date)

        response = post_on_facebook(confession, epoch_time, fb_page_id, fb_token)
        print(str(response.json()))
        if str(response.json())[2:7] == "error":
            break

        sheet.update(range_name="G" + str(conf_index), values=[[str(conf_num)]])
        sheet2.update(range_name="C1", values=[[str(day) + "-" + str(month) + "-" + str(year) + " " + post_time[0] + ":" + post_time[1]]])
        sheet.update(range_name="F" + str(conf_index), values=[[str(day) + "-" + str(month) + "-" + str(year) + " " + post_time[0] + ":" + post_time[1]]])
        sheet.format("B" + str(conf_index), {"backgroundColor": {"red": 0.0, "green": 1.0, "blue": 0.0}})
        print("Posted successfully!")
        conf_num += 1

    sheet2.update(range_name="A1", values=[[str(conf_num)]])
    sheet2.update(range_name="B1", values=[[str(conf_index + 1)]])

    conf_index += 1
    c += 1

    print("--------------------------------------------------\n")

print("Program End!")
input("")
