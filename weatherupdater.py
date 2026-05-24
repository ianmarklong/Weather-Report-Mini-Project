import requests
import win32com.client
import sys

def getDaysData(i, data):
    try:
        day = data['data']['records'][0]['forecasts'][i]

        weather = day['forecast']['summary']

        tempLow = day['temperature']['low']
        tempHigh = day['temperature']['high']
        temp = f"{tempLow} - {tempHigh}"

        if i == 0:
            windLow = day['wind']['speed']['low']
            windHigh = day['wind']['speed']['high']
            windDirection = day['wind']['direction']
            wind = f"{windLow} - {windHigh} {windDirection}"

            dayData = {
                "weather": weather,
                "temp": temp,
                "wind": wind
            }

        elif i == 1:
            date = day['timestamp'][:10]
            dayData = {
                "weather": weather,
                "temp": temp,
                "date": date
            }

        else:
            dayData = {
                "weather": weather,
                "temp": temp
            }

        return dayData

    except KeyError as error:
        print(f"Missing expected API field: {error}")
        sys.exit()

    except IndexError:
        print("The API did not return 4 forecast days.")
        sys.exit()

def exportData(daysData):

    #find opened excel file containing 'DOB.xlsm'
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        print("Microsoft Excel is not currently open.")
        print("Please open the DOB workbook first, then run this script again.")
        return

    targetWorkBook = None
    for workbook in excel.Workbooks:
        if 'DOB.xlsm' in workbook.Name or 'DOB.xlsx' in workbook.Name:
            print(workbook.Name)
            targetWorkBook = workbook 

    if targetWorkBook is None:
        print("No open workbook containing 'DOB.xlsm' or 'DOB.xlsx' was found.")
        print("Please open the correct daily update Excel file first.")
    
        return
    else:
        try:
            targetSheet = targetWorkBook.Worksheets("Weather")
        except Exception:
            print("The workbook was found, but the worksheet 'Weather' was not found.")
            print("Please check the worksheet tab name.")
            return

        try:
            targetSheet.Range("K4").Value = daysData[0]["weather"]
            targetSheet.Range("K5").Value = daysData[0]["temp"]
            targetSheet.Range("K7").Value = daysData[0]["wind"]

            targetSheet.Range("J10").Value = daysData[1]["date"]
            targetSheet.Range("K10").Value = daysData[1]["temp"]
            targetSheet.Range("L10").Value = daysData[1]["weather"]

            targetSheet.Range("K11").Value = daysData[2]["temp"]
            targetSheet.Range("L11").Value = daysData[2]["weather"]

            targetSheet.Range("K12").Value = daysData[3]["temp"]
            targetSheet.Range("L12").Value = daysData[3]["weather"]

            targetWorkBook.Save()

            print("Weather data written successfully.")

        except Exception as error:
            print("Something went wrong while writing to Excel.")
            print(f"Technical error: {error}")
            return


url = "https://api-open.data.gov.sg/v2/real-time/api/four-day-outlook"
try:
    response = requests.get(url, timeout=10)
except requests.exceptions.RequestException as error:
    print("Could not connect to the NEA weather API.")
    print(f"Technical error: {error}")
    sys.exit()

if response.status_code != 200:
    print(f"Failed to retrieve data from NEA API.")
    print(f"Status code: {response.status_code}")
    sys.exit()

elif response.status_code == 200: # Check if the request was successful
    try:
        data = response.json()
    except ValueError:
        print("The API responded, but the data was not valid JSON.")
        sys.exit()

    #collate relevant data
    daysData =[]
    for i in range(4):
        dayData = getDaysData(i,data)
        daysData.append(dayData)
    
    exportData(daysData)

input("Press Enter to close...")

