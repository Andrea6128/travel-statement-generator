from openpyxl import Workbook, load_workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment
import random
import datetime
import sys


# generate random times
def randomWorkTime():
    startHour, startMinute = 6, 0
    randomHour, randomMinute = random.randint(0, 3), random.randint(0, 59)
    startTime = datetime.datetime(1970, 1, 1, hour=startHour+randomHour, minute=startMinute+randomMinute)

    endHour, endMinute = 22, 0
    randEndHour, randEndMinute = random.randint(0, 1), random.randint(0, 59)
    endTime = datetime.datetime(1970, 1, 1, hour=endHour+randEndHour, minute=endMinute+randEndMinute)

    totalTime = endTime - startTime

    startTimeStr = str(startTime)
    startTimeFinal = startTimeStr[12:16]

    endTimeStr = str(endTime)
    endTimeFinal = endTimeStr[11:16]

    totalTimeStr = str(totalTime)
    totalTimeFinal = totalTimeStr[:-3]

    return [startTimeFinal, endTimeFinal, totalTimeFinal]

# generate random start city
def dayRoute(routeType):
    dayRouteResult = []
    ws = wb1.active  # set input excel active
    randomStartCity = random.randint(1, 71)  # select random city combination from the base city
    randomStartCity2 = random.randint(72, 2556)  # select random city combination from the other cities except first

    if routeType == 4:
        cityFrom = ws['A'+str(randomStartCity)].value
        cityTo = ws['B'+str(randomStartCity)].value
        distance = ws['C'+str(randomStartCity)].value
        travelTime = ws['D'+str(randomStartCity)].value

        dayRouteResult.append(cityFrom)
        dayRouteResult.append(cityTo)
        dayRouteResult.append(str(distance))
        dayRouteResult.append(str(travelTime))

        return dayRouteResult
    
    # if routeType == 6:
    #     cityFrom = ws['A'+str(randomStartCity)].value
    #     cityTo1 = ws['B'+str(randomStartCity)].value
    #     cityFrom2 = cityTo1
    #     cityTo2 = ws['A'+str(randomStartCity2)].value
    #     cityFrom3 = cityTo2
    #     cityTo3 = cityFrom
    #     distance1 = ws['C'+str(randomStartCity)].value
    #     distance2 = ws['C'+str(randomStartCity2)].value
    #     distance3 = distance1
    #     travelTime1 = ws['D'+str(randomStartCity)].value
    #     travelTime2 = ws['D'+str(randomStartCity2)].value
    #     travelTime3 = travelTime1

    #     dayRouteResult.append(cityFrom)
    #     dayRouteResult.append(cityTo1)
    #     dayRouteResult.append(str(distance1))
    #     dayRouteResult.append(str(travelTime1))
    #     dayRouteResult.append(cityFrom2)
    #     dayRouteResult.append(cityTo2)
    #     dayRouteResult.append(str(distance2))
    #     dayRouteResult.append(str(travelTime2))
    #     dayRouteResult.append(cityFrom3)
    #     dayRouteResult.append(cityTo3)
    #     dayRouteResult.append(str(distance3))
    #     dayRouteResult.append(str(travelTime3))

    #     return dayRouteResult


def drawLines(startRow):
    bottomBorder = Border(left=Side(style='none'),
                          right=Side(style='none'), 
                          top=Side(style='none'), 
                          bottom=Side(style='thin'))
    bottomRightBorder = Border(left=Side(style='none'),
                               right=Side(style='medium'), 
                               top=Side(style='none'), 
                               bottom=Side(style='thin'))
    bottomLeftBorder = Border(left=Side(style='medium'),
                              right=Side(style='none'), 
                              top=Side(style='none'), 
                              bottom=Side(style='thin'))

    ws.cell(row=startRow, column=1).border = bottomLeftBorder
    for col in range(2, 14):
        ws.cell(row=startRow, column=col).border = bottomBorder
        if col == 13:
            ws.cell(row=startRow, column=col).border = bottomRightBorder


def fillSheet(startRow, startColumn, startDate, numberOfDays):
    for day in range(numberOfDays):

        # singleDay generates how many routes per day will be done (1, 2 or 3)
        # singleDay = random.randrange(4, 8, 2)  # (start, stop, step)

        singleDay = 4  # only 4 lines per route to save space on sheet

        getStartRoute = dayRoute(singleDay)  # get dayRoute function result
        getRandomWorkTime = randomWorkTime()  # get start hour, end hour and division

        print(getStartRoute)
        print(getRandomWorkTime)

        if singleDay == 4:  # if route between 2 cities
            for event in range(singleDay):
                # startDateStr = startDate.strftime('%d.%m.%Y')  # starting date (usualy 1.1.2020)
                # ws.cell(row=startRow, column=startColumn, value=startDateStr)  # fill cell with date

                # 1st route
                if event == 0:
                    startDateStr = startDate.strftime('%d.%m.%Y')  # starting date (usualy 1.1.2020)
                    ws.cell(row=startRow, column=startColumn, value=startDateStr)  # fill cell with date

                    ws.cell(row=startRow, column=startColumn+1, value="odchod")  # fill odchod
                    ws.cell(row=startRow, column=startColumn+2, value=getStartRoute[0])  # fill start city
                    ws.cell(row=startRow, column=startColumn+4, value="AUS").alignment = Alignment(vertical="center", horizontal="center")
                    ws.cell(row=startRow, column=startColumn+5, value=getStartRoute[2]).alignment = Alignment(vertical="center", horizontal="center")  # fill km
                    ws.cell(row=startRow, column=startColumn+6, value=getRandomWorkTime[0]).alignment = Alignment(vertical="center", horizontal="center")  # fill start work time
                if event == 1:
                    ws.cell(row=startRow, column=startColumn+1, value="príchod")  # fill prichod
                    ws.cell(row=startRow, column=startColumn+2, value=getStartRoute[1])  # fill destination city
                    ws.cell(row=startRow, column=startColumn+6, value=getRandomWorkTime[1]).alignment = Alignment(vertical="center", horizontal="center")  # fill end work time

                # 2nd route
                if event == 2:
                    ws.cell(row=startRow, column=startColumn+1, value="odchod")  # fill odchod
                    ws.cell(row=startRow, column=startColumn+2, value=getStartRoute[1])  # fill destination city
                    ws.cell(row=startRow, column=startColumn+4, value="AUS").alignment = Alignment(vertical="center", horizontal="center")
                    ws.cell(row=startRow, column=startColumn+5, value=getStartRoute[2]).alignment = Alignment(vertical="center", horizontal="center")  # fill km
                if event == 3:
                    ws.cell(row=startRow, column=startColumn+1, value="príchod")  # fill prichod
                    ws.cell(row=startRow, column=startColumn+2, value=getStartRoute[0])  # fill start city

                    # draw lines                    
                    # drawLines(startRow)

                # ws.merge_cells('A'+str(startRow-3)+':A'+str(startRow))

                # if startRow == 80:
                #     print("konec stranky")
                #     break

                startRow += 1

            # increment day by 1
            startDate = startDate + datetime.timedelta(days=1)
            print(startDate)

        # if singleDay == 6:  # if route between 3 cities
        #     for event in range(singleDay):
        #         startDateStr = startDate.strftime('%d.%m.%Y')  # starting date (usualy 1.1.2020)
        #         ws.cell(row=startRow, column=startColumn, value=startDateStr)  # fill cell with date

        #         # 1st route
        #         if event == 0:
        #             ws.cell(row=startRow, column=startColumn+1, value="odchod")  # fill odchod
        #             ws.cell(row=startRow, column=startColumn+2, value=getStartRoute[0])  # fill start city
        #             ws.cell(row=startRow, column=startColumn+4, value="AUS").alignment = Alignment(vertical="center", horizontal="center")
        #             ws.cell(row=startRow, column=startColumn+5, value=getStartRoute[2]).alignment = Alignment(vertical="center", horizontal="center")  # fill km
        #             ws.cell(row=startRow, column=startColumn+6, value=getRandomWorkTime[0]).alignment = Alignment(vertical="center", horizontal="center")  # fill start work time

        #         if event == 1:
        #             ws.cell(row=startRow, column=startColumn+1, value="príchod")  # fill prichod
        #             ws.cell(row=startRow, column=startColumn+2, value=getStartRoute[1])  # fill destination city
        #             ws.cell(row=startRow, column=startColumn+6, value=getRandomWorkTime[1]).alignment = Alignment(vertical="center", horizontal="center")  # fill end work time

        #         # 2nd route
        #         if event == 2:
        #             ws.cell(row=startRow, column=startColumn+1, value="odchod")  # fill odchod
        #             ws.cell(row=startRow, column=startColumn+2, value=getStartRoute[4])  # fill destination city
        #             ws.cell(row=startRow, column=startColumn+4, value="AUS").alignment = Alignment(vertical="center", horizontal="center")
        #             ws.cell(row=startRow, column=startColumn+5, value=getStartRoute[6]).alignment = Alignment(vertical="center", horizontal="center")  # fill km
        #         if event == 3:
        #             ws.cell(row=startRow, column=startColumn+1, value="príchod")  # fill prichod
        #             ws.cell(row=startRow, column=startColumn+2, value=getStartRoute[5])  # fill start city

        #         # 3rd route
        #         if event == 4:
        #             ws.cell(row=startRow, column=startColumn+1, value="odchod")  # fill odchod
        #             ws.cell(row=startRow, column=startColumn+2, value=getStartRoute[8])  # fill destination city
        #             ws.cell(row=startRow, column=startColumn+4, value="AUS").alignment = Alignment(vertical="center", horizontal="center")
        #             ws.cell(row=startRow, column=startColumn+5, value=getStartRoute[10]).alignment = Alignment(vertical="center", horizontal="center")  # fill km
        #         if event == 5:
        #             ws.cell(row=startRow, column=startColumn+1, value="príchod")  # fill prichod
        #             ws.cell(row=startRow, column=startColumn+2, value=getStartRoute[9])  # fill start city

        #             # draw lines
        #             drawLines(startRow)

        #         if startRow == 64:
        #             print("konec stranky")
        #             break

        #         startRow += 1

        #     # increment day by 1
        #     startDate = startDate + datetime.timedelta(days=1)
        #     print(startDate)


if __name__ == '__main__':
    wb1 = load_workbook("mesta_input.xlsx")
    wb2 = load_workbook("output.xlsx")
    ws = wb2.active  # set 2nd excel active

    # generateDates params: startRow, startColumn, startDate, numberOfDays
    fillSheet(6, 1, datetime.datetime.strptime('2020-02-01', '%Y-%m-%d'), 31)
    wb2.save("output.xlsx")
