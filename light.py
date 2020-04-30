# Created by N.Kazakov ver 1.00

# python imports
import sys
import calendar
import datetime
import re
from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL

# classes
class Worker:
    __gross_salary = 0
    __com_gross_salary = 0
    __net_salary = 0
    __card = 0
    __costs = 0
    __prepayment = 0
    __working_days = 0
    __days_worked = 0
    count_red_rate = [0] * 31
    count_yellow_rate = [0] * 31
    count_green_rate = [0] * 31
    count_blue_rate = [0] * 31
    sum_red = [0] * 31
    sum_yellow = [0] * 31
    sum_green = [0] * 31
    sum_blue = [0] * 31

    def __init__(self, name):
        if name:
            self.__name = name.strip()
            self.rate = []

    def getName(self):
        return self.__name

    def getGrossSalary(self):
        return self.__gross_salary, self.__com_gross_salary

    def getNetSalary(self):
        return self.__net_salary

    def getCard(self):
        return self.__card

    def getCosts(self):
        return self.__costs

    def getPrepayment(self):
        return self.__prepayment

    def getDaysWorked(self):
        return self.__days_worked

    def setGrossSalary(self, daypay, com_daypay):
        self.__gross_salary += round(daypay, 0)
        self.__com_gross_salary += round(com_daypay, 0)

    def setNetSalary(self, net_salary):
        self.__net_salary += round(net_salary, 0)

    def setCard(self, card):
        self.__card += round(card, 0)

    def setCosts(self, costs):
        self.__costs += round(costs, 0)

    def setPrepayment(self, prepayment):
        self.__prepayment += round(prepayment, 0)

    def setWorkingDays(self, working_days):
        self.__working_days = round(working_days, 1)

    def setDaysWorked(self):
        self.__days_worked += 1

    def setRate(self, cell, day):
        self.rate.append({cell.CellBackColor : round(cell.Value, 1)})
        if cell.CellBackColor == 16711680:
            Worker.count_red_rate[day] += round(cell.Value, 1)
            self.setDaysWorked()
        elif cell.CellBackColor == 16776960:
            Worker.count_yellow_rate[day] += round(cell.Value, 1)
            self.setDaysWorked()
        elif cell.CellBackColor == 43315:
            Worker.count_green_rate[day] += round(cell.Value, 1)
            self.setDaysWorked()
        elif cell.CellBackColor == 2201331:
            Worker.count_blue_rate[day] += round(cell.Value, 1)
            self.setDaysWorked()

    def setSum(self, cell, day):
        a_CellBackColor = cell.CellBackColor
        if cell.CellBackColor == 16711680 and Worker.sum_red[day] == 0:
            Worker.sum_red[day] = round(cell.Value, 1)
        elif cell.CellBackColor == 16776960 and Worker.sum_yellow[day] == 0:
            Worker.sum_yellow[day] = round(cell.Value, 1)
        elif cell.CellBackColor == 43315 and Worker.sum_green[day] == 0:
            Worker.sum_green[day] = round(cell.Value, 1)
        elif cell.CellBackColor == 2201331 and Worker.sum_blue[day] == 0:
            Worker.sum_blue[day] = round(cell.Value, 1)

    def countDaypay(self, day):
        daypay = 0
        com_daypay = 0
        a_rate = self.rate
        rate = self.rate[day]
        com_sum = Worker.sum_red[day] + Worker.sum_yellow[day] + Worker.sum_green[day] + Worker.sum_blue[day]
        com_count_rate = Worker.count_red_rate[day] + Worker.count_yellow_rate[day] + Worker.count_green_rate[day] + Worker.count_blue_rate[day]
        if rate.get(16711680):
            daypay = round(rate[16711680] * Worker.sum_red[day] / Worker.count_red_rate[day], 0)
            com_daypay = round(rate[16711680] * com_sum / com_count_rate, 0)
        elif rate.get(16776960):
            daypay = round(rate[16776960] * Worker.sum_yellow[day] / Worker.count_yellow_rate[day], 0)
            com_daypay = round(rate[16776960] * com_sum / com_count_rate, 0)
        elif rate.get(43315):
            daypay = round(rate[43315] * Worker.sum_green[day] / Worker.count_green_rate[day], 0)
            com_daypay = round(rate[43315] * com_sum / com_count_rate, 0)
        elif rate.get(2201331):
            daypay = round(rate[2201331] * Worker.sum_blue[day] / Worker.count_blue_rate[day], 0)
            com_daypay = round(rate[2201331] * com_sum / com_count_rate, 0)
        self.setGrossSalary(daypay, com_daypay)
        return daypay, com_daypay

    @staticmethod
    def getWorkingDays(date_str):
        year = int(date_str[6:10])
        month = int(date_str[3:5])
        day = int(date_str[0:2])
        date = datetime.date(year, month, day)
        cal = calendar.Calendar()
        working_days = len([x for x in cal.itermonthdays2(date.year, date.month) if x[0] != 0 and x[1] < 5])
        return working_days

    @staticmethod
    def clearLists():
        Worker.count_red_rate = [0] * 31
        Worker.count_yellow_rate = [0] * 31
        Worker.count_green_rate = [0] * 31
        Worker.count_blue_rate = [0] * 31
        Worker.sum_red = [0] * 31
        Worker.sum_yellow = [0] * 31
        Worker.sum_green = [0] * 31
        Worker.sum_blue = [0] * 31

    @staticmethod
    def showMessage(parentwin, error):
        box = parentwin.getToolkit().createMessageBox(
            parentwin, ERRORBOX, BUTTONS_OK, "Ошибка", error)
        box.execute()

def count(*args):
    # get the doc from the scripting context.which is made available to all scripts
    desktop = XSCRIPTCONTEXT.getDesktop()
    model = desktop.getCurrentComponent()
    sheets = model.Sheets
    parentwin = model.CurrentController.Frame.ContainerWindow
    workers = []
    num_sheet = [1]
    num_worker = list(range(20))
    cells_days_month = [
        "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R",
        "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG"]
    i = 0
    for i in [*num_worker]:  # Creating object of worker and loading name of worker
        if sheets[1].getCellRangeByName("B" + str(i + 5)).String:
            if not re.search('[^А-я .]', sheets[1].getCellRangeByName("B" + str(i + 5)).String):
                worker = Worker(sheets[1].getCellRangeByName("B" + str(i + 5)).String)
                workers.append(worker)
            else:
                error = "Неккоректное имя работника"
                Worker.showMessage(parentwin, error)
                try:
                    sys.exit()
                except SystemExit:
                    return None
    for i in [*num_sheet]:
        sheet = sheets[i]
        for i in range(len(workers)):
            worker = workers[i]
            for i in range(len(workers)):
                name = sheet.getCellRangeByName("B" + str(i + 5)).String
                name = name.strip()
                a_name_worker = worker.getName()
                if worker.getName() == name:
                    day = 0
                    for letter in [*cells_days_month]:
                        cell_rate = sheet.getCellRangeByName(letter + str(i + 5))
                        worker.setRate(cell_rate, day)
                        a_worker_rate = worker.rate
                        for j in range(26, 30):
                            cell_sum = sheet.getCellRangeByName(letter + str(j))
                            worker.setSum(cell_sum, day)
                        day += 1
        # Zeroing summation cells
        sheet.getCellRangeByName("AK52").Value = 0
        sheet.getCellRangeByName("AL52").Value = 0
        # Fill in the salary table
        for i in range(len(workers)):
            worker = workers[i]
            sheet.getCellRangeByName("A" + str(i + 32)).Value = i + 1
            sheet.getCellRangeByName("A" + str(i + 55)).Value = i + 1
            sheet.getCellRangeByName("B" + str(i + 32)).String = worker.getName()
            sheet.getCellRangeByName("B" + str(i + 55)).String = worker.getName()
            day = 0
            for letter in [*cells_days_month]:
                daypay, com_daypay = worker.countDaypay(day)
                sheet.getCellRangeByName(letter + str(i + 32)).Value = daypay
                sheet.getCellRangeByName(letter + str(i + 55)).Value = com_daypay
                day += 1
            worker.setCard(sheet.getCellRangeByName("AH" + str(i + 32)).Value)
            worker.setCosts(sheet.getCellRangeByName("AI" + str(i + 32)).Value)
            worker.setPrepayment(sheet.getCellRangeByName("AJ" + str(i + 32)).Value)
            gross_salary, com_gross_salary = worker.getGrossSalary()
            net_salary = gross_salary - worker.getCard() - worker.getCosts() - worker.getPrepayment()
            worker.setNetSalary(net_salary)
            sheet.getCellRangeByName("AK" + str(i + 32)).Value = gross_salary
            sheet.getCellRangeByName("AK52").Value += gross_salary
            sheet.getCellRangeByName("AL" + str(i + 32)).Value = net_salary
            sheet.getCellRangeByName("AL52").Value += net_salary
    # Zeroing summation cells
    sheets[0].getCellRangeByName("F2").FormulaLocal = sheets[1].getCellRangeByName("L2").FormulaLocal
    sheets[0].getCellRangeByName("K26").Value = 0
    sheets[0].getCellRangeByName("L26").Value = 0
    sheets[0].getCellRangeByName("M26").Value = 0
    sheets[0].getCellRangeByName("N26").Value = 0
    sheets[0].getCellRangeByName("O26").Value = 0
    # Fill in the salary table
    for i in range(len(workers)):
        sheets[0].getCellRangeByName("A" + str(i + 6)).Value = i + 1
        sheets[0].getCellRangeByName("B" + str(i + 6)).String = workers[i].getName()
        sheets[0].getCellRangeByName("E" + str(i + 6)).Value = Worker.getWorkingDays(
            sheets[0].getCellRangeByName("F2").FormulaLocal)
        sheets[0].getCellRangeByName("F" + str(i + 6)).Value = workers[i].getDaysWorked()
        sheets[0].getCellRangeByName("G" + str(i + 6)).Value = workers[i].getCard()
        sheets[0].getCellRangeByName("H" + str(i + 6)).Value = workers[i].getCosts()
        sheets[0].getCellRangeByName("J" + str(i + 6)).Value = workers[i].getPrepayment()
        # Correct salary worker
        add_salary = sheets[0].getCellRangeByName("D" + str(i + 6)).Value
        correct_salary = add_salary - sheets[0].getCellRangeByName("C" + str(i + 6)).Value
        correct_salary -= sheets[0].getCellRangeByName("I" + str(i + 6)).Value
        # Gross salary worker
        gross_salary, com_gross_salary = workers[i].getGrossSalary()
        val_gross_salary = gross_salary + correct_salary
        sheets[0].getCellRangeByName("K" + str(i + 6)).Value = val_gross_salary
        sheets[0].getCellRangeByName("K26").Value += val_gross_salary
        # Net salary worker
        val_net_salary = workers[i].getNetSalary() + correct_salary
        sheets[0].getCellRangeByName("L" + str(i + 6)).Value = val_net_salary
        sheets[0].getCellRangeByName("L26").Value += val_net_salary
        # Sharing salary worker on percent
        if sheets[0].getCellRangeByName("M5").Value < 0:
            sheets[0].getCellRangeByName("M5").Value = 0
        val_percent_salary = sheets[0].getCellRangeByName("M5").Value
        if val_percent_salary > 1:
            sheets[0].getCellRangeByName("M5").Value = 1
            val_percent_salary = 1
        val_salary = (workers[i].getNetSalary() + correct_salary) * val_percent_salary
        sheets[0].getCellRangeByName("M" + str(i + 6)).Value = val_salary
        sheets[0].getCellRangeByName("M26").Value += val_salary
        val_percent_salary = 1 - val_percent_salary
        if sheets[0].getCellRangeByName("N5").Value < 0:
            sheets[0].getCellRangeByName("N5").Value = 0
        if sheets[0].getCellRangeByName("N5").Value <= val_percent_salary:
            val_percent_salary -= sheets[0].getCellRangeByName("N5").Value
            val_salary = (workers[i].getNetSalary() + correct_salary) * sheets[0].getCellRangeByName("N5").Value
            sheets[0].getCellRangeByName("N" + str(i + 6)).Value = val_salary
            sheets[0].getCellRangeByName("N26").Value += val_salary
        else:
            sheets[0].getCellRangeByName("N5").Value = val_percent_salary
            val_salary = (workers[i].getNetSalary() + correct_salary) * val_percent_salary
            sheets[0].getCellRangeByName("N" + str(i + 6)).Value = val_salary
            sheets[0].getCellRangeByName("N26").Value += val_salary
        sheets[0].getCellRangeByName("O5").Value = val_percent_salary
        val_percent_salary = sheets[0].getCellRangeByName("O5").Value
        val_salary = (workers[i].getNetSalary() + correct_salary) * val_percent_salary
        sheets[0].getCellRangeByName("O" + str(i + 6)).Value = val_salary
        sheets[0].getCellRangeByName("O26").Value += val_salary
    Worker.clearLists()
    return None