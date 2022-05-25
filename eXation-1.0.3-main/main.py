import sys
from PyQt5.uic import loadUi
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
from PyQt5.QtCore import QThread, pyqtSignal
import win32com.client
from pathlib import Path
import pandas as pd
import numpy as np
import time
import gc


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        loadUi("./main.ui", self)
        self.stackedWidget.setCurrentIndex(0)
        self.rempli()
        self.check()

    # All Rempli
    def rempli(self):
        self.analyics = {}
        self.xl = ""
        self.wb = ""
        self.wbName = ""
        self.incri = 1
        self.incri2 = 1
        self.DiaP = ""
        self.DiaE = ""
        self.Export = ""
        self.worker_1 = Worker_1()
        self.worker_2 = Worker_2()
        self.worker_3 = Worker_3()
        self.worker_6 = Worker_6()
        self.worker_13 = Worker_13()

    # All Pages Checks
    def check(self):
        # Page 0
        self.pushButton.clicked.connect(self.chooseFile)
        # Page 1
        self.pushButton_3.clicked.connect(self.goToHome)
        self.pushButton_2.clicked.connect(self.beginInitial)
        # # Page 2
        self.pushButton_7.clicked.connect(self.doneAll)
        self.pushButton_4.clicked.connect(self.goto_lastPage2)
        self.pushButton_8.clicked.connect(self.goToHome2)
        self.pushButton_9.clicked.connect(self.goToHome5)

        self.pushButton_5.clicked.connect(self.goToAbout)
        self.pushButton_6.clicked.connect(self.goToAbout2)

        # Threads Trigers
        self.worker_3.n.connect(self.setLabelText)

        self.worker_1.finished.connect(self.filtreSheets)
        self.worker_3.finished.connect(self.clearLabelText)

        self.worker_2.finished.connect(self.incrMent)

        self.worker_6.finished.connect(self.startWorker7)

        self.worker_2.notif.connect(self.setLabelText2)
        self.worker_2.progress.connect(self.setprogress)

        self.worker_6.notif.connect(self.setLabelText2)
        self.worker_6.progress.connect(self.setprogress)

        self.worker_13.seco.connect(self.setTimeSeco)

    def goToAbout(self):
        self.stackedWidget.setCurrentIndex(4)

    def goToAbout2(self):
        self.stackedWidget.setCurrentIndex(6)

    def doneAll(self):
        if QMessageBox.Yes == QMessageBox.question(
            self,
            " Done Progress",
            "Are you sure you want to done\nthe progress?",
            QMessageBox.No | QMessageBox.Yes,
        ):
            self.stackedWidget.setCurrentIndex(4)

    def setTimeSeco(self, seco):
        s = str(int(seco[1]))
        m = str(int(seco[0]))

        if len(s) == 1:
            s = "0" + s
        if len(m) == 1:
            m = "0" + m

        self.label_10.setText(m + ":" + s)

    def incrMent(self):
        self.worker_6.start()

    def startWorker7(self):
        self.pushButton_4.setEnabled(True)
        self.worker_13.status = False
        self.goto_lastPage()

    def goto_lastPage2(self):
        self.stackedWidget.setCurrentIndex(3)

    def goto_lastPage(self):

        path = Path().absolute()
        path = str(path) + "\\out\\"
        self.out_path.setText(path)

        self.textBrowser.clear()

        try:
            self.textBrowser.append(
                "Strat progres: " + str(mainwindow.analyics["timeStart"])
            )
            self.textBrowser.append(
                "workbook name: " + str(mainwindow.analyics["workbookName"])
            )

            self.textBrowser.append("Sheets names: ")
            self.textBrowser.append(str(mainwindow.analyics["sheetsNams"]))
            self.textBrowser.append("Export sheets names: ")
            self.textBrowser.append(str(mainwindow.analyics["exportsSheets"]))
            self.textBrowser.append("DiaE sheets names: ")
            self.textBrowser.append(str(mainwindow.analyics["diaESheets"]))
            self.textBrowser.append("DiaP sheets names: ")
            self.textBrowser.append(str(mainwindow.analyics["diaPSheets"]))

            self.textBrowser.append(str(mainwindow.analyics["initDiapTime"]))
            self.textBrowser.append(str(mainwindow.analyics["initaExport"]))
            self.textBrowser.append(str(mainwindow.analyics["initaDiaE"]))

            self.textBrowser.append(str(mainwindow.analyics["calculationTime"]))
            self.textBrowser.append(str(mainwindow.analyics["creatOutTime"]))
            self.textBrowser.append(str(mainwindow.analyics["timeToFinish"]))
        except:
            pass

    def chooseFile(self):
        fileName = QFileDialog.getOpenFileName(
            self, "Select Exel File", "", "Exel File (*.xlsx *.xlsm *.xltx *.xltm) "
        )
        try:
            QFileDialog.open
            self.fileDirectory = fileName[0]
            if self.fileDirectory:
                self.worker_1.start()
                self.worker_3.start()
                self.pushButton.setEnabled(False)
        except:
            pass

    def setLabelText(self, n):
        self.label_9.setText(n)

    def setLabelText2(self, n):
        self.label_6.setText(n)

    def setprogress(self, n):
        if int(n) == 0:
            self.stackedWidget.setCurrentIndex(5)

        self.progressBar.setValue(n)

    def clearLabelText(self):
        self.label_9.setText("")

    def filtreSheets(self):
        self.worker_3.status = False
        self.pushButton.setEnabled(True)
        self.goto_next_page()

    def goto_next_page(self):
        self.stackedWidget.setCurrentIndex(1)

    def goToHome(self):
        if QMessageBox.Yes == QMessageBox.question(
            self,
            " Cancel Progress",
            "Are you sure you want to cancel\nthe progress?",
            QMessageBox.No | QMessageBox.Yes,
        ):
            self.stackedWidget.setCurrentIndex(0)

    def goToHome2(self):
        self.stackedWidget.setCurrentIndex(0)

    def goToHome5(self):
        self.stackedWidget.setCurrentIndex(5)

    def beginInitial(self):
        if QMessageBox.Yes == QMessageBox.question(
            self,
            " Start Progress",
            "Are you sure you want to Start\nthe progress?",
            QMessageBox.No | QMessageBox.Yes,
        ):
            self.stackedWidget.setCurrentIndex(2)
            # disable able button
            self.pushButton_4.setEnabled(False)
            self.worker_2.start()
            self.worker_13.start()


class Worker_1(QThread):
    def run(self):
        print("start worker 1")

        wb_url = mainwindow.fileDirectory.split("/")
        mainwindow.wbName = wb_url[-1]

        mainwindow.analyics["workbookName"] = mainwindow.wbName

        mainwindow.xl = win32com.client.Dispatch("Excel.Application")

        try:
            mainwindow.xl.Visible = True
        except:
            mainwindow.xl.visible = True

        mainwindow.xl.DisplayAlerts = False
        mainwindow.xl.Workbooks.Open(Filename=mainwindow.fileDirectory, ReadOnly=1)

        mainwindow.wb = mainwindow.xl.Workbooks(mainwindow.wbName)

        self.sheets = [sheet.Name for sheet in mainwindow.wb.Sheets]

        print(self.sheets)

        mainwindow.analyics["sheetsNams"] = self.sheets

        exportsSheets = []
        diaESheets = []
        diaPSheets = []

        for i in self.sheets:
            if "diae" in i.lower():
                diaESheets.append(i)
            elif "diap" in i.lower():
                diaPSheets.append(i)
            elif not "execute" in i.lower():
                exportsSheets.append(i)

        mainwindow.analyics["exportsSheets"] = exportsSheets
        mainwindow.analyics["diaESheets"] = diaESheets
        mainwindow.analyics["diaPSheets"] = diaPSheets

        mainwindow.comboBox.clear()
        mainwindow.comboBox_2.clear()
        mainwindow.comboBox_3.clear()

        mainwindow.comboBox.addItems(exportsSheets)
        mainwindow.comboBox_2.addItems(diaESheets)
        mainwindow.comboBox_3.addItems(diaPSheets)

        print("finish worker 1")


class Worker_2(QThread):
    notif = pyqtSignal(str)
    progress = pyqtSignal(int)

    def run(self):
        print("start worker 2")

        mainwindow.analyics["timeStart"] = time.asctime()

        export_index = (
            int(mainwindow.worker_1.sheets.index(mainwindow.comboBox.currentText())) + 1
        )
        diap_index = (
            int(mainwindow.worker_1.sheets.index(mainwindow.comboBox_3.currentText()))
            + 1
        )
        diae_index = (
            int(mainwindow.worker_1.sheets.index(mainwindow.comboBox_2.currentText()))
            + 1
        )

        self.progress.emit(2)

        try:
            self.notif.emit("Converting Export to csv")

            path = Path().absolute()
            path = str(path) + "\\csv_files\\Export.csv"
            mainwindow.wb.Sheets(int(export_index)).SaveAs(
                Filename=path, FileFormat="62"
            )
            print("export finish csv")

            self.progress.emit(4)
            self.notif.emit("converting DiaP to csv")

            path = Path().absolute()
            path = str(path) + "\\csv_files\\DiaP.csv"

            mainwindow.wb.Sheets(int(diap_index)).SaveAs(Filename=path, FileFormat="62")
            print("diap finish csv")

            self.progress.emit(6)
            self.notif.emit("converting DiaE to csv")

            path = Path().absolute()
            path = str(path) + "\\csv_files\\DiaE.csv"

            mainwindow.wb.Sheets(int(diae_index)).SaveAs(Filename=path, FileFormat="62")
            print("diae finish csv")

            self.progress.emit(8)

            print("finish worker 2")
        except:
            try:
                self.notif.emit("Converting Export to csv")

                path = Path().absolute()
                path = str(path) + "\\csv_files\\Export.csv"
                mainwindow.wb.Sheets(int(export_index)).SaveAs(
                    Filename=path, FileFormat="6"
                )
                print("export finish csv")

                self.progress.emit(4)

                self.notif.emit("converting DiaP to csv")

                path = Path().absolute()
                path = str(path) + "\\csv_files\\DiaP.csv"

                mainwindow.wb.Sheets(int(diap_index)).SaveAs(
                    Filename=path, FileFormat="6"
                )
                print("diap finish csv")

                self.notif.emit("converting DiaE to csv")
                self.progress.emit(6)

                path = Path().absolute()
                path = str(path) + "\\csv_files\\DiaE.csv"

                mainwindow.wb.Sheets(int(diae_index)).SaveAs(
                    Filename=path, FileFormat="6"
                )
                print("diae finish csv")

                self.progress.emit(8)

                print("finish worker 2")
            except:
                print("Failed converting to csv")
                self.notif.emit("Failed converting to csv, Error!!")
                self.progress.emit(0)
                mainwindow.stackedWidget.setCurrentIndex(5)


class Worker_13(QThread):
    seco = pyqtSignal(list)
    status = True

    def run(self):
        self.status = True
        start = time.time()

        while self.status:
            end = time.time()
            temp = end - start

            hours = temp // 3600
            temp = temp - 3600 * hours
            minutes = temp // 60
            seconds = temp - 60 * minutes
            time.sleep(0.8)

            self.seco.emit([minutes, seconds])

        print("finish the timer state")


class Worker_3(QThread):

    n = pyqtSignal(str)
    status = True

    def run(self):
        self.status = True
        while self.status:
            self.n.emit("Wait File to connect.")
            time.sleep(0.15)
            self.n.emit("Wait File to connect..")
            time.sleep(0.15)
            self.n.emit("Wait File to connect...")
            time.sleep(0.15)
            self.n.emit("Wait File to connect....")
            time.sleep(0.15)
            self.n.emit("Wait File to connect.....")
            time.sleep(0.15)


class Worker_6(QThread):
    notif = pyqtSignal(str)
    progress = pyqtSignal(int)

    def run(self):

        print("start worker 6")
        gc.collect()
        tic = time.time()

        self.progress.emit(10)
        self.notif.emit("read DiaP data")

        DiaP = pd.read_csv(
            ".\csv_files\DiaP.csv",
            low_memory=False,
            dtype=str,
            encoding="unicode_escape",
            skiprows=3,
            header=0,
        )
        print("Read DiaP done")

        self.progress.emit(12)
        self.notif.emit("read DiaE data")

        DiaE = pd.read_csv(
            ".\csv_files\DiaE.csv",
            low_memory=False,
            dtype=str,
            encoding="unicode_escape",
            skiprows=3,
            header=0,
        )
        print("Read DiaE done")

        self.progress.emit(14)
        self.notif.emit("read Export data")

        Export = pd.read_csv(
            ".\csv_files\Export.csv",
            low_memory=False,
            dtype=str,
            encoding="unicode_escape",
        )
        # print(Export, DiaE, DiaP)
        print("Read Export done")

        self.progress.emit(15)
        self.notif.emit("clean all df from data")

        # clean the dataframe :
        Export.replace(np.nan, "", inplace=True)
        DiaP.replace(np.nan, "", inplace=True)
        DiaE.replace(np.nan, "", inplace=True)
        print("clean all df done")

        try:
            self.progress.emit(17)
            self.notif.emit("initialize DiaP sheet")

            # initalize the DiaP Sheet
            print("initialze DiaP sheet")
            t1 = time.time()
            DiaP["ID"] = ""
            for row in range(len(DiaP.index)):
                DiaP["ID"][
                    row
                ] = f'{DiaP["Teil"][row]}{DiaP["Submodul"][row]}{DiaP["POS"][row]}{DiaP["Codebedingung"][row]}{DiaP["AA"][row]}'
            t2 = time.time()
            print(f"initaize DiaP done in {round(t2-t1,2)} seconds")

            mainwindow.analyics[
                "initDiapTime"
            ] = f"initaize DiaP done in {round(t2-t1,2)} seconds"

            print("done")
        except:
            print(
                "Failed Initialize DiaP, please check if the DiaP sheet have the correct header syntax"
            )
            print("The courant header of the DiaP have this columns : ")
            print(DiaP.columns)
            self.progress.emit(0)
            self.notif.emit("Failed Initialize DiaP, Erorr!")
            mainwindow.stackedWidget.setCurrentIndex(5)
        finally:
            gc.collect()

        print("ba9i")
        try:
            # initialize export sheet
            self.progress.emit(30)
            self.notif.emit("initialize Export sheet")

            print("initalize export sheet...")
            t1 = time.time()
            Export["KEM AGG"] = ""
            Export["ID"] = ""
            Export["KEM NEU VEW FW"] = ""
            Export["KEM NEU VEW FS"] = ""
            Export["KEM NEU VEW FV"] = ""
            Export["KEM UNG"] = "-"
            Export["PEM NEU VEW"] = ""
            Export["PEM UNG"] = ""
            Export["EIN"] = ""
            Export["AUS"] = ""
            Export["Comment"] = ""
            for row in range(len(Export.index)):
                if Export["KEM_kems"][row] == Export["KEM_kemv"][row]:
                    Export["KEM AGG"][row] = f'{Export["KEM_kems"][row]}'
                else:
                    Export["KEM AGG"][
                        row
                    ] = f'{Export["KEM_kems"][row]}{Export["KEM_kemv"][row]}'
                Export["ID"][
                    row
                ] = f'{Export["Sachnummer_snrb"][row]}{Export["Submodul_sollv"][row]}{Export["POS_sollv"][row]}{Export["Code_sollv"][row]}{Export["AA_sollv"][row]}'
            t2 = time.time()
            print(f"initaize Export done in {round(t2-t1,2)} seconds")

            mainwindow.analyics[
                "initaExport"
            ] = f"initaize Export done in {round(t2-t1,2)} seconds"

        except:
            print(
                "Failed to initialize the Export Sheet, please check if the Export sheet have the correct header syntax"
            )
            print("The courant header of the Export have this columns : ")
            print(Export.columns)
            self.progress.emit(0)
            mainwindow.stackedWidget.setCurrentIndex(5)
            self.notif.emit("Failed Initialize Export, Erorr!")
        finally:
            gc.collect()

        # Initialize the DiaE Sheet
        print("initalize DiaE sheet ...")
        self.progress.emit(70)
        self.notif.emit("initialize DiaE sheet")
        try:
            t1 = time.time()
            DiaE["ID FW"] = ""
            DiaE["ID FS"] = ""
            DiaE["ID FV"] = ""
            self.progress.emit(73)
            for row in range(len(DiaE.index)):
                DiaE["ID FW"][
                    row
                ] = f'{DiaE["Teil"][row]}{DiaE["Submodul"][row]}{DiaE["POS"][row]}{DiaE["Codebedingung"][row]}{DiaE["FW MG"][row]}'
                DiaE["ID FS"][
                    row
                ] = f'{DiaE["Teil"][row]}{DiaE["Submodul"][row]}{DiaE["POS"][row]}{DiaE["Codebedingung"][row]}{DiaE["FS MG"][row]}'
                DiaE["ID FV"][
                    row
                ] = f'{DiaE["Teil"][row]}{DiaE["Submodul"][row]}{DiaE["POS"][row]}{DiaE["Codebedingung"][row]}{DiaE["FV MG"][row]}'
            t2 = time.time()
            print(f"initaize DiaE done in {round(t2-t1,2)} seconds")

            mainwindow.analyics[
                "initaDiaE"
            ] = f"initaize DiaE done in {round(t2-t1,2)} seconds"

        except:
            print(
                "Failed to initialize the DiaE Sheet, please check if the Export sheet have the correct header syntax"
            )
            print("The courant header of the DiaE have this columns : ")
            print(DiaE.columns)
            self.progress.emit(0)
            self.notif.emit("Failed Initialize DiaE, Erorr!")
            mainwindow.stackedWidget.setCurrentIndex(5)
        finally:
            gc.collect()

        #  thread
        self.progress.emit(75)
        self.notif.emit("start Calculation Match part")

        print("Start Calculation Match part...")
        try:
            t1 = time.time()
            # Calculation part for Export, DiaP
            temp = pd.DataFrame()
            temp["ID"] = DiaP["ID"]
            temp["PEM ab"] = DiaP["PEM ab"]
            temp["PEM bis"] = DiaP["PEM bis"]
            temp["Termin ab"] = DiaP["Termin ab"]
            temp["Termin bis"] = DiaP["Termin bis"]
            temp.drop_duplicates(
                subset="ID", keep="last", ignore_index=True, inplace=True
            )
            Export = Export.merge(temp, on="ID", how="left")
            Export.replace(np.nan, "", inplace=True)
            Export["PEM NEU VEW"] = Export["PEM ab"]
            Export["PEM UNG"] = Export["PEM bis"]
            Export["EIN"] = Export["Termin ab"]
            Export["AUS"] = Export["Termin bis"]
            del Export["PEM ab"]
            del Export["PEM bis"]
            del Export["Termin bis"]
            del Export["Termin ab"]
            del temp
            # Calculation part for Export,DiaE
            # for fw
            temp = pd.DataFrame()
            temp["ID"] = DiaE["ID FW"]
            temp["KEM ab"] = DiaE["KEM ab"]
            temp.drop_duplicates(
                subset="ID", keep="first", ignore_index=True, inplace=True
            )
            Export = Export.merge(temp, on="ID", how="left")
            Export.replace(np.nan, "", inplace=True)
            Export["KEM NEU VEW FW"] = Export["KEM ab"]
            del Export["KEM ab"]
            del temp
            temp = pd.DataFrame()
            temp["ID"] = DiaE["ID FW"]
            temp["KEM bis"] = DiaE["KEM bis"]
            temp.drop_duplicates(
                subset="ID", keep="last", ignore_index=True, inplace=True
            )
            Export = Export.merge(temp, on="ID", how="left")
            Export.replace(np.nan, "", inplace=True)
            for row in range(len(Export.index)):
                if Export["KEM UNG"][row] == "-":
                    Export["KEM UNG"][row] = Export["KEM bis"][row]
            Export["KEM UNG"].replace("", "-", inplace=True)
            del Export["KEM bis"]
            del temp
            # for fs
            temp = pd.DataFrame()
            temp["ID"] = DiaE["ID FS"]
            temp["KEM ab"] = DiaE["KEM ab"]
            temp.drop_duplicates(
                subset="ID", keep="first", ignore_index=True, inplace=True
            )
            Export = Export.merge(temp, on="ID", how="left")
            Export.replace(np.nan, "", inplace=True)
            Export["KEM NEU VEW FS"] = Export["KEM ab"]
            del Export["KEM ab"]
            del temp
            temp = pd.DataFrame()
            temp["ID"] = DiaE["ID FS"]
            temp["KEM bis"] = DiaE["KEM bis"]
            temp.drop_duplicates(
                subset="ID", keep="last", ignore_index=True, inplace=True
            )
            Export = Export.merge(temp, on="ID", how="left")
            Export.replace(np.nan, "", inplace=True)
            for row in range(len(Export.index)):
                if Export["KEM UNG"][row] == "-":
                    Export["KEM UNG"][row] = Export["KEM bis"][row]
            Export["KEM UNG"].replace("", "-", inplace=True)
            del Export["KEM bis"]
            del temp
            # for fv
            temp = pd.DataFrame()
            temp["ID"] = DiaE["ID FV"]
            temp["KEM ab"] = DiaE["KEM ab"]
            temp.drop_duplicates(
                subset="ID", keep="first", ignore_index=True, inplace=True
            )
            Export = Export.merge(temp, on="ID", how="left")
            Export.replace(np.nan, "", inplace=True)
            Export["KEM NEU VEW FV"] = Export["KEM ab"]
            del Export["KEM ab"]
            del temp
            temp = pd.DataFrame()
            temp["ID"] = DiaE["ID FV"]
            temp["KEM bis"] = DiaE["KEM bis"]
            temp.drop_duplicates(
                subset="ID", keep="last", ignore_index=True, inplace=True
            )
            Export = Export.merge(temp, on="ID", how="left")
            Export.replace(np.nan, "", inplace=True)
            for row in range(len(Export.index)):
                if Export["KEM UNG"][row] == "-":
                    Export["KEM UNG"][row] = Export["KEM bis"][row]
            Export["KEM UNG"].replace("", "-", inplace=True)
            del Export["KEM bis"]
            del temp
            Export["KEM UNG"].replace("-", "", inplace=True)
            t2 = time.time()
            print(f"Calculation match done in {round(t2-t1,2)} seconds.")

            mainwindow.analyics[
                "calculationTime"
            ] = f"Calculation match done in {round(t2-t1,2)} seconds."

            self.progress.emit(78)
            self.notif.emit("start Calculation Match part")

        except:
            print("Failed to Calculate, Retry later.")
            self.notif.emit("Failed to Calculate, Retry later.")
            self.progress.emit(0)
            mainwindow.stackedWidget.setCurrentIndex(5)
        finally:
            gc.collect()

        # final thread
        self.progress.emit(85)
        self.notif.emit("start create output")

        print("Start Create output...")
        try:
            t1 = time.time()
            # output as csv files in out directory
            Export.to_excel("./out/Export.xlsx", index="False")
            self.progress.emit(90)
            DiaP.to_excel("./out/DiaP.xlsx", index="False")
            self.progress.emit(93)
            DiaE.to_excel("./out/DiaE.xlsx", index="False")
            self.progress.emit(96)
            t2 = time.time()
        except: 
            try:
                t1 = time.time()
                # output as csv files in out directory
                Export.to_csv("./out/Export.csv", encoding="utf-8", index="False")
                self.progress.emit(90)
                DiaP.to_csv("./out/DiaP.csv", encoding="utf-8", index="False")
                self.progress.emit(93)
                DiaE.to_csv("./out/DiaE.csv", encoding="utf-8", index="False")
                self.progress.emit(96)
                t2 = time.time()
            except:
                print(
                    "Failed to create output, please check if folder out exist in the same directory, if not create it"
                )
                self.notif.emit("Failed to create output, Erorr")
                self.progress.emit(0)
                mainwindow.stackedWidget.setCurrentIndex(5)
            else:
                print("output created as csv in out directory")
                print(f"done all in {round(t2-t1,2)} seconds")

                mainwindow.analyics[
                    "creatOutTime"
                ] = f"creat output csv in {round(t2-t1,2)} seconds"
        else:
            print("output created as Excel in out directory")
            print(f"done all in {round(t2-t1,2)} seconds")

            mainwindow.analyics[
                "creatOutTime"
            ] = f"creat output excel in {round(t2-t1,2)} seconds"
            
        toc = time.time()
        print(f"time to finish all {toc-tic} seconds")

        mainwindow.analyics["timeToFinish"] = f"time to finish all {toc-tic} seconds"

        print("finish worker 6")
        self.progress.emit(100)
        self.notif.emit("finish progress whit success")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainwindow = MainWindow()
    mainwindow.setFixedHeight(413)
    mainwindow.setFixedWidth(399)
    mainwindow.show()
    try:
        sys.exit(app.exec_())
    except:
        print("Exiting")
