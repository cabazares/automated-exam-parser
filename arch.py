#!/usr/bin/python
import sys
import os
from datetime import datetime
from random import randint
import sqlite3
import logging

from PyQt4 import QtGui, QtCore

import examparser
import xls


def getMainWindow():
    """
    Retrieve main window
    """
    widgets = QtGui.QApplication.topLevelWidgets()
    try:
        return [x for x in widgets if isinstance(x, QtGui.QMainWindow)][0]
    except IndexError:
        return None


class ArchMainWindow(QtGui.QMainWindow):
    """
    Main window
    """
    frames = []

    def __init__(self):
        super(ArchMainWindow, self).__init__()
        self.initUI()

    def center(self):
        """
        Center window on screen
        """
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def initUI(self):
        """
        Initialize UI window
        """
        # exit action
        exitAction = QtGui.QAction(QtGui.QIcon('exit.png'), '&Exit', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('Exit application')
        exitAction.triggered.connect(QtGui.qApp.quit)
        
        aboutAction = QtGui.QAction(QtGui.QIcon('about.png'), '&About', self)
        aboutAction.setStatusTip('About Tatlong Aso')
        aboutAction.triggered.connect(self.showAbout)
        
        # init menu bar
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('&File')
        fileMenu.addAction(exitAction)
        aboutMenu = menubar.addMenu('&Help')
        aboutMenu.addAction(aboutAction)
        
        # init widget stack
        self.stackedWidget = QtGui.QStackedWidget()
        self.setCentralWidget(self.stackedWidget)
        
        # init main menu
        self.mainMenu = MainMenu()
        self.stackedWidget.addWidget(self.mainMenu)
        self.frames.append(self.mainMenu)
        
        # init students list
        self.studentsList = StudentsList()
        self.stackedWidget.addWidget(self.studentsList)
        self.frames.append(self.studentsList)
        
        # init exams manager
        self.examsManager = ExamsManager()
        self.stackedWidget.addWidget(self.examsManager)
        self.frames.append(self.examsManager)
        
        # init window
        self.setGeometry(500, 500, 500, 400)
        self.setWindowTitle('Automated Exam Evaluation')
        self.setWindowIcon(QtGui.QIcon('./resources/icon.png'))
        
        # init statusbar
        self.statusBar()
        
        # center window
        self.center()
        # show window
        self.show()


    def closeEvent(self, event):
        reply = QtGui.QMessageBox.question(self, 'Confirm Exit',
            "Are you sure to quit?", QtGui.QMessageBox.Yes | 
            QtGui.QMessageBox.No, QtGui.QMessageBox.No)

        if reply == QtGui.QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


    def showAbout(self, event):
        self.dialog = AboutWindow()
        self.dialog.setGeometry(QtCore.QRect(100, 100, 150, 80))
        self.dialog.center()
        self.dialog.show()

    def switchFrame(self, frame):
        self.stackedWidget.setCurrentIndex(self.frames.index(frame))

    def showStudentsList(self, event):
        self.switchFrame(self.studentsList)
        self.studentsList.populateStudentsTable()

    
    def showExamsManager(self, event):
        self.switchFrame(self.examsManager)

    def showMainMenu(self, event):
        self.switchFrame(self.mainMenu)


 
class AboutWindow(QtGui.QDialog):
    def __init__(self):
        super(AboutWindow, self).__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(500, 500, 250, 250)
        self.setWindowTitle('Tatlong Aso')
        pixmap = QtGui.QPixmap('./resources/tatlongaso.png')
        imgLabel = QtGui.QLabel('asd')
        imgLabel.setPixmap(pixmap)
        textLabel = QtGui.QLabel('Tatlong Aso\nhttp:\\\\www.tatlongaso.com')
        textLabel.setStyleSheet("qproperty-alignment: AlignCenter;")
        
        
        hbox = QtGui.QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(imgLabel)
        hbox.addStretch(1)
        
        hbox2 = QtGui.QHBoxLayout()
        hbox2.addStretch(1)
        hbox2.addWidget(textLabel)
        hbox2.addStretch(1)
        
        vbox = QtGui.QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(hbox)
        vbox.addLayout(hbox2)
        vbox.addStretch(1)
        
        
        self.setLayout(vbox)
        
    def center(self):
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())



class MainMenu(QtGui.QFrame):

    def __init__(self):
        super(MainMenu, self).__init__()
        self.initUI()
        
    def initUI(self):
        # init buttons
        studentsButton = QtGui.QPushButton("Manage Students")
        icon = QtGui.QPixmap("./resources/students.png");
        studentsButton.setIconSize(QtCore.QSize(64, 64))
        studentsButton.setIcon(QtGui.QIcon(icon))
        studentsButton.clicked.connect(getMainWindow().showStudentsList)
        studentsButton.setStatusTip('Add, edit, update list of students')
        examsButton = QtGui.QPushButton("Manage Exams")
        icon = QtGui.QPixmap("./resources/exams.png");
        examsButton.setIconSize(QtCore.QSize(64, 64))
        examsButton.setIcon(QtGui.QIcon(icon))
        examsButton.setStatusTip('Check papers and generate reports')
        examsButton.clicked.connect(getMainWindow().showExamsManager)

        vbox = QtGui.QVBoxLayout()
        vbox.addStretch(1)
        vbox.addWidget(studentsButton)
        vbox.addWidget(examsButton)
        vbox.addStretch(1)
        
        self.setLayout(vbox)


class QListExamItem(QtGui.QListWidgetItem):
    customData = None
    
    def setCustomData(self, customData):
        self.customData = customData
        
    def getCustomData(self):
        return self.customData


class ExamDialog(QtGui.QDialog):
    semesterChoices = ["1", "2", "SUMMER"]

    def __init__(self, parent = None):
        super(ExamDialog, self).__init__(parent)
        
        self.initUI()
        
    def initUI(self):
        
        name = QtGui.QLabel('Name')
        date = QtGui.QLabel('Date')
        semester = QtGui.QLabel('Semester')

        nameEdit = QtGui.QLineEdit()
        self.nameEdit = nameEdit
        dateEdit = QtGui.QDateTimeEdit(self)
        dateEdit.setCalendarPopup(True)
        dateEdit.setDateTime(QtCore.QDateTime.currentDateTime())
        self.dateEdit = dateEdit
        semesterEdit = QtGui.QComboBox()
        self.semesterEdit = semesterEdit
        for item in ExamDialog.semesterChoices:
            semesterEdit.addItem(item)

        grid = QtGui.QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(name, 1, 0)
        grid.addWidget(nameEdit, 1, 1)

        grid.addWidget(date, 2, 0)
        grid.addWidget(dateEdit, 2, 1)

        grid.addWidget(semester, 3, 0)
        grid.addWidget(semesterEdit, 3, 1)

        self.addButton = QtGui.QPushButton("Add")
        self.addButton.clicked.connect(self.accept)
        cancelButton = QtGui.QPushButton("Cancel")
        cancelButton.clicked.connect(self.reject)
        
        hbox = QtGui.QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(self.addButton)
        hbox.addWidget(cancelButton)
        
        vbox = QtGui.QVBoxLayout()
        vbox.addLayout(grid)
        vbox.addLayout(hbox)

        self.setLayout(vbox)
        
        self.setGeometry(300, 300, 200, 200)
        self.setWindowTitle('Exam')

    def center(self):
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
        
    def getData(self):
        # get form data
        name = self.nameEdit.text()
        dtime = self.dateEdit.dateTime().toPyDateTime()
        semester = self.semesterEdit.currentText()
        return (str(name), dtime, str(semester))
            
    # static method to create the dialog and return data
    @staticmethod
    def getExamData(parent=None):
        dialog = ExamDialog(parent)
        dialog.center()
        dialog.addButton.setText("Add")
        result = dialog.exec_()
        return result == QtGui.QDialog.Accepted, dialog.getData()


    @staticmethod
    def updateExamData(name, date, semester, parent=None):
        dialog = ExamDialog(parent)
        dialog.center()
        dialog.addButton.setText("Update")
        # set data
        index = dialog.semesterChoices.index(semester)
        dialog.nameEdit.setText(name)
        dialog.dateEdit.setDateTime(date)
        dialog.semesterEdit.setCurrentIndex(index)
        result = dialog.exec_()
        return result == QtGui.QDialog.Accepted, dialog.getData()


class ExamsManager(QtGui.QFrame):
    selectedExam = None

    def __init__(self):
        super(ExamsManager, self).__init__()
        self.initUI()

    def initUI(self):
        self.db = Database()
        
        self.list = QtGui.QListWidget(self)
        self.list.setFixedWidth(150)
        self.list.itemSelectionChanged.connect(self.selectExam)
        
        exams = self.db.getExams()
        for exam in exams:
            self.addExamToList(**exam)
        
        # labels for selected 
        nameLabel = QtGui.QLabel('Name:')
        dateLabel = QtGui.QLabel('Date:')
        semesterLabel = QtGui.QLabel('Semester:')
        
        self.nameField = QtGui.QLabel('')
        self.dateField = QtGui.QLabel('')
        self.semesterField = QtGui.QLabel('')
        
        self.studCountField = QtGui.QLabel('# of Students checked:')
        
        grid = QtGui.QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(nameLabel, 1, 0)
        grid.addWidget(self.nameField, 1, 1)

        grid.addWidget(dateLabel, 2, 0)
        grid.addWidget(self.dateField, 2, 1)

        grid.addWidget(semesterLabel, 3, 0)
        grid.addWidget(self.semesterField, 3, 1)
        
        # controls
        answersButton = QtGui.QPushButton("Set Answers")
        answersButton.setStatusTip('Edit answer key for exam')
        answersButton.clicked.connect(self.setAnswers)
        checkButton = QtGui.QPushButton("Read Exam Papers")
        checkButton.clicked.connect(self.getImageFiles)
        checkButton.setStatusTip('Check given exam images')
        reportButton = QtGui.QPushButton("Generate Spreadsheet")
        reportButton.setStatusTip('Generate spreadsheet of all data')
        reportButton.clicked.connect(self.generateSpreadsheet)
        
        self.checkImagesTask = ImageReadThread()
        self.checkImagesTask.taskFinished.connect(self.onImageReadFinished)
        
        controls = QtGui.QVBoxLayout()
        controls.addLayout(grid)
        controls.addWidget(answersButton)
        controls.addWidget(checkButton)
        controls.addWidget(self.studCountField)
        controls.addStretch(1)
        controls.addWidget(reportButton)
        controls.addStretch(1)
        
        mainLayout = QtGui.QHBoxLayout()
        mainLayout.addWidget(self.list)
        mainLayout.addLayout(controls)
        mainLayout.addStretch(1)
        
        addButton = QtGui.QPushButton("Add")
        addButton.setStatusTip('Add new exam')
        addButton.clicked.connect(self.addExam)
        self.addButton = addButton
        editButton = QtGui.QPushButton("Edit")
        editButton.setStatusTip('Update exam information')
        editButton.clicked.connect(self.updateExam)
        self.editButton = editButton
        delButton = QtGui.QPushButton("Delete")
        delButton.setStatusTip('Remove selected exam')
        delButton.clicked.connect(self.removeExam)
        self.delButton = delButton
        self.disableControls()
        
        controls = QtGui.QHBoxLayout()
        controls.addWidget(addButton)
        controls.addWidget(editButton)
        controls.addWidget(delButton)
        controls.addStretch(1)
        
        backButton = QtGui.QPushButton("Go Back")
        backButton.setStatusTip('Back to main menu')
        backButton.clicked.connect(getMainWindow().showMainMenu)
        
        hbox = QtGui.QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(backButton)
        
        vbox = QtGui.QVBoxLayout()
        vbox.addLayout(mainLayout)
        vbox.addLayout(controls)
        vbox.addLayout(hbox)
        
        self.setLayout(vbox)

        # select first item
        first_item = self.list.item(0)
        if first_item:
            self.list.setItemSelected(first_item, True)

    def isBusy(self):
        return self.checkImagesTask.isRunning()
    
    def readImageSuccess(self, text=""):
        msg = "%s\n\n%s" % (text, "Generate Spreadsheet file now?")
        reply = QtGui.QMessageBox.question(self, 'Success', msg,
            QtGui.QMessageBox.Yes | QtGui.QMessageBox.No,
            QtGui.QMessageBox.Yes)

        if reply == QtGui.QMessageBox.Yes:
            self.generateSpreadsheet()

    def enableControls(self):
        self.toggleControls()

    def disableControls(self):
        self.toggleControls(False)

    def toggleControls(self, enable=True):
        #self.addButton.setEnabled(enable)
        self.editButton.setEnabled(enable)
        self.delButton.setEnabled(enable)

    def getSelectedItem(self):
        try:
            item = self.list.selectedItems()[0]
        except IndexError:
            item = None
        return item

    def addExamToList(self, exam_id="", name="", date="", semester=""):
        item = QListExamItem(self.list)
        item.setText(name)  
        item.setCustomData({"exam_id": exam_id,
                            "name": name,
                            "date": date,
                            "semester": semester})
        return item

    def selectExam(self):
        if self.isBusy():
            if self.selectedExam:
                self.list.blockSignals(True)
                self.list.setItemSelected(self.selectedExam, True)
                self.list.blockSignals(False)
            self.showBusyLoading()
            return
        item = self.getSelectedItem()
        if not item:
            return
        self.selectedExam = item
        data = item.getCustomData()
        self.nameField.setText(data["name"])
        self.dateField.setText(data["date"][:19])
        self.semesterField.setText(data["semester"])
        studCount = self.db.countExamStudents(data["exam_id"])
        self.studCountField.setText("# of Students checked: %s" % studCount)
        
        self.enableControls()

    def showBusyLoading(self, msg=None):
        if msg is None:
            msg = ("Please wait while the program is busy reading images.\n")
        QtGui.QMessageBox.information(self, 'Busy', msg)

    def addExam(self):
        if self.isBusy():
            self.showBusyLoading()
            return
        accept, examData = ExamDialog.getExamData(self)
        if accept:
            # insert to exams
            name, date, semester = examData
            eid = self.db.insertExam(name, date, semester)
            item = self.addExamToList(eid, name,
                                      date.strftime("%Y-%m-%d %H:%M:%S"),
                                      semester)
            # select Exam
            self.list.setItemSelected(item, True)

    def updateExam(self):
        if self.isBusy():
            self.showBusyLoading()
            return
        item = self.getSelectedItem()
        if not item:
            return
        data = item.getCustomData()
        date = datetime.strptime(data["date"][:19], '%Y-%m-%d %H:%M:%S')
        accept, examData = ExamDialog.updateExamData(data["name"], date,
                                                     data["semester"], self)
        if accept:
            # insert to exams
            name, date, semester = examData
            eid = data["exam_id"]
            self.db.updateExam(eid, name, date, semester)
            self.db.commit()
            date = date.strftime("%Y-%m-%d %H:%M:%S")
            item.setCustomData({"exam_id": eid,
                                "name": name,
                                "date": date,
                                "semester": semester})
            item.setText(name)


    def removeExam(self, event):
        if self.isBusy():
            self.showBusyLoading()
            return
        item = self.getSelectedItem()
        if not item:
            return
        else:
            reply = QtGui.QMessageBox.question(self, 'Confirm Delete',
                "Are you sure you want to delete the exam.\n"
                "This can not be undone?",
                QtGui.QMessageBox.Yes | QtGui.QMessageBox.No,
                QtGui.QMessageBox.Yes)

            if reply == QtGui.QMessageBox.Yes:
                # delete selected exam
                eid = item.getCustomData()["exam_id"]
                self.db.deleteExams(eid)
                self.db.commit()
                # remove from list
                self.list.takeItem(self.list.row(item))

    def setAnswers(self):
        if self.isBusy():
            self.showBusyLoading()
            return
        item = self.getSelectedItem()
        if not item:
            return
        eid = item.getCustomData()["exam_id"]
        AnswersDialog.setAnswersData(eid, self)

    def onImageReadProgress(self, i):
        self.showProgressBar(i)

    def getImageFiles(self):
        if self.isBusy():
            self.showBusyLoading()
            return
        item = self.getSelectedItem()
        if not item:
            return
        examId = item.getCustomData()["exam_id"]
    
        files = QtGui.QFileDialog.getOpenFileNames(self, 'Open Image files')
        
        # setup thread
        self.checkImagesTask.init(examId, files)
        
        dialog = ProgressDialog()
        dialog.setProgressBarMax(len(files))
        dialog.setTask(self.checkImagesTask)
        dialog.run()

        
    def onImageReadFinished(self, examId, files, data):
            
        student_ids = {}
        for student_num, stud_data in data.iteritems():
            parts = stud_data["parts"]
            # retrieve student id
            try:
                student_id = student_ids[student_num]
            except KeyError:
                student_id = self.db.getStudent(student_num)
                # if not exists insert
                if student_id is None:
                    student_id = self.db.insertStudents(student_num,
                                                        str(student_num),
                                                        "", "")
                else:
                    student_id = student_id["student_id"]
                student_ids[student_num] = student_id
            
            for part, items in parts.iteritems():
                for item, ans in items.iteritems():
                    if len(ans) > 0 and ans.upper() in "ABCDE":
                        rpart = part + 1
                        self.db.insertAnswer(examId, rpart, item, ans,
                                             student_id, commit=False)
        self.db.commit()
        
        # update student count
        studCount = self.db.countExamStudents(examId)
        self.studCountField.setText("# of Students checked: %s" % studCount)
        
        self.readImageSuccess("Successfully Read %s image(s)" % len(files))


    def computeRawScores(self, examId):
        # get correct answers
        answerKey = self.db.getCorrectExamAnswers(examId)
        students = self.db.getExamStudentAnswers(examId)
        
        scores = {}
        for stud_num, student in students.iteritems():
            scores[stud_num] = dict([(x, 0) for x in range(1, 10)])
            for part_num, part in student.iteritems():
                for item, ans in part.iteritems():
                    try:
                        correct = answerKey[part_num][item]
                    except KeyError:
                        # no correct answer given
                        correct = ""
                    if len(ans) > 0 and correct.upper() == ans.upper():
                        scores[stud_num][part_num] += 1
        # get students in exam
        raw = []
        db_studs = dict([(x["student_id"], x) for x in self.db.getStudents()])
        
        for stud_id, score in scores.iteritems():
            sdata = db_studs[stud_id]
            data = {"name": sdata["name"],
                    "bsa_code": sdata["bsa_code"],
                    "student_num": sdata["student_num"]}
            total = 0
            # format part scores for generation
            for prt, scre in score.iteritems():
                data["part" + str(prt)] = scre
                total += scre
            data["total"] = total
            raw.append(data)
        return raw


    def getNoTakeStudents(self, examId):
        # add no take students
        data = self.db.getNoTakeStudents(examId)
        result = []
        for row in data:
            result.append({"name": row[0],
                           "bsa_code": row[2],
                           "student_num": row[1]})
        return result

    def getTopPerSubject(self, raw_scores, num=3):
        parts = {}
        for i in range(1, 10):
            parts[i] = sorted(raw_scores, key=lambda k: k['part' + str(i)],
                              reverse=True)

        for i, part in parts.iteritems():
            top = []
            for row in part:
                score = row["part" + str(i)]
                if (len(top) < num or
                    (len(top) and top[-1]["part" + str(i)] == score)):
                    top.append(row)
            parts[i] = top
        return parts

    def getTopStudents(self, raw_scores, num=10, num2=30):
        ranked = []
        ltotal = sorted(raw_scores, key=lambda k: k["total"], reverse=True)
        for t in ltotal:
            score = t["total"]
            if (len(ranked) < num or
                (len(ranked) and ranked[-1]["total"] == score)):
                ranked.append(t)
        return ranked, ltotal


    def generateSpreadsheet(self):
        if self.isBusy():
            self.showBusyLoading()
            return
        item = self.getSelectedItem()
        if not item:
            return
        examData = item.getCustomData()
        date = datetime.strptime(examData["date"][:19], '%Y-%m-%d %H:%M:%S')
        examId = examData["exam_id"]
        name = "%s_%s" % (examData["name"], date.strftime("%Y-%m-%d"))
        
        TOP_SUBJECT_COUNT = 3
        raw_scores = self.computeRawScores(examId)
        
        if len(raw_scores) == 0:
            QtGui.QMessageBox.warning(self, 'No Exams Found!',
                "No test papers read. Please make sure exam papers have\n"
                "been read using the 'Read Exam Papers' button above.")
            return
        
        top, others = self.getTopStudents(raw_scores)
        
        # compute success (top 30)
        top30 = []
        for t in others:
            score = t["total"]
            if (len(top30) < 30 or
                (len(top30) and top30[-1]["total"] == score)):
                top30.append(t)
        
        # compute scores and ranking
        data = {"raw_scores": raw_scores,
                "no_takes": self.getNoTakeStudents(examId),
                "subject_top": self.getTopPerSubject(raw_scores),
                "top_students": top,
                "success_students": top30,
                "students": others}
        
        fname = QtGui.QFileDialog.getSaveFileName(self, "Save Spreadsheet",
                                                  '%s.xls' % name, "*.xls")
        if not fname:
            return
        try:
            xls.generate('resources/rank.xls', fname, data)
            QtGui.QMessageBox.information(self, 'Spreadsheet Generated',
                "Spreadsheet saved to:\n %s" % fname)
        except IOError as err:
            msg = "Please make sure the the excel file is not open.\n\n%s" % \
                  str(err)
            QtGui.QMessageBox.warning(self, 'Error', msg)

            
class ImageReadThread(QtCore.QThread):
    notifyProgress = QtCore.pyqtSignal(int)
    taskFinished = QtCore.pyqtSignal(int, list, dict)
    
    def init(self, examId, files):
        self.examId = examId
        self.files = files
    
    def run(self):
        files = self.files
        data = None
        cnt = 0
        for tdata in examparser.check_images_generator(map(str, files)):
            data = tdata
            cnt += 1
            self.notifyProgress.emit(cnt)
        
        self.taskFinished.emit(self.examId, list(files), data)


class ProgressDialog(QtGui.QDialog):
    def __init__(self, parent = None):
        super(ProgressDialog, self).__init__(parent)
        
        self.task = None
        self.progressbar = QtGui.QProgressBar()
        self.progressbar.setMinimum(0)
        self.progressbar.setMaximum(0)
        
        vbox = QtGui.QVBoxLayout()
        vbox.addWidget(self.progressbar)

        self.setLayout(vbox)
        self.setFixedWidth(300)
        self.setFixedHeight(40)
        self.setWindowTitle('Processing image files... Please Wait')
    
    def run(self):
        self.task.notifyProgress.connect(self.setProgressBarVal)
        self.task.taskFinished.connect(self.done)
        self.task.start()
        self.exec_()
    
    def setTask(self, task):
        self.task = task
    
    def setProgressBarVal(self, val=0):
        self.progressbar.setValue(val)
        
    def setProgressBarMax(self, max):
        self.progressbar.setMaximum(max)

    def closeEvent(self, evnt):
        evnt.ignore()


class AnswersDialog(QtGui.QDialog):
    def __init__(self, parent = None):
        super(AnswersDialog, self).__init__(parent)
        self.initUI()
        
    def initUI(self):
        self.db = Database()
        self.table = AnswersTable({}, 120, 9)
        self.table.itemChanged.connect(self.table.rowChanged)
                
        self.addButton = QtGui.QPushButton("Save")
        self.addButton.clicked.connect(self.accept)
        
        hbox = QtGui.QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(self.addButton)
        
        vbox = QtGui.QVBoxLayout()
        vbox.addWidget(self.table)
        vbox.addLayout(hbox)

        self.setLayout(vbox)
        
        self.setGeometry(300, 300, 600, 400)
        self.setWindowTitle('Exam')


    def center(self):
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())
    
    def setData(self, examId):
        data = self.db.getCorrectExamAnswers(examId)
        self.table.setData(data);
    
    def getData(self):
        # get form data
        return self.table.getData()
        
    # static method to create the dialog and return data
    @staticmethod
    def setAnswersData(examId, parent=None):
        dialog = AnswersDialog(parent)
        dialog.setData(examId)
        result = dialog.exec_()
        data = dialog.getData()
        dialog.db.saveExamAnswers(examId, data)
        dialog.db.commit()
        return result == QtGui.QDialog.Accepted, data



class AnswersTable(QtGui.QTableWidget):
    def __init__(self, struct, *args):
        QtGui.QTableWidget.__init__(self, *args)
        # set header labels
        self.setHorizontalHeaderLabels(['Part %s' % i for i in range(1, 10)])
        # set multi row select behavior
        self.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
        # strech last column
        header = self.horizontalHeader()
        header.setStretchLastSection(True)

        self.setData(struct)
        self.deleted = []
        # set as dirty
        self.dirty = False
    
    def isDirty(self):
        return self.dirty
        
    def setDirty(self, dirty=True):
        self.dirty = dirty
        
    def setData(self, data):
        self.data = data
        for part in data.keys():
            for item in data[part].keys():
                val = data[part][item]
                val = "" if val is None else val.upper()
                nitem = QtGui.QTableWidgetItem(str(val))
                self.setItem(item - 1, part - 1, nitem)

    def rowChanged(self, item):
        self.data[item.column() + 1][item.row() + 1] = str(item.text())
        self.setDirty()
        
    def getData(self):
        return self.data


class StudentsList(QtGui.QFrame):
    def __init__(self):
        super(StudentsList, self).__init__()
        self.initUI()
        self.db = Database()

    def initUI(self):
        students = []
        studentsTable = StudentsTable(students, len(students), 5)
        self.studentsTable = studentsTable
        studentsTable.itemChanged.connect(studentsTable.rowChanged)
        studentsTable.itemPressed.connect(self.enableButtons)

        addButton = QtGui.QPushButton("Add")
        addButton.setStatusTip('Add new row for student')
        addButton.clicked.connect(studentsTable.addRow)
        self.addButton = addButton
        delButton = QtGui.QPushButton("Delete")
        delButton.clicked.connect(studentsTable.deleteSelectedRows)
        delButton.setStatusTip('Remove selected student rows')
        self.delButton = delButton
        self.disableButtons()

        controls = QtGui.QHBoxLayout()
        controls.addWidget(addButton)
        controls.addWidget(delButton)
        controls.addStretch(1)
        
        okButton = QtGui.QPushButton("Save")
        okButton.setStatusTip('Save students data')
        okButton.clicked.connect(self.saveStudents)
        cancelButton = QtGui.QPushButton("Go Back")
        cancelButton.setStatusTip('Back to main menu')
        cancelButton.clicked.connect(self.backToMenu)
        
        hbox = QtGui.QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(cancelButton)
        hbox.addWidget(okButton)
        
        vbox = QtGui.QVBoxLayout()
        vbox.addWidget(studentsTable)
        vbox.addLayout(controls)
        vbox.addLayout(hbox)
        
        self.setLayout(vbox)

    def readDataFromDatabase(self):
        return [[x["student_id"], x["student_num"], x["name"], x["bsa_code"],
                    x["year_level"]] for x in self.db.getStudents()]

    def backToMenu(self, event):
        if self.studentsTable.isDirty():
            reply = QtGui.QMessageBox.question(self, 'Confirm action',
                "You have unsaved changes. Save data first?",
                QtGui.QMessageBox.Yes | QtGui.QMessageBox.No,
                QtGui.QMessageBox.Yes)

            if reply == QtGui.QMessageBox.Yes:
                self.saveStudents()
            else:
                self.populateStudentsTable()
                self.studentsTable.setDirty(False)
        getMainWindow().showMainMenu(self)

    def populateStudentsTable(self):
        students = self.readDataFromDatabase()
        self.studentsTable.setData(students)
        self.studentsTable.setDirty(False)

    def saveStudents(self):
        data = self.studentsTable.getData()
        error = False
        # save items
        for row in data:
            cellIndex, sid, stud_num, name, bsa, year = row
            try:
                if isinstance(sid, int) or len(sid):
                    self.db.updateStudents(sid, stud_num, name, bsa, year)
                elif len(stud_num) or len(name) or len(bsa) or len(year):
                    sid = self.db.insertStudents(stud_num, name, bsa, year)
                    nitem = QtGui.QTableWidgetItem(str(sid))
                    self.studentsTable.setItem(cellIndex, 0, nitem)
            except self.db.DuplicateError as err:
                QtGui.QMessageBox.warning(self, "Error", str(err))
                error = True

        if not error:
            # remove deleted
            data = self.studentsTable.getDeleted()
            self.db.deleteStudents(data)
            self.db.commit()
            del data[:]
            # set table as clean
            self.studentsTable.setDirty(False)

    def disableButtons(self):
        self.toggleButtons(False)

    def enableButtons(self):
        self.toggleButtons(True)

    def toggleButtons(self, enabled=True):
        #self.addButton.setEnabled(enabled)
        self.delButton.setEnabled(enabled)


class StudentsTable(QtGui.QTableWidget):
    def __init__(self, struct, *args):
        QtGui.QTableWidget.__init__(self, *args)
        # hide id column
        self.setColumnHidden(0, True)
        # set header labels
        self.setHorizontalHeaderLabels(['id', 'Student Num', 'Name', 'BSA Code',
                                        'Year'])
        # set multi row select behavior
        self.setSelectionBehavior(QtGui.QAbstractItemView.SelectRows)
        # strech last column
        header = self.horizontalHeader()
        header.setStretchLastSection(True)

        self.setData(struct)
        self.deleted = []
        # set as dirty
        self.dirty = False
    
    def isDirty(self):
        return self.dirty
        
    def setDirty(self, dirty=True):
        self.dirty = dirty
    
    def setData(self, data):
        # disable sorting while setting rows
        self.setSortingEnabled(False)
        while self.rowCount() < len(data):
            self.insertRow(self.rowCount())
        for i, row in enumerate(data):
            for j, item in enumerate(row):
                item = "" if item is None else item
                nitem = QtGui.QTableWidgetItem(str(item))
                self.setItem(i, j, nitem)
        # enable sorting
        self.setSortingEnabled(True)

    def addRow(self):
        self.insertRow(self.rowCount())
        # scroll to bottom
        vBar = self.verticalScrollBar()
        vBar.setValue(vBar.maximum())

    def deleteSelectedRows(self):
        for row in reversed(self.selectionModel().selectedRows()):
            index = row.row()
            # add to deleted list
            sid = self.item(index, 0).text()
            if isinstance(sid, int) or len(sid):
                self.deleted.append(sid)
            # remove item from ui
            self.removeRow(index)
        self.setDirty(True)
        
    def getData(self):
        rows = []
        for i in range(self.rowCount()):
            row = [i]
            for j in range(self.columnCount()):
                cell = self.item(i, j)
                row.append(str(cell.text() if cell else ""))
            rows.append(row);
        return rows
                
    def rowChanged(self,  item):
        self.setDirty()
    
    def getDeleted(self):
        return self.deleted


class Database():
    _instance = None
    database = 'resources/data.sqlite'
    
    class DuplicateError(Exception):
        pass
    
    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(Database, cls).__new__(cls, *args, **kwargs)
        return cls._instance

    def __init__(self):
        self.connect()
    
    def connect(self):
        self.conn = sqlite3.connect(self.database,
                                    detect_types=sqlite3.PARSE_DECLTYPES)
        self.cursor = self.conn.cursor()

    def commit(self):
        self.conn.commit()

    # STUDENTS *****************
    def getStudents(self):
        self.cursor.execute("SELECT * FROM `students`")
        students = self.cursor.fetchall()
        rows = []
        for row in students:
            rows.append({"student_id": row[0],
                         "student_num": row[1],
                         "name": row[2],
                         "bsa_code": row[3],
                         "year_level": row[4]})
        return rows
        
    def getStudent(self, student_num):
        self.cursor.execute("SELECT * FROM `students` WHERE `student_num`=?",
                            (student_num,))
        row = self.cursor.fetchone()
        if row is None:
            return row
        return {"student_id": row[0],
                 "student_num": row[1],
                 "name": row[2],
                 "bsa_code": row[3],
                 "year_level": row[4]}

    def insertStudents(self, student_num, name, bsa_code, year_level):
        try:
            self.cursor.execute("INSERT INTO `students` (`student_num`, `name`, "
                                "`bsa_code`, `year_level`) "
                                "VALUES (?, ?, ?, ?)", (student_num, name,
                                                        bsa_code, year_level))
            self.conn.commit()
        except sqlite3.IntegrityError as err:
            raise self.DuplicateError("The student Number '%s' is already "
                                      "used by another student" % student_num)
        return self.cursor.lastrowid

    def updateStudents(self, sid, student_num, name, bsa_code, year_level):
        try:
            self.cursor.execute("UPDATE `students` SET `student_num`=?, "
                                "`name`=?, `bsa_code`=?, `year_level`=? "
                                "WHERE student_id=?", (student_num, name,
                                                       bsa_code, year_level,
                                                       sid))
        except sqlite3.IntegrityError as err:
            raise self.DuplicateError("The student Number '%s' is already "
                                      "used by another student" % student_num)

    def deleteStudents(self, sid):
        if isinstance(sid, list):
            quer = ", ".join([str(x) for x in sid])
            self.cursor.execute("DELETE FROM `students` "
                                "WHERE `student_id` in (%s)" % quer)
        else:
            self.cursor.execute("DELETE FROM `students` "
                                "WHERE `student_id`=?", (sid,))


    # EXAMS *****************
    def getExams(self):
        self.cursor.execute("SELECT * FROM `exams`")
        exams = self.cursor.fetchall()
        rows = []
        for row in exams:
            rows.append({"exam_id": row[0],
                         "name": row[1],
                         "date": row[2],
                         "semester": row[3]})
        return rows

    def insertExam(self, name, date, semester):
        self.cursor.execute("INSERT INTO `exams` (`name`, `date`, `semester`) "
                            "VALUES (?, ?, ?)", (name, date, semester))
        self.conn.commit()
        return self.cursor.lastrowid

    def updateExam(self, eid, name, date, semester):
        self.cursor.execute("UPDATE `exams` SET `name`=?, "
                            "`date`=?, `semester`=? WHERE exam_id=?",
                            (name, date, semester, eid))

    def deleteExams(self, eid):
        if isinstance(eid, list):
            quer = ", ".join([str(x) for x in eid])
            self.cursor.execute("DELETE FROM `exams` "
                                "WHERE `exam_id` in (%s)" % quer)
        else:
            self.cursor.execute("DELETE FROM `exams` "
                                "WHERE `exam_id`=?", (eid,))


    def countExamStudents(self, examId):
        self.cursor.execute("SELECT `answer_id` FROM `answers` "
                            "WHERE `student_id` != 0 AND `exam_id`=? "
                            "GROUP BY `student_id`", (examId,))
        return len(self.cursor.fetchall())
            

    def getNoTakeStudents(self, examId):
        sql = ("SELECT `name`, `student_num`, `bsa_code` "
               "FROM students WHERE student_id NOT IN "
               "(SELECT DISTINCT student_id "
               " FROM `answers` "
               " WHERE `exam_id`=? "
               " AND `student_id`!=0)")
        self.cursor.execute(sql, (examId,))
        return self.cursor.fetchall()

    # EXAM ANSWERS *****************
    def getCorrectExamAnswers(self, examId):
        self.cursor.execute("SELECT * FROM `answers`"
                            "WHERE `exam_id`=? AND `student_id`=0", (examId,))
        answers = self.cursor.fetchall()
        data = dict([(x, {}) for x in range(1, 10)])
        for row in answers:
            aid, eid, stud_num, part, item, answer = row
            data[part][item] = answer
        return data
    
    def saveExamAnswers(self, examId, data):
        for part in data.keys():
            for item in data[part].keys():
                answer = data[part][item]
                self.insertAnswer(examId, part, item, answer, commit=False)


    def getExamStudentAnswers(self, examId):
        self.cursor.execute("SELECT * FROM `answers` "
                            "WHERE `exam_id`=? "
                            "AND `student_id`!=0", (examId,))
        answers = self.cursor.fetchall()
        students = {}
        for row in answers:
            aid, eid, stud_num, part, item, answer = row
            if stud_num not in students.keys():
                students[stud_num] = dict([(x, {}) for x in range(1, 10)])
            students[stud_num][part][item] = answer
        return students

    # ANSWERS **********************
    def insertAnswer(self, exam_id, part, item, answer, student_id=0,
                     commit=True):
        self.cursor.execute("INSERT INTO `answers` (`exam_id`, `student_id`, "
                            "`part`, `item`, `answer`) "
                            "VALUES (?, ?, ?, ?, ?)", (exam_id, student_id,
                                                       part, item, answer))
        if commit:
            self.conn.commit()
        return self.cursor.lastrowid


class LoggerWriter:
    def __init__(self, logger, level):
        self.logger = logger
        self.level = level

    def write(self, message):
        message = message.strip("\n")
        if len(message):
            logging.log(self.level, message)

def initLogger():
    logging.basicConfig(level=logging.DEBUG)
    logger = logging.getLogger()
    sys.stdout = LoggerWriter(logger, logging.INFO)
    sys.stderr = LoggerWriter(logger, logging.DEBUG)
    hdlr = logging.FileHandler('./logs/error.log')
    formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
    hdlr.setFormatter(formatter)
    logger.addHandler(hdlr)

def main():
    app = QtGui.QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon('./resources/icon.png'))
    window = ArchMainWindow()
    sys.exit(app.exec_())

if __name__ == '__main__':
    initLogger()
    main()
