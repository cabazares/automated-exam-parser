import math
from decimal import Decimal

from xlrd import open_workbook
from xlwt import easyxf, Formula
from xlutils.copy import copy


# sheets
SHEET_PER_SUBJECT = 0
SHEET_TOP = 1
SHEET_SUCCESS = 2
SHEET_OTHERS = 3
SHEET_NO_TAKES = 4
SHEET_RAW_SCORES = 5

# strings
HEADER_TITLE = "UNIVERSITY OF THE PHILIPPINES MINDANAO"
HEADER_SUBTITLE = ("COLLEGE OF HUMANTIES AND SOCIAL SCIENCES "
                   "DEPARTMENT OF ARCHITECTURE")
SUBHEADER_TITLE = "AIEX (ARCHITECTURAL INSTRUCTIONAL EXAMINATION)"
SUBHEADER_SUBTITLE = ""

#styles
STYLE_DEFAULT = easyxf('font: name Calibri, bold true, height 240;'
                       'borders: left thin, right thin, top thin, bottom thin;')
STYLE_PERCENTAGE = easyxf('font: name Calibri, bold true, height 240;'
                          'borders: left thin, right thin, top thin, '
                          'bottom thin;',
                           num_format_str='0.00%')
STYLE_PERFECT = easyxf('font: name Calibri, bold true, height 240;'
                       'borders: left thin, right thin, top thin, '
                       'bottom thin;',
                       num_format_str='000%')
STYLE_DECIMAL = easyxf('font: name Calibri, bold true, height 240;'
                       'borders: left thin, right thin, top thin, bottom thin;',
                        num_format_str='0.00')
STYLE_HEADER = easyxf('font: name Calibri, bold true, height 240;')
STYLE_SUBHEADER = easyxf('font: name Calibri, bold true, height 240;'
                         'pattern: pattern solid, fore_colour gray25;'
                         'borders: bottom thin;')
STYLE_MAINTITLE = easyxf('font: name Calibri, bold true, '
                         'height 380, color white;'
                         'pattern: pattern solid, fore_colour black;'
                         'align: vertical center;'
                         'borders: bottom thin;')
STYLE_COLHEADER = easyxf('font: name Calibri, bold true, height 200;'
                         'pattern: pattern solid, fore_colour gray25;'
                         'align: vertical center, horizontal center, wrap true;'
                         'borders: left thin, right thin, bottom thin;')
STYLE_PERCENT = easyxf('font: name Calibri, bold true, height 240;'
                       'borders: left thin, right thin, top thin, bottom thin;')

class ExcelGenerator(object):
    def copy(self, filename):
        rb = open_workbook(filename)
        wb = copy(rb)
        return wb

    def setHeaders(self, sheet, title=""):
        sheet.row(0).write(0, HEADER_TITLE, STYLE_HEADER)
        sheet.row(1).write(0, HEADER_SUBTITLE, STYLE_HEADER)
        sheet.row(2).write(0, SUBHEADER_TITLE, STYLE_SUBHEADER)
        sheet.row(0).set_style(STYLE_HEADER)
        sheet.row(1).set_style(STYLE_HEADER)
        sheet.row(2).set_style(STYLE_SUBHEADER)
        
        sheet.row(4).write(0, "", STYLE_MAINTITLE)
        sheet.row(4).write(1, title, STYLE_MAINTITLE)
        sheet.row(4).set_style(STYLE_MAINTITLE)
        sheet.row(4).height_mismatch = 1
        sheet.row(4).height = 480
        
    def setSubheaders(self, sheet, percentage=True):
        labels = ["NO.",
                  "NAME",
                  "STUDENT NUMBER",
                  "BSA NUMBER",
                  "PART 1: ARCHT'L COMM",
                  "PART 2: HISTORY, THEORY, CRI",
                  "PART 3: BLDG SCI & CONS",
                  "PART 4: ARCHT'L STRUC",
                  "PART 5: BLDG UTILTIIES",
                  "PART 6: PRACTICE & GOVERN",
                  "PART 7: PLANNING & URBN DSN",
                  "PART 8: SPECIAL CLUSTER",
                  "PART 9: ARCHT'L DESIGN",
                  "TOTAL SCORES",
                  "NO. OF ITEMS",
                  "",
                  "",
                  "PERCENTAGE"]
    
        sheet.col(1).width =  10500
        sheet.col(2).width =  4000
        sheet.row(5).set_style(STYLE_COLHEADER)
        sheet.row(5).height_mismatch = 1
        sheet.row(5).height = 1000
        
        for col, txt in enumerate(labels):
            if percentage or (not percentage and col < 13):
                sheet.row(5).write(col, txt, STYLE_COLHEADER)
        
    def addRawScores(self, wb, data):
        sheet = wb.get_sheet(SHEET_RAW_SCORES)
        start_row = 6
        
        self.setHeaders(sheet, "ROSTER OF SCORES")
        self.setSubheaders(sheet, False)
        
        for rcnt, row in enumerate(data, start_row):
            srow = sheet.row(rcnt)
            srow.write(0, rcnt - start_row + 1, STYLE_DEFAULT)
            srow.write(1, row["name"], STYLE_DEFAULT)
            srow.write(2, row["student_num"], STYLE_DEFAULT)
            srow.write(3, row["bsa_code"], STYLE_DEFAULT)
            srow.write(4, row["part1"], STYLE_DEFAULT)
            srow.write(5, row["part2"], STYLE_DEFAULT)
            srow.write(6, row["part3"], STYLE_DEFAULT)
            srow.write(7, row["part4"], STYLE_DEFAULT)
            srow.write(8, row["part5"], STYLE_DEFAULT)
            srow.write(9, row["part6"], STYLE_DEFAULT)
            srow.write(10, row["part7"], STYLE_DEFAULT)
            srow.write(11, row["part8"], STYLE_DEFAULT)
            srow.write(12, row["part9"], STYLE_DEFAULT)
            srow.write(13, row["total"], STYLE_DEFAULT)


    def addNoTakes(self, wb, data):
        sheet = wb.get_sheet(SHEET_NO_TAKES)
        start_row = 6
        
        self.setHeaders(sheet, "ROSTER OF SCORES")
        self.setSubheaders(sheet, False)
        
        for rcnt, row in enumerate(data, start_row):
            srow = sheet.row(rcnt)
            srow.write(0, rcnt - start_row + 1, STYLE_DEFAULT)
            srow.write(1, row["name"], STYLE_DEFAULT)
            srow.write(2, row["student_num"], STYLE_DEFAULT)
            srow.write(3, row["bsa_code"], STYLE_DEFAULT)
            for cnt in range(4, 14):
                srow.write(cnt, 0, STYLE_DEFAULT)


    def addSubjectTop(self, wb, data):
        labels = ["ARCHITECTURAL COMMUNICATIONS",
                  "HISTORY, THEORY & CRITICISM",
                  "BUILDING SCIENCE & CONSTRUCTION",
                  "ARCHITECTURAL STRUCTURES",
                  "BUILDING UTILITIES",
                  "PRACTICE & GOVERNANCE",
                  "PLANNING & URBAN DESIGN",
                  "SPECIAL CLUSTER",
                  "ARCHITECTURAL DESIGN"]
    
        sheet = wb.get_sheet(SHEET_PER_SUBJECT)
        start_row = 5
        rcnt = start_row
        
        self.setHeaders(sheet, "TOP 3 PER SUBJECT AREA")
        sheet.col(1).width =  15500
        
        style = easyxf('font: name Calibri, bold true, height 240;'
                       'pattern: pattern solid, fore_colour gray25;'
                       'align: vertical center;'
                       'borders: bottom thin;')
        
        for p in range(1, 10):
            rank = 0
            ls = "PART %s: %s" % (p, labels[p - 1])
            srow = sheet.row(rcnt)
            srow.write(0, ls, style)
            srow.set_style(style)
            
            srow.height_mismatch = 1
            srow.height = 420
            
            rcnt += 1
            prev_score = None
            for student in data[p]:
                score = student["part" + str(p)]
                # ranking (handle same scores)
                if prev_score != score:
                    prev_score = score
                    rank += 1
                # write to sheet.
                srow = sheet.row(rcnt)
                srow.write(0, rank, STYLE_DEFAULT)
                srow.write(1, student["name"], STYLE_DEFAULT)
                rcnt += 1

    def addOthers(self, wb, data):
        sheet = wb.get_sheet(SHEET_OTHERS)
        start_row = 6
        # compute best TOTAL ITEMS
        base = 50.0 # round to nearest 50
        items = int(base * math.ceil(float(data[0]["total"]) / base))
        
        self.setHeaders(sheet, "AIEX ROSTER OF EXAM PERCENTAGES")
        self.setSubheaders(sheet)
        
        for rcnt, row in enumerate(data, start_row):
            formula = Formula("M%s/N%s" % (rcnt + 1, rcnt + 1))
            srow = sheet.row(rcnt)
            srow.write(0, rcnt - start_row + 1, STYLE_DEFAULT)
            srow.write(1, row["name"], STYLE_DEFAULT)
            srow.write(2, row["student_num"], STYLE_DEFAULT)
            srow.write(3, row["bsa_code"], STYLE_DEFAULT)
            srow.write(4, row["part1"], STYLE_DEFAULT)
            srow.write(5, row["part2"], STYLE_DEFAULT)
            srow.write(6, row["part3"], STYLE_DEFAULT)
            srow.write(7, row["part4"], STYLE_DEFAULT)
            srow.write(8, row["part5"], STYLE_DEFAULT)
            srow.write(9, row["part6"], STYLE_DEFAULT)
            srow.write(10, row["part7"], STYLE_DEFAULT)
            srow.write(11, row["part8"], STYLE_DEFAULT)
            srow.write(12, row["part9"], STYLE_DEFAULT)
            srow.write(13, row["total"], STYLE_DEFAULT)
            srow.write(14, items, STYLE_DEFAULT)
            srow.write(15, formula, STYLE_DECIMAL)
            srow.write(16, 1, STYLE_PERFECT)
            srow.write(17, formula, STYLE_PERCENTAGE)


    def addSuccess(self, wb, data):
        sheet = wb.get_sheet(SHEET_SUCCESS)
        start_row = 6
        # compute best TOTAL ITEMS
        base = 50.0 # round to nearest 50
        items = int(base * math.ceil(float(data[0]["total"]) / base))
        
        self.setHeaders(sheet, "AIEX ROSTER OF SUCCESSFUL EXAMINEES")
        self.setSubheaders(sheet)
        
        for rcnt, row in enumerate(data, start_row):
            formula = Formula("M%s/N%s" % (rcnt + 1, rcnt + 1))
            srow = sheet.row(rcnt)
            srow.write(0, rcnt - start_row + 1, STYLE_DEFAULT)
            srow.write(1, row["name"], STYLE_DEFAULT)
            srow.write(2, row["student_num"], STYLE_DEFAULT)
            srow.write(3, row["bsa_code"], STYLE_DEFAULT)
            srow.write(4, row["part1"], STYLE_DEFAULT)
            srow.write(5, row["part2"], STYLE_DEFAULT)
            srow.write(6, row["part3"], STYLE_DEFAULT)
            srow.write(7, row["part4"], STYLE_DEFAULT)
            srow.write(8, row["part5"], STYLE_DEFAULT)
            srow.write(9, row["part6"], STYLE_DEFAULT)
            srow.write(10, row["part7"], STYLE_DEFAULT)
            srow.write(11, row["part8"], STYLE_DEFAULT)
            srow.write(12, row["part9"], STYLE_DEFAULT)
            srow.write(13, row["total"], STYLE_DEFAULT)
            srow.write(14, items, STYLE_DEFAULT)
            srow.write(15, formula, STYLE_DECIMAL)
            srow.write(16, 1, STYLE_PERFECT)
            srow.write(17, formula, STYLE_PERCENTAGE)


    def addTopStudents(self, wb, data):
        sheet = wb.get_sheet(SHEET_TOP)
        start_row = 6
        # compute best TOTAL ITEMS
        base = 50.0 # round to nearest 50
        items = int(base * math.ceil(float(data[0]["total"]) / base))
        
        self.setHeaders(sheet, "AIEX TOP TEN PERFORMERS")
        self.setSubheaders(sheet)
        
        for rcnt, row in enumerate(data, start_row):
            formula = Formula("M%s/N%s" % (rcnt + 1, rcnt + 1))
            srow = sheet.row(rcnt)
            srow.write(0, rcnt - start_row + 1, STYLE_DEFAULT)
            srow.write(1, row["name"], STYLE_DEFAULT)
            srow.write(2, row["student_num"], STYLE_DEFAULT)
            srow.write(3, row["bsa_code"], STYLE_DEFAULT)
            srow.write(4, row["part1"], STYLE_DEFAULT)
            srow.write(5, row["part2"], STYLE_DEFAULT)
            srow.write(6, row["part3"], STYLE_DEFAULT)
            srow.write(7, row["part4"], STYLE_DEFAULT)
            srow.write(8, row["part5"], STYLE_DEFAULT)
            srow.write(9, row["part6"], STYLE_DEFAULT)
            srow.write(10, row["part7"], STYLE_DEFAULT)
            srow.write(11, row["part8"], STYLE_DEFAULT)
            srow.write(12, row["part9"], STYLE_DEFAULT)
            srow.write(13, row["total"], STYLE_DEFAULT)
            srow.write(14, items, STYLE_DEFAULT)
            srow.write(15, formula, STYLE_DECIMAL)
            srow.write(16, 1, STYLE_PERFECT)
            srow.write(17, formula, STYLE_PERCENTAGE)


    def saveExcelFile(self, wb, filename):
        wb.save(filename)



def generate(src, out, data):
    excel = ExcelGenerator()
    wb = excel.copy(src)
    
    no_takes = data["no_takes"]
    students = data["students"]
    no_takes = no_takes[:]
    for row in no_takes:
        for part in range(1,10):
            row["part" + str(part)] = 0
        row["total"] = 0
    students.extend(no_takes)
    
    excel.addSubjectTop(wb, data["subject_top"])
    excel.addTopStudents(wb, data["top_students"])
    excel.addSuccess(wb, data["success_students"])
    excel.addOthers(wb, students)
    excel.addNoTakes(wb, no_takes)
    excel.addRawScores(wb, data["raw_scores"])
    # write excel
    wb.save(out)


if __name__ == "__main__":
    data = [{
        "name": "Keno Test",
        "bsa_code": "Test",
        "part1": 1,
        "part2": 2,
        "part3": 3,
        "part4": 4,
        "part5": 5,
        "part6": 6,
        "part7": 7,
        "part8": 8,
        "part9": 9,
        "total": 101,
    }]
    generate('rank.xls', 'out.xls', data)


#for sheet_index in range(len(rb.sheets())):
#    sheet = wb.get_sheet(sheet_index)
#    sheet.write(0,0, "TESTERS %s" % sheet_index)

