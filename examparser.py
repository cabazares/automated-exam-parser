import math
import sqlite3

import cv2
import numpy as np


# marker image for page one
PAGE_MARKER = 'resources/page_marker.png'
# size of virtual page
PAGE_SIZE = (2000, 3200)
# color to check if black
# values less than the color threshold is considered black
COLOR_THRESHOLD = 50
# size of circles to check for
CIRC_SIZE = (32, 32)
# distance between four points when checking inside the circle
DIST_X, DIST_Y = map(lambda x: x * (3.0 / 8.0), CIRC_SIZE)


def otsu(img):
    """
    Threshold image using otsu algorithm

    retuurns modified image object
    """
    blur = cv2.GaussianBlur(img,(5,5),0)
    ret3,th3 = cv2.threshold(blur,0,255,cv2.THRESH_BINARY+cv2.THRESH_OTSU)
    return th3

def thresholding(img):
    """
    Threshold image using adaptive gaussian algorithm

    returns modified image object
    """
    ret,th1 = cv2.threshold(img,127,255,cv2.THRESH_BINARY)
    return th1
    return cv2.adaptiveThreshold(img, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                 cv2.THRESH_BINARY, 11, 2)

def show_img(img, window='image'):
    """
    Show image file in a window
    """
    cv2.imshow(window, img)
    cv2.waitKey(0)

def save_img(img, name):
    """
    save image object to file
    """
    cv2.imwrite(name,img)

def isPixel(img, x, y):
    """
    check if pixel is black
    """
    return img.item(y, x) < COLOR_THRESHOLD

def distance(x1, y1, x2, y2):
    """
    retrieve distance between two points
    """
    return math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)

def isFilled(img, coors):
    """
    count filled pixels
    """
    return len([c for c in coors if isPixel(img, c[0], c[1])])

def retrieve_relevant_area(fname):
    """
    Retrieve relevant exam area from image
    """
    img = cv2.imread(fname)
    img_size = img.shape[:2][::-1]
    gray = cv2.imread(fname, 0)
    #threshold using otsu
    thresh = otsu(gray)
    #save_img(thresh, 'otsu.png')

    # find shapes
    ret, th = cv2.threshold(gray,127,255,1)
    contours, h = cv2.findContours(th, 1, cv2.CHAIN_APPROX_SIMPLE)
    # draw all contours
    #cv2.drawContours(img, contours, -1, (0, 255,0), 1)

    # get trianles
    triangles = []
    for i, cnt in enumerate(contours):
        vertices = cv2.approxPolyDP(cnt, 0.05*cv2.arcLength(cnt,True), True)
        if len(vertices) == 3:
            triangles.append(vertices)
            # draw detected triangles
            #cv2.drawContours(img, contours, i, (255,0,0), 2)
    
    infinity = 999999999
    top_left_cnt = infinity
    top_right_cnt = infinity
    bottom_left_cnt = infinity
    bottom_right_cnt = infinity
    top_left = (0, 0)
    top_right = (0, 0)
    bottom_left = (0, 0)
    bottom_right = (0, 0)
    
    # calculate corners
    for tri in [item[0] for sublist in triangles for item in sublist]:
        x, y = tri
        dst = distance(0, 0, x, y)
        if top_left_cnt > dst:
            top_left_cnt = dst
            top_left = (x, y)
        dst = distance(img_size[0], 0, x, y)
        if top_right_cnt > dst:
            top_right_cnt = dst
            top_right = (x, y)
        dst = distance(0, img_size[1], x, y)
        if bottom_left_cnt > dst:
            bottom_left_cnt = dst
            bottom_left = (x, y)
        dst = distance(img_size[0], img_size[1], x, y)
        if bottom_right_cnt > dst:
            bottom_right_cnt = dst
            bottom_right = (x, y)
    
    # draw page border
    #cv2.rectangle(img, top_left, bottom_right, (0,255,0), 10)
    #cv2.line(img, top_left, top_right,(255,0,255),5)
    #cv2.line(img, bottom_right, top_right,(0, 255,255),5)
    #cv2.line(img, bottom_left, bottom_right,(255,255,0),5)
    #cv2.line(img, top_left, bottom_left,(255,0,0),5)
    
    #save_img(img, 'find_contours.png')
    
    # skew perspective to corners
    pts1 = np.float32([top_left, top_right, bottom_left, bottom_right])
    pts2 = np.float32([[0,0],[PAGE_SIZE[0],0],[0,PAGE_SIZE[1]],PAGE_SIZE])
    M = cv2.getPerspectiveTransform(pts1,pts2)
    img = cv2.warpPerspective(thresh, M, PAGE_SIZE)
    
    # check if page one or two
    # find marker on page
    template = cv2.imread(PAGE_MARKER,0)
    w, h = template.shape[::-1]
    res = cv2.matchTemplate(img, template, cv2.TM_SQDIFF_NORMED)
    min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
    top_left = min_loc
    bottom_right = (top_left[0] + w, top_left[1] + h)

    # draw page one marker 
    cv2.rectangle(img,top_left, bottom_right, 128, 10)
    
    # check position of marker to determine if page one or two
    is_top_right = (top_left[0] > PAGE_SIZE[0] / 2 and
                    top_left[1] < PAGE_SIZE[1] / 2)
    
    #save_img(img, 'find_page_one.png')
    return img, is_top_right


def read_student_number(img, pos):
    """
    Read student number from boxes given image and position
    
    Returns student number as string in the format YYYY-NNNNN
    """
    num_pos = (pos[0] + 165, pos[1])
    # size of the circle filled

    def read_nums(pos, length):
        vals = ""
        for j in range(length):
            val = -1
            cnt = 0
            for i in range(10):
                x = pos[0] + (j * (CIRC_SIZE[0] * 1.045) +
                              (CIRC_SIZE[0] * 0.35))
                y = pos[1] + (i * (CIRC_SIZE[1] * 1.1868) +
                              (CIRC_SIZE[1] * 0.37))
                coor1 = (int(x), int(y))
                coor2 = (int(x), int(y + DIST_Y))
                coor3 = (int(x +DIST_X), int(y))
                coor4 = (int(x +DIST_X), int(y + DIST_Y))

                fillCnt = isFilled(img, [coor1, coor2, coor3, coor4])
                if fillCnt >= cnt:
                    cnt = fillCnt
                    val = i
                cv2.rectangle(img, coor1, coor1, 128, 1)
                cv2.rectangle(img, coor2, coor2, 128, 1)
                cv2.rectangle(img, coor3, coor3, 128, 1)
                cv2.rectangle(img, coor4, coor4, 128, 1)
            if val != - 1:
                vals += str(val)
        return vals
                
    # get student number
    return "%s-%s" % (read_nums(pos, 4), read_nums(num_pos, 5))


def read_choice(img, pos, length, start = 0):
    """
    Read answer from multiple choice questions,
    given image, position and question length
    
    return dictionary of answers 
    """
    answers = {}
    for j in range(length):
        val = -1
        cnt = 1
        for i in range(5):
            x = pos[0] + (i * (CIRC_SIZE[0] * 1.038) +
                          (CIRC_SIZE[0] * 0.334))
            y = pos[1] + (j * (CIRC_SIZE[1] * 1.1868) +
                          (CIRC_SIZE[1] * 0.34))
            coor1 = (int(x), int(y))
            coor2 = (int(x), int(y + DIST_Y))
            coor3 = (int(x +DIST_X), int(y))
            coor4 = (int(x +DIST_X), int(y + DIST_Y))
            
            fillCnt = isFilled(img, [coor1, coor2, coor3, coor4])
            if fillCnt > cnt:
                cnt = fillCnt
                val = i
            cv2.rectangle(img, coor1, coor1, 128, 1)
            cv2.rectangle(img, coor2, coor2, 128, 1)
            cv2.rectangle(img, coor3, coor3, 128, 1)
            cv2.rectangle(img, coor4, coor4, 128, 1)
        if val != -1:
            answers[start + j+1] = "ABCDE"[val]
        else:
            answers[start + j+1] = " "
    return answers


def read_page_one(img):
    """
    Retrieve student number and answers (arranged by part) in page one
    
    returns dictionary
    """
    data = {"parts": {}}

    # read student number
    # position of the top left circle
    pos = (1632, 488)
    data["student_num"] = read_student_number(img, pos)

    item_total = 55
    box_y = 1090

    def read_parts(points):
        ans = {}
        total = 0
        for point in points:
            x, y, cnt = point
            ta = read_choice(img, (x, y), cnt, total)
            ans.update(ta)
            total += cnt
        return ans

    # set positions for each parts
    parts = [[70, 285], [500, 716], [930, 1147], [1362, 1578], [1794]]
    
    for i, part in enumerate(parts):
        args = [(x, box_y, item_total) for x in part]
        data["parts"][i] = read_parts(args)
    
    return data


def read_page_two(img):
    """
    Retrieve student number and answers (arranged by part) in page two

    return dictionary
    """
    data = {"parts": {}}

    # read student number
    # position of the top left circle
    pos = (1632, 2786)
    data["student_num"] = read_student_number(img, pos)

    box_y = 225
    item_total = 55

    # read part 5 continuation
    data["parts"][4] = read_choice(img, (67, box_y), item_total, item_total)
    
    def read_parts(points):
        ans = {}
        total = 0
        for point in points:
            x, y, cnt = point
            ta = read_choice(img, (x, y), cnt, total)
            ans.update(ta)
            total += cnt
        return ans

    # set positions for each parts (6-9)
    parts = [[285, 500], [716, 930], [1147, 1362], [1578, 1794]]
    
    for i, part in enumerate(parts):
        args = [(x, box_y, item_total) for x in part]
        data["parts"][5 + i] = read_parts(args)

    # read part 1 continuation
    box_y = 2423
    parts = [67, 285, 500, 716, 930, 1147, 1362, 1578, 1794]

    data["parts"][0] = {}
    for i, x in enumerate(parts):
        data["parts"][0].update(read_choice(img, (x, box_y), 1, 110 + i))
    # read part 1 item 120
    data["parts"][0].update(read_choice(img, (1794, 2461), 1, 119))

    return data


def check_images_generator(filenames):
    """
    get data from given set of filenames

    return dictionary of data per student num
    """
    students = {}
    for fname in filenames:
        img, is_page_one = retrieve_relevant_area(fname)
        if is_page_one:
            student_data = read_page_one(img)
        else:
            student_data = read_page_two(img)
        if student_data["student_num"] not in students.keys():
            students[student_data["student_num"]] = student_data
        else:
            data = students[student_data["student_num"]]
            data["parts"].update(student_data["parts"])
        yield students

def check_images(filenames):
    return [x for x in check_images_generator(filenames)][-1]


if __name__ == "__main__":
    fname = 'PAGE02.jpg'
    img, is_page_one = retrieve_relevant_area(fname)

    if is_page_one:
        student_data = read_page_one(img)
    else:
        student_data = read_page_two(img)

    save_img(img, 'result.png')
    print student_data
