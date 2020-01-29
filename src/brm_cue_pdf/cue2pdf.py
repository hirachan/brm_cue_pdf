#!/usr/local/bin/python2.7
# coding: utf-8

import os
from xml.dom import minidom

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.units import cm
from reportlab.graphics.shapes import *
from reportlab.graphics.charts.lineplots import LinePlot
from reportlab.graphics.widgets.markers import makeMarker
from reportlab.lib.colors import HexColor
from math import sin, cos, acos, radians
# import geopy
# import geopy.distance
import datetime
import xlrd

landmarks = [
    ("PTT 711", 15),
    ("711", 15.7),
    ("BigC", 79.9),
    ("711", 86.4),
    ("PTT711", 117.5),
    ("Tesco", 119),
    ("711", 118.8),
    ("711", 147.9),
    ("711", 197.7),
    ("711", 253.6),
    ("MiniBigC", 277.7),
    ("711", 302),
    ("PTT711", 404.3),
    ("711", 405.8),
    ("PTT711", 479.4),
    ("Tesco", 481.8),
    ("711", 481.8),
    ("BigC", 483.4),
    ("PTT711", 749),
    ("711", 749.1),
    ("Tesco", 749.4),
    ("711", 785.4),
    ("711", 786.8),
    ("711", 800.3),
    ("PTT711", 836.7),
    ("711", 837.5),
    ("PTT711", 886.9),
    ("Tesco Lotus", 888.0),
    ("PTT711", 897.8),
    ("PTT711", 937.2),
    ("711", 938.5),
    ("711", 940),
    ("PTT711", 946),
    ("711", 952.4),
    ("Tesco", 970),
    ("PTT711", 970.7),
    ("711", 1035.6),
    ("PT", 1073.1),
    ("CALTEX", 1098.1),
    ("Tesco", 1123),
    ("711", 1124),
    ("PTT711", 1125),
    ("711", 1170),
    ("PTT711", 1170.8),
    ("PTT711", 1188.5),
    ("711", 1189),
    ("PT", 1189.4),
    ("Tesco", 1214.3),
    ("PTT711", 1215),
    ("PTT711", 1217.7),
    ("PTT711", 1252.9),
    ("PTT711", 1257.6),
    ("PTT711", 1286.7),
    ("711", 1287.3),
    ("711", 1288.9),
    ("PTT711", 1301.9),
    ("711", 1302.9),
    ("Tesco", 1304),
    ("PTT711", 1341),
    ("711", 1404.7),
    ("711", 1446.1),
    ("711", 1446.8),
    ("PT", 1510),
    ("711", 1687.8),
    ("Tesco Express", 1732.5),
    ("MiniBigC", 1850.6),
    ("Tesco Express", 1850.6),
    ("PTT711", 1878.2),
    ("711", 1879.5),
    ("Tesco", 1879.7),
    ("CALTEX", 1883.6),
    ("711", 1896.8),
    ("PT", 1906.1),
    ("PTT711", 1915.0),
    ("711", 1922.2),
    ("PTT711", 1923.5),
    ("BigC", 1924.2),
    ("PTT711", 1929.1),
    ("Baangchaak", 1994.3),
    ("711", 2015.1),
    ("PTT711", 2020.9),
]

DIR = "/home/hirano/BRM/2020/ISAN2020"

CSV = os.path.join(DIR, "ISAN2020-q.xlsx")

TCX = (("ISAN2020.tcx", 0, 2024000),
)

START = datetime.datetime(2020, 2, 2, 7, 0, 0)
DISTANCE = 2020

HOSEI = float(DISTANCE) / (float(TCX[-1][2]) / 1000)

def get_limit_time(distance):
    distance = distance * HOSEI
    if distance <= 600:
        return START + datetime.timedelta(distance / (15.0 * 24))
    if distance <= 1000:
        return START + datetime.timedelta(40.0 / 24 + (distance - 600) / (11.428 * 24))
    if distance <= 1200:
        return START + datetime.timedelta(75.0 / 24 + (distance - 1000) / (13.333 * 24))
    if distance <= 1400:
        return START + datetime.timedelta(90.0 / 24 + (distance - 1200) / (11 * 24))
    if distance <= 1800:
        return START + datetime.timedelta(108.18 / 24 + (distance - 1400) / (10 * 24))
    return START + datetime.timedelta(148.18 / 24 + (distance - 1800) / (9 * 24))

def get_limit_time2(distance):
    distance = distance * HOSEI
    return START + datetime.timedelta(distance / (10.0 * 24))

def get_landmark(d1, d2):
    msg = ""
    for landmark, d in landmarks:
        if d1 <= d <=d2:
            msg += "*%.1fkm:%s " % (d - d1, landmark)

    return msg


# Q = []
# f = open(CSV, "r")
# csv_reader = csv.reader(f)
# next(csv_reader)
# prev_time = START
# for row in csv_reader:
#     if not row[1]:
#         continue

#     lap_distance = float(row[2])
#     if row[12]:
#         distance_to_next_pc = float(row[12])
#     else:
#         distance_to_next_pc = 0.0

#     if lap_distance > 10.0:
#         lap = 0.0
#         distance=float(row[1]) - lap_distance

#         while lap_distance > 10.0:
#             lap += 10.0
#             distance += 10.0
#             plan_time = prev_time + datetime.timedelta(lap / (float(row[8]) * 24))
#             q = dict(no=row[0],
#                      distance=distance,
#                      lap_distance=lap,
#                      comment1=u"%d/%.1f" % (lap, float(row[2])),
#                      direction="",
#                      signal="",
#                      road="",
#                      comment2="",
#                      speed=row[8].decode("shift-jis","ignore"),
# #                     plan_time=row[9].decode("shift-jis","ignore"),
#                      plan_time=plan_time.strftime("%-H:%M"),
#                      limit_time=get_limit_time(distance).strftime("%-H:%M"),
#                      rest=row[11].decode("shift-jis","ignore"),
#                      distance_to_next_pc=distance_to_next_pc)

#             Q.append(q)

#             lap_distance -= 10.0
#             distance_to_next_pc -= 10.0


#     plan_time = prev_time + datetime.timedelta(lap_distance / (float(row[8]) * 24))
#     q = dict(no=row[0],
#              distance=float(row[1]),
#              lap_distance=lap_distance,
#              comment1=row[3].decode("shift-jis","ignore"),
#              direction=row[4].decode("shift-jis","ignore"),
#              signal=row[5].decode("shift-jis","ignore"),
#              road=row[6].decode("shift-jis","ignore"),
#              comment2=row[7].decode("shift-jis","ignore"),
#              speed=row[8].decode("shift-jis","ignore"),
#              plan_time=row[9].decode("shift-jis","ignore"),
# #             limit_time=row[10].decode("shift-jis","ignore"),
#              limit_time=get_limit_time(float(row[1])).strftime("%-H:%M"),
#              rest=row[11].decode("shift-jis","ignore"),
#              distance_to_next_pc=distance_to_next_pc)

#     prev_time = datetime.datetime.strptime(row[9], "%H:%M")


#     Q.append(q)


class CueSheet(object):
    def floor(self, s):
        return int(float(s) * 10 + 0.5) / 10.0
    def read(self, filepath):
        Q = []
        _, ext = os.path.splitext(filepath)
        if ext in (".xls", ".xlsx"):
            book = xlrd.open_workbook(filepath)
            sheet = book.sheet_by_index(0)
            Q = []
            prev_time = START
            for row in range(1, sheet.nrows):
                try:
                    lap_distance = self.floor(sheet.cell(row, 2).value)
                except ValueError:
                    break

                if sheet.cell(row, 12).value:
                    distance_to_next_pc = self.floor(sheet.cell(row, 12).value)
                else:
                    distance_to_next_pc = 0.0

                if sheet.cell(row, 13).value:
                    distance_to_drop = self.floor(sheet.cell(row, 13).value)
                else:
                    distance_to_drop = 0.0

                if lap_distance > 10.0:
                    lap = 0.0
                    distance=self.floor(sheet.cell(row, 1).value) - lap_distance

                    while lap_distance > 10.0:
                        landmark = get_landmark(distance, distance + 10)
                        lap += 10.0
                        distance += 10.0
                        plan_time = prev_time + datetime.timedelta(lap / (float(sheet.cell(row, 8).value) * 24))
                        q = dict(no=sheet.cell(row, 0).value,
                                 distance=distance,
                                 lap_distance=lap,
                                 comment1="%d/%.1f" % (lap, self.floor(sheet.cell(row, 2).value)),
                                 direction="",
                                 signal="",
                                 road="",
                                 comment2=landmark,
                                 speed=sheet.cell(row, 8).value,
                                 plan_time=plan_time.strftime("%-H:%M"),
                                 limit_time=get_limit_time2(distance).strftime("%-H:%M"),
                                 # limit_time2=get_limit_time2(distance).strftime("%-H:%M"),
                                 rest=sheet.cell(row, 11).value,
                                 distance_to_next_pc=distance_to_next_pc,
                                 distance_to_drop=distance_to_drop,
                                 )

                        Q.append(q)

                        lap_distance -= 10.0
                        distance_to_next_pc -= 10.0
                        distance_to_drop -= 10.0


                plan_time = datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell(row, 9).value + 367, book.datemode))
                landmark = get_landmark(float(sheet.cell(row, 1).value) - lap_distance, float(sheet.cell(row, 1).value))

                q = dict(no=sheet.cell(row, 0).value,
                         distance=self.floor(sheet.cell(row, 1).value),
                         lap_distance=lap_distance,
                         comment1=sheet.cell(row, 3).value,
                         direction=sheet.cell(row, 4).value,
                         signal=sheet.cell(row, 5).value,
                         road=sheet.cell(row, 6).value,
                         comment2=sheet.cell(row, 7).value + landmark,
                         speed=sheet.cell(row, 8).value,
                         # plan_time=sheet.cell(row, 9).value,
                         plan_time=plan_time.strftime("%-H:%M"),
            #             limit_time=row[10].decode("shift-jis","ignore"),
                         limit_time=get_limit_time2(float(sheet.cell(row, 1).value)).strftime("%-H:%M"),
                         # limit_time2=get_limit_time2(float(sheet.cell(row, 1).value)).strftime("%-H:%M"),
                         rest=sheet.cell(row, 11).value,
                         distance_to_next_pc=distance_to_next_pc,
                         distance_to_drop=distance_to_drop,

                         )

#                print sheet.cell(row, 9).value
#                print xlrd.xldate_as_tuple(sheet.cell(row, 9).value + 367, book.datemode)
                prev_time = datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell(row, 9).value + 367, book.datemode))


                Q.append(q)


        return Q

cuesheet = CueSheet()
Q = cuesheet.read(CSV)


pdfFile = canvas.Canvas(os.path.join(DIR, "ISAN2020.pdf"), pageCompression=1)
pdfFile.saveState()
pdfFile.setAuthor("hirano")
pdfFile.setTitle("ISAN2020-012901")
pdfFile.setSubject("ISAN2020")

pdfFile.setPageSize((29.7*cm,21.0*cm))

fontname = "IPA Gothic"


earth_rad = 6378.137

def latlng_to_xyz(lat, lng):
    rlat, rlng = radians(lat), radians(lng)
    coslat = cos(rlat)
    return coslat*cos(rlng), coslat*sin(rlng), sin(rlat)

def dist_on_sphere(pos0, pos1, radious=earth_rad):
    xyz0, xyz1 = latlng_to_xyz(*pos0), latlng_to_xyz(*pos1)
    return acos(sum(x * y for x, y in zip(xyz0, xyz1)))*radious


def put_text(pdf, y, text, size, x=0, font=fontname):
    max_len = int(1550 / size)
    y_size = size
    line_no = 0

    l = 0
    line = ""
    for c in text:
        if ord(c) in (0x0a, 0x3001, 0x3002):
            c = " "
        if ord(c) < 0x80:
            l += 1
        elif ord(c) in (0x03e1, 0x0e34, 0x0e35, 0x0e36, 0x0e37, 0x0e38, 0x0e39, 0x0e3a, 0x0e47, 0x0e48, 0x0e49, 0x0e4a, 0x0e4b, 0x0e4c, 0x0e4d, 0x0e4f):
            l += 0
        elif 0x0e01 <= ord(c) <= 0x0e5b:
            l += 1
        else:
            l += 2

        line += c

        if l > max_len:
            pdf.setFont(font, size)
            pdf.drawString(1*cm, y - line_no * y_size, line)
            line = ""
            l = 0
            line_no += 1


    if line:
        pdf.setFont(font, size)
        pdf.drawString(1*cm + x, y - line_no * y_size, line)

    return y - (line_no + 1) * y_size

def chart(distance, altitude, angle, next_point, qn):
    data1 = []
    data2 = []
    data3 = []
    alt_min = alt_max = alt_min_all = alt_max_all = altitude[0]
    for i in range(len(distance)):
        data1.append((distance[i], altitude[i]))

        if distance[i] <= next_point:
            if alt_min > altitude[i]:
                alt_min = altitude[i]

            if alt_max < altitude[i]:
                alt_max = altitude[i]

        if alt_min_all > altitude[i]:
            alt_min_all = altitude[i]

        if alt_max_all < altitude[i]:
            alt_max_all = altitude[i]

    _alt_min = int(alt_min / 100) * 100
    _alt_max = int(alt_max / 100) * 100 + 100
    _alt_min_all = int(alt_min_all / 100) * 100
    _alt_max_all = int(alt_max_all / 100) * 100 + 100

    rate = (_alt_max_all - _alt_min_all) / 15
    for i in range(len(distance)):
        data2.append((distance[i], angle[i] * rate + _alt_min_all))

#    data = (data1, data2)
    data = (data1,)

    drawing = Drawing(400, 0)

    lp = LinePlot()
    lp.x = 25
    lp.y = 20
    lp.height = 400
    lp.width = 815
    lp.data = data
    lp.joinedLines = 1
#    lp.lines[0].symbol = makeMarker('FilledCircle')
#    lp.lines[1].symbol = makeMarker('Circle')
    lp.lines[0].strokeColor = HexColor(0x808080)
    lp.lines[0].strokeWidth = 3
    lp.lines[1].strokeColor = HexColor(0x808080)
    lp.lines[1].strokeWidth = 1


    lp.xValueAxis.valueMin = 0
    lp.xValueAxis.valueMax = 10
    lp.xValueAxis.valueSteps = list(range(1, 11)) + [next_point]
    lp.xValueAxis.labelTextFormat = '%.1f'
#    lp.xValueAxis.strokeColor = HexColor(0x808080)

    lp.yValueAxis.valueMin = _alt_min_all
    lp.yValueAxis.valueMax = _alt_max_all
    lp.yValueAxis.valueSteps = [alt_min_all, alt_max_all] + list(range(_alt_min_all, _alt_max_all + 1, 100))
    lp.yValueAxis.labelTextFormat = '%d'

    drawing.add(lp)

    x = 815 * next_point / 10 + 25
    drawing.add(Line(x, 20, x, 400 + 20, strokeColor=HexColor(0x808080)))


    prev_distance = distance[0]
    prev_alt = altitude[0]
    for i in range(1, len(distance)):
        high = max(prev_alt, altitude[i])
        low = min(prev_alt, altitude[i])
        for alt in range(int(low) // 3 + 1, int(high) // 3 + 1):
            rate = (alt * 3 - low) * 1.0 / (high - low)
            d = (distance[i] - prev_distance) * rate + prev_distance
            x = 815 * d / 10 + 25
            if low == prev_alt:
                drawing.add(Line(x, 30 + 20, x, 80 + 20, strokeColor=HexColor(0x000000)))
            else:
                drawing.add(Line(x, 10 + 20, x, 30 + 20, strokeColor=HexColor(0x808080)))


        prev_distance = distance[i]
        prev_alt = altitude[i]





    drawing.drawOn(pdfFile, 0, 0)

    pdfFile.line(0, 0, 0, 100)


    pdfmetrics.registerFont(TTFont(fontname, '/usr/share/fonts/ipa-gothic/ipag.ttf'))
    pdfmetrics.registerFont(TTFont("waree", '/usr/share/fonts/thai-scalable/Waree.ttf'))

    pdfFile.setFont(fontname, 40)
    pdfFile.drawString(1*cm, 19.5*cm, str(round(Q[qn]["distance"], 1)))

    pdfFile.setFont(fontname, 50)
    pdfFile.drawString(7*cm, 19.5*cm, str(Q[qn]["limit_time"]))

    pdfFile.setFont(fontname, 50)
    pdfFile.drawString(13.5*cm, 19.5*cm, str(round(Q[qn]["distance_to_next_pc"], 1)))

    pdfFile.setFont(fontname, 50)
    pdfFile.drawString(19*cm, 19.5*cm, str(Q[qn]["plan_time"]))

    pdfFile.setFont(fontname, 40)
    pdfFile.drawString(25*cm, 19.5*cm, str(Q[qn - 1]["plan_time"]))
    # pdfFile.drawString(25*cm, 19.5*cm, str(int(Q[qn]["no"])))


    direction = "%s %s %s" % (round(Q[qn]["lap_distance"], 1), Q[qn]["direction"], Q[qn]["road"])
    y = 16*cm
    y = put_text(pdfFile, y, direction, 110, 0.5*cm)
#    pdfFile.drawString(0*cm, 17*cm, direction)

#    y = 14*cm
    if Q[qn]["comment1"]:
        y = put_text(pdfFile, y, Q[qn]["comment1"], 90, font=fontname)

    y = put_text(pdfFile, y, "%s%s" % (Q[qn]["signal"], Q[qn]["comment2"]), 60, font=fontname)

    y = put_text(pdfFile, y, "%d - %d - %d - %d" % (altitude[0], alt_min, alt_max, alt_max_all), 50)

    pdfFile.setFont(fontname, 50)
    pdfFile.drawString(25*cm, 1.0*cm, str(round(Q[qn]["distance_to_drop"], 1)))




    pdfFile.showPage()

def plot(data, n, qn):
    if qn > 1:
        point = Q[qn - 1]["distance"] * 1000
    else:
        point = 0

    try:
        next_point = Q[qn]["distance"] * 1000
    except IndexError:
        next_point = point

    print("%.1f - %.1f" % (point / 1000, next_point / 1000))

    distance = []
    altitude = []
    angle = []

    for _n in range(n, len(data)):
        if data[_n][0] > point + 10000:
            break

        distance.append((data[_n][0] - point) * 1.0 / 1000.0)
        altitude.append(data[_n][1])
#        altitude.append(data[_n][1] % 100)
        angle.append(data[_n][2])

    chart(distance, altitude, angle, (next_point - point) / 1000, qn)



def main():
    res = []
    prev_distance = 0
    prev_altitude = 0

    data = []

#    ddd = 0
#    ddd2 = 0
    prev_pos = None
    for tcx, base_distance, base_end_distance in TCX:
        _data = []
        dom = minidom.parse(os.path.join(DIR, tcx))
        for x in dom.getElementsByTagName("Trackpoint"):
            distance = float(x.getElementsByTagName("DistanceMeters")[0].firstChild.data)
            altitude = float(x.getElementsByTagName("AltitudeMeters")[0].firstChild.data)

            pos = x.getElementsByTagName("Position")
            latitude = float(x.getElementsByTagName("LatitudeDegrees")[0].firstChild.data)
            longitude = float(x.getElementsByTagName("LongitudeDegrees")[0].firstChild.data)

#            if prev_pos:
#                if prev_pos != (latitude, longitude):
#                    ddd += dist_on_sphere(prev_pos, (latitude, longitude))
#                    ddd2 += geopy.distance.distance(geopy.Point(prev_pos[0], prev_pos[1]), geopy.Point(latitude, longitude)).km

#                    print distance, ddd, ddd2

            prev_pos = (latitude, longitude)

#            _data.append((base_distance + distance, altitude))
            _data.append((distance, altitude))

        end_distance = _data[-1][0]

        hosei = (base_end_distance - base_distance) * 1.0 / end_distance

        for distance, altitude in _data:
            print(hosei, distance, distance * hosei)
            data.append((base_distance + distance * hosei, altitude))


    for distance, altitude in data:
        if distance == 0:
            res.append((distance, altitude, 0))

        elif distance - prev_distance < 100:
            continue

        else:
            angle = 100.0 * (altitude - prev_altitude) / (distance - prev_distance)

            if angle > 15:
                angle = 15
            elif angle < -1:
                angle = -1

            res.append((distance, altitude, angle))

        prev_distance = distance
        prev_altitude = altitude

    p = -1
    point = Q[0]["distance"] * 1000
    for n in range(len(res)):
        if res[n][0] > point:
            plot(res, n - 1, p + 1)
            p += 1
            try:
                point = Q[p]["distance"] * 1000
                if p + 1 >= len(Q):
                    break
            except IndexError:
                break


def convert():
    main()
    pdfFile.save()
