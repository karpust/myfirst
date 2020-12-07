import sys
import time
import copy
import math
import win32com.client
from pythoncom import VT_R8, VT_ARRAY, VT_DISPATCH, VT_BSTR, VT_I2, VT_VARIANT

sys.path.append("C:\\Users\\UserGeo\\AppData\Local\Programs\Python\Python37-32\Lib\site-packages\win32")
sys.path.append("C:\\Users\\UserGeo\\AppData\Local\Programs\Python\Python37-32\Lib\site-packages")
sys.path.append("C:\\Users\\UserGeo\\AppData\Local\Programs\Python\Python37-32\Lib\site-packages\win32\lib")
sys.path.append("C:\\Users\\UserGeo\\AppData\Local\Programs\Python\Python37-32\Lib\site-packages\Pythowin")

app = win32com.client.Dispatch("AutoCAD.Application")
aDoc = app.ActiveDocument
msp = aDoc.ModelSpace
pfss = aDoc.PickfirstSelectionSet

# print(dir(win32com.client))


def to_var(coord_po: list):
    """преобразует список координат в вариант"""
    coord_po_variant = win32com.client.VARIANT(VT_ARRAY | VT_R8, coord_po)
    return coord_po_variant


def two_to_three(coord_po: list):
    """преобразует двумерные координаты в трехмерные"""
    if len(coord_po) < 3:
        coord_po.append(0)
    return coord_po


def to_real(var):
    """преобразует вариант в действительное число"""
    real = win32com.client.VARIANT(VT_R8, var)
    return real


def select_last(data, type=0):
    """находит последний отрисованный объект
    например по типу дуга: data = 'ARC' """
    sset = aDoc.SelectionSets.Add('1arc')  # создали пустой набор под номером
    ftype = [type]  # dxf код
    fdata = [data]
    ftype = win32com.client.VARIANT(VT_ARRAY | VT_I2, ftype)  # вариант из целого числа
    fdata = win32com.client.VARIANT(VT_ARRAY | VT_VARIANT, fdata)  # вариант из варианта из строки
    sset.Select(5, None, None, ftype, fdata)  # добавили в набор
    last_elem = sset.Item(0)
    sset.Delete()
    return last_elem

# def number_p(a, b):                   # - номер точки, вход-список коорд точки
#     pt1 = copy.copy(a)
#     pt1[:] = [i+.05 for i in pt1]
#     pt1[2] = 0
#     pt1_v = to_var(pt1)
#     msp.AddText(text[b], pt1_v, .1)
#     return


# def draw_po(a, b):                        # - отрисовка точки
#     p0_v = to_var(a)
#     number_p(a, b)
#     msp.AddPoint(p0_v)
#     return()


def to_nam_crd(coord: list):
    """Добавляет в словарь имя и координаты"""
    nam_crd.setdefault(text[1], coord)
    return coord


# def to_nam_obj(nam1, nam2):               # рисует отрезок по именам точек, добавляет в словарь объект
#     nam_obj.setdefault(nam1 + nam2, draw_line_nam(nam1, nam2))
#     return


def to_nam_obj(*args):  # names of points
    """Рисует отрезок по именам точек, добавляет в словарь отрезок-объект"""
    lst = []
    for n in args:
        lst.append(n)
    for i in range(len(lst) - 1):
        p1 = to_var(nam_crd[lst[i]])
        p2 = to_var(nam_crd[lst[i + 1]])
        lin = msp.AddLine(p1, p2)
        nam_obj.setdefault(lst[i] + lst[i+1], lin)
    return


def coord_p(dim, ang, coord_start_po):
    """Вычисляет координаты по расстоянию, углу в градусах
    и координатам начальной точки"""
    x1 = coord_start_po[0]
    y1 = coord_start_po[1]
    x2, y2 = x1+dim*(math.cos(math.radians(ang))), y1+dim*(math.sin(math.radians(ang)))
    f = [x2, y2, 0]
    return f


def draw_line_po(coord_p1, coord_p2):
    """Рисует отрезок по варианту координат"""
    coord_p1 = to_var(coord_p1)
    coord_p2 = to_var(coord_p2)
    line = msp.AddLine(coord_p1, coord_p2)
    return line


# def draw_line_nam(nam1, nam2):       # + рисует отрезок по именам точек
#     p1 = to_var(nam_crd[nam1])
#     p2 = to_var(nam_crd[nam2])
#     lin = msp.AddLine(p1, p2)
#     return lin


# def draw_line_nam(*args):            # + рисует отрезок по именам точек
#     lst = []
#     for n in args:
#         lst.append(n)
#     for i in range(len(lst)-1):
#         p1 = to_var(nam_crd[lst[i]])
#         p2 = to_var(nam_crd[lst[i+1]])
#         lin = msp.AddLine(p1, p2)
#     return lin


# def draw_pline(*ls):                 # соединяет точки с выбранными номерами полилинией
#     m = copy.deepcopy(ls)
#     lst = []
#     for n in m:
#         n.pop()                      # приводит трехмерные координаты к двумерным
#         lst.append(n[0])
#         lst.append(n[1])
#     lst = to_var(lst)
#     lst = msp.AddLightWeightPolyline(lst)
#     return lst


def main_fun():
    """Создает объект используя списки длин, углов, координат начальных точек,
    добавляет в словарь объектов"""
    p0 = nam_crd[text[0]]
    p1 = coord_p(lenght[0], angle[0], p0)
    if text[1] == 'v':                    # если имя = v то запись в словарь и выйти из ф в против выполнить
        nam_obj.setdefault(text[0]+text[1], draw_line_po(p0, p1))
    else:
        if text[1] not in nam_crd:
            to_nam_crd(p1)
        draw_line_po(p0, p1)
    del (lenght[0], angle[0], text[0:2])
    return


def intersect_to_po(line1, line2, number):
    """Определяет точки пересечения отрезков и дуг"""
    inters_po_var = line1.IntersectWith(line2, number)
    return inters_po_var


def inters_to_dict(new_name_po, name_po1, name_po2, number):
    """добавляет в словарь точку пересечения заданных отрезков/дуг
    а - имя новой точки
    в,с - имя сущ точек"""
    nam_crd.setdefault(new_name_po, list(intersect_to_po(nam_obj[name_po1], nam_obj[name_po2], number)))
    return()


def mid_po(a, b=[]):
    """Определяет 3х-мерные координаты центра отрезка a или середину между точками а, b
    возвращает список 3х-мерных координат"""
    if a in nam_obj:
        a1 = nam_obj[a].StartPoint
        b1 = nam_obj[a].EndPoint
        x = (a1[0]+b1[0])/2
        y = (a1[1]+b1[1])/2
        coord_po_mid = [x, y, 0]
        return coord_po_mid
    else:
        x = (a[0]+b[0])/2
        y = (a[1]+b[1])/2
        coord_po_mid = [x, y, 0]
        return coord_po_mid


# def two_thir_po(a, b):
#     a, b, c, d = nam_crd[a], nam_crd[b], 0, 0
#     if a[0] < b[0]:
#         c = a[0] + abs(a[0]) - abs(b[0]) / 3 * 2
#     if a[0] > b[0]:
#         c = a[0] - abs(a[0]) - abs(b[0]) / 3 * 2
#     if a[1] < b[1]:
#         d = a[1] + abs(a[1]) - abs(b[1]) / 3 * 2
#     if a[1] > b[1]:
#         d = a[1] - abs(a[1]) - abs(b[1]) / 3 * 2
#     f = [c, d, 0]
#     return f


def add_arc(name_center_po, r, start_ang, end_ang):
    """Рисует дугу по варианту координат центра и радиусу
    nam - имя центра"""
    c = msp.AddArc(to_var(nam_crd[name_center_po]), r, start_ang, end_ang)
    return c


def ang_bis(name_obj1, name_obj2):
    """вычисляет угол биссектрисы в градусах
    по имени двух объектов"""
    l1, l2 = nam_obj[name_obj1], nam_obj[name_obj2]
    if l1.Angle > l2.Angle:
        bis = (l1.Angle - l2.Angle) / 2
    else:
        bis = (l1.Angle + l2.Angle) / 2
    return bis*grad


def po_at_arc(name_center_po, name_r_obj, nam_line_obj, nam_dim1, nam_dim2):
    """Вычисляет координаты точки по длине дуги"""
    r = main_dic[name_r_obj]
    line = nam_obj[nam_line_obj]
    len_circle = math.pi * r * 2
    len_arc = main_dic[nam_dim1] - main_dic[nam_dim2]
    coord_center = nam_crd[name_center_po]
    x_center = coord_center[0]
    y_center = coord_center[1]
    ang = line.Angle + (len_arc*2*math.pi)/len_circle
    x1 = math.cos(ang) * r + x_center
    y1 = math.sin(ang) * r + y_center
    coord = [x1, y1, 0]
    return coord


def add_pos_names():
    """Добавляет имена точек в чертеж"""
    nam = list(nam_crd.keys())
    for k in nam:
        p0 = nam_crd[k]
        p0_v = to_var(p0)
        msp.AddPoint(p0_v)
        pt0 = copy.copy(p0)
        pt0[:] = [i+0.01 for i in pt0]
        pt0[2] = 0
        pt0_v = to_var(pt0)
        msp.AddText(k, pt0_v, .1)
    return()


def create_third_po_at_line(object_name, dim, change_angle, name_new_po='Ar'):
    """ищет третью точку для построения дуги по имени
    объекта-отрезка и расстоянию до дуги"""
    mid = mid_po(nam_obj[object_name].StartPoint, nam_obj[object_name].EndPoint)
    ang = nam_obj[object_name].Angle*grad + change_angle
    c = coord_p(dim, ang, mid)
    nam_crd[name_new_po] = c
    return c


def mirror_by_two_po(fdata, name_po1, name_po2):
    """отражает объект по имени последнего созданного
     объекта и именам двух точек оси отражения"""
    point1 = to_var(two_to_three(nam_crd[name_po1]))
    point2 = to_var(two_to_three(nam_crd[name_po2]))
    select_last(fdata).Mirror(point1, point2)
    return


def make_curve7(lst, num_p0, n=2):
    """Строит дугу на 7 см выше линии Т
    на вытачке и линию к середине этой дуги"""
    nam_crd['c1'] = coord_p(7, nam_obj[lst[n]].Angle*grad, nam_obj[lst[n]].StartPoint)
    nam_crd['c2'] = mid_po(nam_obj[lst[n]].StartPoint, nam_crd['c1'])
    nam_crd['c3'] = coord_p(0.5, nam_obj[lst[n]].Angle*grad-90, nam_crd['c2'])
    main_arc('c1', 'c3', '4' + num_p0)
    mirror_by_two_po('ARC', '1' + num_p0, num_p0)
    nam_obj.setdefault('arc1', select_last('ARC'))
    to_nam_obj('1' + num_p0, "c3")
    mirror_by_two_po('LINE', '1' + num_p0, num_p0)
    return


def vit(num_p0, dim1, dim2, dim3, dim4, num_p_end):
    """строит вытачки по измеренным глубинам
    vit('T1', 'Гтс1i', 'Гтс2i', 'Гтс1', 'Гтс2', 'B1')
    vit('T2', 'Гб1i', 'Гб2i', 'Гб1', 'Гб2', 'B2')
    vit('T4', 'Гт1i', 'Гт2i', 'Гт1', 'Гт2', 'B4')"""

    nam_crd.setdefault('1' + num_p0, coord_p(main_dic[dim1], 90, nam_crd[num_p0]))
    nam_crd.setdefault('2' + num_p0, coord_p(main_dic[dim2], 270, nam_crd[num_p0]))
    nam_crd.setdefault('3' + num_p0, coord_p(main_dic[dim3] / 2, 180, nam_crd[num_p0]))
    nam_crd.setdefault('4' + num_p0, coord_p(main_dic[dim3] / 2, 0, nam_crd[num_p0]))
    nam_crd.setdefault('5' + num_p0, coord_p(abs(main_dic[dim3] - main_dic[dim4]) / 2, 180, nam_crd['2' + num_p0]))
    nam_crd.setdefault('6' + num_p0, coord_p(abs(main_dic[dim3] - main_dic[dim4]) / 2, 0, nam_crd['2' + num_p0]))
    nam_crd.setdefault('7' + num_p0, coord_p(abs(main_dic[dim3] - main_dic[dim4]) / 2, 180, nam_crd[num_p_end]))
    nam_crd.setdefault('8' + num_p0, coord_p(abs(main_dic[dim3] - main_dic[dim4]) / 2, 0, nam_crd[num_p_end]))

    if (main_dic[dim3] - main_dic[dim4]) > 0:
        # to_nam_obj('7' + num_p0, '5' + num_p0, '3' + num_p0, '1' + num_p0, '4' + num_p0, '6' + num_p0, '8' + num_p0)
        to_nam_obj('8' + num_p0, '6' + num_p0, '4' + num_p0, '1' + num_p0, '3' + num_p0, '5' + num_p0, '7' + num_p0)
        # lst = ('7' + num_p0, '5' + num_p0, '3' + num_p0, '1' + num_p0, '4' + num_p0, '6' + num_p0, '8' + num_p0)
        lst = ('8' + num_p0, '6' + num_p0, '4' + num_p0, '1' + num_p0, '3' + num_p0, '5' + num_p0, '7' + num_p0)
        lst = make_list_name_obj(lst)

        # #  построение дуг вытачки
        # create_third_po_at_line(lst[1], 0.1, 90)
        # main_arc('5' + num_p0, 'Ar', '3' + num_p0)
        # mirror_by_two_po('ARC', '1' + num_p0, num_p0)
        # create_third_po_at_line(lst[2], 0.3, 90)
        # main_arc('3' + num_p0, 'Ar', '1' + num_p0)
        # mirror_by_two_po('ARC', '1' + num_p0, num_p0)

        # построение верхней дуги вытачки 7см и линии
        make_curve7(lst, num_p0)
        # построение нижней дуги вытачки
        # create_third_po_at_line(lst[1], 0.4, 90)
        # main_arc('3'+num_p0, 'Ar', '5'+num_p0)
        # mirror_by_two_po('ARC', '1'+num_p0, num_p0)



    elif (main_dic[dim3] - main_dic[dim4]) < 0:
        to_nam_obj('7' + num_p0, '5' + num_p0, '4' + num_p0, '1' + num_p0, '3' + num_p0, '6' + num_p0, '8' + num_p0)
        lst = ('7' + num_p0, '5' + num_p0, '4' + num_p0, '1' + num_p0, '3' + num_p0, '6' + num_p0, '8' + num_p0)
        lst = make_list_name_obj(lst)

        # построение верхней дуги вытачки 7см и линии
        make_curve7(lst, num_p0)
        # построение нижней дуги вытачки
        create_third_po_at_line(lst[1], 0.4, 90)
        main_arc('5'+num_p0, 'Ar', '4'+num_p0)
        mirror_by_two_po('ARC', '1'+num_p0, num_p0)

    else:
        to_nam_obj('2' + num_p0, '4' + num_p0, '1' + num_p0, '3' + num_p0, '2' + num_p0)
        lst = ('2' + num_p0, '4' + num_p0, '1' + num_p0, '3' + num_p0, '2' + num_p0)
        lst = make_list_name_obj(lst)

        # построение верхней дуги вытачки 7см и линии
        make_curve7(lst, num_p0, 1)
        # построение нижней дуги вытачки
        create_third_po_at_line(lst[0], 0.4, 90)
        main_arc('2'+num_p0, 'Ar', '4'+num_p0)
        mirror_by_two_po('ARC', '1'+num_p0, num_p0)
    return


def make_list_name_obj(list_name_po):
    list_name_obj = []
    for i in range(len(list_name_po)-1):
        name_obj = list_name_po[i] + list_name_po[i+1]
        list_name_obj.append(name_obj)
    return list_name_obj


def adds_to_vit(a, b):
    """Определяет раствор вытачек в соответствии с прибавкой"""
    up_min = (a, b)
    measure = main_dic['Гт1'] + main_dic['Гтс1'] + main_dic['Гб1']
    comput = nam_obj['GG3'].length - nam_obj['Gg'].length - main_dic['Олт'] - main_dic['Ст'] - main_dic['Пт']
    p = comput - measure
    m = 100
    k = 0
    for i in range(2):
        if main_dic[up_min[i]] < m:
            m = main_dic[up_min[i]]
            k = up_min[i]
    m += p  # main_dic['Пт']
    main_dic[k] = m


# class Vitochka:
#     """строит вытачки по измеренным глубинам
#             vit('T1', 'Гтс1i', 'Гтс2i', 'Гтс1', 'Гтс2', 'B1')
#             vit('T2', 'Гб1i', 'Гб2i', 'Гб1', 'Гб2', 'B2')
#             vit('T4', 'Гт1i', 'Гт2i', 'Гт1', 'Гт2', 'B4')"""
#     def __init__(self, num_p0, dim1, dim2, dim3, dim4, num_po_end):
#         self.num_po = num_p0
#         self.dim1 = dim1
#         self.dim2 = dim2
#         self.dim3 = dim3
#         self.dim4 = dim4
#         self.num_po_end = num_po_end
#
#     def make_po(self, num_p0, dim1, dim2, dim3, dim4, num_po_end):
#         for i in range(8):
#
#         nam_crd.setdefault('1' + num_p0, coord_p(main_dic[dim1], 90, nam_crd[num_p0]))
#         nam_crd.setdefault('2' + num_p0, coord_p(main_dic[dim2], 270, nam_crd[num_p0]))
#         nam_crd.setdefault('3' + num_p0, coord_p(main_dic[dim3] / 2, 180, nam_crd[num_p0]))
#         nam_crd.setdefault('4' + num_p0, coord_p(main_dic[dim3] / 2, 0, nam_crd[num_p0]))
#         nam_crd.setdefault('5' + num_p0, coord_p(abs(main_dic[dim3] - main_dic[dim4]) / 2, 180, nam_crd['2' + num_p0]))
#         nam_crd.setdefault('6' + num_p0, coord_p(abs(main_dic[dim3] - main_dic[dim4]) / 2, 0, nam_crd['2' + num_p0]))
#         nam_crd.setdefault('7' + num_p0, coord_p(abs(main_dic[dim3] - main_dic[dim4]) / 2, 180, nam_crd[num_po_end]))
#         nam_crd.setdefault('8' + num_p0, coord_p(abs(main_dic[dim3] - main_dic[dim4]) / 2, 0, nam_crd[num_po_end]))



class MakeArcFromThree:
    """строит дугу по именам трех двумерных точек.
    Создает строку вида ('_arc 28.7,116.5 32.9,118.1 34.5,122.3 ')
    и отправляет ее командой в автокад"""
    def __init__(self, start):
        self.start = start
        start = nam_crd[start]
        start = start.copy()
        self.start = start

    def threeToTwo(self):
        """make три координаты в две"""
        num = self.start.copy()
        num.pop()
        return num

    def listToString(self):
        """convert list to string"""
        return ','.join([str(el) for el in (self.threeToTwo())])

    def writeCommand(self):
        """make string of command to draw arc"""
        return self.listToString() + ' '

    def sendCommand(self, a, b, c):
        """send the command to AutoCad"""
        com = '_Arc ' + a + b + c + ' '
        aDoc.SendCommand(com)


def main_arc(start, center, end):
    d1 = MakeArcFromThree(start)
    d2 = MakeArcFromThree(center)
    d3 = MakeArcFromThree(end)
    a = d1.writeCommand()
    b = d2.writeCommand()
    c = d3.writeCommand()
    d1.sendCommand(a, b, c)


def check_measure():
    """проверяет правильность снятия мерок условием
    Ст = Сг - Олт - Сумма вытачек
    Сб = Сг + Пг + (Пб - Пг) - Олб + Пб
    Ст = сумме отрезков не входящих в вытачку по линии талии и
    Сб = сумме отрезков не входящих в вытачку по линии бедер"""
    """мерки глубин от талии до груди"""
    to_nam_obj('Ti', '3T1')
    to_nam_obj('4T1', '3T2')
    to_nam_obj('4T2', '3T4')
    to_nam_obj('4T4', 'T3')

    """мерки глубин от талии до бедер"""
    to_nam_obj('Bi', '8T1')
    to_nam_obj('7T1', '8T2')
    to_nam_obj('7T2', '7T4')
    to_nam_obj('8T4', 'B3')
    measure_up = nam_obj['Ti3T1'].Length + nam_obj['4T13T2'].Length + nam_obj['4T23T4'].Length + nam_obj['4T4T3'].Length
    measure_down = nam_obj['Bi8T1'].Length + nam_obj['7T18T2'].Length + nam_obj['7T27T4'].Length + nam_obj['8T4B3'].Length
    print('Прибавка по талии факт = ', round(measure_up - main_dic['Ст'], 1))
    print('Прибавка по бедрам факт = ', round(measure_down - main_dic['Сб'], 1))

    up = main_dic['Ст'] + main_dic['Пт']
    down = main_dic['Сб'] + main_dic['Пб']

    if abs(measure_up - up) <= 0.5 and abs(measure_down - down) <= 0.5:
        print('\nПоздравляю маэстро!!!!!!\nВы постигли глубины!', '\n', up - measure_up, down - measure_down)
    elif abs(measure_up - up) > 0.5 and abs(measure_down - down) <= 0.5:
        print('\nКлассная грудь! \nМожет все-таки замерим ее?', '\n', up - measure_up)
    elif abs(measure_up - up) <= 0.5 and abs(measure_down - down) > 0.5:
        print('\nКажется попочка у нее что надо! \nА теперь замерь ее правильно!', '\n', down - measure_down)
    else:
        print('\nТы мерки снимала или че делала там?? '
              '\nА ну перемеряй все нах!', up - measure_up, down - measure_down)


'''-----------------------------------------------------------------------------------------------------------------'''
# """создание таблицы один раз"""
# strt_tabl = [-65, 160, 0]
# fortable = ['Тип фигуры', 'Сш', 'Сг', 'Ст', 'Сб', 'Шг1', 'Шг2', 'Цг', 'Шс', 'Дпл',
#             'Дтп', 'Вг', 'Вг2', 'Впкп', 'Дтс', 'ВПРЗ', 'Впк', 'Вб', 'Гт1', 'Гт1i',
#             'Гт2', 'Гт2i', 'Гтс1', 'Гтс1i', 'Гтс2', 'Гтс2i', 'Гб1', 'Гб1i', 'Гб2',
#             'Гб2i']
# tabl = msp.AddTable(to_var(strt_tabl), 31, 2, 10, 30)
# i = 0
# for _ in fortable:
#     tabl.SetCellValueFromText(i+1, 0, fortable[i], 0)
#     i += 1

#  считывание с таблицы
table = select_last('ACAD_TABLE')  # выбор таблицы

main_dic = dict()
k = 1
while k < table.Rows:
    if k == 1:
        main_dic.setdefault(table.GetCellValue(k, 0), (table.GetCellValue(k, 1)))
    else:
        main_dic.setdefault(table.GetCellValue(k, 0), float(table.GetCellValue(k, 1)))  # взяли из таблицы значение по номеру ячейки и столбца
    k += 1


# main_dic = dict(Сш=19, Сг=50, Ст=42, Сб=52, Шг1=17.5, Шг2=22, Цг=10, Шс=18, Дпл=13,
#                 Дтп=47, Вг=29, Вг2=10.5, Впкп=28, Дтс=44, ВПРЗ=21, Впк=44.5, Вб=20,
#                 Гт1=3.5, Гт1i=12, Гт2=1, Гт2i=6, Гтс1=3, Гтс1i=13, Гтс2=5.5, Гтс2i=12,
#                 Гб=2.5, Гбi=14, Гб2=2.5, Гб2i=10, Пг=3, Пт=1, Пб=2, Пшгор=0.5, Пгпр=2.5,
#                 Типфигуры='нормальная')


main_dic.setdefault('Пг', 3)
main_dic.setdefault('Пт', 1)
main_dic.setdefault('Пб', 2)
main_dic.setdefault('Пшгор', 0.5)
main_dic.setdefault('Пгпр', 2.5)
#  Пшс = 30% от Пг = 0,9, Пшп = 20% от Пг = 0,6"
main_dic.setdefault('Пшс', (main_dic['Пг']*0.3))
main_dic.setdefault('Пшп', (main_dic['Пг']*0.2))
main_dic.setdefault('v', 0.01)                      # для отрезка

# отводы
if main_dic['Тип фигуры'] in ['нормальная', 'норм', 'normal', 'norm']:
    main_dic.setdefault('Олг', 0.5)
    main_dic.setdefault('Олт', 1)
    main_dic.setdefault('Олб', 0)
    main_dic.setdefault('Олс', 0.5)
    main_dic.setdefault('Выт', 2)
    main_dic.setdefault('Пос', 0.5)
elif main_dic['Тип фигуры'] in ['сутулая', 'сут', 'сутул', 'stooped', 'stoop']:
    main_dic.setdefault('Олг', 1)
    main_dic.setdefault('Олт', 0.5)
    main_dic.setdefault('Олб', 1)
    main_dic.setdefault('Олс', 0)
    main_dic.setdefault('Выт', 2.5)
    main_dic.setdefault('Пос', 0.5)
elif main_dic['Тип фигуры'] in ['перегиб', 'пер', 'перегибистая', 'straight', 'str']:
    main_dic.setdefault('Олг', 0)
    main_dic.setdefault('Олт', 0)
    main_dic.setdefault('Олб', 0)
    main_dic.setdefault('Олс', 1)
    main_dic.setdefault('Выт', 1.5)
    main_dic.setdefault('Пос', 0.5)
else:
    main_dic.setdefault('Олг', 0.5)
    main_dic.setdefault('Олт', 1)
    main_dic.setdefault('Олб', 0)
    main_dic.setdefault('Олс', 0.5)
    main_dic.setdefault('Выт', 2)
    main_dic.setdefault('Пос', 0.5)


i = 1
grad = 57.29577951308233

text = ['A0', 'T',    # Дтс
        'A0', 'G',    # ВПРЗ
        'T', 'B',     # Вб
        'A0', 'U',    # Дтс*0.4
        'A0', "A0i",  # Олг
        'T', "Ti",    # Олт
        'B', 'Bi',    # Олб
        'G', 'v',     # 0.1
        'g', 'G3',    # Сг+Пг
        'G3', 'v',    # 0.1
        'A0i', 'v',   # 0.1
        'Ti', 'v',    # 0.1
        'B', 'v',     # 0.1
        'A0', 'a',    # Шс+Пшс
        'a', 'v',     # 0.1
        'a1', 'a2',   # Шг2+Пшп
        'a2', 'v',    # 0.1
        'G2', 'v',    # 0.1
        'G1', 'G1i',  # Пгпр
        'G4', 'G4i',  # Пгпр
        'A0i', 'A2',  # Сш/3+0.5+Пшгор
        'A1', 'v',    # 0.1
        'P', 'v',     # 0.1
        'G5', 'v',    # 0.1
        'b1', 'v',    # 0.1
        'gi', 'v']    # 0.1

lenght = [main_dic['Дтс'], main_dic['ВПРЗ'], main_dic['Вб'], main_dic['Дтс']*0.4, main_dic['Олг'], main_dic['Олт'],
          main_dic['Олб'], main_dic['v'], main_dic['Сг'] + main_dic['Пг'], main_dic['v'], main_dic['v'],
          main_dic['v'], main_dic['v'], main_dic['Шс'] + main_dic['Пшс'], main_dic['v'],
          main_dic['Шг2'] + main_dic['Пшп'], main_dic['v'], main_dic['v'], main_dic['Пгпр'],
          main_dic['Пгпр'], main_dic['Сш'] / 3 + 0.5 + main_dic['Пшгор'],
          main_dic['v'], main_dic['v'], main_dic['v'], main_dic['v'], main_dic['v']]

angle = [270, 270, 270, 270, 0, 0, 0, 0, 0, 90, 0, 0, 0, 0, 270, 180, 270, 270, 270,
         270, 0, 180, 180, 270, 270, 270]  # углы поворота

nam_crd = dict(A0=[4, 140, 0])  # координаты начальной точки построения чертежа
n = copy.copy(lenght)
nam_obj = dict()

nam_crd.setdefault('c1', 0)
nam_crd.setdefault('c2', 0)
nam_crd.setdefault('c3', 0)
nam_crd.setdefault('Ar', 0)


for _ in range(8):
    main_fun()
to_nam_obj('A0i', 'U', 'Ti', 'Bi')

inters_to_dict('g', 'Gv', 'UTi', 3)
for _ in range(3):
    main_fun()
inters_to_dict('Ti', 'G3v', 'A0iv', 3)
for _ in range(2):
    main_fun()
inters_to_dict('T3', 'G3v', 'Tiv', 3)
inters_to_dict('B3', 'G3v', 'Bv', 3)
for _ in range(2):
    main_fun()
inters_to_dict('G1', 'av', 'Gv', 3)
inters_to_dict('a1', 'G3v', 'A0iv', 3)
for _ in range(2):
    main_fun()
inters_to_dict('G4', 'Gv', 'a2v', 3)
to_nam_obj('G1', 'G4')
nam_crd.setdefault('G2', (mid_po('G1G4')))
for _ in range(1):
    main_fun()
inters_to_dict('T2', 'G2v', 'Tiv', 3)
inters_to_dict('B2', 'G2v', 'Bv', 3)
for _ in range(3):
    main_fun()
to_nam_obj('A0i', 'A2')
nam_crd.setdefault('A1', coord_p(((nam_obj['A0iA2']).Length / 3), 270, nam_crd['A2']))
main_fun()
inters_to_dict('A', 'A1v', 'A0iU', 3)

nam_obj.setdefault('A2r', add_arc('A2', main_dic['Дпл'] + main_dic['Выт'] + main_dic['Пос'], 0, 1.5))
nam_obj.setdefault('Tir', add_arc('Ti', main_dic['Впк'], 0, 1.5))
inters_to_dict('P', 'A2r', 'Tir', 1)
nam_obj['A2r'].delete()
nam_obj['Tir'].delete()

to_nam_obj('A2', 'P')

nam_crd.setdefault('b', coord_p(4.5, nam_obj['A2P'].Angle*grad, nam_crd['A2']))
nam_crd.setdefault('b1', coord_p(7, nam_obj['A0iU'].Angle*grad, nam_crd['b']))
nam_crd.setdefault('b0', coord_p(2, nam_obj['A2P'].Angle*grad, nam_crd['b']))
to_nam_obj('b', 'b1')
to_nam_obj('b1', 'b0')
nam_crd.setdefault('b2', coord_p(nam_obj['bb1'].Length, nam_obj['b1b0'].Angle*grad, nam_crd['b1']))

to_nam_obj('b1', 'b2')
to_nam_obj('b2', 'P')

for _ in range(1):
    main_fun()
to_nam_obj('a', 'G1')
inters_to_dict('P1', 'Pv', 'aG1', 3)
to_nam_obj('P1', 'G1i')
nam_crd.setdefault('P2', coord_p(nam_obj['P1G1i'].Length / 3 + 2, 90, nam_crd['G1i']))

to_nam_obj('G1i', 'P1')
to_nam_obj('G1i', 'G4i')
nam_crd.setdefault('O1', coord_p(nam_obj['G1iG4i'].Length * 0.2 + 0.3, ang_bis('G1iP1', 'G1iG4i'), nam_crd['G1i']))

to_nam_obj('P2', 'P')

inters_to_dict('G2i', 'G2v', 'G1iG4i', 3)

nam_crd.setdefault('A3', coord_p(main_dic['Дтп'], 90, nam_crd['T3']))
to_nam_obj('A3', 'T3')

nam_crd.setdefault('A3i', coord_p(main_dic['Олс'], 180, nam_crd['A3']))
to_nam_obj('A3', 'A3i')
to_nam_obj('A3i', 'G3')

nam_crd.setdefault('A4', coord_p(main_dic['Сш'] / 3 + main_dic['Пшгор'], 180, nam_crd['A3i']))
to_nam_obj('A3i', 'A4')

nam_crd.setdefault('A5', coord_p(nam_obj['A3iA4'].Length + 1, nam_obj['A3iG3'].Angle*grad, nam_crd['A3i']))

nam_crd.setdefault('A4', coord_p(main_dic['Сш'] / 3 + main_dic['Пшгор'], 180, nam_crd['A3i']))
to_nam_obj('A3i', 'A4')

to_nam_obj('G', 'G3')
to_nam_obj('g', 'G3')
to_nam_obj('G', 'g')

nam_crd.setdefault('G5', coord_p(main_dic['Цг'], 180, nam_crd['G3']))
to_nam_obj('G3', 'G5')

to_nam_obj('g', 'G1')
nam_crd.setdefault('gi', (mid_po('gG1')))

for _ in range(3):
    main_fun()
to_nam_obj('Ti', 'T3')
inters_to_dict('T4', 'G5v', 'TiT3', 3)
to_nam_obj('G5', 'T4')
inters_to_dict('T1', 'giv', 'TiT3', 3)
to_nam_obj('gi', 'T1')

nam_obj.setdefault('A4r', add_arc('A4', main_dic['Вг'], 4.5, 4.8))
inters_to_dict('G6', 'A4r', 'G5T4', 1)
nam_obj['A4r'].delete()
to_nam_obj('G6', 'A4')

nam_obj.setdefault('G6r', add_arc('G6', main_dic['Вг2'], 1, 2))
inters_to_dict('G7', 'G6r', 'G6A4', 1)
nam_obj['G6r'].delete()

nam_crd.setdefault('G8', po_at_arc('G6', 'Вг2', 'G6A4', 'Шг2', 'Шг1'))

to_nam_obj('G6', 'G8')
nam_crd.setdefault('G9', coord_p(nam_obj['G6A4'].Length, nam_obj['G6G8'].Angle*grad, nam_crd['G6']))
nam_obj['G6G8'].delete()
to_nam_obj('G6', 'G9')

nam_obj.setdefault('G9r', add_arc('G9', main_dic['Дпл'], 3, 5))
nam_obj.setdefault('G6ri', add_arc('G6', main_dic['Впкп'], 1, 3))
inters_to_dict('P5', 'G9r', 'G6ri', 0)
nam_obj['G9r'].delete()
nam_obj['G6ri'].delete()
to_nam_obj('P5', 'G9')

to_nam_obj('G4i', 'a2')
nam_crd.setdefault('P4', coord_p((nam_obj['G1iP1'].Length - 1) / 3, nam_obj['G4ia2'].Angle*grad, nam_crd['G4i']))
to_nam_obj('P5', 'P4')

to_nam_obj('G4i', 'G4')
to_nam_obj('G4i', 'G2i')
nam_crd.setdefault('O2', coord_p(nam_obj['G1iG4i'].Length * 0.2, ang_bis('G4iG4', 'G4iG2i'), nam_crd['G4i']))
to_nam_obj('G4i', 'O2')

to_nam_obj('B', 'B3')
inters_to_dict('B1', 'giv', 'BB3', 3)
inters_to_dict('B4', 'G5v', 'BB3', 3)

nam_crd.setdefault('gi', (mid_po('gG1')))


adds_to_vit('Гтс1', 'Гб1')

# вытачки
vit('T1', 'Гтс1i', 'Гтс2i', 'Гтс1', 'Гтс2', 'B1')
vit('T2', 'Гб1i', 'Гб2i', 'Гб1', 'Гб2', 'B2')
vit('T4', 'Гт1i', 'Гт2i', 'Гт1', 'Гт2', 'B4')


# скругления
main_arc('G2i', 'O2', 'P4')

# добираю точки для построения дуг
to_nam_obj('G1i', 'G2i')
nam_crd.setdefault('Ga', coord_p(nam_obj['G1iG2i'].Length, nam_obj['G1iP1'].Angle*grad, nam_crd['G1i']))
main_arc('Ga', 'O1', 'G2i')

to_nam_obj('A1', 'A')
to_nam_obj('A1', 'A2')
nam_crd.setdefault('Aa', coord_p(nam_obj['A1A2'].Length, nam_obj['A1A'].Angle*grad, nam_crd['A1']))

to_nam_obj('Aa', 'A2')
nam_crd.setdefault('Ar1', coord_p(nam_obj['AaA2'].Length*0.21, nam_obj['AaA2'].Angle*grad-90, mid_po('AaA2')))
main_arc('Aa', 'Ar1', 'A2')

to_nam_obj('A4', 'A5')
nam_crd.setdefault('Ar2', coord_p(nam_obj['A4A5'].Length*0.145, nam_obj['A4A5'].Angle*grad-90, mid_po('A4A5')))
main_arc('A5', 'Ar2', 'A4')

nam_crd.setdefault('Ar3', coord_p(0.7, nam_obj['P5P4'].Angle*grad+90, mid_po('P5P4')))
main_arc('P5', 'Ar3', 'P4')

to_nam_obj('P', 'P2')
nam_crd.setdefault('Ar4', coord_p(0.3, nam_obj['P2P'].Angle*grad+90, mid_po('PP2')))
main_arc('P2', 'Ar4', 'P')

to_nam_obj('T3', 'B3')

create_third_po_at_line('A0iU', 0.2, -90)
main_arc('A', 'Ar', 'U')

# проверяет правильность снятых мерок глубин
check_measure()


# добавляет номера точек в рисунок
# add_pos_names()

# print('\n'+'nam_crd:', list(nam_crd.keys()))
# print('\n'+'nam_obj:', list(nam_obj.keys()))


