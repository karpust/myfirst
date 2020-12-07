"""использовать декоратор стоит если код повторяется и
если что-то можно вынести за скобки"""
import math
import sys
import os
from datetime import datetime
'------------------------------------------------------------------------------------------'
# """пример декоратора
# оборачивание функции которая принимает аргументы"""
#
# def timeit(func):
#     def wrapper(*args, **kwargs):
#         start = datetime.now()
#         result = func(*args, **kwargs)
#         print(datetime.now() - start)
#         return result
#     return wrapper
#
#
# @timeit
# def one(n):
#     l = [x for x in range(n) if x % 2 == 0]
#     return l
#
#
# l1 = timeit(one)(1000)  # => wrapper(1000) => one(1000)
# print(l1)
#
# l1 = timeit(one)
# print(type(l1), l1.__name__)
'------------------------------------------------------------------------------------------'
# """пример декоратора который принимает аргументы и
#  функции которая принимает аргументы"""
#
#
# def timeit(arg):
#     print(arg)
#
#     def outer(func):
#         def wrapper(*args, **kwargs):
#             start = datetime.now()
#             result = func(*args, **kwargs)
#             print(datetime.now() - start)
#             return result
#         return wrapper
#     return outer
#
#
# @timeit('name')
# def one(n):
#     l = [x for x in range(n) if x % 2 == 0]
#     return l
#
# # l1 = one(10)
# l1 = timeit('name1')(one)(10)
# print(l1)
'------------------------------------------------------------------------------------------'


# def makebold(fn):
#     def wrapped():
#         return "<b>" + fn() + "</b>"
#
#     return wrapped
#
#
# def makeitalic(fn):
#     def wrapped():
#         return "<i>" + fn() + "</i>"
#
#     return wrapped
#
#
# @makebold
# # @makeitalic
# def hello():
#     return "hello habr"
#
#
# hello = hello()
#
# print(hello)  # выведет <b><i>hello habr</i></b>
'-----------------------------------------------------------------------------------------------'
# a = [1, 2, 3, 4, 6, 7, 99, 88, 999]
# max = 0
# for i in a:
#     if i > max:
#         max = i
# print(max)

# import functools
#
#
# def greet(greeting, name1, name):
#     print('%s, %s %s!' % (greeting, name1, name))
#
#
# greet = functools.partial(greet, 'превед', 'негодяй')
# greet('красавчик')

#
# l = ['Коля', 'Маша', 'Витя']
#
# def people(*args):
#    return args
#
# print(people(l))
#
# help(math)

import sys
import copy
import math
import win32com.client
from pythoncom import VT_R8, VT_ARRAY, VT_DISPATCH, VT_BSTR, VT_I2, VT_VARIANT

sys.path.append("C:\\Users\\UserGeo\\AppData\Local\Programs\Python\Python37-32\Lib\site-packages\win32")
sys.path.append("C:\\Users\\UserGeo\\AppData\Local\Programs\Python\Python37-32\Lib\site-packages")
sys.path.append("C:\\Users\\UserGeo\\AppData\Local\Programs\Python\Python37-32\Lib\site-packages\win32\lib")
sys.path.append("C:\\Users\\UserGeo\\AppData\Local\Programs\Python\Python37-32\Lib\site-packages\Pythowin")
