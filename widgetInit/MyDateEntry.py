import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from datetime import datetime

root = tk.Tk()

# 테마를 clam으로 기본 설정
style = ttk.Style(root)
style.theme_use('clam')

# 현재 시스템 날짜 (date)
sYear = int(datetime.today().strftime('%Y'))
sMonth = int(datetime.today().strftime('%m'))
sDay = int(datetime.today().strftime('%d'))

class MyDateEntry(DateEntry):
    def __init__(self, master=None, **kw):
        DateEntry.__init__(self, master=None, **kw)
        # add black border around drop-down calendar
        self._top_cal.configure(bg='black', bd=1)
        # add label displaying today's date below
        tk.Label(self._top_cal, bg='gray90', anchor='w',
                 text='Today: %s' % datetime.today().strftime('%x')).pack(fill='x')

# create the entry and configure the calendar colors
de = MyDateEntry(root, year=sYear, month=sMonth, day=sDay,
                 selectbackground='gray80',
                 selectforeground='black',
                 normalbackground='white',
                 normalforeground='black',
                 background='gray90',
                 foreground='black',
                 bordercolor='gray90',
                 othermonthforeground='gray50',
                 othermonthbackground='white',
                 othermonthweforeground='gray50',
                 othermonthwebackground='white',
                 weekendbackground='white',
                 weekendforeground='black',
                 headersbackground='white',
                 headersforeground='gray70')
de.pack()   # 위젯을 상위(parent)에 배치하기 전에 block 으로 구성한다.

# root.mainloop()