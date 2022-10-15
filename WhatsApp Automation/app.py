from bdikot import *
from maacabi import *

from tkinter import *

class MyWindow:
    def __init__(self, win):
        self.btn1 = Button(win, text='Ramat Hahayal')
        self.btn2 = Button(win, text='Raanana')
        self.btn3 = Button(win, text='Ramat Hahayal & Rannana')  # , command=bkidot())
        self.btn4 = Button(win, text='Maccabi Office')  # , command=office())
        self.btn5 = Button(win, text='Single Reminder Ramat hahayal')
        self.btn6 = Button(win, text='Single Reminder Raanana')
        self.btn7 = Button(win, text='Kill Program')

        self.b1 = Button(win, text='Ramat Hahayal', command=self.ramatHahayal)
        self.b2 = Button(win, text='Raanana', command=self.raanana)
        self.b3 = Button(win, text='Ramat Hahayal & Raanana', command=self.bdikot)
        self.b4 = Button(win, text='Maacabi', command=self.office)
        self.b5 = Button(win, text='Single Reminder Ramat hahayal', command=self.singleReminderRamatHahayal)
        self.b6 = Button(win, text='Single Reminder Raanana', command=self.singleReminderRaanana)
        self.b7 = Button(win, text='Kill Program', command=self.stop)


        self.b1.place(x=100, y=50)
        self.b2.place(x=100, y=100)
        self.b3.place(x=100, y=150)
        self.b4.place(x=100, y=200)
        self.b5.place(x=100, y=250)
        self.b6.place(x=100, y=300)
        self.b7.place(x=100, y=400)

    def ramatHahayal(self):
        ramatHahyal()
    def raanana(self):
        raanana()
    def bdikot(self):
        bdikot()
    def office(self):
        office()
    def singleReminderRamatHahayal(self):
        singleReminderRamatHahayal()
    def singleReminderRaanana(self):
        singleReminderRaanana()
    def stop(self):
        exit(1)

def main():
    window=Tk()
    mywin=MyWindow(window)
    window.title('Dr Rosner Office')
    window.geometry("400x450")
    window.mainloop()


if __name__ == "__main__":
    main()
