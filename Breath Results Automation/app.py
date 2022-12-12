from mail import *
from papa import *
from tkinter import *
from tkinter import filedialog
from os import walk


class MyWindow:
    def __init__(self, win):
        #self.b1 = Button(win, text='Choose Directory', command=self.chooseDir)
        self.b2 = Button(win, text='Download XLSX', command=self.scan)
        self.b6 = Button(win, text='Run SIBO Automator', command=self.runSibo)
        self.b7 = Button(win, text='Kill Program', command=self.stop)

        #self.b1.place(x=100, y=50)
        self.b2.place(x=100, y=100)
        self.b6.place(x=100, y=300)
        self.b7.place(x=100, y=400)

    """def chooseDir(self):
        global attachment_dir
        attachment_dir = filedialog.askdirectory()"""
    def scan(self):
        scan()
    def runSibo(self):
        filenames = next(walk(attachment_dir), (None, None, []))[2]
        for file in filenames:
            sibo(f"{attachment_dir}{file}")
    def stop(self):
        exit(1)

def main():
    window=Tk()
    mywin = MyWindow(window)
    window.title('Dr Rosner Office')
    window.geometry("400x450")
    window.mainloop()

if __name__ == "__main__":
    main()
