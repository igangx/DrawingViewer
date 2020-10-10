# -*- coding: utf-8 -*-
# Advanced zoom example. Like in Google Maps.
# It zooms only a tile, but not the whole image. So the zoomed tile occupies
# constant memory and not crams it with a huge resized image for the large zooms.

import os
import shutil
import tempfile
import time
import tkinter
import tkinter as tk
import winreg
from tkinter import *
from tkinter import ttk

import PIL
import fitz
import win32api
import win32con
import win32gui
import win32print
from PIL import Image, ImageTk


class AutoScrollbar(ttk.Scrollbar):
    """ A scrollbar that hides itself if it's not needed.
        Works only if you use the grid geometry manager """

    def set(self, lo, hi):
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            self.grid_remove()
        else:
            self.grid()
            ttk.Scrollbar.set(self, lo, hi)

    def pack(self, **kw):
        raise tk.TclError('Cannot use pack with this widget')

    def place(self, **kw):
        raise tk.TclError('Cannot use place with this widget')


class Zoom_Advanced(ttk.Frame):
    """ Advanced zoom of the image """

    def __init__(self, mainframe, path):
        global page_total_num, current_page

        ''' Initialize the main Frame '''
        ttk.Frame.__init__(self, master=mainframe)
        self.master.title('DrawingViewer')

        # 判断文件类型
        ext = os.path.splitext(path)[1]
        if (ext == ".tif") or (ext == ".TIF") or (ext == ".tiff") or (ext == ".TIFF"):
            file_type = 'tif'
        elif (ext == ".pdf") or (ext == ".PDF"):
            file_type = 'pdf'

        # Vertical and horizontal scrollbars for canvas
        vbar = AutoScrollbar(self.master, orient='vertical')
        hbar = AutoScrollbar(self.master, orient='horizontal')
        vbar.grid(row=0, column=2, sticky='ns')
        hbar.grid(row=1, column=0, sticky='we')
        # Create canvas and put image on it
        self.canvas = tk.Canvas(self.master, highlightthickness=0,
                                xscrollcommand=hbar.set, yscrollcommand=vbar.set, cursor='hand2')
        self.canvas.grid(row=0, column=0, sticky='nswe', columnspan=2)
        self.canvas.update()  # wait till canvas is created
        vbar.configure(command=self.scroll_y)  # bind scrollbars to the canvas
        hbar.configure(command=self.scroll_x)
        # Make the canvas expandable
        self.master.rowconfigure(0, weight=1)
        self.master.columnconfigure(0, weight=1)
        # Bind events to the Canvas
        self.canvas.bind('<Configure>', self.show_image)  # canvas is resized
        self.canvas.bind('<ButtonPress-1>', self.move_from)
        self.canvas.bind('<B1-Motion>', self.move_to)
        self.canvas.bind('<MouseWheel>', self.wheel)  # with Windows and MacOS, but not Linux
        self.canvas.bind('<Button-5>', self.wheel)  # only with Linux, wheel scroll down
        self.canvas.bind('<Button-4>', self.wheel)  # only with Linux, wheel scroll up
        if file_type == 'tif':
            self.image = PIL.Image.open(path).convert('RGB')  # open imag
            page_total_num = countTifPages(path)
        elif file_type == 'pdf':
            self.image = self.page_pdf(path, 0)
            page_total_num = countPdfPages(path)
        current_page = 0

        self.width, self.height = self.image.size
        self.imscale = 1.0  # scale for the canvaas image
        self.delta = 1.3  # zoom magnitude
        # Put image into container rectangle and use it to set proper coordinates to the image
        self.container = self.canvas.create_rectangle(0, 0, self.width, self.height, width=0)

        # 创建翻页按钮
        fm2 = tk.Frame(self.master)
        PrevPage = tkinter.Button(fm2, text="上一页", command=lambda: self.pageUp(path, file_type)).pack(side=LEFT)
        NextPage = tkinter.Button(fm2, text="下一页", command=lambda: self.pageDown(path, file_type)).pack(side=RIGHT)
        # 创建页码标签        
        self.textvar = tk.StringVar()
        label = tkinter.Label(fm2, textvariable=self.textvar).pack(side=BOTTOM)
        self.textvar.set(str(current_page + 1) + "/" + str(page_total_num))
        fm2.grid(row=2, column=0)

        # 创建一个打印按钮
        fm3 = tk.Frame(self.master)
        if file_type == 'tif':
            Bcom = tkinter.Button(fm3, text="打印", command=lambda: self.print_tif(path)).pack()
        elif file_type == 'pdf':
            Bcom = tkinter.Button(fm3, text="打印", command=lambda: self.print_pdf(path)).pack()
        fm3.grid(row=3, column=0)

        self.show_image()
        self.show_all()

    def scroll_y(self, *args, **kwargs):
        """ Scroll canvas vertically and redraw the image """
        self.canvas.yview(*args, **kwargs)  # scroll vertically
        self.show_image()  # redraw the image

    def scroll_x(self, *args, **kwargs):
        """ Scroll canvas horizontally and redraw the image """
        self.canvas.xview(*args, **kwargs)  # scroll horizontally
        self.show_image()  # redraw the image

    def move_from(self, event):
        """ Remember previous coordinates for scrolling with the mouse """
        self.canvas.scan_mark(event.x, event.y)

    def move_to(self, event):
        ''' Drag (move) canvas to the new position '''
        self.canvas.scan_dragto(event.x, event.y, gain=1)
        self.show_image()  # redraw the image

    def wheel(self, event):
        ''' Zoom with mouse wheel '''
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        bbox = self.canvas.bbox(self.container)  # get image area
        if bbox[0] < x < bbox[2] and bbox[1] < y < bbox[3]:
            pass  # Ok! Inside the image
        else:
            return  # zoom only inside image area
        scale = 1.0
        # Respond to Linux (event.num) or Windows (event.delta) wheel event
        if event.num == 5 or event.delta == -120:  # scroll down
            i = min(self.width, self.height)
            if int(i * self.imscale) < 30: return  # image is less than 30 pixels
            self.imscale /= self.delta
            scale /= self.delta
        if event.num == 4 or event.delta == 120:  # scroll up
            i = min(self.canvas.winfo_width(), self.canvas.winfo_height())
            if i < self.imscale: return  # 1 pixel is bigger than the visible area
            self.imscale *= self.delta
            scale *= self.delta
        self.canvas.scale('all', x, y, scale, scale)  # rescale all canvas objects

        self.show_image()

    def show_image(self, event=None):
        """ Show image on the Canvas """
        bbox1 = self.canvas.bbox(self.container)  # get image area
        # Remove 1 pixel shift at the sides of the bbox1
        bbox1 = (bbox1[0] + 1, bbox1[1] + 1, bbox1[2] - 1, bbox1[3] - 1)
        bbox2 = (self.canvas.canvasx(0),  # get visible area of the canvas
                 self.canvas.canvasy(0),
                 self.canvas.canvasx(self.canvas.winfo_width()),
                 self.canvas.canvasy(self.canvas.winfo_height()))
        bbox = [min(bbox1[0], bbox2[0]), min(bbox1[1], bbox2[1]),  # get scroll region box
                max(bbox1[2], bbox2[2]), max(bbox1[3], bbox2[3])]
        if bbox[0] == bbox2[0] and bbox[2] == bbox2[2]:  # whole image in the visible area
            bbox[0] = bbox1[0]
            bbox[2] = bbox1[2]
        if bbox[1] == bbox2[1] and bbox[3] == bbox2[3]:  # whole image in the visible area
            bbox[1] = bbox1[1]
            bbox[3] = bbox1[3]
        self.canvas.configure(scrollregion=bbox)  # set scroll region
        x1 = max(bbox2[0] - bbox1[0], 0)  # get coordinates (x1,y1,x2,y2) of the image tile
        y1 = max(bbox2[1] - bbox1[1], 0)
        x2 = min(bbox2[2], bbox1[2]) - bbox1[0]
        y2 = min(bbox2[3], bbox1[3]) - bbox1[1]
        if int(x2 - x1) > 0 and int(y2 - y1) > 0:  # show image if it in the visible area
            x = min(int(x2 / self.imscale), self.width)  # sometimes it is larger on 1 pixel...
            y = min(int(y2 / self.imscale), self.height)  # ...and sometimes not
            image = self.image.crop((int(x1 / self.imscale), int(y1 / self.imscale), x, y))
            imagetk = ImageTk.PhotoImage(image.resize((int(x2 - x1), int(y2 - y1))))
            imageid = self.canvas.create_image(max(bbox2[0], bbox1[0]), max(bbox2[1], bbox1[1]),
                                               anchor='nw', image=imagetk)
            self.canvas.lower(imageid)  # set image into background
            self.canvas.imagetk = imagetk  # keep an extra reference to prevent garbage-collection
        self.textvar.set(str(current_page + 1) + "/" + str(page_total_num))

        # 显示比例
        view_scale = str(int(self.imscale * 100)) + "%"
        #print("self.imscale=%s" % view_scale)

    def page_tif(self, path):
        img = PIL.Image.open(path)
        img.seek(current_page)
        return img

    def page_pdf(self, path, i):
        doc = fitz.open(path)
        page = doc[i]
        rotate = int(0)
        zoom_x = 4.0
        zoom_y = 4.0
        trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotate)
        pix = page.getPixmap(matrix=trans, alpha=False)
        img = PIL.Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img

    def print_tif(self, path):
        filename = path
        win32api.ShellExecute(
            0,
            "print",
            filename,
            #
            # If this is None, the default printer will
            # be used anyway.
            #
            '/d:"%s"' % win32print.GetDefaultPrinter(),
            ".",
            0
        )

    def print_pdf(self, file):
        key = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r".pdf")
        name, value, type = winreg.EnumValue(key, 0)
        key2 = winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, value + "\\shell\\open\\command")
        name, path, type = winreg.EnumValue(key2, 0)
        cmd = path[:-5]
        win32api.ShellExecute(0, '', cmd, " /P " + file, '', 0)
        self.foo()

    def pageUp(self, path, f_type):
        global current_page
        if current_page == 0:
            current_page = page_total_num - 1
        else:
            current_page = current_page - 1
        if f_type == 'tif':
            self.image = self.page_tif(path)
        elif f_type == 'pdf':
            self.image = self.page_pdf(path, current_page)
        self.show_image()

    def pageDown(self, path, f_type):
        global current_page
        if current_page < page_total_num - 1:
            current_page = current_page + 1
        else:
            current_page = 0
        if f_type == 'tif':
            self.image = self.page_tif(path)
        elif f_type == 'pdf':
            self.image = self.page_pdf(path, current_page)
        self.show_image()

    def foo(self):
        time.sleep(1)
        hwnd_jp = win32gui.FindWindow(None, "印刷")
        hwnd_cn = win32gui.FindWindow(None, "打印")
        # print("1.hWnd1=%s" % hWnd1)
        if hwnd_jp == 0 and hwnd_cn == 0:
            time.sleep(5)
        while hwnd_jp or hwnd_cn:
            time.sleep(0.1)
            hwnd_jp = win32gui.FindWindow(None, "印刷")
            hwnd_cn = win32gui.FindWindow(None, "打印")
            # print("2.hWnd1=%s" % hWnd1)
        i = 1
        # print("print closed")
        time.sleep(1)
        hwnd_reader = win32gui.FindWindow(None, "Adobe Reader")
        hwnd_dc = win32gui.FindWindow(None, "Adobe Acrobat Reader DC")
        # print("1.hWnd2=%s" % hWnd2)
        if hwnd_reader == 0 and hwnd_dc == 0:
            time.sleep(5)
        while 1:
            hwnd_reader = win32gui.FindWindow(None, "Adobe Reader")
            hwnd_dc = win32gui.FindWindow(None, "Adobe Acrobat Reader DC")
            # print("2.hWnd2=%s" % hWnd2)
            if hwnd_reader or hwnd_dc:
                win32gui.PostMessage(hwnd_reader, win32con.WM_CLOSE, 0, 0)
                win32gui.PostMessage(hwnd_dc, win32con.WM_CLOSE, 0, 0)
                break
            i = i + 1
            if i == 10:
                break

    def show_all(self):
        """ Zoom with mouse wheel """
        # x = self.canvas.canvasx(self.canvas.winfo_width()/2)
        # y = self.canvas.canvasy(self.canvas.winfo_height()/2)
        bbox = self.canvas.bbox(self.container)
        image_size = (bbox[2] - bbox[0], bbox[3] - bbox[1])
        # delta = min(self.canvas.winfo_width()/self.image.size[0], self.canvas.winfo_height()/self.image.size[1])
        canvas_width = 1000
        canvas_height = 716
        delta = min(canvas_width / self.image.size[0], canvas_height / self.image.size[1])
        ca = self.canvas.winfo_width()
        im = self.image.size[0]
        a = self.image.size[0]
        b = a * delta
        k = (canvas_width - b) / 2
        x = a * k / (a - b - k)
        y = 0
        scale = 1.0
        self.imscale *= delta
        scale *= delta
        self.canvas.scale('all', x, y, scale, scale)  # rescale all canvas objects
        self.show_image()


def countTifPages(file):
    # 统计tif页数
    img = PIL.Image.open(file)
    i = 1
    try:
        img.seek(1)
        for i in range(30):
            try:
                img.seek(i)
            except EOFError:
                break
    except EOFError:
        1 + 1
    return i


def countPdfPages(the_file):
    # 统计pdf页数
    pdf_doc = fitz.open(the_file)
    i = pdf_doc.pageCount
    return i


def loadFile(patha):
    path1 = patha
    root = tk.Tk()
    root.geometry("1024x768")
    app = Zoom_Advanced(root, path=path1)
    root.mainloop()


if __name__ == '__main__':
    PIL.Image.MAX_IMAGE_PIXELS = 1000000000
    tmp = ''
    with tempfile.TemporaryDirectory() as tmpdir:
        path_ori = sys.argv[1]
        fileName = os.path.split(path_ori)[1]
        extname = os.path.splitext(path_ori)[1]
        shutil.move(path_ori, tmpdir)
        file = os.path.join(tmpdir, fileName)
        tmpfile = os.path.join(tmpdir, "82218868" + extname)
        os.rename(file, tmpfile)
        loadFile(tmpfile)

#    top.mainloop()
