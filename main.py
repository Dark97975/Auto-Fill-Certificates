import tkinter as tk
from tkinter import filedialog,ttk
from tkinter import *
import pandas as pd
import os
from tkinter.filedialog import askopenfile
import PIL
from PIL import ImageTk,Image,ImageDraw,ImageFont
import math

template_path = ''
data_path = ''
save_folder_path = ''
font_path = 'C:\Windows\Fonts'

font_name = []
pos = []
pos_arr = []
align_arr = []
data_arr = []
able_key = []

pos_dict = {}
align_dict = {}
img_dict = {}
font_dict = {}
size_dict = {}
color_dict = {}

full_fontname = {}

size = 20
diagnal_small = 596
mid_x = 400
mid_y = 225

# function when start a new project
def file_choice(selected):
    global img_first,img_dict,pos_dict,align_dict,able_key,data_arr,color_dict
    s_choice.set('File')
    if selected == 'New':
        if find_template_path() == '':
            return
        if find_data_path() == '':
            return
        img_first = Image.open(template_path)
        img_dict = dict.fromkeys(list(data),img_first)
        color_dict = dict.fromkeys(list(data), (255,255,255,0))
        find_ratio()
        update_pic(img_first)

    elif selected == 'Exit':
        win.destroy()

    elif selected == 'Clear':
        img = Image.open(template_path)
        submit['state'] = DISABLED
        img_dict = dict.fromkeys(list(data), img)
        able_key = []
        data_arr = []
        '''pos_dict = dict.fromkeys(list(data), [0, 0])
        align_dict = dict.fromkeys(list(data), 'mm')
        font_dict = dict.fromkeys(list(data), 'font/SuperspaceBold.otf')
        size_dict = dict.fromkeys(list(data), 20)'''
        update_pic(img)
# create an array which contain all data
def create_data_arr():
    global data_arr
    for i in columns:
        if i in able_key:
            for j in data[i].tolist():
                data_arr.append(j)
# find necessary path
def find_data_path():
    global data_path,data,columns
    file = filedialog.askopenfile(mode='r',filetypes =[('*','.xlsx')],title = 'Choose Data File')
    if file:
        data_path = os.path.abspath(file.name)
        data = pd.read_excel(data_path)
        columns = list(data)  # make a list of columns name
        column_combo['values'] = columns

def find_template_path():
    global template_path
    file = filedialog.askopenfile(mode='r', filetypes=[('*', '.png .jpg .jpeg')],title = 'Choose Template')
    if file:
        template_path = os.path.abspath(file.name)
# using for update the display pic to other
def update_pic(img):
    img_re_size = img.resize((487,344), Image.ANTIALIAS)
    # change the image that display right now, have to use imagetk.photoimage
    img2 = ImageTk.PhotoImage(img_re_size)
    display.configure(image=img2)
    display.image = img2
    display.place(x = 157,y = 38)
# find the ratio for scaling the font size
def find_ratio():
    global ratio,width,height
    img = Image.open(template_path)
    # scaling coord
    width, height = img.size
    diagnal = math.sqrt(width ** 2+height ** 2)
    ratio = diagnal / diagnal_small
# choose a spot and collect necessary data
def spot(event):
    global pos,size,img_dict,key,using_font,able_key,color
    if display.cget('image') != 'pyimage1' and column_combo.get() != '' \
       and align_combo.get() != '' and font_combo.get() != '':
        submit['state'] = NORMAL
        key = column_combo.get()
        if not key in able_key:
            able_key.append(key)
        # get clicked coord and scale it to be fit for real image
        x, y = event.x, event.y
        # scaling coord
        scaled_x = x/(487/width)
        scaled_y = y/(344/height)
        # set position to scaled position
        pos = [scaled_x,scaled_y]
        using_font = f'{font_path}\{full_fontname[font_combo.get()]}'
        if size_combo.get().isdigit():
            size = int(size_combo.get())
        place_font = ImageFont.truetype(font=using_font, size=size)
        # open previous image
        img = img_dict[column_combo.get()]
        color = set_color()
        set_dict(img)
        # write lines on previous image
        img = img.resize((487, 344), Image.ANTIALIAS)
        lined = ImageDraw.Draw(img)
        lined.line((x, 0, x, 450), fill='#ff0000')
        lined.line((0, y, 800, y), fill='#ff0000')
        # write text on previous image
        lined.text((x, y), column_combo.get(), color, font=place_font, anchor=align)
        # change the display image
        img2 = ImageTk.PhotoImage(img)
        display.configure(image=img2)
        display.image = img2
        display.place(x=157, y=38)
        # set current image to previous image
        img_dict[column_combo.get()] = img_first
# save the information in a form of dictionary with columns of excel file as a key
def set_dict(img):
    global align#, align_dict
    align = align_combo.get()
    if len(align) > 1:
        if not 'baseline' in align:
            align = align[0] + align[align.find('-') + 1]
        else:
            align = align[0] + 's'
    if column_combo.get() in list(data):
        pos_dict[key] = pos
        img_dict[key] = img
        align_dict[key] = align
        font_dict[key] = using_font
        size_dict[key] = size
        color_dict[key] = color
# create and save image
def draw():
    save_folder_path = filedialog.askdirectory(title='Choose Folder For Saving')
    create_data_arr()
    distance = int(len(data_arr)/len(able_key))
    for i in range(0,distance,1):
        img = Image.open(template_path)
        result = ImageDraw.Draw(img)
        for j in range(i,len(able_key)*distance,distance):
            run_key = able_key[j//distance]
            size = size_dict[run_key]
            use_color = color_dict[run_key]
            font = ImageFont.truetype(font=font_dict[run_key], size=int(size * ratio))
            result.text( (int(pos_dict[run_key][0]), int(pos_dict[run_key][1])), str(data_arr[j])
                        ,use_color, font=font, anchor=align_dict[run_key])
        img.save(f'{save_folder_path}/{i}.png')
        print(f'save {i} successfully')
    print('done!')
# collect the font from device and create font list with no format for display
# and create a dictionary of font name with format for usage/write text on image
def get_font_name():
    global font_name
    for file in os.listdir(font_path):
        if not '.fon' in file:
            this_font = file[0:file.find('.')-1]
            font_name.append(this_font)
            full_fontname[this_font] = file
def set_color():
    if color_combo.get() == 'white':
        return (255,255,255,0)
    elif color_combo.get() == 'black':
        return (0,0,0,255)
    elif color_combo.get() == 'red':
        return (255,0,0,255)
    elif color_combo.get() == 'green':
        return (0,255,0,255)
    elif color_combo.get() == 'blue':
        return (0,0,255,255)

win = tk.Tk()

win.geometry("800x450")
win.resizable(width=False, height=False)
win.title('Auto Fill Certificates')
get_font_name()
# place mockup image
emty_img = Image.open('emty.jpg')
emty_img = emty_img.resize((487,344), Image.ANTIALIAS)
emty_img = ImageTk.PhotoImage(emty_img)
display = Label(image=emty_img)
display.place(x = 157,y = 38)

# create  abd place label
column_label = Label(win,text = 'Selected Column : ',anchor = 'w')
align_label = Label(win,text = 'Alignment : ',anchor = 'w')
font_label = Label(win,text = 'Font : ',anchor = 'w')
size_label = Label(win,text = 'Text Size : ',anchor = 'w')
color_label = Label(win,text = 'Color : ',anchor = 'w')

column_label.place(x = 649,y = 38,height = 25,width = 146)
align_label.place(x = 649,y = 98,height = 25,width = 146)
font_label.place(x = 649,y = 158,height = 25,width = 146)
size_label.place(x = 649,y = 218,height = 25,width = 146)
color_label.place(x = 649,y = 278,height = 25,width = 146)


# create dropdown menu
s_choice = StringVar(win)
s_choice.set('File')
drop = OptionMenu(win,s_choice,'New','Clear','Exit',command=file_choice)
drop.place(x = 0,y = 0,width = 70)

# combo box
column_combo = ttk.Combobox(win)
align_combo = ttk.Combobox(win)
size_combo = ttk.Combobox(win)
font_combo = ttk.Combobox(win)
color_combo = ttk.Combobox(win)

column_combo['values'] = 'Emty'
align_combo['values'] = ['left-top','left-middle','left-baseline',
                       'middle-top','middle-middle','middle-baseline',
                       'right-top','right-middle','right-baseline']
size_combo['values'] = [6,8,10,12,14,18,24,36,40,48,60,72]
font_combo['values'] = font_name
color_combo['values'] = ['white','black','red','green','blue']

column_combo['state'] = 'readonly'
align_combo['state'] = 'readonly'
font_combo['state'] = 'readonly'
color_combo['state'] = 'readonly'

column_combo.place(x = 649,y = 68,height = 25,width = 146)
align_combo.place(x = 649,y = 128,height = 25,width = 146)
font_combo.place(x = 649,y = 188,height = 25,width = 146)
size_combo.place(x = 649,y = 248,height = 25,width = 146)
color_combo.place(x = 649,y = 308,height = 25,width = 146)

display.bind('<Button 1>', spot)

# create and place btn
submit = Button(win,text = 'save',command = lambda : draw(),state = DISABLED)
submit.place(x = 649,y = 357, width=146, height=25)

win.mainloop()
