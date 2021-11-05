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

font_name = []
pos = []
pos_arr = []
align_arr = []
able_key = []

pos_dict = {}
align_dict = {}
img_dict = {}
font_dict = {}
size_dict = {}
color_dict = {}
unscaled_pos_dict = {}
text_dict = {}

full_fontname = {}

size = 20
mid_x = 400
mid_y = 225

# function when start a new project
def file_choice(selected):
    global img_first,img_dict,able_key,text_dict,color_dict
    s_choice.set('File')
    if selected == 'New':
        if find_template_path() == '':
            return
        if find_data_path() == '':
            return
        img_first = Image.open(template_path)
        img_dict = dict.fromkeys(columns,img_first)
        color_dict = dict.fromkeys(columns, (255,255,255,0))
        able_key = []
        text_dict = {}
        find_ratio()
        update_pic(img_first)

    elif selected == 'Exit':
        win.destroy()

    elif selected == 'Clear':

        submit['state'] = DISABLED
        img_dict = dict.fromkeys(columns, img_first)
        able_key = []
        text_dict = {}
        update_pic(img_first)
# create an array which contain all data
def create_text_dict():
    global text_dict
    for i in able_key:
        text_dict[i] = list(data[i])
# find necessary path
def find_data_path():
    global data_path,data,columns
    file = filedialog.askopenfile(mode='r',filetypes =[('*','.xlsx')],title = 'Choose Data File')
    if file:
        data_path = os.path.abspath(file.name)
        data = pd.read_excel(data_path)
        columns = list(data)  # make a list of columns name
        if 'Unnamed: 0' in columns:
            columns.remove('Unnamed: 0')
        column_combo['values'] = columns
def find_template_path():
    global template_path
    file = filedialog.askopenfile(mode='r', filetypes=[('*', '.png .jpg .jpeg')],title = 'Choose Template')
    if file:
        template_path = os.path.abspath(file.name)
# using for update the display pic to other
def update_pic(img):
    img_re_size = img.resize((487,344), Image.ANTIALIAS)
    img2 = ImageTk.PhotoImage(img_re_size)
    display.configure(image=img2)
    display.image = img2
    display.place(x = 157,y = 38)
# find the ratio for scaling the font size
def find_ratio():
    global ratio,width,height
    width, height = img_first.size
    ratio = (math.sqrt(width ** 2+height ** 2)) / 596
# choose a spot and collect necessary data
def spot(event):
    global pos,size,img_dict,key,using_font,able_key,color,place_font
    if column_combo.get() != '' and align_combo.get() != '' and font_combo.get() != '' \
       and color_combo.get() != '':

        key = column_combo.get()

        if submit['state'] != NORMAL:
            submit['state'] = NORMAL
        if not key in able_key:
            able_key.append(key)
        if size_combo.get().isdigit():
            size = int(size_combo.get())
        using_font = f'C:\Windows\Fonts\{full_fontname[font_combo.get()]}'
        place_font = ImageFont.truetype(font=using_font, size=size)
        x, y = event.x, event.y
        scaled_up_x = x*(width/487)
        scaled_up_y = y*(height/344)
        pos = [scaled_up_x,scaled_up_y]

        img_show = img_first
        img_show = img_show.resize((487, 344), Image.ANTIALIAS)
        set_dict(img_show)
        for run_key in able_key:
            if run_key != column_combo.get():
                img_draw = ImageDraw.Draw(img_show)
                the_font = ImageFont.truetype(font=font_dict[run_key], size=size_dict[run_key])
                img_draw.text(unscaled_pos_dict[run_key], run_key, color_dict[run_key],font=the_font
                              ,anchor=align_dict[run_key])

        lined = ImageDraw.Draw(img_show)
        lined.line((x, 0, x, 450), fill='#ff0000')
        lined.line((0, y, 800, y), fill='#ff0000')
        lined.text((x, y), column_combo.get(), set_color(), font=place_font, anchor=align)

        img_new = ImageTk.PhotoImage(img_show)
        display.configure(image=img_new)
        display.image = img_new
        display.place(x=157, y=38)

        img_dict[column_combo.get()] = img_first
# save the information in a form of dictionary with columns of excel file as a key
def set_dict(img):
    global align
    align = align_combo.get()
    if len(align) > 1:
        if not 'baseline' in align:
            align = align[0] + align[align.find('-') + 1]
        else:
            align = align[0] + 's'
    if column_combo.get() in columns:
        pos_dict[key] = pos
        img_dict[key] = img
        align_dict[key] = align
        font_dict[key] = using_font
        size_dict[key] = size
        color_dict[key] = set_color()
        unscaled_pos_dict[key] = (pos[0]*(487 / width),pos[1]*(344 / height))
# create and save image
def draw():
    save_folder_path = filedialog.askdirectory(title='Choose Folder For Saving')
    create_text_dict()
    for i in range(len(data)):
        img = Image.open(template_path)
        result = ImageDraw.Draw(img)
        for j in able_key:
            font = ImageFont.truetype(font=font_dict[j], size=int(size_dict[j] * ratio))
            result.text((int(pos_dict[j][0]), int(pos_dict[j][1])), str(text_dict[j][i])
                        ,color_dict[j], font=font, anchor=align_dict[j])
        img.save(f'{save_folder_path}/{i+1}.png')
        print(f'Save at {save_folder_path}')
        print(f'save {i+1} successfully')
    print('done!')
# collect all fonts from device and create font list with no format for display
# and create a dictionary of font name with format for usage/write text on image
def get_font_name():
    global font_name
    for file in os.listdir('C:\Windows\Fonts'):
        if not '.fon' in file:
            this_font = file[0:file.find('.')]
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
# create canvas
canvas = Canvas(win,width = 800,height = 450)
canvas.place(x = 0,y = 0)
rect = canvas.create_rectangle(157,38,644,382,outline = 'black')
# place mockup image
display = Label(win,text = 'emty',anchor = 'center')
display.place(x = 390,y = 205)

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
