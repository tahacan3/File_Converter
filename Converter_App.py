import os
import random
import rawpy
import imageio
import docx
import tkinter as tk
import ttkbootstrap as tkb
from PIL import Image
from psd_tools import PSDImage
from tkinter import messagebox, ttk, Toplevel

desktop = os.path.join(os.path.join(os.environ["USERPROFILE"]), "Desktop") # Kullanıcının masaüstü yolu.
os.chdir(desktop)

numbers = "01234567890123456789"

files = [".jpeg", ".png", ".webp", ".ico", ".gif", ".psd", ".hdr", ".cr3"]
files2 = [".jpeg", ".png", ".webp", ".ico", ".gif"]

def start():
    global top, cb1, cb2, button_convert
    top = Toplevel()
    form.withdraw()
    top.geometry(f"{form_width}x{form_height}+{int(x)}+{int(y)}")
    top.resizable(False, False)

    input_frame = ttk.Frame(master = top) 
    cb1 = ttk.Combobox(master = input_frame, values = files, font = "calibri 15", width = 15)
    lb1 = tk.Label(master = input_frame, font = "calibri 25", width = 15, text = "to")
    cb2 = ttk.Combobox(master = input_frame, values = files2, font = "calibri 15", width = 15)
    button_create_file = tkb.Button(master = top, text = "Create Files", style = "success.Outline.TButton", width = 15, command = create_file)
    button_convert = tkb.Button(master = top, text = "Convert", style = "success.Outline.TButton", width = 15, command = convert)
    input_frame.pack(pady = 20)
    cb1.pack(pady = 5)
    lb1.pack()
    cb2.pack(pady = 5)
    button_create_file.pack(pady = 10)
    button_convert.pack()

    button_convert.config(state = tk.DISABLED)

    buutton_back = tkb.Button(master = top, text = "Back", style = "2.success.Outline.TButton", width = 10, command = back1)
    buutton_back.pack(side = tk.BOTTOM, anchor = tk.SW, padx = 10, pady = 10)

    top.protocol("WM_DELETE_WINDOW", on_close)

def help():
    global top2
    top2 = Toplevel()
    form.withdraw()
    top2.geometry(f"{form_width}x{form_height}+{int(x)}+{int(y)}")
    top2.resizable(False, False)

    lb1 = tk.Label(master = top2, text = "How to Use?", font = "calibri 30 bold")
    lb1.pack()

    long_text = """First select the file types you want to convert. Then press the create file button. Two empty folders will appear on your desktop. At the beginning of the names of the folders will be the names of the file types you have selected and randomly assigned names. Put all the files you want to convert into the folder with the corresponding name. And then press the convert button."""

    lb1 = tk.Label(master = top2, text = long_text, justify = "left", wraplength = 250, font = "calibri 11")
    lb1.pack(pady=5)

    buutton_back = tkb.Button(master = top2, text = "Back", style = "2.success.Outline.TButton", width = 10, command = back2)
    buutton_back.pack(side = tk.BOTTOM, anchor = tk.SW, padx = 10, pady = 10)

    top2.protocol("WM_DELETE_WINDOW", on_close2)

def on_close():
    top.destroy() 
    form.destroy()   

def on_close2():
    top2.destroy() 
    form.destroy()

def flatten_psd(psd):
    img = Image.new('RGB', psd.size)
    for layer in reversed(psd.layers):
        if layer.is_visible():
            img.paste(layer.topil(), (0, 0), layer.topil())
    return img

def convert():
    source_dir = str(f"{cb1_text}-{name1}")
    output_dir = str(f"{cb2_text}-{name2}")

    print(f"Source Directory: {source_dir}")
    print(f"Source Directory Exists: {os.path.exists(source_dir)}")


    for file in os.listdir(source_dir):
        try:
            file_path = os.path.join(source_dir, file)
            output_filename = f"{os.path.splitext(file)[0]}{cb2_text}"
            output_path = os.path.join(output_dir, output_filename)

            if cb1_text == ".cr3":
                # Process RAW files using rawpy
                raw = rawpy.imread(file_path)
                rgb = raw.postprocess()
                image_converted = Image.fromarray(rgb)
            elif cb1_text == ".psd":
                # Process PSD files using Pillow directly
                psd_image = Image.open(file_path)
                # You may need to add further processing steps specific to your PSD files
                image_converted = psd_image
            elif cb1_text == ".hdr":
                # Process HDR files using imageio
                hdr_image = imageio.imread(file_path)
                image_converted = Image.fromarray(hdr_image)
            else:
                # Process non-RAW files using PIL
                image_converted = Image.open(file_path).convert("RGB")

            # Save the converted image
            image_converted.save(output_path)

        except Exception as e:
            print(f"Skipping {file} as an error occurred: {e}")

def create_file():
    global cb1_text, cb2_text
    cb1_text = str(cb1.get())
    cb2_text = str(cb2.get())

    for i in files:
        if(cb1_text == i):
            break
    for j in files2:
        if(cb2_text == j):
            break
    else:
        messagebox.showerror(title = "Error", message = "You should choose one of the files.")
        return

    button_convert.config(state = tk.NORMAL)

    global name1, name2
    name1 = "".join(random.sample(numbers, 4))
    os.mkdir(f"{cb1_text}-{name1}")
    name2 = "".join(random.sample(numbers, 4))
    os.mkdir(f"{cb2_text}-{name2}")

def back1():
    top.destroy()
    form.deiconify()

def back2():
    top2.destroy()
    form.deiconify()

# Main Form
form = tkb.Window(themename = "darkly")
form.title("Converter") #formun ismini değiştirir

form_width = 300
form_height = 400

screen_width = form.winfo_screenwidth() #ekranın genişliğini ölçer
screen_height = form.winfo_screenheight() #ekranın yüksekliğini ölçer

x = (screen_width - form_width) / 2
y = (screen_height - form_height) / 2

form.geometry(f"{form_width}x{form_height}+{int(x)}+{int(y)}") #formun boyutunu ayarlar ve ekranın ortasına hizalar
form.resizable(False, False) #formun boyutunu ayarlamayı devre dışı bırakır

# Style
main_style = tkb.Style()
main_style.configure("success.Outline.TButton", font = ("Calibri", 15))
back_style = tkb.Style()
main_style.configure("2.success.Outline.TButton", font = ("Calibri", 10))

# title
label1 = tk.Label(master = form, text = "Converter", font = "calibri 40 bold")
label1.pack()

# button frame
input_frame = ttk.Frame(master = form) 
button1 = tkb.Button(master = input_frame, text = "Start", width = 15, style = "success.Outline.TButton", command = start)
button2 = tkb.Button(master = input_frame, text = "Read", style = "success.Outline.TButton", width = 15, command = help)
input_frame.pack(pady = 30)
button1.pack(pady = 10)
button2.pack(pady = 10)

form.mainloop()