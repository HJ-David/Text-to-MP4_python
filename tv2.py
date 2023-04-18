from tkinter import *
from tkinter import filedialog
import pandas as pd
import openpyxl
import xlrd
import PyPDF2
import docx
from PIL import Image, ImageDraw, ImageFont
from moviepy.editor import *
from gtts import gTTS
import os
import textwrap

# Clear the textbox
def clear_text_box():
    my_text.delete(1.0, END)
    successful_label.config(text="")
    path_label.config(text="")
    
def open_file():
    global contents
    global file_path
    global top_text
    # Ask the user to select a file
    file_path = filedialog.askopenfilename()
    # Check the file type and read its contents accordingly
    if file_path.endswith('.txt'):
        # Open the selected file and read its contents
        with open(file_path, 'r', encoding='utf-8') as file:
            contents = file.readlines()
            top_text = file.readline()
        # Convert the contents list to a string
        contents_str = "".join(contents)
        # Add the string to the textbox
        my_text.insert(1.0, contents_str)
    elif file_path.endswith(('.xlsx', '.xls')):
        # Read the Excel file using pandas
        df = pd.read_excel(file_path, 'Sheet1')
        contents = df[df.columns[0]].values.tolist()
        # Convert the dataframe to a string
        contents_str = df.to_string(index=False)
        # Add the string to the textbox
        my_text.insert(1.0, contents_str)
        top_text = df.iloc[0, 0]
    elif file_path.endswith('.numbers'):
        # Read the Numbers file using pandas
        df = pd.read_excel(file_path, engine='odf')
        # Convert the dataframe to a string
        contents_str = df.to_string(index=False)
        # Add the string to the textbox
        my_text.insert(1.0, contents_str)
    elif file_path.endswith('.docx'):
        # Read the Word file using docx
        doc = docx.Document(file_path)
        # Extract the text from the document
        contents = [p.text for p in doc.paragraphs]
        # Convert the contents list to a string
        contents_str = "".join(contents)
        # Add the string to the textbox
        my_text.insert(1.0, contents_str)
    elif file_path.endswith('.pdf'):
        # Read the PDF file using PyPDF2
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfFileReader(file)
            contents = []
            for i in range(reader.numPages):
                page = reader.getPage(i)
                contents.append(page.extractText())
        # Convert the contents list to a string
        contents_str = "".join(contents)
        # Add the string to the textbox
        my_text.insert(1.0, contents_str)

def mp3_create_loop():
    # create mp3 loop
    for index, line in enumerate(contents, 1):
        speak = gTTS(line, lang= "en")
        speak.save(str(new_dir_path) + "//text{0}.mp3".format(index))
        # speak.save("./Tkinter_app/audio//text{0}.mp3".format(index))
    successful_label.config(text="MP3 created!")
    path_label.config(text="Location: " + str(new_dir_path))

def create_mp3():
    global mp3_directory_path
    
    initialdir = os.path.expanduser(file_path)
    mp3_directory_path = filedialog.askdirectory(initialdir=initialdir)
    # create mp3 loop
    for index, line in enumerate(contents, 1):
        speak = gTTS(line, lang= "en")
        speak.save(str(mp3_directory_path) + "//text{0}.mp3".format(index))
        # speak.save("./Tkinter_app/audio//text{0}.mp3".format(index))
    successful_label.config(text="MP3 created!")
    path_label.config(text="Location: " + str(mp3_directory_path))

def img_create_loop():
    # Open the text file
    for index, line in enumerate(contents, 1):
        # Strip any whitespace from the beginning and end of the line
        line = line.strip()

        # Create the image with the background
        image = Image.new(color_mode, image_size, background_color)

        # Create a drawing object
        draw = ImageDraw.Draw(image)

        # Get the size of the text
        text_size = draw.textsize(line, font=font)
        
        # Get the text to add to the top left of the image
        text_top = str(top_text)
        

        # Get the size of the text to add
        text_size_top = draw.textsize(text_top, font=font_top)
        
        # Calculate the x and y coordinates for centering the English text
        x = (image.width - text_size[0]) / 2
        y = (image.height - text_size[1] ) / 2
      
        
        # Wrap the English text onto multiple lines if it's too long for the image
        english_lines = textwrap.wrap(line, width=30)
        # Calculate the y-coordinate for the English text
        y_english = y - (len(english_lines) - 1) * text_size[1] / 2
        # Draw each line of English text on a separate line in the image
        for english_line in english_lines:
            english_size = draw.textsize(english_line, font=font)
            x_english = (image.width - english_size[0]) / 2
            draw.text((x_english, y_english), english_line, font=font, fill=(255, 255, 255))
            y_english += english_size[1]
            
        # Calculate the x and y coordinates for the top left of the image
        x_top = 40
        y_top = 40
        # Draw the text to add in the top left corner of the image
        draw.text((x_top, y_top), text_top, font=font_top, fill=(255, 255, 255))
        
        # Calculate the x and y coordinates for the top left of the image
        x_logo = 1600
        y_logo = 40
        text_logo = "RLL English"
        # Draw the text to add in the top left corner of the image
        draw.text((x_logo, y_logo), text_logo, font=font_top, fill=(255, 255, 255))
        
        # Save the image file with the name of the line
        image.save(str(new_dir_path) + "//img.{0:03d}.jpg".format(index))
    successful_label.config(text="Img created!")
    path_label.config(text="Location: " + str(new_dir_path))
    
def create_img():
    global img_directory_path
    initialdir = os.path.expanduser(file_path)
    img_directory_path = filedialog.askdirectory(initialdir=initialdir)
    # Open the text file
    for index, line in enumerate(contents, 1):
        # Strip any whitespace from the beginning and end of the line
        line = line.strip()

        # Create the image with the background
        image = Image.new(color_mode, image_size, background_color)

        # Create a drawing object
        draw = ImageDraw.Draw(image)

        # Get the size of the text
        text_size = draw.textsize(line, font=font)
        
        # Get the text to add to the top left of the image
        text_top = str(top_text)
        

        # Get the size of the text to add
        text_size_top = draw.textsize(text_top, font=font_top)
        
        # Calculate the x and y coordinates for centering the English text
        x = (image.width - text_size[0]) / 2
        y = (image.height - text_size[1] ) / 2
      
        
        # Wrap the English text onto multiple lines if it's too long for the image
        english_lines = textwrap.wrap(line, width=30)
        # Calculate the y-coordinate for the English text
        y_english = y - (len(english_lines) - 1) * text_size[1] / 2
        # Draw each line of English text on a separate line in the image
        for english_line in english_lines:
            english_size = draw.textsize(english_line, font=font)
            x_english = (image.width - english_size[0]) / 2
            draw.text((x_english, y_english), english_line, font=font, fill=(255, 255, 255))
            y_english += english_size[1]
            
        # Calculate the x and y coordinates for the top left of the image
        x_top = 40
        y_top = 40
        # Draw the text to add in the top left corner of the image
        draw.text((x_top, y_top), text_top, font=font_top, fill=(255, 255, 255))
        
        # Calculate the x and y coordinates for the top left of the image
        x_logo = 1600
        y_logo = 40
        text_logo = "RLL English"
        # Draw the text to add in the top left corner of the image
        draw.text((x_logo, y_logo), text_logo, font=font_top, fill=(255, 255, 255))
        
        # Save the image file with the name of the line
        image.save(str(img_directory_path) + "//img.{0:03d}.jpg".format(index))
    successful_label.config(text="Img created!")
    path_label.config(text="Location: " + str(img_directory_path))

    
def create_movie():
    # file_path = ""
    # if not file_path:
    #     mp4_directory_path=filedialog.askdirectory()
    # else:
        # open_file()
        initialdir = os.path.expanduser(file_path)
        mp4_directory_path = filedialog.askdirectory(initialdir=initialdir)
        i=1
        y=[]
        for file in os.listdir(mp3_directory_path):
            if file.endswith('.mp3'):
                audio = AudioFileClip(str(mp3_directory_path) + "/text%s.mp3"%i)
                clip_a = ImageClip(str(img_directory_path) + "/img.%03d.jpg"%i).set_duration(audio.duration)
                clip_b = clip_a.set_audio(audio)
            
                y.append(clip_b)
                y.append(clip_a)
                y.append(clip_a)
                # y.append(clip_b)
                # y.append(clip_a)
                # y.append(clip_a)
                # y.append(clip_b)
                # y.append(clip_a)
                # y.append(clip_a)
                        
                i += 1

        # 기본 디렉토리에 render 비디오를 합쳐서 body.mp4로 저장
        final_clip = concatenate_videoclips(y) 
        final_clip.write_videofile(str(mp4_directory_path) + "/body.mp4", codec='libx264', audio=True, audio_codec='aac', fps=24)
        
        successful_label.config(text="MP4 created!")
        path_label.config(text="Location: " + str(mp4_directory_path))
        
def create_movie_oneclick():
        global new_dir_path
        open_file()
        # Extract directory from file path
        directory_path = "/".join(file_path.split("/")[:-1])
        # Create a directory in the same directory as the file
        new_dir_name = "mp4_created"
        new_dir_path = os.path.join(directory_path, new_dir_name)
        # Check if directory already exists and add number if it does
        num = 2
        while os.path.exists(new_dir_path):
            new_dir_name = "mp4_created" + str(num)
            new_dir_path = os.path.join(directory_path, new_dir_name)
            num += 1
        os.mkdir(new_dir_path)
        
        # Create MP3
        mp3_create_loop()
        
        # Create Image
        img_create_loop()
    
        # Create MP4
        i=1
        y=[]
        for file in os.listdir(new_dir_path):
            if file.endswith('.mp3'):
                audio = AudioFileClip(str(new_dir_path) + "/text%s.mp3"%i)
                clip_a = ImageClip(str(new_dir_path) + "/img.%03d.jpg"%i).set_duration(audio.duration)
                clip_b = clip_a.set_audio(audio)
            
                y.append(clip_b)
                y.append(clip_a)
                # y.append(clip_a)
                y.append(clip_b)
                y.append(clip_a)
                # y.append(clip_a)
                # y.append(clip_b)
                # y.append(clip_a)
                # y.append(clip_a)
                        
                i += 1

        # 기본 디렉토리에 render 비디오를 합쳐서 body.mp4로 저장
        final_clip = concatenate_videoclips(y) 
        final_clip.write_videofile(str(directory_path) + "/body.mp4", codec='libx264', audio=True, audio_codec='aac', fps=24)
        
        #Delete MP3 & Image
        for file_name in os.listdir(new_dir_path):
            if file_name.endswith(".mp3"):
                os.remove(os.path.join(new_dir_path, file_name))
        for file_name in os.listdir(new_dir_path):
            if file_name.endswith(".jpg"):
                os.remove(os.path.join(new_dir_path, file_name))
        # Check if directory is empty
        if not os.listdir(new_dir_path):
            # Remove directory if it's empty
            os.rmdir(new_dir_path)
        #Success Message
        successful_label.config(text="MP4 created!")
        path_label.config(text="Location: " + str(new_dir_path))

def intro_outro():
        global new_dir_path
        open_file()
        # Extract directory from file path
        directory_path = "/".join(file_path.split("/")[:-1])
        # Create a directory in the same directory as the file
        new_dir_name = "mp4_created"
        new_dir_path = os.path.join(directory_path, new_dir_name)
        # Check if directory already exists and add number if it does
        num = 2
        while os.path.exists(new_dir_path):
            new_dir_name = "mp4_created" + str(num)
            new_dir_path = os.path.join(directory_path, new_dir_name)
            num += 1
        os.mkdir(new_dir_path)
        
        # Create MP3
        mp3_create_loop()
        # Create Image
        img_create_loop()
        
        # Create MP4
        i=1
        y=[]
        for file in os.listdir(new_dir_path):
            if file.endswith('.mp3'):
                audio = AudioFileClip(str(new_dir_path) + "/text%s.mp3"%i)
                clip_a = ImageClip(str(new_dir_path) + "/img.%03d.jpg"%i).set_duration(audio.duration)
                clip_b = clip_a.set_audio(audio)
            
                y.append(clip_b)
                y.append(clip_a)
                # y.append(clip_a)
                # y.append(clip_b)
                # y.append(clip_a)
                # y.append(clip_a)
                # y.append(clip_b)
                # y.append(clip_a)
                # y.append(clip_a)
                        
                i += 1
        # intro와 outro를 y 배열 맨 앞과 뒤에 추가
        intro = VideoFileClip(str(selected_intro_path))
        outro = VideoFileClip(str(selected_outro_path))
        y.append(outro)
        y.insert(0,intro)
        
        # 모든 비디오를 합쳐서 final.mp4로 저장
        final_clip = concatenate_videoclips(y) 
        final_clip.write_videofile(str(directory_path) + "/final.mp4", codec='libx264', audio=True, audio_codec='aac', fps=24)
        
        #Delete MP3 & Image
        for file_name in os.listdir(new_dir_path):
            if file_name.endswith(".mp3"):
                os.remove(os.path.join(new_dir_path, file_name))
        for file_name in os.listdir(new_dir_path):
            if file_name.endswith(".jpg"):
                os.remove(os.path.join(new_dir_path, file_name))
        # Check if directory is empty
        if not os.listdir(new_dir_path):
            # Remove directory if it's empty
            os.rmdir(new_dir_path)
            
        successful_label.config(text="MP4 created with Intro and Outro!")
        path_label.config(text="Location: " + str(new_dir_path))
    
    
# Create a window with a button to open the file
root = Tk()
root.title('English Text to MP4 Converter App')
root.geometry("800x450")

#Title
title = Label(root, text = "English Text to MP4 Converter App")
title.grid(row = 0, column = 0, sticky = W, padx = 15, pady = 5)

# Create open and clear buttons
open_button = Button(root, text="Open", command=open_file)
open_button.grid(row = 0, column = 1, sticky = W, columnspan = 1)
clear_button = Button(root, text="Clear", command=clear_text_box)
clear_button.grid(row = 0, column = 2, sticky = W, columnspan = 1)

# Create a textbox
my_text = Text(root, height=30, width=60)
my_text.grid(row = 1, column = 0, rowspan=20, columnspan=5, sticky = W, padx=10, pady=0)

# Create the Successful_label
successful_label = Label(root, text="")
successful_label.grid(row = 1, column = 7, sticky = W, padx = 10, columnspan = 3)
# Create the path_label
path_label = Label(root, text="")
path_label.grid(row = 2, column = 7, sticky = W, padx = 10, columnspan = 3)

# Create a conver MP# button
mp3_button = Button(root, text="Create MP3", command=create_mp3)
mp3_button.grid(row = 3, column = 7, sticky = W, columnspan = 1)

# Create a button widget in the root window with the text "Create img"
# and assign it to the variable image_button. When the button is clicked,
# the function create_img will be called.
image_button = Button(root, text="Create img", command=create_img)
image_button.grid(row = 3, column = 8, sticky = W, columnspan = 1)

# Create the MP4
movie_button = Button(root, text="Create MP4", command=create_movie)
movie_button.grid(row = 3, column = 9, sticky = W, columnspan = 1)

# Create the MP4_oneclick
movie_button = Button(root, text="One Click MP4 Create", command=create_movie_oneclick)
movie_button.grid(row = 4, column = 7, sticky = W, columnspan = 2)

def choose_intro():
    intro_file_path = filedialog.askopenfilename()
    if intro_file_path:
        file_name = intro_file_path.split("/")[-1]  # extract the file name from the path
        intro_entry_var.set(file_name)
        global selected_intro_path
        selected_intro_path = intro_file_path
def choose_outro():
    outro_file_path = filedialog.askopenfilename()
    if outro_file_path:
        file_name = outro_file_path.split("/")[-1]  # extract the file name from the path
        outro_entry_var.set(file_name)
        global selected_outro_path
        selected_outro_path = outro_file_path


intro_button = Button(root, text="Intro", command=choose_intro)
intro_button.grid(row = 6, column = 7, sticky = W, columnspan = 2)

intro_entry_var = StringVar()
intro_entry = Entry(root, textvariable=intro_entry_var)
intro_entry.grid(row = 7, column = 7, sticky = W, columnspan = 2)

selected_intro_path = None

outro_button = Button(root, text="Outro", command=choose_outro)
outro_button.grid(row = 8, column = 7, sticky = W, columnspan = 2)

outro_entry_var = StringVar()
outro_entry = Entry(root, textvariable=outro_entry_var)
outro_entry.grid(row = 9, column = 7, sticky = W, columnspan = 2)

selected_outro_path = None

# Create the MP4 with Intro and Outro
final_button = Button(root, text="MP4 with Intro & Outro", command=intro_outro)
final_button.grid(row = 10, column = 7, sticky = W, columnspan = 2)

# Set the image size and color mode
image_size = (1920, 1080)
color_mode = 'RGB'
background_color = (25, 124, 163)

# Set the font size and type
font = ImageFont.truetype("/System/Library/Fonts/Supplemental/DIN Alternate Bold.ttf", 100)

# Create a new font object for the top text
font_top = ImageFont.truetype("/Users/hj/Library/Fonts/KakaoRegular.ttf", 50)

#Create A Menu
my_menu = Menu(root)
root.config(menu=my_menu)

# Add some dropdown menus
file_menu = Menu(my_menu, tearoff=False)
my_menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Open", command=open_file)
file_menu.add_command(label="Clear", command=clear_text_box)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)



root.mainloop()
