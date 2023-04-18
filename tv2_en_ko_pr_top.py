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
import requests


# 자신의 REST_API_KEY를 입력하세요!
REST_API_KEY = "52223979b01c7f819c42682ee1b0f3fa"


class KakaoTTS:

	def __init__(self, text, API_KEY=REST_API_KEY):
		self.resp = requests.post(
			url="https://kakaoi-newtone-openapi.kakao.com/v1/synthesize",
			headers={
				"Content-Type": "application/xml",
				"Authorization": f"KakaoAK {API_KEY}"
			},
			data=f"<speak>{text}</speak>".encode('utf-8')
		)

	def save(self, filename="output.mp3"):
		with open(filename, "wb") as file:
			file.write(self.resp.content)


# Clear the textbox
def clear_text_box():
    my_text.delete(1.0, END)
    successful_label.config(text="")
    path_label.config(text="")
    
def open_file():
    global contentsen
    global contentsko
    global contentspr
    global file_path
    # Ask the user to select a file
    file_path = filedialog.askopenfilename()
    # Read the Excel file using pandas
    df = pd.read_excel(file_path, 'Sheet1')
    contentsen = df[df.columns[0]].values.tolist()
    contentsko = df[df.columns[1]].values.tolist()
    contentspr = df[df.columns[2]].values.tolist()
 
def create_mp3():
    global mp3_directory_path
    
    initialdir = os.path.expanduser(file_path)
    mp3_directory_path = filedialog.askdirectory(initialdir=initialdir)
    # Create En mp3 with the directory_path
    for index, line in enumerate(contentsen, 1):
        speak = gTTS(line, lang= "en")
        speak.save(str(mp3_directory_path) + "//texten{0}.mp3".format(index))
        # speak.save("./Tkinter_app/audio//text{0}.mp3".format(index))
    successful_label.config(text="MP3 created!")
    path_label.config(text="Location: " + str(mp3_directory_path))
    # Create Ko mp3 with the directory_path
    for index, line in enumerate(contentsko, 1):
        tts = KakaoTTS(line)
        tts.save(str(mp3_directory_path) + "//textko{0}.mp3".format(index))

    
def create_img():
    global img_directory_path
    initialdir = os.path.expanduser(file_path)
    img_directory_path = filedialog.askdirectory(initialdir=initialdir)
    # create img
    for index, line in enumerate(contentsen, 1):
        # Define a constant value for the space between sentences
        sentence_space = 50
        # Strip any whitespace from the beginning and end of the line
        line = line.strip()
        # Strip any whitespace from the beginning and end of the Korean line
        line_ko = contentsko[index-1].strip()
         # Strip any whitespace from the beginning and end of the pronounciation line
        line_pr = contentspr[index-1].strip()
        # Create the image with the background
        image = Image.new(color_mode, image_size, background_color)
        # Create a drawing object
        draw = ImageDraw.Draw(image)
        # Get the size of the English text
        text_size = draw.textsize(line, font=font)
        # Get the size of the Korean text
        text_size_ko = draw.textsize(line_ko, font=font_ko)
        # Get the size of the Korean pronounciation text
        text_size_pr = draw.textsize(line_pr, font=font)
        # Calculate the x and y coordinates for centering the English text
        x = (image.width - text_size[0]) / 2
        y = (image.height - text_size[1] - text_size_ko[1] - text_size_pr[1]) / 2 - sentence_space
        # Draw the English text in the center of the image
        draw.text((x, y), line, font=font, fill=(255, 255, 255))
        # Calculate the x and y coordinates for centering the Korean text
        x_ko = (image.width - text_size_ko[0]) / 2
        y_ko = y + text_size[1] + sentence_space
        # Calculate the x and y coordinates for centering the Korean pronounciation
        x_pr = (image.width - text_size_pr[0]) / 2
        y_pr = y + text_size[1] + text_size_ko[1] + sentence_space * 2
        # Place the Korean text below the English text
        # Draw the Korean text in the center of the image below the English text
        draw.text((x_ko, y_ko), line_ko, font=font_ko, fill=(255, 255, 255))
        # Place the Korean text below the English text
        # Draw the Korean text in the center of the image below the English text
        draw.text((x_pr, y_pr), line_pr, font=font, fill=(255, 255, 255))
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
            if "en" in file:
                audio_en = AudioFileClip(str(mp3_directory_path) + "/texten%s.mp3"%i)
                audio_ko = AudioFileClip(str(mp3_directory_path) + "/textko%s.mp3"%i)
                clip_a = ImageClip(str(img_directory_path) + "/img.%03d.jpg"%i).set_duration(audio_en.duration)
                clip_b = ImageClip(str(img_directory_path) + "/img.%03d.jpg"%i).set_duration(audio_ko.duration)
                clip_en = clip_a.set_audio(audio_en)
                clip_ko = clip_a.set_audio(audio_ko)
            
                y.append(clip_en)
                y.append(clip_a)
                y.append(clip_ko)
                y.append(clip_b)
                y.append(clip_en)
                y.append(clip_a)
                y.append(clip_ko)
                y.append(clip_b)
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
         # Create En mp3 with the directory_path
        for index, line in enumerate(contentsen, 1):
            speak = gTTS(line, lang= "en")
            speak.save(str(mp3_directory_path) + "//texten{0}.mp3".format(index))
            # speak.save("./Tkinter_app/audio//text{0}.mp3".format(index))
        successful_label.config(text="MP3 created!")
        path_label.config(text="Location: " + str(mp3_directory_path))
        # Create Ko mp3 with the directory_path
        for index, line in enumerate(contentsko, 1):
            tts = KakaoTTS(line)
            tts.save(str(mp3_directory_path) + "//textko{0}.mp3".format(index))
        successful_label.config(text="MP3 created!")
        
        # Create Image
        # create img
        for index, line in enumerate(contentsen, 1):
            # Define a constant value for the space between sentences
            sentence_space = 50
            # Strip any whitespace from the beginning and end of the line
            line = line.strip()
            # Strip any whitespace from the beginning and end of the Korean line
            line_ko = contentsko[index-1].strip()
            # Strip any whitespace from the beginning and end of the pronounciation line
            line_pr = contentspr[index-1].strip()
            # Create the image with the background
            image = Image.new(color_mode, image_size, background_color)
            # Create a drawing object
            draw = ImageDraw.Draw(image)
            # Get the size of the English text
            text_size = draw.textsize(line, font=font)
            # Get the size of the Korean text
            text_size_ko = draw.textsize(line_ko, font=font_ko)
            # Get the size of the Korean pronounciation text
            text_size_pr = draw.textsize(line_pr, font=font)
            # Calculate the x and y coordinates for centering the English text
            x = (image.width - text_size[0]) / 2
            y = (image.height - text_size[1] - text_size_ko[1] - text_size_pr[1]) / 2 - sentence_space
            # Draw the English text in the center of the image
            draw.text((x, y), line, font=font, fill=(255, 255, 255))
            # Calculate the x and y coordinates for centering the Korean text
            x_ko = (image.width - text_size_ko[0]) / 2
            y_ko = y + text_size[1] + sentence_space
            # Calculate the x and y coordinates for centering the Korean pronounciation
            x_pr = (image.width - text_size_pr[0]) / 2
            y_pr = y + text_size[1] + text_size_ko[1] + sentence_space * 2
            # Place the Korean text below the English text
            # Draw the Korean text in the center of the image below the English text
            draw.text((x_ko, y_ko), line_ko, font=font_ko, fill=(255, 255, 255))
            # Place the Korean text below the English text
            # Draw the Korean text in the center of the image below the English text
            draw.text((x_pr, y_pr), line_pr, font=font, fill=(255, 255, 255))
            # Save the image file with the name of the line
            image.save(str(img_directory_path) + "//img.{0:03d}.jpg".format(index))
        successful_label.config(text="Image created!")
    
        # Create MP4
        i=1
        y=[]
        for file in os.listdir(mp3_directory_path):
            if "en" in file:
                audio_en = AudioFileClip(str(mp3_directory_path) + "/texten%s.mp3"%i)
                audio_ko = AudioFileClip(str(mp3_directory_path) + "/textko%s.mp3"%i)
                clip_a = ImageClip(str(img_directory_path) + "/img.%03d.jpg"%i).set_duration(audio_en.duration)
                clip_b = ImageClip(str(img_directory_path) + "/img.%03d.jpg"%i).set_duration(audio_ko.duration)
                clip_en = clip_a.set_audio(audio_en)
                clip_ko = clip_a.set_audio(audio_ko)
            
                y.append(clip_en)
                y.append(clip_a)
                y.append(clip_ko)
                y.append(clip_b)
                y.append(clip_en)
                y.append(clip_a)
                y.append(clip_ko)
                y.append(clip_b)
                # y.append(clip_a)
                # y.append(clip_a)
                # y.append(clip_b)
                # y.append(clip_a)
                # y.append(clip_a)
                        
                i += 1

        # 기본 디렉토리에 render 비디오를 합쳐서 body.mp4로 저장
        final_clip = concatenate_videoclips(y) 
        final_clip.write_videofile(str(new_dir_path) + "/body.mp4", codec='libx264', audio=True, audio_codec='aac', fps=24)
        
        #Delete MP3 & Image
        for file_name in os.listdir(new_dir_path):
            if file_name.endswith(".mp3"):
                os.remove(os.path.join(new_dir_path, file_name))
        for file_name in os.listdir(new_dir_path):
            if file_name.endswith(".jpg"):
                os.remove(os.path.join(new_dir_path, file_name))
        #Success Message
        successful_label.config(text="MP4 created!")
        path_label.config(text="Location: " + str(new_dir_path))

def intro_outro():
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
        for index, line in enumerate(contentsen, 1):
            speak = gTTS(line, lang= "en")
            speak.save(str(mp3_directory_path) + "//texten{0}.mp3".format(index))
            # speak.save("./Tkinter_app/audio//text{0}.mp3".format(index))
        successful_label.config(text="MP3 created!")
        path_label.config(text="Location: " + str(mp3_directory_path))
        # Create Ko mp3 with the directory_path
        for index, line in enumerate(contentsko, 1):
            tts = KakaoTTS(line)
            tts.save(str(mp3_directory_path) + "//textko{0}.mp3".format(index))
        
        # Create Image
        # create img
        for index, line in enumerate(contentsen, 1):
            # Define a constant value for the space between sentences
            sentence_space = 50
            # Strip any whitespace from the beginning and end of the line
            line = line.strip()
            # Strip any whitespace from the beginning and end of the Korean line
            line_ko = contentsko[index-1].strip()
            # Strip any whitespace from the beginning and end of the pronounciation line
            line_pr = contentspr[index-1].strip()
            # Create the image with the background
            image = Image.new(color_mode, image_size, background_color)
            # Create a drawing object
            draw = ImageDraw.Draw(image)
            # Get the size of the English text
            text_size = draw.textsize(line, font=font)
            # Get the size of the Korean text
            text_size_ko = draw.textsize(line_ko, font=font_ko)
            # Get the size of the Korean pronounciation text
            text_size_pr = draw.textsize(line_pr, font=font)
            # Calculate the x and y coordinates for centering the English text
            x = (image.width - text_size[0]) / 2
            y = (image.height - text_size[1] - text_size_ko[1] - text_size_pr[1]) / 2 - sentence_space
            # Draw the English text in the center of the image
            draw.text((x, y), line, font=font, fill=(255, 255, 255))
            # Calculate the x and y coordinates for centering the Korean text
            x_ko = (image.width - text_size_ko[0]) / 2
            y_ko = y + text_size[1] + sentence_space
            # Calculate the x and y coordinates for centering the Korean pronounciation
            x_pr = (image.width - text_size_pr[0]) / 2
            y_pr = y + text_size[1] + text_size_ko[1] + sentence_space * 2
            # Place the Korean text below the English text
            # Draw the Korean text in the center of the image below the English text
            draw.text((x_ko, y_ko), line_ko, font=font_ko, fill=(255, 255, 255))
            # Place the Korean text below the English text
            # Draw the Korean text in the center of the image below the English text
            draw.text((x_pr, y_pr), line_pr, font=font, fill=(255, 255, 255))
            # Save the image file with the name of the line
            image.save(str(img_directory_path) + "//img.{0:03d}.jpg".format(index))
        successful_label.config(text="Image created!")
        
        # Create MP4
        i=1
        y=[]
        for file in os.listdir(mp3_directory_path):
            if "en" in file:
                audio_en = AudioFileClip(str(mp3_directory_path) + "/texten%s.mp3"%i)
                audio_ko = AudioFileClip(str(mp3_directory_path) + "/textko%s.mp3"%i)
                clip_a = ImageClip(str(img_directory_path) + "/img.%03d.jpg"%i).set_duration(audio_en.duration)
                clip_b = ImageClip(str(img_directory_path) + "/img.%03d.jpg"%i).set_duration(audio_ko.duration)
                clip_en = clip_a.set_audio(audio_en)
                clip_ko = clip_a.set_audio(audio_ko)
            
                y.append(clip_en)
                y.append(clip_a)
                y.append(clip_ko)
                y.append(clip_b)
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
        final_clip.write_videofile(str(new_dir_path) + "/final.mp4", codec='libx264', audio=True, audio_codec='aac', fps=24)
        
        #Delete MP3 & Image
        for file_name in os.listdir(new_dir_path):
            if file_name.endswith(".mp3"):
                os.remove(os.path.join(new_dir_path, file_name))
        for file_name in os.listdir(new_dir_path):
            if file_name.endswith(".jpg"):
                os.remove(os.path.join(new_dir_path, file_name))
        
        successful_label.config(text="MP4 created with Intro and Outro!")
        path_label.config(text="Location: " + str(new_dir_path))
    
    
# Create a window with a button to open the file
root = Tk()
root.title('Converter App')
root.geometry("1000x500")

#Title
title = Label(root, text = "Converter App")
title.grid(row = 0, column = 0, sticky = W, padx = 15, pady = 5)

# Create open and clear buttons
open_button = Button(root, text="Open Excel", command=open_file)
open_button.grid(row = 0, column = 1, sticky = W, columnspan = 2)
clear_button = Button(root, text="Clear", command=clear_text_box)
clear_button.grid(row = 0, column = 3, sticky = W, columnspan = 2)

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
font = ImageFont.truetype("/System/Library/Fonts/Supplemental/DIN Alternate Bold.ttf", 60)
font_ko = ImageFont.truetype("/Users/hj/Library/Fonts/KakaoRegular.ttf", 60)

#Create A Menu
my_menu = Menu(root)
root.config(menu=my_menu)

# Add some dropdown menus
file_menu = Menu(my_menu, tearoff=False)
my_menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Open Excel", command=open_file)
file_menu.add_command(label="Clear", command=clear_text_box)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.quit)



root.mainloop()
