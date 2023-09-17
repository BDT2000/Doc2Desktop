import time
import textwrap
import ctypes
from docx import Document
from PIL import Image, ImageDraw, ImageFont
from apscheduler.schedulers.background import BackgroundScheduler
"""
Features to add:
* XMLParser
 - paragraph lines
 - tabs and indentations
 - bold and italics
 onedrive
"""
# In settings change background fit to 'Fit'

directory = "C:\\Users\\Bryce\\AppData\\Local\\Programs\\Python\\Python311\\d2d\\"
word_docx_file = "desktopbkgd.docx"

width = 1000
height = 1000
margin = 50

font_name = "arial.ttf"
font_size = 20
fnt = ImageFont.truetype("arial.ttf", font_size)
text_color = "black"

output_file="output.png"

# Function to extract text from a Word document
def extract_text_from_docx(docx_file):
    doc = Document(docx_file)
    extracted_text = ""
    for paragraph in doc.paragraphs:
        extracted_text += paragraph.text + "\n"
    return extracted_text

# Function to convert text to an image
def text_to_image(text):
    lines = []
    paragraphs = text.split('\n')
    whitespace_size = fnt.getlength(" ")

    for paragraph in paragraphs:
        words = paragraph.split()
        line = ""
        line_width = 0

        for word in words:
            word_width = sum([fnt.getlength(c) for c in word])
            if line_width + word_width <= width-margin*2:
                line += word + " "
                line_width += word_width+whitespace_size
            else:
                lines.append(line.strip())
                line = word + " "
                line_width = word_width

        if line:
            lines.append(line.strip())

    image = Image.new("RGB", (width,height), "white")
    draw = ImageDraw.Draw(image)
    y = margin
    for line in lines:
        draw.text((margin, y), line, font=fnt, fill=text_color)
        y += font_size + 10
    image.save(output_file)
    return output_file

# Function to set the desktop background
def set_wallpaper(image_path):
    ctypes.windll.user32.SystemParametersInfoW(20, 0, directory+image_path, 3)

# Function to periodically update the desktop background
def update_desktop_background():
    text = extract_text_from_docx(word_docx_file)
    output_image = text_to_image(text)
    set_wallpaper(output_image)

def main():
    scheduler = BackgroundScheduler()
    print('starto!')
    scheduler.add_job(update_desktop_background, 'interval', seconds=5)  # Run every 5 seconds
    scheduler.start()

    try:
        while True:
            time.sleep(3600)  # Keep the program running
    except (KeyboardInterrupt, SystemExit):
        scheduler.shutdown()

if __name__ == "__main__":
    main()
