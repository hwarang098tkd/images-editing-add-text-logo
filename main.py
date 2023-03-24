import string
from random import choice

from numpy import random

from install_packages import install_packages

# Install required packages if not already installed
install_packages()

from PIL import Image, ImageDraw, ImageFont
from PIL.ExifTags import TAGS
import pandas as pd
import os
from datetime import datetime


def main():
    print("-----Script Started-----")
    rows = read_excel_file("images_settings.xlsx")
    copy_images(rows)
    print("-----Script End-----")


def get_image_metadata(image_path):
    # Reads the metadata and file info of an image file and returns it as a dictionary.
    with Image.open(image_path) as img:
        # Extract the raw Exif data
        exif_data = img._getexif()

        # Decode the Exif data into human-readable tags
        metadata = {}
        if exif_data:
            for tag_id, value in exif_data.items():
                tag = TAGS.get(tag_id, tag_id)
                metadata[tag] = value

    # Get file information using os module
    file_info = {}
    file_stats = os.stat(image_path)
    file_info['file_size'] = file_stats.st_size
    file_info['creation_time'] = file_stats.st_ctime
    file_info['modified_time'] = file_stats.st_mtime

    # Merge metadata and file information into a single dictionary
    file_data = {**metadata, **file_info}
    return file_data


# check if the path is valid ( folder or file) or not exists( invalid)
def process_paths(path):
    if os.path.isdir(path):
        return "folder"
    elif os.path.isfile(path):
        return "file"
    else:  # Path is not valid
        return 'invalid'


# find the png images when the path is a directory
def find_png_images(folder_path, settings):
    # Create a list to store the dictionaries
    rows = []

    # Loop through each file in the directory and check if it's a PNG image
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.jpg'):
            # Create a new dictionary for the image
            image_dict = {}

            # Add the settings to the dictionary
            for key, value in settings.items():
                image_dict[key] = value

            # Add the file path to the dictionary
            image_dict['Path'] = os.path.join(folder_path, file_name)

            # Add the image dictionary to the list of rows
            rows.append(image_dict)

    print(f"In the folder '{folder_path}' found {len(rows)} images")
    # Return the list of image dictionaries
    return rows


def read_excel_file(file_name):
    # Get the full path of the Excel file in the same directory as the script
    script_dir = os.path.dirname(__file__)
    file_path = os.path.join(script_dir, file_name)
    print(f"Excel filepath: {file_path}")
    # Read the Excel file into a pandas dataframe
    df = pd.read_excel(file_path)
    print(f"In the excel file found: {len(df)} rows.")
    # Create a list to store the dictionaries
    rows = []

    # Loop through each row in the dataframe and create a dictionary
    for index, row in df.iterrows():
        # Create an empty dictionary for the row
        row_dict = {}
        # Loop through each column and add it to the dictionary
        for column in df.columns:
            row_dict[column] = row[column]

        path = row_dict['Path']

        # Call the path function to examine if it is a folder or not, else not valid
        folder_type = process_paths(path)
        row_dict['folder'] = folder_type

        # If it's a directory, find all PNG images and add them to the list of rows
        if folder_type == 'folder':
            folder_rows = find_png_images(path, row_dict)
            rows.extend(folder_rows)
            for folder_row in folder_rows:
                metadata = get_image_metadata(folder_row['Path'])
                folder_row['metadata'] = metadata
        # If it's a file, add it to the list of rows
        elif folder_type == 'file':
            rows.append(row_dict)
            metadata = get_image_metadata(row_dict['Path'])
            row_dict['metadata'] = metadata

    print(f"Generally found {len(rows)} images.")
    # Return the list of dictionaries
    return rows


def copy_images(rows):
    for row in rows:
        print("Image processing...")
        path = row["Path"]
        plant = row["Plant"]
        quality = row["Quality"]

        # Create the output directory if it doesn't exist
        output_dir = os.path.join(os.path.expanduser("~"), "Downloads", "output", plant, quality)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        # Get the date from the metadata, or use the current time as a fallback
        date_keys = ["DateTimeOriginal", "DateTime", "DateTimeDigitized", "creation_time"]
        date = None
        for key in date_keys:
            if key in row["metadata"]:
                # print(f"Timestamp created by {key} key.")
                # date_str = row["metadata"][key]
                # date = datetime.strptime(date_str, "%Y:%m:%d %H:%M:%S")
                # break
                date_str = row["metadata"][key]
                try:
                    # Try to parse the date string as a Unix timestamp
                    timestamp = float(date_str)
                    date = datetime.fromtimestamp(timestamp)
                    print(f"Timestamp created by {key} key. (float)")
                    break
                except ValueError:
                    pass
                else:
                    # Parse the date string using the specified format
                    date = datetime.strptime(date_str, "%Y:%m:%d %H:%M:%S")
                    print(f"Timestamp created by {key} key.")
                    break
        if date is None:
            date = datetime.now()

        print(f"Date : {date}")
        # Construct the new filename
        timestamp = date.strftime("%Y%m%d_%H%M%S")
        print(f"timestamp : {timestamp}")
        # timestamp = date.strftime("%H%M%S")
        filename, extension = os.path.splitext(os.path.basename(path))
        title = row["Title"]
        new_filename = f"{title}_{timestamp}{extension}"

        # Copy the image to the plant subdirectory with the new filename
        output_path = os.path.join(output_dir, new_filename)
        text = {}
        text['Title'] = row["Title"]
        text['Description'] = row["Description"]
        text['Logo'] = row["Logo"]
        text['Compress(%)'] = row["Compress(%)"]
        text['Timestamp'] = date.strftime("%d%m%y%H%M%S")


        add_text_to_image(path, output_path, text)
        print("Image Saved !!!")
        print("~~~~~~~~~~~~~~~~~~~~~~~")
        # shutil.copyfile(path, output_path) # just copy the original for debugging prorposes


def add_text_to_image(image_path, output_path, text):
    # Open image
    print("Image modification...")
    with Image.open(image_path) as img:
        # Define font and text color
        font = ImageFont.truetype("arial.ttf", 40)
        text_color_title = (255, 255, 255)
        text_color_desc = (212, 251, 121)
        text_color_timestamp = (115, 253, 255)

        # Define background color
        bg_color = (128, 128, 128)
        if not text['Logo'].lower() == 'yes' and isinstance(text['Description'], float):
            x1 = 75
        else:
            x1 = 150
        # Add background color to image
        bg = Image.new('RGB', (img.width, img.height + x1), bg_color)
        bg.paste(img, (0, 0))

        # Add text to image
        draw = ImageDraw.Draw(bg)
        x = 20
        y = img.height + 10
        draw.text((x, y), text['Title'], font=font, fill=text_color_title)

        if not isinstance(text['Description'],
                          float):  # Caution, if the cell in excel is empty the type become float, if it has values it become string
            y += 30
            draw.text((x, y + 35), str(text['Description']), font=font, fill=text_color_desc)

        # Add timestamp to image

        timestamp = text['Timestamp']
        # timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        timestamp_bbox = draw.textbbox(xy=(0, 0), text='Timestamp(' + timestamp + ')', font=font)
        timestamp_width = timestamp_bbox[2] - timestamp_bbox[0]
        timestamp_height = timestamp_bbox[3] - timestamp_bbox[1]
        x = img.width - timestamp_width #timestamp SPACE ex. x = img.width - timestamp_width +/- 100
        y = img.height + 10
        draw.text((x, y), 'Timestamp(' + timestamp + ')', font=font, fill=text_color_timestamp)

        if text['Logo'].lower() == 'yes':
            script_dir = os.path.dirname(__file__)
            logo_path = os.path.join(script_dir, "david_logo.png")
            logo = Image.open(logo_path).convert("RGBA")
            logo_resized = logo.resize((300, 100))
            logo_width, logo_height = logo_resized.size
            x = img.width - logo_width + 40 #LOGO SPACE
            y = img.height - logo_height
            bg.paste(logo_resized, (x - 50, y + 150), logo_resized)

        print("Image modification Done")
        # Save modified image
        if os.path.exists(output_path):
            # If the file already exists, add a random string of one letter and one number before the extension
            filename, ext = os.path.splitext(output_path)
            new_filename = f"{filename}({choice(string.ascii_lowercase)}{random.randint(0, 9)}){ext}"
            print(f"Warning: {output_path} already exists. Saving the file as {new_filename} instead.")
            output_path = new_filename

        # Save the image to the output path
        print(f"Compressed to {text['Compress(%)']}%.")
        bg.save(output_path, quality=text['Compress(%)'])


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()
