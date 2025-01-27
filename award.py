from PIL import Image, ImageDraw, ImageFont
import os
from datetime import datetime
from openpyxl import load_workbook


# Function to validate input with repeated prompts
def get_valid_input(prompt, valid_options):
    while True:
        value = input(prompt).strip().lower()
        if value in valid_options:
            return value
        print(f"Invalid input. Choose from {', '.join(valid_options)}.")


# Function to draw text with a border (outline)
def draw_text_with_border(draw, text, position, font, text_color, border_color, border_thickness):
    # Draw border by rendering text multiple times around the original position
    x, y = position
    for offset_x in range(-border_thickness, border_thickness + 1):
        for offset_y in range(-border_thickness, border_thickness + 1):
            # Skip the center position (to avoid drawing the text multiple times in the exact same spot)
            if offset_x == 0 and offset_y == 0:
                continue
            # Draw the border (outline)
            draw.text((x + offset_x, y + offset_y), text, font=font, fill=border_color)

    # Draw the main text over the border (in the desired text color)
    draw.text(position, text, font=font, fill=text_color)


# Ask for input details
color_input = get_valid_input("\nEnter the color (options: b for Bronze, s for Silver, g for Gold): ", ["b", "s", "g"])

# Map the color input to the full color name
if color_input == "b":
    color = "bronze"
elif color_input == "s":
    color = "silver"
elif color_input == "g":
    color = "gold"

# Ask for the activator/hunter option
type_option = get_valid_input("\n\t  Is it an activator (a) or a hunter (h)? (options: a, h): ", ["a", "h"])

# Map single-letter input to full words
type_option = "activator" if type_option == "a" else "hunter"

name = input("\n\t      Enter the name to be added (e.g., 'Erwin - PA3EFR'): ").strip()
serial_number = input("\n\t\t\t\t\t  Enter the serial number: ").strip()

# Get today's date
today_date = datetime.now().strftime("%Y-%m-%d")

# Construct the correct input file path based on color and activator/hunter
input_directory = color.capitalize()  # Directory name: Bronze, Silver, Gold
input_image_name = f"{color.capitalize()}{type_option.capitalize()}.jpg"
input_image_path = os.path.join(input_directory, input_image_name)

# Check if the file exists
if not os.path.exists(input_image_path):
    print(f"The file '{input_image_path}' does not exist. Please check the name and try again.")
    exit()

# Create the output file name
sanitized_name = name.replace(" ", "-").replace("/", "-")
output_pdf_name = f"{serial_number}_{color}_{type_option}_{sanitized_name}.pdf"
output_pdf_path = os.path.join(input_directory, output_pdf_name)

try:
    # Open the image
    img = Image.open(input_image_path)

    # Create a drawing object
    draw = ImageDraw.Draw(img)

    # Choose a font and size (adjust the path if necessary)
    font_size_name = int(150)  # Fixed font size
    font_size_number = int(150)  # Fixed font size
    font_size_date = int(150)  # Fixed font size

    try:
        font_name = ImageFont.truetype("Bodoni Bd BT Bold.ttf", font_size_name)
        font_number = ImageFont.truetype("arial.ttf", font_size_number)
        font_date = ImageFont.truetype("arial.ttf", font_size_date)
    except IOError:
        # Fallback to a default font if arial.ttf is not available
        font_name = ImageFont.load_default()
        font_number = ImageFont.load_default()
        font_date = ImageFont.load_default()

    # Calculate the position of the name
    text_name_bbox = draw.textbbox((0, 0), name, font=font_name)  # (left, top, right, bottom)
    text_name_width = text_name_bbox[2] - text_name_bbox[0]
    name_position = (
        1500,  # 1500 pix from left
        200  # 140 pix from the top (140 pixels at 300 DPI)
    )

    # Calculate the position of the serial number
    text_number_bbox = draw.textbbox((0, 0), serial_number, font=font_number)  # (left, top, right, bottom)
    text_number_width = text_number_bbox[2] - text_number_bbox[0]
    number_position = (
        img.size[0] - text_number_width - 10,  # 10 pix from the right edge
        img.size[1] - text_number_bbox[3] - 10  # 10 pix above the bottom
    )

    # Calculate the position of the date
    text_date_bbox = draw.textbbox((0, 0), today_date, font=font_date)  # (left, top, right, bottom)
    date_position = (
        100,  # 100 pixels from the left edge
        img.size[1] - text_date_bbox[3] - 100  # 100 pixels from the bottom edge
    )

    # Add the name with a white border and black text
    draw_text_with_border(draw, name, name_position, font_name, (0, 0, 0), (255, 255, 255), 5)  # Border thickness = 5

    # Add the serial number with a white border and black text
    draw_text_with_border(draw, serial_number, number_position, font_number, (0, 0, 0), (255, 255, 255), 5)

    # Add the date with a white border and red text
    draw_text_with_border(draw, today_date, date_position, font_date, (255, 0, 0), (255, 255, 255), 5)

    # Convert the image to RGB (necessary for PDF export)
    if img.mode in ("RGBA", "P"):  # Check if conversion is needed
        img = img.convert("RGB")

    # Save the image as a PDF in the correct directory
    img.save(output_pdf_path, "PDF")
    print(f"\n\n>  The image with added text has been saved as '{output_pdf_path}'.")

    # Now add the information to the Excel file
    try:
        # Load the existing workbook
        workbook = load_workbook("AwardGrantsOverview.xlsx")
        sheet = workbook.active

        # Add a new row with the data
        sheet.append([color.capitalize(), type_option.capitalize(), name, serial_number, output_pdf_name])

        # Save the changes to the Excel file
        workbook.save("AwardGrantsOverview.xlsx")
        print(">  The data has been added to 'AwardGrantsOverview.xlsx'.")

    except Exception as e:
        print(f"An error occurred while updating the Excel file: {e}")

except Exception as e:
    print(f"An error occurred: {e}")
