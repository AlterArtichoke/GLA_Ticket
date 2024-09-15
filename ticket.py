import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Load the registration data from an Excel file
excel_file = 'email_details.xlsx'  # Ensure this file is in the same directory or provide the path
df = pd.read_excel(excel_file)

# Create a "Tickets" folder if it doesn't exist
output_folder = 'Tickets'
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Path to your predesigned ticket template
template_path = 'final.png'

# Function to add name and roll number to the ticket template
def create_ticket_from_template(roll_number, name, output_path):
    # Open the predesigned ticket template
    ticket = Image.open(template_path)
    draw = ImageDraw.Draw(ticket)

    # Get the dimensions of the image
    img_width, img_height = ticket.size

    # Define the text to be added
    text_name = f"{name}"
    text_roll = f"{roll_number}"

    # Manually increase font size to make text larger
    font_size = 100  # You can adjust this value further
    font_path = "KaushanScript-Regular.ttf"  # Path to your font

    # Load the font with the updated size
    try:
        font = ImageFont.truetype(font_path, font_size)
    except IOError:
        font = ImageFont.load_default()  # Fallback to default font if custom font not found

    # Calculate the text box dimensions (for centering)
    name_bbox = draw.textbbox((0, 0), text_name, font=font)
    roll_bbox = draw.textbbox((0, 0), text_roll, font=font)

    name_width = name_bbox[2] - name_bbox[0]
    roll_width = roll_bbox[2] - roll_bbox[0]

    # Define the coordinates for "Name" and "Reg.No" to be lower and centered
    # Adjust the spacing between Name and Reg.No by modifying the y-coordinates
    name_position = ((img_width - name_width) // 2, img_height - 610)  # Adjust this value for name
    roll_number_position = ((img_width - roll_width) // 2, img_height - 610 + 130)  # Adjust this value for roll number

    # Define stroke properties (border)
    stroke_width = 5  # Thickness of the border
    stroke_color = (0, 0, 0)  # Brown color for the stroke

    # Draw bordered "Name" text
    for offset in [(x, y) for x in range(-stroke_width, stroke_width+1) for y in range(-stroke_width, stroke_width+1)]:
        draw.text((name_position[0] + offset[0], name_position[1] + offset[1]), text_name, font=font, fill=stroke_color)

    # Draw bordered "Reg.No" text
    for offset in [(x, y) for x in range(-stroke_width, stroke_width+1) for y in range(-stroke_width, stroke_width+1)]:
        draw.text((roll_number_position[0] + offset[0], roll_number_position[1] + offset[1]), text_roll, font=font, fill=stroke_color)

    # Draw the actual "Name" and "Roll No." text in orange, centered
    draw.text(name_position, text_name, font=font, fill=(255, 235, 0))  # Orange text color for Name
    draw.text(roll_number_position, text_roll, font=font, fill=(255, 235, 0))  # Orange text color for Reg.No

    # Save the new ticket with the name and roll number
    ticket.save(output_path)

# Generate image tickets for all participants
for _, row in df.iterrows():
    roll_number = row['Roll Number']
    name = row['Name']
    
    # Generate a unique ticket image and save it in the "Tickets" folder
    ticket_image = os.path.join(output_folder, f"{roll_number}_ticket.png")
    create_ticket_from_template(roll_number, name, ticket_image)

print(f"All tickets saved in the {output_folder} folder.")
