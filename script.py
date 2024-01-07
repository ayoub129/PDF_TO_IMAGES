import fitz  # PyMuPDF
from openpyxl import Workbook
import os

def extract_images(pdf_path, output_directory):
    images_and_links = []

    # Open the PDF file
    pdf_document = fitz.open(pdf_path)

    # Create a new Excel workbook
    workbook = Workbook()
    worksheet = workbook.active

    # Add headers to the Excel sheet
    worksheet.append(["Image File", "Link"])

    for page_number in range(pdf_document.page_count):
        page = pdf_document[page_number]

        # Extract images
        images = page.get_images(full=True)
        for img_index, image_info in enumerate(images):
            img_index += 1  # Start index from 1

            # Extract image data
            base_image = pdf_document.extract_image(image_info[0])
            image_bytes = base_image["image"]

            # Save image to file
            image_filename = f"{output_directory}/image_{page_number + 1}_{img_index}.png"
            with open(image_filename, "wb") as image_file:
                image_file.write(image_bytes)

            # Extract links
            links = page.get_links()
            link_url = None  # Initialize link_url as None
            for link_index, link_info in enumerate(links):
                    link_url = link_info['uri']
                    break

            images_and_links.append((image_filename, link_url))

    # Populate the Excel sheet
    for entry in images_and_links:
        worksheet.append(entry)

    # Save the Excel workbook
    excel_filename = "images_and_links.xlsx"
    workbook.save(excel_filename)

    print(f"Images and links extracted and saved to {excel_filename}")

if __name__ == "__main__":
    pdf_path = "gift.pdf"  # Replace with your PDF file path
    output_directory = "output_images"  # Replace with the desired output directory
    os.makedirs(output_directory, exist_ok=True)
    extract_images(pdf_path, output_directory)
