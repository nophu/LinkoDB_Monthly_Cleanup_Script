import os
import easyocr
import pandas as pd

# Initialize the reader (English language; set gpu=True if you have a graphics card)
reader = easyocr.Reader(['en'], gpu=False)

image_folder = "screenshots_folder/"
output_data = []

# Loop through all screenshots in your directory
for filename in os.listdir(image_folder):
    if filename.endswith((".png", ".jpg", ".jpeg")):
        image_path = os.path.join(image_folder, filename)

        # Read the text from the image
        results = reader.readtext(image_path)

        # Combine all extracted fragments into one text line, or map them out
        extracted_text = " ".join([res[1] for res in results])

        output_data.append({
            "File Name": filename,
            "Extracted Data": extracted_text
        })

# Export directly to an Excel sheet for easy study management
df = pd.DataFrame(output_data)
df.to_excel("study_results.xlsx", index=False)
print("Data extraction complete! Saved to study_results.xlsx")