import os
import comtypes.client

# Since the script is inside the folder, use the current working directory
folder_path = os.getcwd()

# Start PowerPoint application
powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
powerpoint.Visible = 1

# Loop through all files in the current folder
for filename in os.listdir(folder_path):
    if filename.endswith(".pptx") or filename.endswith(".ppt"):
        input_file = os.path.join(folder_path, filename)
        
        # Strip the old extension and add .pdf
        base_name = os.path.splitext(filename)[0]
        output_file = os.path.join(folder_path, f"{base_name}.pdf")
        
        if os.path.exists(output_file):
            print(f"Skipping {filename}, PDF already exists.")
            continue
            
        print(f"Converting {filename}...")
        try:
            deck = powerpoint.Presentations.Open(input_file)
            deck.SaveAs(output_file, 32)  # 32 is the format code for PDF
            deck.Close()
        except Exception as e:
            print(f"Error converting {filename}: {e}")

powerpoint.Quit()
print("All conversions finished!")
