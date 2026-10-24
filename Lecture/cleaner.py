import os
import comtypes.client


def convert_to_pdf_with_notes(powerpoint, inp_file, out_file):
    try:
        # Open the presentation
        deck = powerpoint.Presentations.Open(inp_file)

        # Export as PDF with speaker notes
        deck.ExportAsFixedFormat(
            out_file,
            FixedFormatType = 2,    # PDF
            Intent = 1,             # Screen quality
            OutputType = 5          # Notes pages
        )
        deck.Close()
        print(f'Converted: {os.path.basename(inp_file)}')
    except Exception as e:
        print(f'Error converting {os.path.basename(inp_file)}: {e}')


def main():
    # Since the script is inside the folder, use the current working directory
    folder_path = os.getcwd() 

    # Start PowerPoint application
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    # Loop through all files in the current folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.pptx') or filename.endswith('ppt'):
            inp_file = os.path.join(folder_path, filename)

            # Strip the old extension and add .pdf
            base_name = os.path.splitext(filename)[0]
            out_file = os.path.join(folder_path, f"{base_name}.pdf")

            if os.path.exists(out_file):
                print(f"Skipping {filename}, PDF already exists.")
                continue
            convert_to_pdf_with_notes(powerpoint, inp_file, out_file)

    powerpoint.Quit()
    print("All conversions finished!")
     
if __name__ == '__main__':
    main()
