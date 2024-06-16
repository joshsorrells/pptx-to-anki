import PySimpleGUI as sg
import aspose.slides as slides
import aspose.pydrawing as drawing
import sys
import os

# def pptx_to_images(pptx_filename):
#     print("Converting pptx to images...")
#     with slides.Presentation(pptx_filename) as presentation:
#         for slide in presentation.slides:
#             slide.get_thumbnail(2, 2).save("images/presentation_slide_{0}.png".format(str(slide.slide_number)), drawing.imaging.ImageFormat.png)

def pptx_to_images(pptx_filename):
    print("Converting pptx to images...")
    # Desired thumbnail dimensions
    desired_width = 960
    desired_height = 540

    # Load the presentation
    with slides.Presentation(pptx_filename) as presentation:
        # Loop through each slide in the presentation
        for i, slide in enumerate(presentation.slides):
            # Calculate scale ratios for width and height
            scale_width = desired_width / presentation.slide_size.size.width
            scale_height = desired_height / presentation.slide_size.size.height
            
            # Use the smaller scale to ensure the image fits within the desired dimensions
            scale = min(scale_width, scale_height)
            
            # Get the thumbnail of the slide with the specified scale
            thumbnail = slide.get_thumbnail(scale, scale)
            
            # Save the thumbnail image as a PNG file
            thumbnail.save(f"images/presentation_slide_{i + 1}.png", drawing.imaging.ImageFormat.png)

def clean_up_images():
    os.system("rm -rf images/*.png")

def display_gui():
    print("Displaying GUI...")
    try:
        while True:
            layout = [ 
                [sg.Text('Term: '), sg.Multiline(key='input', write_only=True, size=(60,10), reroute_cprint=True), sg.Image('images/presentation_slide_70.png')]
                 ]
            window = sg.Window('PPTX to Anki Console', layout, keep_on_top=True)
            event, values = window.read()
            if event == sg.WIN_CLOSED:
                # clean_up_images()
                break
        window.close()
    except:
        clean_up_images()


#pptx_to_images(sys.argv[1])
display_gui()







