import PySimpleGUI as sg
import sys
import os
import genanki
from random import randrange
import glob
from pptx_tools.utils import save_pptx_as_png

def pptx_to_images(pptx_filename):
    print("Creating image files from pptx...")
    current_dir = os.getcwd()
    pptx = os.path.join(current_dir, pptx_filename)
    image_dir = "images"
    image_path = os.path.join(current_dir, image_dir)
    save_pptx_as_png(image_path, pptx, overwrite_folder=True)
    return len([name for name in os.listdir(image_path) if os.path.isfile(os.path.join(image_path, name))])

def clean_up_images():
    print("Cleaning up images...")The
    files = glob.glob('images/*.png')
    for file in files:
        os.remove(file)


def display_gui(num_slides):
    print("Displaying GUI...")
    output = []
    try:
        slide = 1
        input = sg.InputText(key='input')
        image = sg.Image(f'images/Slide{slide}.png', key='image')
        back_button = sg.Button('Back')
        next_button = sg.Button('Next')
        same_button = sg.Button('Repeat Image')
        counter = sg.Text(f"1/{num_slides}", key='counter')
        layout = [ 
            [sg.Text('Term: '), input, image],
            [back_button, same_button, next_button, sg.Push(), counter]
                ]
        window = sg.Window('PPTX to Anki Console', layout, finalize=True)
        window.bind('<F1>', 'Back')
        window.bind('<F2>', 'Repeat Image')
        window.bind('<F3>', 'Next')
        
        while True:
            event, values = window.read()
            if event == sg.WIN_CLOSED:
                break
            elif event == 'Back' and slide > 1:
                if values['input'] != '':
                    output.append((values['input'], f'images/Slide{slide}.png'))
                slide = slide - 1
                window['image'].update(f'images/Slide{slide}.png')
                window['input'].update('')
                window['counter'].update(f"{slide}/{num_slides}")
            elif event == 'Repeat Image':
                if values['input'] != '':
                    output.append((values['input'], f'images/Slide{slide}.png'))
                window['input'].update('')
            elif event == 'Next' and slide < num_slides:
                if values['input'] != '':
                    output.append((values['input'], f'images/Slide{slide}.png'))
                slide = slide + 1
                window['image'].update(f'images/Slide{slide}.png')
                window['input'].update('')
                window['counter'].update(f"{slide}/{num_slides}")
            elif event == 'input' and slide < num_slides:
                if values['input'] != '':
                    output.append((values['input'], f'images/Slide{slide}.png'))
                slide = slide + 1
                window['image'].update(f'images/Slide{slide}.png')
                window['input'].update('')
        window.close()
    except Exception as e:
        print(e)
    return output

def anki_generation(arr, pptx_name):
    # Define the model for the flashcards
    print("Generating Anki Deck...")
    model = genanki.Model(
    randrange(9999999999),
    'Term-Image Model',
    fields=[
        {'name': 'Question'},
        {'name': 'Answer'},
    ],
    templates=[
        {
        'name': 'Card 1',
        'qfmt': '{{Question}}',
        'afmt': '{{Answer}}',
        },
    ],
    css="""
        .card {
        font-family: arial;
        font-size: 20px;
        text-align: center;
        color: black;
        background-color: white;
        }
        .card img {
        max-width: 100%;
        height: auto;
        }
    """
    )
    deck = genanki.Deck(
    randrange(9999999999),
    pptx_name.split('.pptx')[0]
    )
    media_files = set()
    for item in arr:
        term = item[0]
        imagepath = item[1]
        print(imagepath)
        media_files.add(imagepath)

        if not os.path.exists(imagepath):
            raise FileNotFoundError(f"The image file {imagepath} does not exist")
        
        # Create the note
        note = genanki.Note(
            model=model,
            fields=[term, f'<img src="{os.path.basename(imagepath)}">']
        )
        deck.add_note(note)

    package = genanki.Package(deck)

    package.media_files = list(media_files)

    package.write_to_file(f"{pptx_name.split('.pptx')[0]}.apkg")
    





num_slides = pptx_to_images(sys.argv[1])
output = display_gui(num_slides)
anki_generation(output, sys.argv[1])
clean_up_images()







