import os
import win32com.client

def hide_non_base_layers(doc):
    for layer in doc.ArtLayers:
        if not 'base' in layer.name.lower():
            layer.Visible = False


def export_layers_as_images(doc, options, folder):
    for layer in doc.ArtLayers:

        if 'base' in layer.name.lower():
            continue

        hide_non_base_layers(doc)
        layer.Visible = True

        outpath = os.path.join(folder, f"{layer.name}.jpg")
        doc.Export(ExportIn=outpath, ExportAs=2, Options=options)

        print(f'saved {layer.name}')

def main():

    print('This script exports an image for each layer of a PSD.')

    filepath = 'unset'
    
    while filepath and not os.path.exists(filepath):
        print(filepath)
        filepath = input("Press enter to use active doc, or paste a path:\n>>>")
    
    psApp = win32com.client.Dispatch("Photoshop.Application")

    if filepath:
        psApp.Open(filepath)

    doc = psApp.Application.ActiveDocument

    options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    options.Format = 6 #JPEG
    options.Quality = 80

    if not filepath:
        folder = os.path.dirname(doc.path)
    else:
        folder = os.path.dirname(filepath)
    folder = os.path.join(folder, 'export')

    if not os.path.exists(folder):
        os.makedirs(folder)

    export_layers_as_images(doc, options, folder)

if __name__ == "__main__":
    main()