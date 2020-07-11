import os
import win32com.client

def hide_non_base_layers(doc):
    for layer in doc.ArtLayers:
        if not 'base' in layer.name.lower():
            layer.Visible = False


def export_layers_as_images(doc, options):
    for layer in doc.ArtLayers:
        hide_non_base_layers(doc)
        layer.Visible = True
        outpath = os.path.join(os.getcwd(), f"{layer.name}.jpg")
        doc.Export(ExportIn=outpath, ExportAs=2, Options=options)

psApp = win32com.client.Dispatch("Photoshop.Application")

filepath = os.path.join(os.getcwd(), 'test.psd')

psApp.Open(filepath)

doc = psApp.Application.ActiveDocument

options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
options.Format = 6 #JPEG
options.Quality = 80

hide_non_base_layers(doc)
export_layers_as_images(doc, options)