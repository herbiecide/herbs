# This script adds a simple text watermark at the bottom of an image in photoshop.


from win32com.client import Dispatch, GetActiveObject


# Start up Photoshop application
# Or get Reference to already running Photoshop application instance
# app = Dispatch('Photoshop.Application')
app = GetActiveObject("Photoshop.Application")

# Save current settings
ogRulerUnits = app.Preferences.RulerUnits
ogTypeUnits = app.Preferences.TypeUnits

# Set to pixels
app.Preferences.RulerUnits = 1 #1=pixels
app.Preferences.TypeUnits = 1 #1=pixels

if len(app.Documents) > 0:

        doc = app.ActiveDocument
        imageHeight=doc.height
        imageWidth=doc.width

        # Set the font color
        watermarkColor=Dispatch("Photoshop.SolidColor")
        watermarkColor.RGB.Red = 225
        watermarkColor.RGB.Green = 225
        watermarkColor.RGB.Blue = 225

        # add a new text layer to document and apply the text color
        newTextLayer = doc.ArtLayers.Add()
        psTextLayer = 2     # from enum PsLayerKind
        newTextLayer.Kind = psTextLayer
        newTextLayer.TextItem.Contents = "@Herbiecide"
        newTextLayer.TextItem.Position = [imageWidth/2, imageHeight*98/100]
        newTextLayer.TextItem.Size = imageHeight/200
        newTextLayer.TextItem.Justification=2 #2 = center justified
        newTextLayer.TextItem.Color=watermarkColor
        newTextLayer.fillOpacity=80

else:
    print("You must have at least one open document to run this script!")




#set the app preference the way it was before the operation
app.Preferences.RulerUnits = ogRulerUnits
app.Preferences.TypeUnits = ogTypeUnits
