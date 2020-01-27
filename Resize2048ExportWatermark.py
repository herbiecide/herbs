# This script resizes images to have a longest edge of 2048 pixels,
# Adds a simple text watermark at the bottom,
# Then exports as a jpg.
# 2048 pixels is the maximum dimension for Facebook, Instagram, and Twitter.

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

maxDim=2048 # Maximum dimension for web size

if len(app.Documents) > 0:

        doc = app.ActiveDocument
        imageHeight=doc.height
        imageWidth=doc.width

        ###### Add Watermark ######
        # Set the font color
        watermarkColor=Dispatch("Photoshop.SolidColor")
        watermarkColor.RGB.Red = 225
        watermarkColor.RGB.Green = 225
        watermarkColor.RGB.Blue = 225

        # add a new text layer to document and apply the text color
        newTextLayer = doc.ArtLayers.Add()
        psTextLayer = 2     # from enum PsLayerKind
        newTextLayer.Kind = psTextLayer
        newTextLayer.TextItem.Font="Roboto" #This must be the PostScript Name
        newTextLayer.TextItem.Contents = "@Herbiecide"
        newTextLayer.TextItem.Position = [imageWidth/2, imageHeight*98/100]
        newTextLayer.TextItem.Size = imageHeight/200
        newTextLayer.TextItem.Justification=2 #2 = center justified
        newTextLayer.TextItem.Color=watermarkColor
        newTextLayer.fillOpacity=80

        ###### Resize #######
        # Check image is bigger than desired web size
        if (imageWidth>maxDim) or (imageHeight>maxDim):
            # Check which side is longer and resize accordingly
            if(imageWidth>=imageHeight):

                doc.ResizeImage(maxDim,None, None, 4) #4 is BICUBIC

            else:
                doc.ResizeImage(None, maxDim, None, 4) #4 is BICUBIC
        else:
            print("No need to resize. Image is too small")

else:
    print("You must have at least one open document to run this script!")


# Web export options
options = Dispatch('Photoshop.ExportOptionsSaveForWeb')
options.quality = 100
options.format = 6 #6 is jpg
options.optimized = False

newName = "web-size_"+doc.name
newPath = doc.path+newName
doc.Export(ExportIn=newPath, ExportAs=2, Options=options)

# OPTIONAL: close and do not save changes
# doc.close(SaveOptions.DONOTSAVECHANGES);


# set the app preference the way it was before the operation
app.Preferences.RulerUnits = ogRulerUnits
app.Preferences.TypeUnits = ogTypeUnits
