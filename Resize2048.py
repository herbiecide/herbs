# This script resizes images to have a longest edge of 2048 pixels
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

        # Check image is bigger than desired web size
        if (imageWidth>maxDim) and (imageHeight>maxDim):
            # Check which side is longer and resize accordingly
            if(imageWidth>=imageHeight):

                doc.ResizeImage(maxDim,None, None,None)

            else:
                doc.ResizeImage(None, maxDim, None,None)
        else:
            print("No need to resize. Image is too small")

else:
    print("You must have at least one open document to run this script!")



# set the app preference the way it was before the operation
app.Preferences.RulerUnits = ogRulerUnits
app.Preferences.TypeUnits = ogTypeUnits
