#To use the image as the icon of the notification, we have to change it's extension to .ico

from PIL import Image

#open the current file
img = Image.open("python_icon.png")

#change the extension and save it
img.save("python_icon.ico")

#print something to confirm that the extension is converted
print("Converted")