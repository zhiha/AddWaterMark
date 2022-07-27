import glob, fitz
from pathlib import Path
import img2pdf

# To get better resolution
zoom_x = 4.0  # horizontal zoom
zoom_y = 4.0  # vertical zoom
mat = fitz.Matrix(zoom_x, zoom_y)  # zoom factor 2 in each dimension

path = 'E:\Projects\Project\PyPDF\CustomWatermark\\'
all_files = glob.glob(path + "*.pdf")

for filename in all_files:
    doc = fitz.open(filename)  # open document
    for page in doc:  # iterate through the pages
        pix = page.get_pixmap(matrix=mat)  # render page to an image
        pix.save("E:\Projects\Project\PyPDF\CustomWatermark/page-%i.png" % page.number)  # store image as a PNG

with open("Stickers.pdf","wb") as f:
    f.write(img2pdf.convert([str(path) for path in Path('./').glob('*.png')]))