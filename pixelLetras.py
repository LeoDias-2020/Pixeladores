'''pixeladorLetras.py
'''
import os
from PIL import Image, ImageDraw, ImageFont

if __name__=='__main__':
    os.chdir('c:/users/leona/Pictures/MonaLisas/')
    im	   = Image.open('mona.png') # arquivo para discretizar
    comprim, altura = im.size

    im2		 = Image.new('RGB', (comprim, altura))
    letras 	 = ImageDraw.Draw(im2)

    fontsFolder = 'FONT_FOLDER'
    arial		= ImageFont.truetype(os.path.join(fontsFolder, 'arial.ttf'), 12)

    passosVert  = altura //100
    passosHoriz = comprim//100
    index = 0
    for y in range(0, altura, passosVert):
        for x in range(0, comprim, passosHoriz):
            letras.text((x, y), 'MonaLisa'[index%8], font=arial, fill=im.getpixel((x, y)))
            index+=1          

    im2.save(f'pixeladorLetras_{im.filename}')