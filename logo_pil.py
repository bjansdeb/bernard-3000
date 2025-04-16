from PIL import Image, ImageDraw, ImageFont

text= "Brussels\nConservatories\nLibrary"
size = (100, 40)

draw = Image.new(mode="RGBA", size=size, color=(255, 255, 255, 0)) 
# (255, 255, 255, 0)
font = ImageFont.truetype("NotoSansBold.ttf", 9)
d = ImageDraw.Draw(draw)

d.text((0, 0), text, font=font, fill=(0, 0, 0))

draw.show()
draw.save("conservatoire.png")