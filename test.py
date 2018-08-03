from reportlab.pdfgen import canvas
import os

os.chdir('C:\\Users\\remyo\\Desktop')
c = canvas.Canvas("test.pdf")
from reportlab.lib.units import inch, cm, mm
from reportlab.lib.colors import pink, black, red, blue, green, gray
from reportlab.lib.pagesizes import letter, A4
width, height = A4

c.setStrokeColor(gray)

x_nodes = 30
y_nodes = 30
spacing = .2

grid_spacing = [spacing*inch*i for i in range(1, x_nodes + 1)]
grid_dimension = spacing*25.4*mm*x_nodes
grid_origin = ((width - grid_dimension)/2, (height - grid_dimension)/2)  

c.translate(grid_origin[0], grid_origin[1])
c.grid(grid_spacing, grid_spacing)
c.showPage()
c.save()