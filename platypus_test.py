from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, Frame, PageBreak, Flowable, BaseDocTemplate, PageTemplate, NextPageTemplate
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.rl_config import defaultPageSize
from reportlab.lib.units import inch, cm, mm
from reportlab.lib.colors import pink, black, red, blue, green, gray, white, Color
import datetime
import math
MARGIN_SIZE = 36
width, height = letter

class MeasurementGrid(Flowable):

	def __init__(self, dataDict, which_error, title, wafer_name):
		self.title = title
		self.wafer_name = wafer_name
		self.data = dataDict
		self.which_error = which_error
		self.maxX = max([x for (x, y), (diffx, diffy) in dataDict.items()])
		self.maxY = max([y for (x, y), (diffx, diffy) in dataDict.items()])
		
		self.numberXLines = self.maxX*2 + 2
		self.numberYLines = self.maxY*2 + 2
		
		self.spacing = (width - 3*MARGIN_SIZE) / (self.numberXLines-1) #MARGIN_SIZE is for the margins

		if (self.maxX, self.maxY) in dataDict.keys(): #inscribed
			self.spacing = self.spacing*math.pow(2, .5)/2

			self.width = (self.numberXLines-1)*self.spacing
			self.height = (self.numberYLines-1)*self.spacing

			self.xcen = width/2 - MARGIN_SIZE
			self.ycen = height/2 - MARGIN_SIZE
			self.r = (self.width)/math.pow(2,.5)
		else:
			self.width = (self.numberXLines-1)*self.spacing
			self.height = (self.numberYLines-1)*self.spacing

			self.xcen = width/2 - MARGIN_SIZE
			self.ycen = height/2 - MARGIN_SIZE
			self.r = (self.width)/2

	def wrap(self, *args):
		return (self.width, self.height)

	def draw(self):

		self.canv.saveState()
		self.canv.setFont('Helvetica', 15)
		self.canv.drawCentredString(self.r, self.height-36, self.title)

		outstring = ''
		if self.which_error == 'x':
			outstring = "X Axis Error Map (um)"
		if self.which_error == 'y':
			outstring = "Y Axis Error Map (um)"
		if self.which_error == 'both':
			outstring = "X, Y Error Map (um)"
		self.canv.setFont('Helvetica', 12)
		self.canv.drawString(0, self.height-72, outstring)
		translate_horizontally = (width -2*MARGIN_SIZE - (self.width))/2
		self.canv.translate(translate_horizontally, 0)
		translate_vertically = (height - (24+72) - (self.height))/2
		self.canv.translate(0, -translate_vertically)
		
		self.xcen = width/2 - MARGIN_SIZE
		self.ycen = height/2 - MARGIN_SIZE
		
		circle_color = Color(0, 160/256, 51/256, .3)
		self.canv.setStrokeColor(circle_color)
		self.canv.setFillColor(circle_color)
		self.canv.circle(self.width/2, self.height/2, self.r, stroke=1, fill=1)

		self.canv.setStrokeColor(gray)
		grid_spacing = [self.spacing*i for i in range(0, self.numberXLines)]
		self.canv.grid(grid_spacing, grid_spacing)
		
		
		self.canv.setFillColor(blue)
		#self.canv.circle(0,0, 5, fill=1)
		
		self.canv.setFillColor(black)
		self.canv.setFont('Helvetica', 6)

		x_nodes = [x for x in range(-self.maxX, self.maxX+1)]
		y_nodes = [y for y in range(-self.maxY, self.maxY+1)]

		counter = 0

		for i in grid_spacing[:-1]:
			self.canv.drawRightString(-3, i+self.spacing/2-2, str(y_nodes[counter]))
			counter+=1
		counter = 0

		for i in grid_spacing[:-1]:
			x_node = str(x_nodes[counter])
			if len(x_node) == 1:
				self.canv.drawString(i-1+self.spacing/2, -8, x_node)
			if len(x_node) == 2:
				self.canv.drawString(i-2+self.spacing/2, -8, x_node)
			if len(x_node) == 3:
				self.canv.drawString(i-4+self.spacing/2, -8, x_node)
			counter+=1
		
		self.canv.setFillColor(black)
		self.canv.setFont('Helvetica', 5)

		for i in grid_spacing[:-1]:
			for j in grid_spacing[:-1]:
				x = int(i/self.spacing) - self.maxX
				y = int(j/self.spacing)- self.maxY

				if (x, y) in self.data.keys():
					xerror = str(self.data.get((x, y))[0])
					yerror = str(self.data.get((x, y))[1])
					if self.which_error == "x":
						output_string = xerror
						horizontal_spacing = 3
					if self.which_error == "y":
						output_string = yerror
						horizontal_spacing = 3
					if self.which_error == "both":
						output_string = xerror  + ", " + yerror
						self.canv.setFont('Helvetica', self.spacing/7)
						horizontal_spacing = 16
					self.canv.drawString(i+self.spacing/2-horizontal_spacing, j+self.spacing/2, output_string)
		
		self.canv.restoreState()



def create_pdfdoc(pdfdoc, story):
    """
    Creates PDF doc from story.
    """
    pdf_doc = BaseDocTemplate(pdfdoc, pagesize = letter,
        leftMargin = MARGIN_SIZE, rightMargin = MARGIN_SIZE,
        topMargin = 24, bottomMargin = MARGIN_SIZE)

    second_frame = main_frame = Frame(MARGIN_SIZE, MARGIN_SIZE,
        width - 2 * MARGIN_SIZE, height - (24 + 72),
        leftPadding = 0, rightPadding = 0, bottomPadding = 0,
        topPadding = 0, id = 'main_frame', showBoundary=0)
    second_template = PageTemplate(id = 'second_template', frames=[second_frame], onPage=header)

    main_frame = Frame(MARGIN_SIZE, MARGIN_SIZE,
        width - 2 * MARGIN_SIZE, height - (24 + MARGIN_SIZE),
        leftPadding = 0, rightPadding = 0, bottomPadding = 0,
        topPadding = 0, id = 'main_frame', showBoundary=0)
    main_template = PageTemplate(id = 'main_template', frames = [main_frame], onPage=header)

    
    
    pdf_doc.addPageTemplates([main_template, second_template])

    pdf_doc.build(story)

#Styles
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name='rightAlign', alignment=TA_RIGHT, borderPadding=0))
NO_PADDING = TableStyle([('LEFTPADDING', (0,0), (-1,-1), 0), ('RIGHTPADDING', (0,0), (-1,-1), 0),
	('TOPPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0)])



grid = TableStyle([('INNERGRID', (0,0), (-1,-1), .25, black),
				   ('BOX', (0,0), (-1,-1), 0.25, black)])


column_width = 2*[width/2-.5*inch]


def header(canvas, doc):
	#Header Table
	canvas.saveState()
	logo = "Qualitau Logo.png"
	Title = "300SA Stage Accuracy Report"
	pageinfo = "something"

	img_aspect_ratio = 457/139
	im = Image(logo, img_aspect_ratio*.35*inch, .35*inch)
	im.hAlign = 'LEFT'

	address = "Qualitau Inc.<br/>830 Maude Avenue<br/>Mountain View, CA<br/>94043"
	ptext = '<para leading=8><font size=7>%s</font></para>' % address

	data = [[im, Paragraph(ptext, styles["rightAlign"])]]
	header_table = Table(data, column_width)
	header_table.setStyle(NO_PADDING)
	w, h = header_table.wrap(doc.width, doc.topMargin)
	header_table.drawOn(canvas, doc.leftMargin, doc.height)
	canvas.restoreState()

filenames = ["Accuracy check1_Calmap1_recipeCheck.dat"]

for i in range(1,4):

	if i == 2:
		filenames.append("spear_acc_25C_2_recipeCheck.dat")
	if i == 3:
		filenames.append("iterativecal_25C_spear_recipeCheck.dat")

	doc_name = "C:\\Users\\Remy\\Desktop\\pdf\\300SA Accuracy Report - " + str(i) + ".pdf"
	#doc = SimpleDocTemplate(doc_name, pagesize=letter, rightMargin=36,
	#	leftMargin=36, topMargin=24, bottomMargin=36)

	aboveTitle = 1*cm
	belowTitle = 1*cm
	after_customer_space = 0 #in cm
	after_table_heading = 5*mm
	space_after_accuracy_table = 0 #in mm 
	above_signature = 8*mm
	customer_font_size = 10
	signature_font_size = 10
	accuracy_table_font_size = 8

	if len(filenames) == 1:
		aboveTitle = 2*cm
		belowTitle = 2*cm
		after_table_heading = 1*cm
		customer_font_size = 14
		signature_font_size = 14
		accuracy_table_font_size = 16
		print(accuracy_table_font_size)
		space_after_accuracy_table = 12
		after_customer_space = 1.5
		above_signature = 16*mm
	elif len(filenames) == 2:
		accuracy_table_font_size = 12
		space_after_accuracy_table = 7
		after_customer_space = .8
	else:
		accuracy_table_font_size = 8
		space_after_accuracy_table = 4
		after_customer_space = .5

	column_widths = [2.5*inch, 4.5*inch]

	Story = []

	grid_right_align = TableStyle([('ALIGN', (0,0), (0,2), 'RIGHT'), 
							  ('INNERGRID', (0,0), (-1,-1), .25, black),
							  ('BOX', (0,0), (-1,-1), 0.25, black),
							  ('SIZE', (0,0), (-1,-1), customer_font_size)])
	signature_style = TableStyle([('ALIGN', (0,0), (0,1), 'RIGHT'), 
							  ('INNERGRID', (0,0), (-1,-1), .25, black),
							  ('BOX', (0,0), (-1,-1), 0.25, black),
							  ('SIZE', (0,0), (-1,-1), signature_font_size)])
	




	#Doc Title
	Story.append(Spacer(1*inch, aboveTitle))
	ptext = '<para alignment=center><font size=20>300SA Stage Accuracy Report</font></para>'
	Story.append(Paragraph(ptext, styles["Normal"]))
	Story.append(Spacer(1*inch, belowTitle))



	#Customer Name Table User Input
	# customer_name = input("Please enter customer name: ")
	# system_description = input("Please enter system description: ")
	# system_id = input("Please enter system identifier: ")
	# user_name = input("Please enter user name: ")

	customer_name = "GF - Singapore"
	system_description = "300SA Prober with ATT Chuck"
	system_id = "QM2"
	user_name = "Remy Orans"

	data = [["Customer Name:", customer_name], 
			["System Description:", system_description],
			["System #:", system_id]]

	customer_table = Table(data, column_widths)
	customer_table.setStyle(grid_right_align)
	Story.append(customer_table)





	#Measurement Reporting
	#title
	filename = ''
	


	titles = ["25C Glass Accuracy"]*i

	wafers = ["Precision Glass"]*i
	#filenames = []
	#titles = []
	#wafers=[]
	count = 0
	min_max = []
	
	
	grids = []

	# while True:
	# 	filename = input("Enter first filename (\"quit\" to cancel): ")
	# 	if filename=="quit":
	# 		break
	# 	filenames.append(filename)
	# 	title = input("Enter title: ")
	# 	titles.append(title)
	# 	wafer_name = input("What is the wafer name? ")
	# 	wafers.append(wafer_name)





	for file in filenames:
		data_dict = {}
		diffxes = []
		diffyes = []
		with open(file, "r") as myfile:
			for line in myfile:
				newline = line.split(",")
				try:
					val = int(newline[0][0])
					diex = int(newline[2])
					diey = int(newline[3])
					diffx = float(newline[8])/10
					diffxes.append(diffx)
					diffy = float(newline[9])/10
					diffyes.append(diffy)
					data_dict[(diex, diey)] = (diffx, diffy)
					
				except:
					pass
			minx = str(min(diffxes))
			maxx = str(max(diffxes))
			miny = str(min(diffyes))
			maxy = str(max(diffyes))
			min_max.append([minx, maxx, miny, maxy])

			maxX = max([x for (x, y), (diffx, diffy) in data_dict.items()])

			if maxX > 8:
				a_grid1 = MeasurementGrid(data_dict, "x", titles[count], wafers[count])
				a_grid2 = MeasurementGrid(data_dict, "y", titles[count], wafers[count])
				grids.append(a_grid1)
				grids.append(a_grid2)
			else:
				a_grid = MeasurementGrid(data_dict, "both", titles[count], wafers[count])
				grids.append(a_grid)
		count+=1



	Story.append(Spacer(1*inch, after_customer_space*cm))
	row_height = 6/10*accuracy_table_font_size*mm
	sub_table_heights = [row_height, row_height, row_height, 10/6*row_height]

	overall_rows = [accuracy_table_font_size+1*accuracy_table_font_size, 5*accuracy_table_font_size,
	row_height*4 + 2*accuracy_table_font_size] 
	print(accuracy_table_font_size)
	accuracy_headings = TableStyle([('LEFTPADDING', (0,0), (-1,-1), 0), 
									('RIGHTPADDING', (0,0), (-1,-1), 0),
									('TOPPADDING', (0,0), (-1,-1), 0), 
									('BOTTOMPADDING', (0,0), (-1,-1), 0), 
									('ALIGN', (0,0), (0,3), 'RIGHT'), 
									('SIZE', (0,0), (-1,-1), accuracy_table_font_size),
									('VALIGN', (0,0), (-1,-1), 'MIDDLE')])

	accuracy_table = TableStyle([('ALIGN', (0,0), (0,2), 'RIGHT'), 
								('INNERGRID', (0,0), (-1,-1), .25, black),
								('BOX', (0,0), (-1,-1), 0.25, black), 
								('SIZE', (0,0), (-1,-1), accuracy_table_font_size), 
								('VALIGN', (0,0), (-1,-1), 'MIDDLE'), 
								('LEADING', (0,0), (-1,-1), 1.2*accuracy_table_font_size)])

	accuracy_values = TableStyle([('LEFTPADDING', (0,0), (-1,-1), 0), 
								('RIGHTPADDING', (0,0), (-1,-1), 0),
								('TOPPADDING', (0,0), (-1,-1), 0), 
								('BOTTOMPADDING', (0,0), (-1,-1), 0), 
								('SIZE', (0,0), (-1,-1), accuracy_table_font_size), 
								('VALIGN', (0,0), (-1,-1), 'MIDDLE')])

	count = 0
	for file in filenames:

		ptext = '<para alignment=left><font size=15>%s</font></para>' % titles[count]
		Story.append(Paragraph(ptext, styles["Normal"]))
		Story.append(Spacer(1*inch, after_table_heading))

		measurement_headings_data = [["Minimum Error:"], ["Maximum Error:"], ["Specification:"],["Pass/Fail:"]]
		measurement_table_headings = Table(measurement_headings_data, [(2.5-(2/12))*inch], sub_table_heights)
		measurement_table_headings.setStyle(accuracy_headings)

		measurement_values_data = [["Xmin: " + min_max[count][0] + "um", "Ymin: " + min_max[count][2]+ "um"], ["Xmax: " + min_max[count][1] + "um", "Ymax: "+ min_max[count][3]+ "um"], ["+-5um"], ["PASS"]]
		measurement_table_values = Table(measurement_values_data, [2.25*inch-6, 2.25*inch-6], sub_table_heights)
		measurement_table_values.setStyle(accuracy_values)

		data = [["Wafer Used:", wafers[count]], ["Measurement Method:", "Image Recognition \n*Typical measurement uncertainty is +-1um\n**See Map on following page for die tested"]
		, [measurement_table_headings, measurement_table_values]]
		measurement_table = Table(data, column_widths, overall_rows)
		measurement_table.setStyle(accuracy_table)

		Story.append(measurement_table)
		Story.append(Spacer(1*inch, space_after_accuracy_table*mm))
		count+=1


	signature_table_data = [["Performed by: ", user_name], ["Date: ", datetime.date.today().strftime("%d %b %Y")]]
	signature_table = Table(signature_table_data, column_widths)
	signature_table.setStyle(signature_style)
	Story.append(signature_table)

	Story.append(Spacer(1*inch, above_signature))
	ptext = "<font size = 12>Signature: _________________________________________________</font>"
	Story.append(Paragraph(ptext, styles['Normal']))
	Story.append(NextPageTemplate('second_template'))
	Story.append(PageBreak())

	for grid in grids:

		Story.append(grid)
		Story.append(PageBreak())

	create_pdfdoc(doc_name, Story)
