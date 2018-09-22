import io, os, sys, openpyxl as px
from importlib import reload

#reload(sys)
#sys.setdefaultencoding('utf8')

fromaddr = "juanluis.rto@gmail.com"
password = "xxxxxx"

subject = "Stiftelsen Pro Cultura - Faktura"
body = "Hola Javi! \n Este mail con factura te lo envía un programa, es un experimento"




class Faktura:
	def __init__(self, A, B, C, D, E, F, G, H, I, J):
		self.fakturanummer = A
		self.fakturadatum = B
		self.forfallodatum = C
		self.mottagare = D
		self.mejladress = E
		self.koncept = F
		self.summa = G
		self.moms = H
		self.skickad = I
		self.betalad = J
		self.latex = None

	def generate_latex_file(self):
	    name = "latex\\" + str(self.fakturanummer)  + ".tex" #+ "-" + self.koncept + ".tex"
	    f= io.open(name,"w+", encoding='utf8')
	    f.write("\\documentclass{faktura}\n\\begin{document}\n\\printStiftelseProCultura")
	    f.write("\\printReceiver{" + self.mottagare + "}\n\\printDates{" + \
	    str(self.fakturanummer)+ "}{" + str(self.fakturadatum.date()).split()[0] + "}{"+ str(self.forfallodatum.date()).split()[0] + "}\n\\begin{invoiceTable}\n")
	    f.write("\\feerow{" + self.koncept + "}{" + str(self.summa) + "}\n")
	    if self.moms is not None:
	        f.write("\\momsrow{" + str(self.moms) + "}\n")
	    f.write("\\end{invoiceTable}\n\\LarkstadenInfo\n\\end{document}")
	    f.close()
	    self.latex = f

	def send_invoice(self):
		send_email(fromaddr, self.mejladress, password, subject, body, "pdf\\" + str(self.fakturanummer) + ".pdf")




def create_dir(path):
	if not os.path.exists(path):
	    os.makedirs(path)

def send_email(fromaddr, toaddr, password, subject, body, path):
	import smtplib
	from email.mime.multipart import MIMEMultipart
	from email.mime.text import MIMEText
	from email.mime.base import MIMEBase
	from email import encoders

	msg = MIMEMultipart()
	 
	msg['From'] = fromaddr
	msg['To'] = toaddr
	msg['Subject'] = subject
	 
	msg.attach(MIMEText(body, 'plain'))
	 
	filename = path.split('\\')[1]
	attachment = open(path, "rb")
	 
	part = MIMEBase('application', 'octet-stream')
	part.set_payload((attachment).read())
	encoders.encode_base64(part)
	part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
	 
	msg.attach(part)
	 
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.starttls()
	server.login(fromaddr, password)
	text = msg.as_string()
	server.sendmail(fromaddr, toaddr, text)
	server.quit()

	
if __name__ == '__main__':

	wb = px.load_workbook('data.xlsx', data_only=True)
	ws = wb.active #wb.get_sheet_by_name('Sheet1')

	#import invoices
	fakturor = []
	for i in range(2,ws.max_row):
		if ws['A'][i].value is None:
			continue
		fakturor.append(Faktura(ws['A'][i].value, ws['B'][i].value, ws['C'][i].value, ws['D'][i].value, ws['E'][i].value, ws['F'][i].value, ws['G'][i].value, ws['H'][i].value, ws['I'][i].value, ws['J'][i].value))

	#creates directories if not already existing
	create_dir('pdf')
	create_dir('latex')

	#generates tex and pdf files
	for faktura in fakturor:
		try:
			faktura.generate_latex_file()
		except Exception as err:
			print("Latex file generation failed for invoice " + str(faktura.fakturanummer))
			print(err)
		try:
			os.system("pdflatex -halt-on-error -output-directory pdf latex/" + str(faktura.fakturanummer) + ".tex")
		except Exception as err:
			print("PDF file compilation failed for invoice " + str(faktura.fakturanummer))
			print(err)
		

	#sends emails with invoices
	for faktura in fakturor:
		try:
			faktura.send_invoice()
		except FileNotFoundError as err:
			print("Invoice " + str(faktura.fakturanummer) + " was not found and therefore not sent to " + faktura.mejladress)
			print(err)
		

	#delets logs and auxiliary files 
	for item in os.listdir('pdf'):
	    if item.endswith(".log") or item.endswith(".aux"):
	        os.remove(os.path.join('pdf', item))

