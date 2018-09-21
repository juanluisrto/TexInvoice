import re, urllib2, sys, openpyxl as px

reload(sys)
sys.setdefaultencoding('utf8')

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


	def generate_latex_file(self):
	    name = "latex\\" + str(self.fakturanummer) + "-" + self.koncept + ".tex"
	    f= open(name,"w+")
	    f.write("\\documentclass{faktura}\n\\begin{document}\n\\printStiftelseProCultura")
	    f.write("\\printReceiver{" + self.mottagare + "}\n\\printDates{" + \
	    str(self.fakturanummer)+ "}{" + str(self.fakturadatum.date()).split()[0] + "}{"+ str(self.forfallodatum.date()).split()[0] + "}\n\\begin{invoiceTable}\n")
	    f.write("\\feerow{" + self.koncept + "}{" + str(self.summa) + "}\n")
	    if self.moms is not None:
	        f.write("\\momsrow{" + str(self.moms) + "}\n")
	    f.write("\\end{invoiceTable}\n\\LarkstadenInfo\n\\end{document}")
	    f.close()


def create_dir(path):
	if not os.path.exists(path):
	    os.makedirs(path)


	
if __name__ == '__main__':

	wb = px.load_workbook('data.xlsx', data_only=True)
	ws = wb.get_sheet_by_name('sheet1')

	#import invoices
	fakturor = []
	for i in range(3,ws.max_row):
	    fakturor.append(Faktura(ws['A'][i].value, ws['B'][i].value, ws['C'][i].value, ws['D'][i].value, ws['E'][i].value, ws['F'][i].value, ws['G'][i].value, ws['H'][i].value, ws['I'][i].value, ws['J'][i].value))

	create_dir('pdf')
	create_dir('latex')

	for faktura in fakturor:
		faktura.generate_latex_file()


