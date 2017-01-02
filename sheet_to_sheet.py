from openpyxl import load_workbook

class Sheet_to_sheet():
	def __init__(self):
		self.inv_to_dell = {'SERIAL': 'Service Tag', 'MODELO':'System Model', 'HOSTNAME':'System Description', 'Start Date': ' Service Start Date', 'End Date': 'Service End Date'}
		self.dell_to_inv = {'Service Tag': 'SERIAL' , 'System Model': 'MODELO' , 'System Description': 'HOSTNAME'}
