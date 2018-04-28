import xlrd
import xlsxwriter

import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import inflect 
ordersh= pd.read_excel('orders.xlsx', sheet_name='orders')
gstsh= pd.read_excel('GST.xlsx', sheet_name='Sheet1')
statesh= pd.read_excel('state.xlsx', sheet_name='Sheet1') 
prev=""
cnt=1
pos=5
lst=[]
z=inflect.engine();
invoice1 = xlsxwriter.Workbook('invoice05.xlsx')
for index, rw in ordersh.iterrows():
	name=rw['Name']
	
	
		
	if name in lst:
		continue
	invoice_number=rw['Name']
	lst.append(name)	
	top=rw['Payment Method']
	#print(rw['Lineitem quantity'])
	dated=rw['Created at']
	BillingName=rw['Billing Name']
	ShippingAddress=rw['Shipping Address1']
	BuyersOrderNumber=rw['Name']
	ShippingPhone=rw['Shipping Phone']
	Province=rw['Shipping Province']
	print(dated)	
	State=statesh.loc[statesh['State Abbreviations'] == Province,'State Name'].iloc[0]
	upperState=State.upper()	
	StateCode=gstsh.loc[gstsh['STATE NAME']==upperState,'STATE CODE'].iloc[0]
	Taxes=rw['Taxes']
	Total=rw['Total']
	Currency=rw['Currency']
	Amount=z.number_to_words(Total)		
	merge_format = invoice1.add_format({'bold': 1,'align': 'center','valign': 'vcenter','font_size':18,'bg_color':'gray'})
	merge_format1 = invoice1.add_format({'align': 'center','valign': 'vcenter','bg_color':'gray'})
	merge_format2 = invoice1.add_format({'valign': 'vcenter','bg_color':'white'})
	merge_format3 = invoice1.add_format({'align': 'center','valign': 'vcenter','border':1,'border_color':'#000000'})
	merge_format4 = invoice1.add_format({'align': 'center','valign': 'vcenter','border':1,'bottom_color':'#FFFFFF','left_color':'#000000','right_color':'#000000'})

	left = invoice1.add_format({'align': 'left','bg_color':'gray'})
	right = invoice1.add_format({'align': 'right','bg_color':'gray'})
	right1 = invoice1.add_format({'align': 'right','bg_color':'gray','num_format': 'dd/mm/yy'})
	bgcolor = invoice1.add_format({'bg_color':'gray'})	
	col=0	
	if(cnt%2==1):
		p=0
		Filename='invoice0'+str(pos)+'.xlsx'
		invoice1 = xlsxwriter.Workbook(Filename)
		worksheet = invoice1.add_worksheet()
		
	else:
		p=12
		pos+=1
	for row in range(20):
		worksheet.write(row,col+p,'',bgcolor)
		worksheet.write(row,col+10+p,'',bgcolor)
	worksheet.merge_range(0,1+p,0,9+p, '',merge_format)
	worksheet.set_row(1, 30)
	worksheet.merge_range(1,1+p,1,9+p, 'ENCELADUS INTERNET PRIVATE LIMITED',merge_format)
	worksheet.merge_range(2,1+p,2,9+p, 'Address of Enceladus Internet Private Limited',merge_format1)
	worksheet.set_row(3, 10)
	worksheet.merge_range(3,1+p,3,9+p, '',merge_format)
	worksheet.merge_range(4,1+p,4,4+p, 'GSTIN/UIN-GSTINXXAAWWKK00Z',left)
	worksheet.merge_range(4,5+p,4,9+p, 'Email- hello@pechkus.co',merge_format1)
	worksheet.merge_range(5,1+p,5,9+p, '',merge_format1)
	worksheet.merge_range(6,1+p,6,2+p, 'Invoice Number',left)
	worksheet.merge_range(6,3+p,6,7+p, 'Model/Terms of Payment',merge_format1)
	worksheet.merge_range(6,8+p,6,9+p, 'Dated',right)
	worksheet.merge_range(7,1+p,7,2+p,invoice_number,left)
	worksheet.merge_range(7,3+p,7,7+p, top,merge_format1)
	#worksheet.merge_range('I8:J8', dated,right)
	worksheet.write(7,8+p, '',right)
		
	worksheet.write_datetime(7,9+p, dated, right1)			
	worksheet.merge_range(8,1+p,8,9+p, '',merge_format1)
	worksheet.merge_range(9,1+p,9,2+p,'Buyers Order Number',left)
	worksheet.merge_range(9,3+p,9,6+p, '',merge_format1)
	worksheet.merge_range(9,7+p,9,9+p, 'Despatched Through',right)
	worksheet.merge_range(10,1+p,10,2+p,BuyersOrderNumber,left)
	worksheet.merge_range(10,3+p,10,6+p, '',merge_format1)
	worksheet.merge_range(10,7+p,10,9+p, 'ECOM EXPRESS',right)
	worksheet.merge_range(11,1+p,11,9+p, '',merge_format1)
	worksheet.merge_range(12,1+p,12,9+p, '',merge_format2)
	
	worksheet.merge_range(13,2+p,13,4+p,BillingName,merge_format2)
	worksheet.merge_range(14,2+p,14,4+p,ShippingAddress,merge_format2)
	worksheet.merge_range(15,2+p,15,8+p, 'Mobile:'+ShippingPhone,merge_format2)
	worksheet.write(16,2+p,'State',merge_format2)
	worksheet.write(16,3+p,State,merge_format2)
	worksheet.write(17,2+p,'State Code',merge_format2)
	worksheet.write(17,3+p,StateCode,merge_format2)
	worksheet.merge_range(18,2+p,18,9+p, '',merge_format2)
	worksheet.merge_range(19,1+p,19,9+p, '',merge_format1)
	worksheet.conditional_format(13,1+p,18,9+p, {'type':     'blanks',
       	                            'format':   merge_format2})
	worksheet.set_column('C:C', 20)
	worksheet.set_column('O:O', 20)
	
	worksheet.merge_range(20,2+p,20,9+p, '',merge_format2)
	worksheet.write(21,0+p,'S.No.',merge_format3)
	worksheet.merge_range(21,1+p,21,4+p,'Description of Goods',merge_format3)
	worksheet.write(21,5+p,'HSN/SAC',merge_format3)
	worksheet.write(21,6+p,'Quantity',merge_format3)
	worksheet.write(21,7+p,'Rate',merge_format3)	
	worksheet.write(21,8+p,'Per',merge_format3)
	worksheet.merge_range(21,9+p,21,10+p,'Amount',merge_format3)
#starting of table phase
	row=22
	Sno=1	
	for ind, rew in ordersh.iterrows():	
		if(rew['Name']==rw['Name']):
					
			Desc=rew['Lineitem name']
			Price=rew['Lineitem price']
			Quantity=rew['Lineitem quantity']
			worksheet.write(row,0+p,Sno,merge_format3)
			worksheet.merge_range(row,1+p,row,4+p,Desc,merge_format3)
			worksheet.write(row,5+p,'6101',merge_format3)
			worksheet.write(row,6+p,Quantity,merge_format3)
			worksheet.write(row,7+p,Price,merge_format3)	
			worksheet.write(row,8+p,'Nos',merge_format3)
			worksheet.merge_range(row,9+p,row,10+p,Price*Quantity,merge_format3)			
			row+=1
			Sno+=1
	for j in range(row,27):
		worksheet.write(j,0+p,'',merge_format4)
		worksheet.merge_range(j,1+p,j,4+p,'',merge_format4)
		worksheet.write(j,5+p,'',merge_format4)
		worksheet.write(j,6+p,'',merge_format4)
		worksheet.write(j,7+p,'',merge_format4)	
		worksheet.write(j,8+p,'',merge_format4)
		worksheet.merge_range(j,9+p,j,10+p,'',merge_format4)			
		
					
	worksheet.write(27,0+p,'',merge_format3)
	worksheet.merge_range(27,1+p,27,4+p,'',merge_format3)
	worksheet.write(27,5+p,'',merge_format4)
	worksheet.write(27,6+p,'',merge_format4)
	worksheet.write(27,7+p,'',merge_format4)	
	worksheet.write(27,8+p,'IGST',merge_format4)
			
	worksheet.merge_range(27,9+p,27,10+p,Taxes,merge_format3)
	worksheet.write(28,0+p,'',merge_format3)
	worksheet.merge_range(28,1+p,28,4+p,'Total',merge_format3)
	worksheet.write(28,5+p,'',merge_format4)
	worksheet.write(28,6+p,'',merge_format4)
	worksheet.write(28,7+p,'',merge_format4)	
	worksheet.write(28,8+p,'',merge_format4)
		
	worksheet.merge_range(28,9+p,28,10+p,Total,merge_format3)
	
	worksheet.merge_range(29,0+p,29,10+p,'',merge_format3)
	worksheet.merge_range(30,0+p,30,1+p,'Tax Amount',merge_format3)
	worksheet.write(30,2+p,Total,merge_format3)
	worksheet.merge_range(30,3+p,30,10+p,'in words:'+Currency+' '+Amount,merge_format3)
	worksheet.merge_range(31,0+p,31,10+p,'',merge_format2)
	worksheet.merge_range(32,0+p,32,10+p,'We declare this invoice shows the actual price of the goods described and all that particulars are true and correct.',merge_format3)
	worksheet.merge_range(33,0+p,33,10+p,'',merge_format2)
	worksheet.merge_range(34,6+p,34,10+p,'For Enceladus Internet PVT LTD',merge_format3)
	worksheet.merge_range(35,6+p,35,10+p,'Authorised Signatory',merge_format3)
	worksheet.merge_range(36,0+p,36,10+p,'This is a Computer Generated Invoice',merge_format3)
	#worksheet.conditional_format(20,0+p,36,10+p, {'type':     'blanks',
         #                           'format':   merge_format2})
		
	cnt+=1		
invoice1.close()
