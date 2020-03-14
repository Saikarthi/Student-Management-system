import os
import platform
import xlsxwriter
import xlrd

global list
list=[]
loc = ("output.xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
for i in range(sheet.nrows): 
    list.append(sheet.cell_value(i, 0))



def manageStudent():

	print(""" 

 

Enter 1 : To View Student's List 
Enter 2 : To Add New Student 
Enter 3 : To Search Student 
Enter 4 : To Remove Student 
		
		""")

	try:
		userInput = int(input("Please Select An Above Option: "))
	except ValueError:
		exit("\nHy! That's Not A Number")
	else:
		print("\n")

		
	if(userInput == 1):
		print("List Students\n")  
		for students in list:
			print("=> {}".format(students))

	elif(userInput == 2):
		newStd = input("Enter New Student: ")
		if(newStd in list):
		    print("\nThis Student {} Already In The Database".format(newStd))
		else:
                    list.append(newStd)
                    workbook = xlsxwriter.Workbook('output.xlsx')
                    worksheet = workbook.add_worksheet()
                    row = 0
                    column = 0
                    for item in list :
          
                        worksheet.write(row, column, item)
                        row += 1
                    workbook.close()

                    

			
	elif(userInput == 3):
                srcStd = input("Enter Student Name To Search: ")
                if(srcStd in list):
                    print("\n=> Record Found Of Student {}".format(srcStd))
                else:
                    print("\n=> No Record Found Of Student {}".format(srcStd))

	elif(userInput == 4):
		rmStd = input("Enter Student Name To Remove: ")
		if(rmStd in list):
			list.remove(rmStd)
			workbook = xlsxwriter.Workbook('output.xlsx')
			worksheet = workbook.add_worksheet()
			row = 0
			column = 0
			for item in list :
			          
				worksheet.write(row, column, item)
				row += 1
			workbook.close()
			

			print("\n=> Student {} Successfully Deleted \n".format(rmStd))
			for students in list:
				print("=> {}".format(students))
		else:
			print("\n=> No Record Found of This Student {}".format(rmStd))
	 
	elif(userInput < 1 or userInput > 4):
		print("Please Enter Valid Option")
						

manageStudent()

	

def runAgain():
	runAgn = input("\nwant To Run Again Y/n: ")
	if(runAgn.lower() == 'y'):
		if(platform.system() == "Windows"):
			print(os.system('cls')) 
		else:
			print(os.system('clear'))
		manageStudent()
		runAgain()
	else:
		print("bye")

runAgain()
