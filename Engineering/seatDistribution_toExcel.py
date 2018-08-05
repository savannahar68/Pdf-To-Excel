import pdftotext
import pandas as pd
import os
import sys
import re
import csv
import xlwt

def getUniversity(no):
	if(no >= 2 and no <= 32):
		return universities_array[0]
	elif(no >= 33 and no <= 44):
		return universities_array[1]
	elif(no >= 45 and no <= 95):
		return universities_array[2]
	elif(no >= 96 and no <= 101):
		return universities_array[3]
	elif(no >= 102 and no <= 168):
		return universities_array[4]
	elif(no >= 169 and no <= 177):
		return universities_array[5]
	elif(no >= 178 and no <= 222):
		return universities_array[6]
	elif(no >= 223 and no <= 227):
		return universities_array[7]
	elif(no >= 228 and no <= 251):
		return universities_array[8]
	elif(no >= 252 and no <= 360):
		return universities_array[9]
	elif(no >= 361 and no <= 376):
		return universities_array[10]
	elif(no == 377):
		return universities_array[11]
	else:
		return universities_array[12]

def checkIfStringIsINT(s):
	try: 
		int(s)
		return True
	except ValueError:
		return False


def csvtoexcel(masterCSV_1, masterCSV_2, collegecode, collegename, universityName):
	csv1 = open("csv1.csv", "w")
	csv1.write(collegecode + '\n' + collegename + '\n' + universityName + '\n')
	csv1.write("Choice Code,Course Name,SI,MSCAP,AI,Minority Seats,Insitute Level Seats,Status,Cap Seats,TFWS Code,TFWS Seats\n\n")
	for c in masterCSV_1:
		csv1.write(','.join(c) + '\n')
	csv1.close();
	csv2 = open("csv2.csv", "w")
	csv2.write("Open-G,Open-L,SC-G,SC-L,ST-G,ST-L,VJ/DT-G,VJ/DT-L,NTB-G,NTB-L,NTC-G,NTC-L,NTD-G,NTD-l,OBC-G,OBC-L,PH1,PH2,PH3,PHc,DF,TOTAL\n\n")		
	r = 0
	while r < len(masterCSV_2):
		csv2.write(','.join(masterCSV_2[r]) + '\n')
		r += 1
		csv2.write(','.join(masterCSV_2[r]) + '\n')
		r += 1
		csv2.write(','.join(masterCSV_2[r]) + '\n\n')
		r += 1
	csv2.close()
	wb = xlwt.Workbook()
	ws = wb.add_sheet('Sheet1')
	with open("csv1.csv", 'r') as f:
		reader = csv.reader(f)
		for r, row in enumerate(reader):
			for c, val in enumerate(row):
				ws.write(r, c, val)
	ws = wb.add_sheet('Sheet2')
	with open("csv2.csv", 'r') as f1:
		reader = csv.reader(f1)
		for r, row in enumerate(reader):
			for c, val in enumerate(row):
				ws.write(r, c, val)
	wb.save('excel/' + collegecode + '.xls')
	os.remove("csv1.csv")
	os.remove("csv2.csv")

filename = 'Seat Distribution.pdf'
with open(filename, "rb") as f:
	pdf = pdftotext.PDF(f)

#print(pdf[361], pdf[362], pdf[363], pdf[364])
universities_array = ['Autonomous Institutes', 'Dr. B. A. Marathwada University', 'Dr. B.A.T. University', 'Gondwana University', 'Mumbai University' , 'North Maharashtra University', 'Rashtrasant Tukadoji Maharaj Nagpur University', 'S. R. T. Marathwada University', 'Sant Gadge Baba Amravati University', 'Savitribai Phule Pune University', 'Shivaji University', 'SNDT Womens University', 'Solapur University']
collegecode = ''
collegename = ''
universityName = ''
coursename = [] #this list will have all the unique courses(for database insertion)
masterCSV_1 = [] #this is going to be list of list saving all the categories of engg(comp,it etc)
masterCSV_2 = [] #this is going to be l of l saving all the seat distribution
for i in range(2, 386):
	alllines = [] #this will contrain all lines of a file
	tempfilename = 'temp_' + str(i) + '.txt' #this file will have the whole pdf page(only 1) at a time
	
	#filling whole data in temporary file, the pdf page number 361-364 have half table on each page,
	#so combine both the pages for it and do checking
	if(i == 361 or i == 363):
		#Combine 2 pages such that remove header footer of page 1 and add page 2 without removing header remove only footer
		tempfile = open(tempfilename, "a+")
		tempfile.write(pdf[i])
		tempfile.close()
		
		temp = open(tempfilename, "r")
		alllines = temp.readlines()
		alllines = alllines[: -3]
		temp.close()
		
		w = open(tempfilename, "w")
		w.writelines([line for line in alllines])
		w.close()
		i += 1;
		
	#work normal as all the file are on the same page now!(witty you)	
	with open(tempfilename, "a") as tempfile:
		tempfile.write(pdf[i])
	
	tempfile.close()
	temp = open(tempfilename, "r")
	alllines = temp.readlines()
	temp.close()
	os.remove(tempfilename)
	#until here we have created a text file having all the data of 1 page of pdf line wise
	#now we'll remove the first line and the last few lines which are irrelevant
	alllines = alllines[1:-3]
	#print(alllines)
	#now all lines have only the data which we need
	line = 0
	while line < len(alllines):
		if(checkIfStringIsINT(alllines[line][:4])):
			if(collegecode != alllines[line][:4]):
				#write code here to save the previous things into 2 csv and convert into excel
				#also clear all the array all the string and all the lists
				if(i != 2):
					csvtoexcel(masterCSV_1, masterCSV_2, collegecode, collegename, universityName)
				print(masterCSV_1, ' \n', masterCSV_2)
				print('\n\n')
				masterCSV_1 = [] #this is going to be list of list saving all the categories of engg(comp,it etc)
				masterCSV_2 = [] #this is going to be l of l saving all the seat distribution
				collegecode = alllines[line][: 4]
				collegename = alllines[line][6:].strip()
				universityName = getUniversity(i)
			line += 1;
			temp_masterCSV_1 = [None]*11 #this will store civil, it comp and all
			temp_masterCSV_2_HU = ['Nil']*22 #this will store seat distribution of civil it comp
			temp_masterCSV_2_OHU = ['Nil']*22 #this will store seat distribution of civil it comp
			temp_masterCSV_2_stateLevel = ['Nil']*22 #this will store seat distribution of civil it comp
			#print(alllines[line].strip())
			alllines[line] = alllines[line].strip()
			flag = 0
			flag1 = 0
			if "Rashtrasant Tukadoji Maharaj" in alllines[line]: #check if the status name of college is not in 2 lines
				statusName = alllines[line] + ' '
				flag = 1
				line += 1
				alllines[line] = alllines[line].strip()
			elif not "CAP Seats" in alllines[line] : #check if the name of college is not in 2 lines
				collegename += alllines[line]
				line += 1
				alllines[line] = alllines[line].strip()
			querystring = 'CAP Seats:' #getting the capseats and the status
			capseats = alllines[line][alllines[line].find(querystring) + len(querystring):].strip()
			if(flag == 1):
				statusName += alllines[line][:alllines[line].find(querystring)].strip()
				flag = 0
			else:
				statusName = alllines[line][:alllines[line].find(querystring)].strip()
			line += 1
			alllines[line] = alllines[line].strip()
			if(alllines[line][:6] != "Choice"): #if not then it is part of statusName
				statusName += alllines[line]
				line += 1
				alllines[line] = alllines[line].strip()
			temp_masterCSV_1[7] = statusName
			temp_masterCSV_1[8] = capseats
			line += 2
			alllines[line] = alllines[line].strip()
			templine = alllines[line].split()
			temp_masterCSV_1[0] = templine[0][:9] #saving the choice code (if #then striped)
			temp_masterCSV_1[6] = templine[len(templine)-1] #Insti level seats
			temp_masterCSV_1[5] = templine[len(templine)-2] #Minority Seats
			temp_masterCSV_1[4] = templine[len(templine)-3] #AI
			temp_masterCSV_1[3] = templine[len(templine)-4] #MSCAP
			temp_masterCSV_1[2] = templine[len(templine)-5] #SI
			if "#" in templine[1]:
				temp_masterCSV_1[1] = ' '.join(templine[2: len(templine)-5]) #Course Name(civil etc)
			else:		
				temp_masterCSV_1[1] = ' '.join(templine[1: len(templine)-5]) #Course Name(civil etc)
			coursename.append(temp_masterCSV_1[1])		
			line += 3
			alllines[line] = alllines[line].strip().split()
			origline = '' #if tfws is not there then there is a problem
			if(line < len(alllines) and alllines[line][0] == 'HU'):
				for k in range(0, 22):
					temp_masterCSV_2_HU[k] = alllines[line][k + 1]
				line += 1
				if line < len(alllines):
					origline = alllines[line]
					alllines[line] = alllines[line].strip().split()
					flag1 = 1;

			if(line < len(alllines) and alllines[line][0] == 'OHU'):
				for k in range(0, 22):
					temp_masterCSV_2_OHU[k] = alllines[line][k + 1]
				line += 1
				if line < len(alllines):
					origline = alllines[line]
					alllines[line] = alllines[line].strip().split()
					flag1 = 1;

			if(line < len(alllines) and alllines[line][0] == 'State'):
				for k in range(0, 22):
					temp_masterCSV_2_stateLevel[k] = alllines[line][k + 2]
				line += 1
				if line < len(alllines):
					origline = alllines[line]
					alllines[line] = alllines[line].strip().split()
					flag1 = 1;
					
			if(line < len(alllines) and flag1 == 1 and "TFWS" in alllines[line]):
				temp_masterCSV_1[9] = alllines[line][2][5:] 
				temp_masterCSV_1[10] = alllines[line][3][6:]
				flag1 = 0
			else:
				if(line < len(alllines)):
					alllines[line] = origline
				line -= 1
				flag1 = 0
				temp_masterCSV_1[9] = 'Nil' 
				temp_masterCSV_1[10] = 'Nil'
			#print(collegecode, '\n', collegename, '\n', universityName)
			#print(temp_masterCSV_1,'\n', temp_masterCSV_2_HU,'\n' ,temp_masterCSV_2_OHU,'\n', temp_masterCSV_2_stateLevel)
			
			#my complete 1 table is made here so mush everything into master tables
			masterCSV_1.append(temp_masterCSV_1)
			masterCSV_2.append(temp_masterCSV_2_HU)
			masterCSV_2.append(temp_masterCSV_2_OHU)
			masterCSV_2.append(temp_masterCSV_2_stateLevel)
		line += 1
	csvtoexcel(masterCSV_1, masterCSV_2, collegecode, collegename, universityName)	
	#print(i)	
coursename = list(set(coursename))
with open("CourseName.txt", "w") as cn:
	for name in coursename:
		cn.write(name + '\n')