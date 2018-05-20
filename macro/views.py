from django.shortcuts import render, HttpResponse
from time import ctime
from .send_mail import send

# Create your views here.

def home_view(request):
	items = {
		"ActiveWorkbook"						: "ThisComponent",
		"ActiveCell"							: "ThisComponent.getCurrentSelection",
		"Application.ScreenUpdating = True"		: "ThisComponent.UnlockControllers",
		"Application.ScreenUpdating = False"	: "ThisComponent.LockControllers",
		"Selection.Value"						: "ThisComponent.getCurrentSelection.Value",
		#vbMsgBoxStyle Arguments
		"vbOKOnly"								: "0",	# MB_OK - OK button
		"vbOKCancel"							: "1",	# MB_OKCANCEL - OK and Cancel button
		"vbAbortRetryIgnore"					: "2",	# MB_ABORTRETRYIGNORE - Abort, Retry, and Ignore buttons
		"vbYesNoCancel"							: "3",	# MB_YESNOCANCEL - Yes, No, and Cancel buttons
		"vbYesNo"								: "4",	# MB_YESNO - Yes and No buttons
		"vbRetryCancel"							: "5",	# MB_RETRYCANCEL - Retry and Cancel buttons
		"vbCritical"							: "16",	# MB_ICONSTOP - Stop sign
		"vbQuestion"							: "32",	# MB_ICONQUESTION - Question mark
		"vbExclamation"							: "48",	# MB_ICONEXCLAMATION - Exclamation point
		"vbInformation"							: "64",	# MB_ICONINFORMATION - Tip icon
		"vbDefaultButton1"						: "0",	# MB_DEFBUTTON1 - First button is default value
		"vbDefaultButton2"						: "256",	# MB_DEFBUTTON2 - Second button is default value
		"vbDefaultButton3"						: "512",	# MB_DEFBUTTON3 - Third button is default value
		"vbDefaultButton4"						: "768",
		"vbApplicationModal"					: "0",
		"vbSystemModal"							: "4096",
		"vbMsgBoxHelpButton"					: "16384",
		"VbMsgBoxSetForeground"					: "65536",
		"vbMsgBoxRight"							: "524288",
		"vbMsgBoxRtlReading"					: "1048576",
		# MsgBox Return Values
		"vbOK"									: "1",	# IDOK - Ok
		"vbCancel"								: "2",	# IDCANCEL - Cancel
		"vbAbort"								: "3",	# IDABORT - Abort
		"vbRetry"								: "4",	# IDRETRY - Retry
		"vbIgnore"								: "5",	# - Ignore
		"vbYes"									: "6",	# IDYES - Yes
		"vbNo"									: "7",	# IDNO - No
		}

	entries = request.GET.get('entries')
	if entries:
		exits = entries

		if exits.find("\n") == -1:		# for some control situation
			exits += "\n"
		try:		# unexpected python error handling
			backup = exits
			#Range("A1").Select
			#Selection.Value = 10
			if exits.find("Range(") != -1:
				x = False
				index_counter = 0
				change_list = {}
				exits_array = exits.splitlines()
				for i in exits_array:
					if (x == True) & (i.find("Selection.Value") == 0):	
						new_str = i.replace("Selection.Value", "ThisComponent.getCurrentSelection.Value")
						change_list[index_counter] = new_str
						x = False
						index_counter += 1
						continue
					if (x == True) & (i.find("Selection.Copy") == 0):		# doesnt work copy and paste  method. must be here for selection method
						new_str = i.replace("Selection.Copy", "ThisComponent.getCurrentSelection.Copy")
						change_list[index_counter] = new_str
						x = False
						index_counter += 1
						continue
					if i.find("Range(") != -1:
						if i.find(".Select") != -1:		# if ".Select" value is hasnt, pass the for cycle
							position = i.find("Range(") + len("Range(")
							for j in range(position,len(i)):
								if (i[j] == ")") & (i[j+1:j+8] == ".Select"):
									parameter = i[position:j]
									new_str = "ThisComponent.CurrentController.select(ThisComponent.CurrentController.ActiveSheet.getCellRangeByName((temp)))"
									new_str = new_str.replace("temp", parameter)
									change_list[index_counter] = new_str
									x = True
									break
						else:
							new_str = i.replace("Range", "ThisComponent.CurrentController.ActiveSheet.getCellRangeByName")
							change_list[index_counter] = new_str
					index_counter += 1
				for keys,values in change_list.items():
					exits = exits.replace(exits_array[keys],values)
		except:
			exits = backup

		try:
			backup = exits
			# Cells(1,1)=1
			if exits.find("Cells(") != -1:
				x = False
				index_counter = 0
				change_list = {}
				exits_array = exits.splitlines()
				for i in exits_array:
					if (x == True) & (i.find("Selection.Value") == 0):	
						new_str = i.replace("Selection.Value", "ThisComponent.getCurrentSelection.Value")
						change_list[index_counter] = new_str
						x = False
						index_counter += 1
						continue
					if (x == True) & (i.find("Selection.Copy") == 0):		# doesnt work copy and paste  method. must be here for selection method
						new_str = i.replace("Selection.Copy", "ThisComponent.getCurrentSelection.Copy")
						change_list[index_counter] = new_str
						x = False
						index_counter += 1
						continue
					if i.find("Cells(") != -1:
						position = i.find("Cells(") + len("Cells(")
						for j in range(position,len(i)):
							if (i[j] == ")") & (i[j+1:j+8] == ".Select"):
								parameter = i[position:j]
								new_str = "ThisComponent.CurrentController.select(ThisComponent.CurrentController.ActiveSheet.getCellByPosition(temp))"
								new_str = new_str.replace("temp", parameter)
								change_list[index_counter] = new_str
								x = True
								break
						else:
							new_str = i.replace("Cells", "ThisComponent.CurrentController.ActiveSheet.getCellByPosition")
							change_list[index_counter] = new_str
					index_counter += 1
				for keys,values in change_list.items():
					exits = exits.replace(exits_array[keys],values)
		except:
			exits = backup

		try:
			backup = exits
			# Range("A3").FormulaR1C1 = "=R1C1 + R2C1" => row column formula method. you can change parameter to A1 notation. R1C1 not supported
			if (exits.find("FormulaR1C1") != -1):
				position = exits.find("FormulaR1C1") + len("FormulaR1C1")
				parameter = ""
				x = False
				while(exits[position+1] != "\n"):		# if entries is one line, there arent "\n" value. index out of range (Server Error (500)). Users usually dont check one line. Example: End Sub bypass this error. we are actually added the new line value at first
					if exits[position] == '"':
						x = True
					if x == True:
						parameter += exits[position]
					position += 1
				parameter += exits[position]
				parameter = parameter.rstrip()		# removed \n
				new_str = "setFormula("+parameter+")"
				exits = exits.replace("FormulaR1C1",new_str)
				first_position = exits.find(new_str) + len(new_str)		# remove unexpected values => ("=R1C1 + R2C1") = "=R1C1 + R2C1"
				last_position = first_position
				while(exits[last_position] != "\n"):
					last_position += 1
				exits = exits.replace(exits[first_position:last_position],"")
		except:
			exits = backup

		try:
			backup = exits
			# all "ActiveSheet" stuations. we are not use "ActiveSheet" in dictionary because is replicate in line
			if exits.find("ActiveSheet") != -1:
				index_counter = 0
				change_list = {}
				exits_array = exits.splitlines()
				for i in exits_array:
					if i.find("ThisComponent.") == -1:	# if not previously prepaired, there arent "ThisComponent" (general value) in line
						new_str = i.replace("ActiveSheet", "ThisComponent.CurrentController.ActiveSheet")
						change_list[index_counter] = new_str
					index_counter += 1
				for keys,values in change_list.items():
					exits = exits.replace(exits_array[keys],values)
		except:
			exits = backup

		try:
			backup = exits
			# all "Formula" stuations. we are not use "Formula" in dictionary because is replicate in line
			if exits.find("Formula") != -1:
				index_counter = 0
				change_list = {}
				exits_array = exits.splitlines()
				for i in exits_array:
					if i.find("setFormula") == -1:
						first_position = i.find("Formula") + len("Formula") -1
						last_position = first_position
						for last_position in range(first_position,len(i)):
							if i[last_position] == '"':
								break
						if i[first_position:last_position+1].find("=") == -1:
							new_str = i.replace("Formula", "Formula = ")
							change_list[index_counter] = new_str
					index_counter += 1
				for keys,values in change_list.items():
					exits = exits.replace(exits_array[keys],values)
		except:
			exits = backup

		# dictionary items
		for i in items:
			exits = exits.replace(i, items[i])


		return render(request,'home.html',{'entries':entries,'exits':exits})
	else:
		return render(request,'home.html')

def about(request):
	return render(request,'about.html')

def benefits(request):
	return render(request,'benefits.html')

def progress(request):
	return render(request,'progress.html')

def used_technologies(request):
	return render(request,'used-technologies.html')

def examples(request):
	return render(request,'examples.html')

def license(request):
	return render(request,'license.html')

def contact(request):
	name = request.GET.get('name')
	email = request.GET.get('email')
	mail = request.GET.get('mail')
	number = request.GET.get('number')
	robot = " + 5 = 7"

	if name and email and mail:
		if number == "2":		# robot verification
			try:
				send(name=name, email=email, mail=mail)
				return render(request,'contact.html',{'status':'Message Sended ('+str(ctime())+')','robot':robot})
			except:
				return render(request,'contact.html',{'status':'Connection Error!','robot':robot})
		else:
			return render(request,'contact.html',{'status':'You entered an incorrect CAPTCHA answer. Try again!','robot':robot})
	elif name:
		return render(request,'contact.html',{'status':'Please fill All Area!','robot':robot})
	elif email:
		return render(request,'contact.html',{'status':'Please fill All Area!','robot':robot})
	elif mail:
		return render(request,'contact.html',{'status':'Please fill All Area!','robot':robot})
	else:
		return render(request,'contact.html',{'robot':robot})

def contribute(request):
	name = request.GET.get('name')
	email = request.GET.get('email')
	excel = request.GET.get('excel')
	libreoffice = request.GET.get('libreoffice')
	number = request.GET.get('number')
	robot = " + 6 = 9"

	if name and email and excel and libreoffice:
		if number == "3":
			try:
				send(name=name, email=email, excel=excel, libreoffice=libreoffice)
				return render(request,'contribute.html',{'status':'Message Sended ('+str(ctime())+')','robot':robot})
			except:
				return render(request,'contribute.html',{'status':'Connection Error!','robot':robot})
		else:
			return render(request,'contribute.html',{'status':'You entered an incorrect CAPTCHA answer. Try again!','robot':robot})
	elif name:
		return render(request,'contribute.html',{'status':'Please fill All Area!','robot':robot})
	elif email:
		return render(request,'contribute.html',{'status':'Please fill All Area!','robot':robot})
	elif excel:
		return render(request,'contribute.html',{'status':'Please fill All Area!','robot':robot})
	elif libreoffice:
		return render(request,'contribute.html',{'status':'Please fill All Area!','robot':robot})
	else:
		return render(request,'contribute.html',{'robot':robot})

def test(request):
	return HttpResponse('<h2>test page </h2>')
