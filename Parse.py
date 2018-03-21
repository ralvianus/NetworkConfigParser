import textfsm
from datetime import datetime
import os,sys,platform
import xlsxwriter
import threading
import Parse_Config

log = open('parse_config.log', 'w')

print ("Start Parsing Configuration")
with open("./File/file.csv", "r") as SourceText:
	thread=[]
	for line in SourceText:
		FileName,hostname,command,device_type = line.split(',')
		input_file = open(FileName)
		raw_text_data = input_file.read()
		input_file.close()
		
		#TextFSM(raw_text_data,hostname,command,device_type)
		t=threading.Thread(target=Parse_Config.TextFSM, args=(raw_text_data, hostname, command, device_type))
		thread.append(t)
		t.start()