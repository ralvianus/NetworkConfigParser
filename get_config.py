from netmiko import ConnectHandler
from netmiko import ssh_exception
from datetime import datetime
import time
import threading
import os,sys,platform
import subprocess
from docx import Document
import csv

import Parse_Config
import Write_Config

log = open('get_config.log', 'w')

def ping(Host):
    #Check Operating System and use ping command accordingly
    if platform.system() is 'Windows':
        ping_command = "ping -n 1 "+Host
    else:
        ping_command = "ping -c 1 "+Host

    #Ping
    try:
        output = subprocess.check_output(ping_command, shell=True)
    except subprocess.CalledProcessError:
        return False
    else:
        if "TTL expired in transit" in str(output):
            print ("TTL expired in transit")
            return False
        else:
            return True

def Connect_SSH(Hostname, device_type, ip, port, username, password, enable_secret, isParsed, isWrite):
    Format = '%Y%m%d'
    DOC_TITLE = "UAT Test"
    DOC_NAME_TEMPLATE = "UAT"
    Commands = []
    #Connect to SSH using ConnectHandler
    try:
        SSH_Connect = ConnectHandler(
            device_type=device_type,
            ip=ip,
		    port=port,
            username=username,
            password=password,
            secret=enable_secret
        )

    #Catch Exception from SSH connect
    except ssh_exception.AuthenticationException:
        #print >> log, ("Wrong Credentials for %s" % Hostname)
        print ("Wrong Credentials for %s" % Hostname)
    except ssh_exception.NetMikoTimeoutException:
        #print >> log,  ("Connection Timeout for %s" % Hostname)
        print ("Connection Timeout for %s" % Hostname)
    else:
        #print >> log,  ("Connecting to %s via SSH......" % Hostname)
        print ("Connecting to %s via SSH......" % Hostname)

        #Check device type, then load the command for each device type
        if device_type == 'cisco_ios':
            #Commands = ['show cdp neighbor', 'show interface status', 'show mac address-table', 'show interface']
            with open("./Command/ios_show_command.csv", "r") as SourceText:
                #CSVSource = csv.reader(SourceText)
                for line in SourceText.readlines():
                    Commands = line.split(',')
        elif device_type == 'cisco_nxos':
            with open("./Command/nxos_show_command.csv", "r") as SourceText:
                #CSVSource = csv.reader(SourceText)
                for line in SourceText:
                    Commands = line.split(',')
        elif device_type == 'juniper':
            with open("./Command/juniper_show_command.csv", "r") as SourceText:
                # CSVSource = csv.reader(SourceText)
                for line in SourceText:
                    Commands = line.split(',')
        elif device_type == 'arista_eos':
            with open("./Command/arista_show_command.csv", "r") as SourceText:
                # CSVSource = csv.reader(SourceText)
                for line in SourceText:
                    Commands = line.split(',')
        else:
            return

    #Enter enable mode
    SSH_Connect.enable()

    #Check Path for output folder
    path = "./output/%s" % Hostname
    if not os.path.exists(path):
        os.mkdir(path)

    #Initializing write to docx function
    if (isWrite == 'y' or isWrite == 'yes'):
        document = Document()
        title_input = DOC_TITLE
        title = "%s - %s" % (Hostname,title_input)
        Write_Config.Write_Title(document, title)
        FileName_template = DOC_NAME_TEMPLATE
        docx_FileName = '%s/%s-%s_%s' % (path, Hostname, FileName_template, datetime.now().strftime(Format))
        Write_Config.Create_Code_Style(document)
    else:
        return

    #Call threading function
    thread = []
    global JSON_Filename, JSON_Command
    JSON_Filename = []
    JSON_Command = []

    for command in Commands:
        try:
            #Append the command and hostname to JSON_Filename variable
            JSON_Filename.append("./output/%s/%s-%s.json" % (Hostname, Hostname, command))
            JSON_Command.append(command)
            #Catch command result and put inside file
            Content = '%s\n' % (SSH_Connect.find_prompt())
            Content += SSH_Connect.send_command(command, delay_factor=2)
            Filename = '%s/%s-%s_%s.txt' % (path, Hostname, str(command), datetime.now().strftime(Format))
            OutputFile = open(Filename, 'w').write(Content)
            #print >> log,  ("File %s succesfully generated" % Filename)
            print ("File %s succesfully generated" % Filename)

            #if Parsed function required, call parsing function
            if (isParsed == 'y' or isParsed == 'yes'):
                Parse_Config.TextFSM(Content,Hostname,command,device_type)
                #t = threading.Thread(target=Parse_Config.TextFSM, args=(Content,Hostname,command,device_type))
                #thread.append(t)
                #t.start()
            else:
                return

            #if Write function required, call write function
            if (isWrite == 'y' or isWrite == 'yes'):
                Subtitle = "%s - %s" % (Hostname,command)
                #y = threading.Thread(target=Write_Config.Write_Code, args=(document,Subtitle,Content,docx_filename))
                #thread.append(y)
                #y.start()
                Write_Config.Write_Code(document,Subtitle,Content,docx_FileName)
                #print >> log,  ("File %s succesfully generated" % docx_FileName)
                print ("File %s succesfully generated" % docx_FileName)
            else:
                return
        except:
            return

    #Disconnect SSH connection
    SSH_Connect.disconnect()
    #print >> log,  ("Get configuration of %s is completed" % Hostname)
    print ("Get configuration of %s is completed" % Hostname)
    return

#main function
def main():
    t0 = time.time()
    #Check output folder
    if not os.path.exists("output"):
        os.mkdir ("output")

    #User intervention for parse to spreadsheet and write to docx selection
    isParsed = input('Do you want to parse the output to xlsx file?(Yes/No) : ').lower()
    isWrite = input('Do you want to write the output to docx file?(Yes/No) : ').lower()
    isSummarised = input('Do you want to summarise interfaces information?(Yes/No) : ').lower()

    #Load device database
    with open("./Device/device.csv", "r") as SourceText:
        thread=[]
        Hostnames = set()

        for line in SourceText:
            #line = line.rstrip('\n')
            #Insert string from database into each variable
            hostname,device_type,ip,port,username,password,enable = line.split(',')
            Hostnames.add(hostname)

            #Ping target host, if succesfull continue with SSH
            if not ping(ip):
                ##print >> log, ("Host %s is not reachable" % hostname)
                print ("Host %s is not reachable" % hostname)
            else:
                ##print >> log, "Host %s is reachable" % hostname
                print ("Host %s is reachable" % hostname)
                #Connect_SSH(hostname, device_type, ip, port, username, password, enable, isParsed, isWrite)
                t=threading.Thread(target=Connect_SSH, args=(hostname, device_type, ip, port, username, password, enable, isParsed, isWrite))
                thread.append(t)
                t.start()

    #Join back the Thread and wait until all thread finish to execute
    for t in thread:
        t.join()

    #try:
    #Writing Summary Page to Excel
    if (isSummarised == 'y' or isSummarised == 'yes'):
        #time.sleep(10)
        for Hostname in Hostnames:
            Parse_Config.Write_Summary_Excel(Hostname)
    #except:
    #    print ("Error writing summary file")

    print ("Total Parsing time :", round(time.time()-t0,3), "s")

    # JSON Analysis
    from json_analyzer import get_json, compare_json
    #JSON_Baseline = [''

    #]
    #print (JSON_Filename)



if __name__ == "__main__":
    #If this Python file runs by itself, run below command. If imported, this section is not run
    main()
