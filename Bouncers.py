# -*- coding: utf-8 -*-
"""
Created on Tue Jul 17 13:51:15 2018
Edit: 20-7-2018
@author: Timo de Wit
"""
from win32com.client import Dispatch

outlook=Dispatch("Outlook.Application").GetNamespace("MAPI")
root_folder = outlook.Folders.Item(1)

def getOutlook():
    for i in range(6):
        try:
            TmpName = outlook.Folders.Item(i).Name
            if TmpName == '....................@......':
                #print('NR: ' + str(i) + ' ' + TmpName)
                return i
        except:
            pass
        
    if TmpName != "....................@......":
       return False


def getMessages(tmpMailbox):
    namespace = Dispatch("Outlook.Application").GetNamespace("MAPI")
    root_folder = namespace.Folders.Item(tmpMailbox) #choose account in this situation we took i
    subfolder = root_folder.Folders['Inbox'] #choose folder, subfolder
    #subfolderO = root_folder.Folders['Inbox'].Folders['Closed'] #choose folder, subfolder
    all_inbox = subfolder.Items
    msg = all_inbox.GetFirst()
    
    #f = open("bouncers.txt","w+")
    
    for msg in all_inbox:
        if msg.Class==43:
            if msg.SenderEmailType=='EX':
                print(msg.Sender.GetExchangeUser().PrimarySmtpAddress + ' ' + msg.Subject)
                # print(msg.Sender.GetExchangeUser().PrimarySmtpAddress,file=f)
                print("Subj: " + msg.Subject)
                print("Body: " + msg.Body)
                print("========")
            else:
                print(msg.SenderEmailAddress + ' ' + msg.Subject)
    #               print(msg.SenderEmailAddress,file=f)
                print("Subj: " + msg.Subject)
                print("Body: " + msg.Body)
                print("========")
        if msg.Class==46: #46 = ReportItem
    #           print("Subj: " + msg.Subject)
            tmpDeelnemernummer = msg.Subject[msg.Subject.find("deelnemernummer")+16:]       
            mySubString=msg.Body[msg.Body.find("<")+1:msg.Body.find(">")]
            print(tmpDeelnemernummer + ";" + mySubString)
    #           print(mySubString,file=f)
        else:
            print("Uitval gevonden")
    ##           print(msg.Sender)
    ##           print(msg.Sender,file=f)
    ##           print("Subj: " + msg.Subject)
    ##           print("Body: " + msg.Body)
            print("========")    
         
    #f.close()

def main():
    tmpMailbox = getOutlook()
    if tmpMailbox != False:
        """ We found the index of the mailbox. 
            Now read all the emails from the specific folder.
        """
        getMessages(tmpMailbox)  
        print("all done")
    else:
        print("Error: Mailbox not found")
        
# Standard boilerplate to call the main() function to begin
# the program.
if __name__ == '__main__':
    main()