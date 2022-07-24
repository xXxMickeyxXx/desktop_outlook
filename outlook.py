"""Class for interacting with Desktop Outlook. Objects of "Outlook" use Window's COM interface and processes, via the "pywin32" package.
                             _______________________________________
                            |                                       |
                            |         *****INFORMATION*****         |
                            |_______________________________________|


__________INSTALLATION_REQUIREMENTS:
• pywin32 package (most current)
• psutil package (most current)
 
"""

import psutil
import win32com.client as client
import subprocess as subpro
import os
import csv  
import tempfile  
from pathlib import Path
from typing import Union


# Check to see if Desktop Outlook application is open (by checking for it's process name)
def ensure_outlook_application():
    for pro in psutil.process_iter(['name']):
        if pro.info["name"].lower() == "outlook.exe":
            return True, pro
    return (False,)
    

# Open Desktop Outlook application
def open_outlook_application():
    if not ensure_outlook_application():
        os.startfile("outlook")
        
        
# Close Desktop Outlook application
def close_outlook_application():
    check_outlook = ensure_outlook_application()
    if check_outlook[0]:
        subpro.run("TASKKILL /F /IM outlook.exe", stdout=subpro.DEVNULL, stderr=subpro.DEVNULL)
    

# Restart Desktop Outlook application in the event that scripting/program/app fails to properly interact with the com process(es)
def restart_outlook_application():
    close_outlook_application()
    os.startfile("outlook")
    
    
class Outlook:
    """ Represents user's Desktop Outlook account. Objects of this class are instantiated with user's account name (typically just the email address of which the mailbox is associated with, such as "workaccount@workdomain.com"). 
    
        _____ATTRIBUTES/PROPERTIES_____
        • "account" - Current account instantiated with "Outlook" object
        • "outlook_application" - COM object itself, used to interact with Desktop Outlook
        • "inbox"
        • "folder"
        • "emails"
        • "attachments"
        
    
        _____METHODS_____
        • "set_folder"
        • "save_attachments"
        • "attachment_data"
        •
    """
    
    _outlook_application = client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
    def __init__(self, account: str) -> str:
        self._account = account
        self._mailbox = self._set_account(account)
        self.inbox = self._set_inbox()
        self._folder = None
        self._emails = None
        self._attachments = None
        
    def __str__(self):
        return f"Outlook Mailbox_____\nACCCOUNT: {self._account}\nFOLDER: {self._folder}\nEMAILS: {self._emails}\nATTACHMENTS: {self._attachments}"
        
    @property
    def account(self):
        # TODO: Determine if additional logic is needed/if a "setter" method is needed
        return self._account
        
    @property
    def folder(self):
        # TODO: Add validatdion for making sure "_set_folder" value is a com object
        if not self._folder:
            # TODO: Create and raise custom error here
            raise ValueError(f"FOLDER IS NOT SET...")
        return self._folder.Name

    @property
    def emails(self):
        # TODO: Add validation, if needed be
        return self._emails
        
    @property
    def attachments(self):
        # TODO: Add validation, if needed be
        return self._attachments
        
    def open_outlook(self):
        open_outlook_application()
        
    def close_outlook(self):
        close_outlook_application()
        
    def restart_outlook(self):
        restart_outlook_application()

    def _set_account(self, account):    
        account = self._outlook_application.Folders(account)
        return account
        
    def _set_inbox(self):
        return self._mailbox.Folders("Inbox")
        
    def set_folder(self, desired_folder, com_object=None):
        """Get folder (as COM object) from account.
        
            _____TO-DOs_____
            • Add caching capabilities.
        """
        if not self._mailbox or not self.inbox:
            raise ValueError(f"No account exists for this object; please set account...")
        
        if com_object is None:
            com_object = self.inbox

        try:
            if com_object.Name == desired_folder:
                self._folder = com_object
                self._get_emails()
                self._get_attachments()
                return com_object
        except (NameError, AttributeError) as error:
            pass
            
        for folder in com_object.Folders:
            result = self.set_folder(desired_folder, folder)
            if isinstance(result, type(com_object)):
                return result
            elif result is None:
                continue
            else:
                return False
                    
    def _get_emails(self):
        email_list = []
        for email in self._folder.Items:
            email_list.append(email)
        if email_list:
            self._emails = email_list
        else:
            self._emails = False
    
    def _get_attachments(self):
        if not self._emails:
            self._attachments = False
        else:
            attachment_list = []
            for email in self.emails:
                for attachment in email.Attachments:
                    attachment_list.append(attachment)
            self._attachments = attachment_list
        
    def save_attachments(self, save_directory):
        if not self._attachments:
            raise ValueError(f"NO ATTACHMENTS FOUND...")
        for attachments in self.attachments:
            attachments.SaveAsFile(os.path.join(save_directory, attachments.FileName))
    
    def attachment_data(self, delimiter=",", skip_header=True):
        if not self._attachments:
            return []
        temp_dir = tempfile.TemporaryDirectory()
        temp_dir_name = temp_dir.name
        data = []
        for attachment in self.attachments:
            file_name = attachment.FileName
            full_path = os.path.join(temp_dir_name, file_name)
            attachment.SaveAsFile(os.path.join(full_path))
            with open(full_path, "r", newline="") as in_file:
                if skip_header:
                    next(in_file)
                reader = csv.reader(in_file, delimiter=delimiter)         
                for line in reader:
                    if not line:
                        continue
                    data.append(line)
        return data
        
                
if __name__ == "__main__":
    pass
        
