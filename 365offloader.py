#!/usr/bin/env python3
import os
from os import path
from configparser import ConfigParser
from exchangelib import (Credentials, DELEGATE, Account, Configuration, Mailbox, FileAttachment, ItemAttachment, Folder)

parser = ConfigParser()
parser.read('exchange_config.ini')

#To access values in config.ini > [SECTION], Key Name
conf_password = parser.get('Invoices','password')
conf_username = parser.get('Invoices','username')
conf_smtp_address = parser.get('Invoices','smtp_address')
conf_server = parser.get('Invoices','server')
conf_folder = parser.get('Invoices','folder')
conf_attachment_file_name = parser.get('Invoices','filename_search')

credentials = Credentials(username=conf_username, password=conf_password)

config = Configuration(server=conf_server, credentials=credentials)
a = Account(primary_smtp_address=conf_smtp_address, config=config, autodiscover=False, access_type=DELEGATE)

if not os.path.exists('tmp'):
    os.makedirs('tmp')
local_path = os.getcwd()
staging_dir = r'tmp'

exchange_folder = a.inbox / conf_folder  # folder inbox to read from under the inbox directory ex. Trash, Junk, etc.


for item in exchange_folder.all():
    for attachment in item.attachments:
        if isinstance(attachment, FileAttachment):
            if conf_attachment_file_name in attachment.name:
                # file name = Message UUID + file_time_stamp + file_extension, In my use case had to add UUID to file name 
                # because attachment names were not unique
                local_path = os.path.join(staging_dir, attachment.content_id[:32] + \
                    '_' + str(attachment.last_modified_time.strftime("%Y%m%d")) + '_' + attachment.name)
                with open(local_path, 'wb') as f:
                    f.write(attachment.content)
                    # if you uncomment below, the message will be marked as read, so re-import won't happen
                    #item.is_read = True
                    #item.save(update_fields=['is_read'])
                    print('Saved attachment to', local_path)


print("Import Done")
