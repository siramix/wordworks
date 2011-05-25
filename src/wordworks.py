#!/usr/bin/python
#
# Copyright (C) 2011 Siramix Labs
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

__author__ = 'cpreynolds@gmail.com (Patrick Reynolds)'

import gdata.spreadsheet.service
import getpass
import sys

class Spreadsheet:

    def __init__(self, email, password):
        """
        Contructor taking an email and a password.
        """
        self.gd_client = gdata.spreadsheet.service.SpreadsheetsService()
        self.gd_client.email = email
        self.gd_client.password = password
        self.gd_client.source = 'WordWorks'
        self.gd_client.ProgrammaticLogin()
        
        self.targetSheet = 'Buzzwords Secret Words'
        self.targetWorksheet = 'Starter Deck'
        
        self.feed = None
        self.wsFeed = None
        self.wsKey = None
        self.sheetKey = None

    def findSheet(self):
        """
        List the sheets that we can use.
        """
        q = gdata.spreadsheet.service.DocumentQuery()
        q['title'] = self.targetSheet
        q['title-exact'] = 'true'
        self.feed = self.gd_client.GetSpreadsheetsFeed(query=q)
        #self._PrintFeed(self.feed)
        
        # Bail out if we cannot find the spreadsheet we seek
        if( len(self.feed.entry) != 1 ):
            print "%s not found in your Spreadsheets Feed" % self.targetSheet
            sys.exit(-1)
        
        self.sheetKey = self.feed.entry[0].id.text.rsplit( '/',1 )[-1]
        self.listFeed = self.gd_client.GetListFeed(self.sheetKey,2)
        return self.listFeed
        
    def _PrintFeed(self, feed):
        for i, entry in enumerate(feed.entry):
            if isinstance(feed, gdata.spreadsheet.SpreadsheetsCellsFeed):
                print '%s %s\n' % (entry.title.text, entry.content.text)
            elif isinstance(feed, gdata.spreadsheet.SpreadsheetsListFeed):
                print '%s %s %s' % (i, entry.title.text, entry.content.text)
                # Print this row's value for each column (the custom dictionary is
                # built using the gsx: elements in the entry.)
                print 'Contents:'
                for key in entry.custom:  
                    print '  %s: %s' % (key, entry.custom[key].text) 
                    print '\n',
            else:
                print '%s %s\n' % (i, entry.title.text)

def check_data(data,word):
    s = "\n".join(data);
    if word in s:
        return True

    return False

def main():
    """
    Main function.
    """
    email = raw_input('Email: ')
    password = getpass.getpass('Password: ')
    sheet = Spreadsheet(email,password)
    entries = sheet.findSheet().entry
    data = [e.custom['title'].text.lower() for e in entries]
    while True:
        print "There are now %i rows in the spreadsheet" % len(sheet.findSheet().entry)
        row = dict()
        row['title'] = raw_input('New Word: ')
        if row['title'].lower() == 'quit':
            break
        if check_data(data,row['title'].lower()):
            print '%s is already a word!' % row['title']
            continue
        for i in range(1,6):
            row['bw%s'%i] = (raw_input('Bad Word: '))
        row['category'] = ''
        row['pack'] = ''
        entry = sheet.gd_client.InsertRow(row, sheet.sheetKey, 2)
        if isinstance(entry, gdata.spreadsheet.SpreadsheetsList):
            print "Insert row succeeded."
        else:
            print "Insert row failed."
        
if __name__ == '__main__':
    main()
