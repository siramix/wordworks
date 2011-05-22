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
        self.curr_key = ''
        self.curr_wksht_id = ''
        self.list_feed = None   




    
    def listSheets(self):
        """
        List the sheets that we can use.
        """
        q = gdata.spreadsheet.service.DocumentQuery()
        q['title'] = 'Buzzwords Secret Words'
        q['title-exact'] = 'true'
        feed = self.gd_client.Query(q.ToUri())
        self._PrintFeed(feed)
      
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

def main():
    """
    Main function.
    """
    email = raw_input('Email: ')
    password = getpass.getpass('Password: ')
    sheet = Spreadsheet(email,password)
    sheet.listSheets()
    

if __name__ == '__main__':
    main()
