#-------------------------------------------------------------------------------
# Name:       Tdriver
# Purpose: Optimised for multi-processing
#
# Author:      vishx008
#
# Created:     14/09/2014
# Copyright:   (c) vishx008 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import os
import re
import csv
import win32com.client
from urllib2 import url2pathname
from time import time, gmtime, strftime

class TDriver:


    def __init__(self, speciality, source, destination):


        self.spec = speciality
        self.src = self.tidy(source)
        self.dst = self.tidy(destination,speciality)

    def sortFiles(self):

        if os.path.isdir(self.src):
            self._files = os.listdir(self.src)
        else:
            return False

        #Create sub-directories
        self._good = self.tidy(self.dst,'good')
        self._bad = self.tidy(self.dst,'bad')
        self._ugly = self.tidy(self.dst,'ugly')
        self._pdf = self.tidy(self.dst,'pdf')

        try:
            for adir in [self.dst, self._good, self._bad, self._ugly, self._pdf]:
                if not os.path.isdir(adir): os.makedirs(adir)
        except:
            return False

        #create/open index files _speciality_good.csv
        self._goodfile = open(self.tidy(self.dst, '_' + self.spec + '_good.csv'),'w+')
        self._badfile = open(self.tidy(self.dst, '_' + self.spec + '_bad.csv'),'w+')
        #create csv writers for index files
        self._goodwriter = csv.writer(self._goodfile, dialect = csv.excel)
        self._badwriter = csv.writer(self._badfile, dialect = csv.excel)

        #create temporary lists to hold good and bad filenames
        self._good_list = list()
        self._bad_list = list()
        #create an instance of msword
        self._word = win32com.client.DispatchEx('Word.Application')
        self._word.Visible = False

        #start iterating through filenames
        for self._aFile in self._files:

            if self._aFile.find('.doc')<0 or self._aFile.find('~')>=0: continue

            #try:

            self._doc = self._word.Documents.Open(self.tidy(self.src, self._aFile), False, False, False)
            self._result = self.scanForNHSIds(self._doc.Content.Text)

            if len(self._result) == 12:
                self._pdfname = self._result.replace(' ','') + '.pdf'
                self._pdfpath = self.tidy(self._pdf,self._pdfname)

                #if the file already exists, save with timestamp
                #this is an unlikely scenario
                if os.path.exists(self._pdfpath):
                    self._pdfname = self._result.replace(' ','') + '_multifile_' + strftime('%H_%M_%S', gmtime()) + '.pdf'
                    self._pdfpath = self.tidy(self._pdf, self._pdfname)
                #FileFormat 17 is pdf
                self._doc.SaveAs(self._pdfpath,FileFormat = 17)
                self._doc.SaveAs(self.tidy(self._good,self._aFile))
                self._good_list.append([self._aFile, self._result])
            else:
                self._doc.SaveAs(self.tidy(self._bad, self._result + '-'+ self._aFile))
                self._bad_list.append([self._aFile, self._result])

            if self._doc:self._doc.Close()

            #except (RuntimeError, TypeError, NameError):
                #self._result = 'rd_err'
               # self._bad_list.append([self._aFile, 'rd_err'])
                #print('error')
                #if self._doc:self._doc.Close()
        #print 'Successful PDF exports to good folder: ' + str(len(tmp_good))
        #print 'Files with errors copied to bad folder: ' + str(len(tmp_bad))
        #save list to index file
        self._goodwriter.writerows(self._good_list)
        self._badwriter.writerows(self._bad_list)

        self._goodfile.close()
        self._badfile.close()
        self._word.Quit()

    def scanForNHSIds(self,t):
        pattern = '[0-9]{3} [0-9]{3} [0-9]{4}'

        if  re.search(pattern, t):
            m = re.findall(pattern, t)
            first_match = m[0]
            for amatch in m:
                if first_match != amatch:
                    return 'n_uni' #non-unique
            return first_match
        else:
            return 'no_id'

    def tidy(self,p1,p2=None):
        if not p2:
            return url2pathname(p1)
        else:
            return url2pathname(os.path.join(p1,p2))
