#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      vishx008
#
# Created:     11/09/2014
# Copyright:   (c) vishx008 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import os
import re
import csv
import win32com.client
from urllib2 import url2pathname
from time import time, gmtime, strftime

# Colorectal Surgery        T:\SurgicalServices-RPH\Patient Letters Surgery, Head & Neck\General Surgery\Colorectal\
# Ophthalmology             T:\SurgicalServices-RPH\Patient Letters Surgery, Head & Neck\Ophthalmology\
# ENT                       T:\SurgicalServices-RPH\Patient Letters Surgery, Head & Neck\ENT\
# General Surgery           T:\SurgicalServices-RPH\Patient Letters Surgery, Head & Neck\General Surgery\General Surgery\
# Restorative Dentistry     T:\SurgicalServices-RPH\Patient Letters Surgery, Head & Neck\Maxfac, Oral, Orthodontics and Restorative Dentistry\Restorative\
# Special Care Dentistry	T:\SurgicalServices-RPH\Patient Letters Surgery, Head & Neck\Special Care Dentistry\
# DiabetesEndocrinology     T:\MedicalDirectorate-CDGH\Secretaries - CDGH\Diabetes Endocrinology\
# Upper GI Surgery          T:\SurgicalServices-RPH\Patient Letters Surgery, Head & Neck\Upper GI\
# Urology                   T:\SurgicalServices-RPH\Patient Letters Surgery, Head & Neck\Urology\
# Orthodontics              T:\SurgicalServices-RPH\Patient Letters Surgery, Head & Neck\Maxfac, Oral, Orthodontics and Restorative Dentistry\Orthodontics\
# Oral & Maxillo-Facial	    T:\SurgicalServices-RPH\Patient Letters Surgery, Head & Neck\Maxfac, Oral, Orthodontics and Restorative Dentistry\Max Fac\

p_base = os.path.join('T:', 'SurgicalServices-RPH', 'Patient Letters Surgery, Head & Neck')
paths = dict (colorectal = os.path.join(p_base, 'General Surgery', 'Colorectal'),
                eye = os.path.join(p_base, 'Ophthalmology'),
                ent = os.path.join(p_base, 'ENT'),
                general = os.path.join(p_base, 'General Surgery', 'General Surgery'),
                dentist_restorative = os.path.join(p_base, 'Maxfac, Oral, Orthodontics and Restorative Dentistry', 'Restorative'),
                dentist_special = os.path.join(p_base, 'Special Care Dentistry'),
                upper_gi = os.path.join(p_base, 'Upper GI'),
                urology = os.path.join(p_base, 'Urology'),
                orthodontics = os.path.join(p_base,  'Maxfac, Oral, Orthodontics and Restorative Dentistry', 'Orthodontics'),
                maxfac = os.path.join(p_base,  'Maxfac, Oral, Orthodontics and Restorative Dentistry', 'Max Fac'))
                #diabetes = os.path.join('T:', 'MedicalDirectorate-CDGH', 'Secretaries - CDGH', 'Diabetes Endocrinology')

p_indices = os.path.join('C:','Users','vishx008', 'td')

#this is for local copy of letters
test_paths = dict (colorectal = os.path.join('C:', 'Users', 'vishx008', 'td', 'colorectal'),
                    upper_gi = os.path.join('C:', 'Users', 'vishx008', 'td', 'upper_gi'),
                    mix = os.path.join('C:', 'Users', 'vishx008', 'td', 'mix'),
                    breast = os.path.join('C:', 'Users', 'vishx008', 'td', 'breast'))

def sortFiles(indexname):
    print indexname

    #get pathname corresponding to indexname
    #check to make sure it is a directory
    #get list of filenames in that directory
    p_files = url2pathname(test_paths.get(indexname))
    #Use this for TDrive original files
    p_source = url2pathname(paths.get(indexname))

    #Use this for local copy of test files
    #p_source = test_paths.get(indexname)
    if os.path.isdir(p_source):
        filenames = os.listdir(p_source)
        print 'Total number of files to be sorted:' + str(len(filenames))
    else:
        return
    #set path and filenames for storing index files
    g = indexname + '_good.csv'
    b = indexname + '_bad.csv'
    #store the index files in the same directory as original files
    p_g = url2pathname(os.path.join(p_files, g))
    p_b = url2pathname(os.path.join(p_files, b))
    #open index files
    goodfile = open(p_g,'w+')
    badfile = open(p_b,'w+')
    #create csv writers for index files
    goodwriter = csv.writer(goodfile, dialect = csv.excel)
    badwriter = csv.writer(badfile, dialect = csv.excel)

    #create temporary lists to hold good and bad filenames
    tmp_good = list()
    tmp_bad = list()
    #create an instance of msword
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False

    #--------Sampler - remove for production------------------
    #filenames = shuffle(filenames)
    i = 0
    #--------End of Sampler - remove for production------------------

    #start iterating through filenames
    for aFile in filenames[10:20]:
        #if this is not a word document or is a temporary file, move on
        if aFile.find('.doc')<0 or aFile.find('~')>=0: continue

        try:
            #set docname to filepath + filename
            #url2pathname avoids %20
            docname = url2pathname(os.path.join(p_files, aFile))
            # open the document
            doc = word.Documents.Open(docname, False, False, False)
            result = scanForNHSIds(doc.Content.Text)

            if len(result) == 12:
                result = result.replace(' ','') #get rid of spaces
                pdfname = result + '.pdf'
                pdfpath = url2pathname(os.path.join(p_files,'good',pdfname))

                #if the file already exists, save with timestamp
                #this is an unlikely scenario
                if os.path.exists(pdfpath):
                    pdfname = result + '_multifile_' + strftime('%H_%M_%S', gmtime()) + '.pdf'
                    pdfpath = url2pathname(os.path.join(p_files,'good',pdfname))
                #FileFormat 17 is pdf
                doc.SaveAs(pdfpath,FileFormat = 17)
                tmp_good.append([aFile,result])
            else:
                docpath = url2pathname(os.path.join(p_files,'bad',result + '-'+ aFile))
                doc.SaveAs(docpath)
                tmp_bad.append([aFile,result])

            #close the word document if it has been opened
            if doc:doc.Close()
            #export count

        except:
            result = 'rd_err'
            tmp_bad.append([aFile,result])
        i += 1
        if i>1000:break

    print 'Successful PDF exports to good folder: ' + str(len(tmp_good))
    print 'Files with errors copied to bad folder: ' + str(len(tmp_bad))
    #save list to index file
    for a in tmp_good:
        goodwriter.writerow(a)

    for a in tmp_bad:
        badwriter.writerow(a)

    goodfile.close()
    badfile.close()
    del doc
    del word

def scanForNHSIds(t):
    #this function scans for all instances of nhs numbers in a string
    #if a unique nhs id is found, it is returned
    #if different nhs ids are found, it returns 'non-unique'
    #if no nhs id is found, it returns 'no_nhs_id'

    pattern = '[0-9]{3} [0-9]{3} [0-9]{4}'

    if  re.search(pattern, t):
        m = re.findall(pattern, t)
        old = m[0]
        for amatch in m:
            if old != amatch:
                return 'n_uni'
        return old
    else:
        return 'no_id'

def main():

   pass
if __name__ == '__main__':
    #main()
    start_time = time()
    print 'Start time: ', strftime("%a, %d %b %Y %H:%M:%S", gmtime())
    #for apath in paths.items(): sortFiles(apath[0])
    sortFiles('colorectal')

    end_time = time()
    print 'End time: ', strftime("%a, %d %b %Y %H:%M:%S", gmtime())
    print 'Elapsed time(minutes):', (end_time - start_time)/60
