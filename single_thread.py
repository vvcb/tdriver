#-------------------------------------------------------------------------------
# Name:     Single_threaded_version
# Purpose:  Demo - Word COM server not thread-safe and prevents easy multiprocessing
#           Use several machines instead!!!
#
# Author:      vishnuvcb
#
# Created:     15/09/2014
# Copyright:   (c) vishnuvcb 2014
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import tdriver
import os
from time import time, gmtime, strftime
def main():
    start_time = time()
    print 'Random sample of 1000 files from colorectal folder.'
    print 'Start time: ', strftime("%a, %d %b %Y %H:%M:%S", gmtime())
    

    source_paths = dict (colorectal = os.path.join('C:', 'Users', 'vishx008', 'td', 'colorectal'),
                        mix = os.path.join('C:', 'Users', 'vishx008', 'td', 'mix'),
                        breast = os.path.join('C:', 'Users', 'vishx008', 'td', 'breast'))
    destination_path = os.path.join('C:', 'Users', 'vishx008', 'td')
    
    machine = tdriver.TDriver(speciality='colorectal',
                source = source_paths['colorectal'],
                destination = destination_path)
    machine.sortFiles()
    end_time = time()
    print 'End time: ', strftime("%a, %d %b %Y %H:%M:%S", gmtime())
    elapsed_time = end_time-start_time
    print 'Elapsed time(seconds): ', elapsed_time
    print 'Elapsed time(minutes) for 1000 files:',elapsed_time/60
    print 'Estimated time(hours) for 20000 colorectal files:', (20000/1000)*(elapsed_time/(60*60))
    print 'If it takes 1 minute for 1 cat to eat 1 mouse, how long will it take 10 cats to eat 10 mice?'
    import this    
if __name__ == '__main__':
    main()
