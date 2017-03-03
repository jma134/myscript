# -*- coding: utf-8 -*-
'''
Created on Jan 21, 2017

@author: SHAP1816
'''

# configurable parameters #
days_old = 500    #specify the number of days old that the file is. If you enter +5, it will move files older than 5 days
#org_folder = "C:\\Users\\shap1816\\Downloads\\EDI\\Output214"    # original folder 
org_folder = "Z:\\IT Share1\\Cocoa"
dst_folder = "C:\\Users\\shap1816\\Downloads\\EDI\\HKG Eclipse\\backup"        # destination folder
LOG_FILE = "C:\\Users\\shap1816\\Downloads\EDI\\HKG Eclipse\\log.log"        # log file
process_mode = 'COPY'   #MOVE or COPY
IgnoreFileType = set(['.pdf', '.jpg', '.xxx'])   #filename end with any element of this collection will be ignored in the move/copy process 
# end of configurable parameters #


import logging
import logging.handlers
import os
from os import walk
import shutil
import time
from datetime import datetime, date, timedelta


def log_init():
    logger = logging.getLogger("eclipse")  
    logger.setLevel(logging.INFO)      
    
    hdlr = logging.handlers.RotatingFileHandler(LOG_FILE,maxBytes=1024*1024,backupCount=30)
    ch = logging.StreamHandler()  
    ch.setLevel(logging.ERROR)

    formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")  
    ch.setFormatter(formatter)  
    #fh.setFormatter(formatter) 
    hdlr.setFormatter(formatter)
    logger.addHandler(ch)  
    logger.addHandler(hdlr)  

# start process #
def run():
    move_logger = logging.getLogger("eclipse.file")
    
    move_date = date.today() - timedelta(days=days_old)
    move_date = time.mktime(move_date.timetuple())
       
#     logger = logging.getLogger("eclipseshare")
#     hdlr = logging.FileHandler(logfile)
#     hdlr.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s: %(message)s'))
#     logger.addHandler(hdlr) 
#     logger.setLevel(logging.INFO)
       
    move_logger.info("========START @ %s=======", time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
    move_logger.info("Operation Mode: %s", process_mode)
    
    count = 0
    size = 0.0
    
    for root, subdirs, files in os.walk(org_folder):
        print('--\nLocation = %s, subdirectories = %s, files = %s' % (root, len(subdirs), len(files)))
        move_logger.info("||Location = %s, subdirectories = %s, files = %s" % (root, len(subdirs), len(files)))
               
        for filename in files:            
            #print('\t- file %s (full path: %s)' % (filename, file_path))
            #print ('\t- file %s' % (filename))            
            
            if os.path.splitext(filename)[1] in IgnoreFileType:
                print ('\t- Ignored %s' % (filename))
                move_logger.warning("Ignored '" + filename + "'." )                
                continue
            
            src_file = os.path.join(root, filename)            
            newDir = os.path.join(dst_folder, root[1+len(org_folder):])
            if not os.path.exists(newDir):
                os.makedirs(newDir)
                move_logger.warning("MakeDirs: (" + newDir + ")" ) 
            dst_file = os.path.join(newDir, filename)
            #dst_file = os.path.join(dst_folder, filename)                        
            if os.stat(src_file).st_mtime < move_date:
                if not os.path.isfile(dst_file):
                    get_size = os.path.getsize(src_file)
                    size = size + (get_size / (1024*1024.0))
                    filesize = str(round ((get_size/(1024)),1)) + 'k'             
                    
                    try:
                        if process_mode == 'MOVE':
                            shutil.move(src_file, dst_file)
                            print ('\t- [MOV] %s' % (src_file)) 
                            move_logger.info("Archived [MOV]'" + src_file + "'," + filesize)                            
                        else:
                            shutil.copy(src_file, dst_file)
                            print ('\t- [CPY] %s' % (src_file)) 
                            move_logger.info("Archived [CPY]'" + src_file + "'," + filesize)
                        #shutil.copytree(src, dst, symlinks, ignore)
                        
                        count = count + 1
                    except OSError, why:
                        if WindowsError is not None and isinstance(why, WindowsError):
                            pass
                        else:
                            move_logger.error(src_file + dst_file + str(why))              
                
    
    move_logger.info("||Summary: " + str(count) + " files, Total Size: " + str(round(size,2)) + " MB.")
    print "Summary: " + str(count) + " files, Total Size: " + str(round(size,2)) + " MB."
    move_logger.info("========END @ %s======="  + "\n", time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))


if __name__ == '__main__':
    log_init()
    run()
