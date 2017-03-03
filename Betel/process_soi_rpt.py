# -*- coding: utf-8 -*-

import os
import logging  
import logging.handlers
import imapy
from imapy.query_builder import Q
import re
import datetime
import time
import smtplib
import socket
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.mime.application import MIMEApplication
import glob
import xlwt
#from xlsxwriter.workbook import Workbook
import platform
import pandas as pd
from pandas import Series, DataFrame
import collections

maxcounter = 1000
sht = "Page1_1"
codelist = ['COPO', 'DNGI']
ProdEnvironment = False

if platform.system() == 'Windows':
    fold_tw_rp1282 = "C:\D\DSC Prj\Apple\Prj Eclipse\input\RP1282"
    fold_fileout_tw = "C:\D\DSC Prj\Apple\Prj Eclipse\output\RP1282"
    fold_LOG = 'C:\D\DSC Prj\Apple\Prj Eclipse\input\Log'
else:
    fold_tw_rp1282 = "/usr/workspace/vmirpt/wmos/in"
    fold_fileout_tw = "/usr/workspace/vmirpt/wmos/out"
    fold_LOG = '/usr/workspace/vmirpt/wmos/Log'
    ProdEnvironment = True
 
    
HOST = '199.40.27.1'    #'KULDCCAS1.dhl.com'
PORT = '25'
USER = 'apple.npi@dhl.com'
PWD = 'Welcome7'


#RP1282 = ['RP1282_Apple_TW_APAC_FG_SOI_A063_DAILY_VIEW', 'RP1282_Apple_TW_APAC_FG_SOI_A058_DAILY_VIEW', 'RP1282_Apple_TW_APAC_FG_SOI_A059_DAILY_VIEW', 'RP1282_Apple_TW_APAC_FG_SOI_A066_DAILY_VIEW', 'RP1282_Apple_TW_APAC_FG_SOI_A060_DAILY_VIEW', 'RP1282_Apple_TW_APAC_FG_SOI_A065_DAILY_VIEW', 'RP1282_Apple_TW_APAC_FG_SOI_A064_DAILY_VIEW', 'RP1282_Apple_TW_APAC_FG_SOI_A068_DAILY_VIEW', 'RP1282_Apple_TW_APAC_FG_SOI_A057_DAILY_VIEW' ]
RP1282 = ['A063', 'A058', 'A059', 'A066', 'A060', 'A065', 'A064', 'A068', 'A057', 'A061']
EmailReceiver = {'A060':['sam.ma@dhl.com,ben.wangs@dhl.com', 'Ming.Li2@dhl.com,da.li@dhl.com'],
                 'A057':['ben.wangs@dhl.com', 'sam.ma@dhl.com,da.li@dhl.com,Ming.Li2@dhl.com'],
                 'A068':['da.li@dhl.com,ben.wangs@dhl.com', 'sam.ma@dhl.com,Ming.Li2@dhl.com'],                 
                 }

#EmailReceiver = {'A099':['sam.ma@dhl.com,da.li@dhl.com,ben.wangs@dhl.com', 'shap1816@gmail.com'], 'A094':['ben.wangs@dhl.com,da.li@dhl.com,sam.ma@dhl.com', 'sam.ma@dhl.com']}

  
RptSender = 'scinoreply@dhl.com'
RptSubject = 'Report: RP1282_Apple_TW_APAC_FG_SOI_'
RptUID = 27800 #27800 2017/2/4


def log_init():
    logger = logging.getLogger("wmos_rpt")  
    logger.setLevel(logging.DEBUG)      
    
    hdlr = logging.handlers.RotatingFileHandler(os.path.join(fold_LOG, 'wmosreport.log'), maxBytes=1024*1024, backupCount=30)
    ch = logging.StreamHandler()  
    ch.setLevel(logging.ERROR) 

    formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")  
    ch.setFormatter(formatter)  
    #fh.setFormatter(formatter) 
    hdlr.setFormatter(formatter)
    logger.addHandler(ch)  
    logger.addHandler(hdlr)     
    #return logger
    
class DHLMail(object):
    """Compose and Send via GMail"""
    def __init__(self, email_address, password):
        super(DHLMail, self).__init__()
        self.email_address = email_address
        self.password = password
    
    def email(self, to, cc, subject, text, attch=None):
        module_logger = logging.getLogger("wmos_rpt.sendmail")
        if attch:
            msg = self.sendexcel(to, cc, subject, text, attch)
        else:
            msg = self.compose(to, cc, subject, text)
        try:
#             mailServer = smtplib.SMTP("smtp.gmail.com", 587)
            mailServer = smtplib.SMTP(HOST, PORT)
            mailServer.ehlo()
            mailServer.starttls()
            mailServer.ehlo()
            mailServer.login(self.email_address, self.password)
            #print msg['To']+msg['cc']
            #mailServer.set_debuglevel(1)
            mailServer.sendmail(self.email_address, [to]+[cc], msg.as_string())
            print "Successfully sent email"
            module_logger.info("Successfully sent email: %s" % subject)
            #mailServer.close()
        except smtplib.SMTPException, err:
            print str(err)
            print "Error: unable to send email" 
            module_logger.error(str(err))
        finally:
            #mailServer.quit()
            mailServer.close()
        
  
    def compose(self, to, cc, subject, text):
        msg = MIMEMultipart()
        msg['From'] = self.email_address
        msg['To'] = to
        msg['cc'] = cc
        msg['Subject'] = subject
        msg.attach(MIMEText(text))
        return msg
    
    def sendexcel(self, to, cc, subject, text, file):
        msg = MIMEMultipart()
        msg['From'] = self.email_address
        msg['To'] = to
        msg['cc'] = cc
        msg['Subject'] = subject + os.path.splitext(os.path.basename(file))[0]
        msg.attach(MIMEText(text))        
        
        with open(file, "rb") as fil:
            part = MIMEApplication(
                fil.read(),
                Name=os.path.basename(file)
            )
            part['Content-Disposition'] = 'attachment; filename="%s"' % os.path.basename(file)
            msg.attach(part)
            
        return msg
    
    
def process_xlsx(inputpath):
    module_logger = logging.getLogger("wmos_rpt.processxlsx") 
    
    for excelfile in sorted(glob.glob(r'%s/*.xlsx' % inputpath)):
        #filename = os.path.splitext(os.path.basename(excelfile))[0]        
        #module_logger.info("processing file: %s" % excelfile)
        try:
            process_pandas(excelfile, sht)
        except Exception as e:
            print "**********ERROR Skip**********", e.message
            module_logger.error("**********ERROR Skip[%s]**********" % e.message)        


def fetchemails(mailset, dstfolder, RptSender, RptSubject, mailUID):
    module_logger = logging.getLogger("wmos_rpt.mailbox")
    #"========START @ %s=======", time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
    module_logger.info('===Operating maibox=== %s', USER)
    #module_logger.info("Dest Folder: %s", dstfolder)
          
    #operate mail box
    try:    
        box = imapy.connect(
            host = HOST,
            username = USER,
            password = PWD,
            ssl=True,
        )
        module_logger.info('LOGIN completed, user: %s', USER)
    except Exception as e:
        print e.message
        module_logger.error(e.message)
        return
    
    # select the required email        
    #print datetime.datetime.now().strftime("%d-%b-%Y")
    q = Q()
    emails = box.folder('INBOX').emails(
          q.uid('%s:*' % mailUID).sender(RptSender).subject(RptSubject)
          #q.uid('%s:*' % mailUID).subject(RptSubject)
          #q.uid('%s:*' % RptUID).subject(RptSubject)
#         q.uid('3:*').since('21 Oct 2014')
    )  
   
    t = set([])
    # get attachment info
    if len(emails):
        for em in emails:
#             print em
#             print em['from_whom']
            print em.uid, em['subject']
            mailuid = em.uid
            module_logger.info('Email UID: %s, Subject: %s' % (mailuid, em['subject']))
            
#             headers = email['headers']
#             print('Ip of a sender: {0}'.format(get_ipv4(headers)))
            
            for attachment in em['attachments']:
                # save each attachment in current directory
                file_name = attachment['filename']
                module_logger.info('Attachment filename: %s', file_name)
    #             content_type = attachment['content_type']
                data = attachment['data']
                
                #fn = ("%s/%s.tmp" % (fold_shasummaryout, file_name))
                att_path = os.path.join(dstfolder, file_name)
                #print os.path.basename(att_path)
                #print os.path.splitext(os.path.basename(att_path))[0]
                #elename= os.path.splitext(file_name)[0]
                elename, ext = os.path.splitext(file_name)
                if ext <> '.xlsx':
                    #log here
                    module_logger.error('Attachment extension: %s', file_name)
                    continue
                
                print elename
                elename = elename.strip()[28:32]                
                if elename in mailset:
                    t.add(elename)
                    print elename
                    with open(att_path, 'wb') as f:
                        f.write(data)
                    
                    #move email from 'INBOX' to 'Archive.Report'
                    if os.path.isfile(att_path):
                        em.mark(['flagged', 'seen'])
                        em.move('Archive.SCi')
                        module_logger.info('Mail(%s) marked as Flagged and moved to Archive folder', mailuid)
    else:
        module_logger.warn('No matches found for Predetermined conditions')
        
        
    box.logout()
    module_logger.info('Logout.')
    
    if len(mailset - t) == 0: 
        return True
    else:
        #TO-DO, if partially received...
        if len(emails):
            sMsg = 'Reports partially received. \r\nMissing: \r\n%s' % ', \r\n'.join(set(mailset) - set(t))
            module_logger.error(sMsg)
            mymail = DHLMail(USER, PWD)
            #mymail.email('sam.ma@dhl.com,Applecnsha-DSC-hub@dhl.com,Ming.Li2@dhl.com,Peter.Li2@dhl.com', 'Missing Reports', 'Hello,\r\n%s' % sMsg)
            mymail.email('sam.ma@dhl.com', '', 'Missing Reports', 'Hello,\r\n%s' % sMsg)
            

def process_pandas(infile, sht):
    module_logger = logging.getLogger("wmos_rpt.pandas")
    module_logger.info('===Processing=== %s', infile)
    #emaillog.append('\n===Processing=== %s' % infile)
    print "============", infile
    
    #dateparse = lambda x: pd.datetime.strptime(x, '%Y-%m-%d')
    parse_dates = ['TransDate', 'TransTime']
    df = pd.read_excel(infile, sheetname=sht, header=0, converters={'ConsignPO':str, 'ConsignPO Item':str, 'TransRef':str, 'TransRef Item':str})
    if df.empty:
        print "-----------EMPTY----------"
        module_logger.warn('------This file is EMPTY, Skip to the next------')
        #emaillog.append('\n===This file is EMPTY, Skip to the next=== %s')
        return
    
    print df.dtypes
    df = df.sort_values(by=['TransDate', 'TransTime'], ascending=[True, True])
    print df[['SKU', 'TransDate', 'TransTime']]
    #exit()
    
    df2 = df[df['TransCode'].isin(codelist)]   #all rows where transcode in ['COPO', 'DNGI'] 
    dfrest = df[~df['TransCode'].isin(codelist)]    ## use something.isin(somewhere) and ~something.isin(somewhere) / (something.isin(somewhere)) == False
    
    
    criterion = df2['ConsignPO'].isnull() & df2['ConsignPO Item'].isnull()  
    #print df2.loc[criterion, ['SKU','Mode']]
    dfrestcopo = df2.loc[~(criterion | (df['TransCode']=='DNGI')), :]   #added in UAT
    df2 = df2.loc[criterion | (df['TransCode']=='DNGI'), :]
    
    #print dfrestcopo
    #print df2
    
    # grouped = df2['Qty'].groupby([df2['SKU'], df2['TransCode']])
    # print grouped
    # print grouped.sum()
    # for name,group in grouped:
    #     print name, group
    #print df2
    
#/********originally assume that COPO = DNGI, but is not the truth, so to use sum.unstack() to do so    
#     dfsum = df2.groupby(['SKU', 'TransCode'])['Qty'].sum()
#     #print df2.groupby(by=['SKU', 'TransCode'])['Qty'].sum()
#         
#     tmpsku = ''
#     tmpcode = ''
#     tmpqty = 0
#     for index_val, series_val in dfsum.iteritems():    
#         #print "mmmmmmmmmm", index_val, series_val, tmpsku, tmpcode, tmpqty
#         if tmpsku == '' and tmpcode == '':
#             tmpsku, tmpcode = list(index_val)
#             tmpqty = series_val
#             #print tmpsku, tmpcode, tmpqty 
#         else:            
#             #if tmpsku == list(index_val)[0] and tmpcode in codelist:
#             if tmpcode <> list(index_val)[1] and tmpsku == list(index_val)[0]:        ###added in UAT, to handle DNGI is null
#                 if tmpqty == series_val:                
#                     print "===Identical=== [%s], Qty: %s(%s) = %s(%s)" % (tmpsku, tmpqty, tmpcode, series_val, list(index_val)[1])
#                     module_logger.info("===Identical=== [%s], Qty: %s(%s) = %s(%s)" % (tmpsku, tmpqty, tmpcode, series_val, list(index_val)[1]))              
#                 else:   
#                     print "---Warning--- [%s], Qty: %s(%s) <> %s(%s)" % (tmpsku, tmpqty, tmpcode, series_val, list(index_val)[1])
#                     module_logger.warning("---Warning--- [%s], Qty: %s(%s) <> %s(%s)" % (tmpsku, tmpqty, tmpcode, series_val, list(index_val)[1]))
#                     #return
#             else:
#                 print "-*-Missing-*- [%s], Qty: %s(%s) no match record" % (tmpsku, tmpqty, tmpcode)
#                 module_logger.warning("-*-Missing-*- [%s], Qty: %s(%s) no match record" % (tmpsku, tmpqty, tmpcode))
#                 tmpsku, tmpcode = list(index_val)
#                 tmpqty = series_val
#                 continue
#             tmpsku = ''
#             tmpcode = '' 

#     #print dfsum.index
#     
#     dfcopo = df2[df2.TransCode == 'COPO'].sort_values(by=['SKU','TransDate', 'TransTime'], ascending=[True, True, True])
#     #dfcopo1 = df[df.TransCode == 'COPO' & df['ConsignPO'].isnull() & df['ConsignPO Item'].isnull()]  # or to have original index
#     dfdngi = df2[df2.TransCode == 'DNGI'].sort_values(by=['SKU','TransDate', 'TransTime', 'ConsignPO', 'ConsignPO Item'], ascending=[True, True, True, True, True])
#     
#     #print dfcopo
#     #print dfdngi
#     
#     dngi_row_iterator = dfdngi.iterrows()
#     
#     iremainder = 0
#     ssku = ''
#     processdf = pd.DataFrame(columns=dfcopo.columns)    
#     rowindex = []
# 
#     for index, row in dfcopo.iterrows():
#         print '--COPO--', index, row['SKU'], row['Qty'], iremainder
#         module_logger.info('--COPO-- %s, %s, %s, %s' % (index, row['SKU'], row['Qty'], iremainder))
#         
#         ###    match logic here (without ConsignPO/ConsignPO Item)
#         tmpset = set()
#         tmpdict = {}
#         if iremainder == 0:
#             idx, datarow = dngi_row_iterator.next()
#             print '\t (DNGI)', idx, datarow['SKU'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']
#             module_logger.info('\t (DNGI) %s, %s, %s, %s, %s' % (idx, datarow['SKU'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']))
#             
#             iremainder = datarow['Qty']
#             ssku = datarow['SKU']
#             if datarow['ConsignPO Item'] not in tmpset:
#                 tmpset.add(datarow['ConsignPO Item'])
#                 tmpdict[datarow['ConsignPO Item']] = (datarow['ConsignPO'], datarow['Qty'])
#         
#         icounter = 0
#         while (row['SKU'] == ssku) and (row['Qty'] < iremainder) and icounter < maxcounter:
#             icounter += 1
#             idx, datarow = dngi_row_iterator.next()
#             print '\t --(DNGI)', idx, datarow['SKU'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']
#             module_logger.info('\t --(DNGI) %s, %s, %s, %s, %s' % (idx, datarow['SKU'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']))
#             
#             if datarow['ConsignPO Item'] not in tmpset:
#                 tmpset.add(datarow['ConsignPO Item'])
#                 tmpdict[datarow['ConsignPO Item']] = (datarow['ConsignPO'], datarow['Qty'])
#             iremainder += datarow['Qty']      
#         
#         print '\t  Length:', len(tmpset)
#         module_logger.info('\t  Length: %s' % len(tmpset))
#         print '\t * remainder:', iremainder
#         module_logger.info('\t * remainder: %s' % iremainder)
#         iremainder = iremainder - row['Qty']
#         print '\t * new remainder:', iremainder
#         module_logger.info('\t * new remainder: %s' % iremainder)
#                            
#         ### update dataframe
#         if len(tmpset) > 1:
#             tmpdf = pd.DataFrame(columns=dfcopo.columns)
#             #print tmpset      
#             #tmpdf.loc[dfcopo.index[0]] = dfcopo.loc[index]        
#             s = dfcopo.loc[[index]]
#             #print s
#             #print dfcopo.loc[index, :]
#     #         print s.iloc[0]
#             #tmpdf.append(s, ignore_index=True)
#             #tmpdf = pd.concat([tmpdf, dfcopo.ix[[11],:]])
#             #tmpdf = pd.concat([tmpdf, s])
#             
#     #         if tmpdf.empty:          
#     #             #tmpdf = pd.concat([s]*len(tmpset), ignore_index=True)
#     #             tmpdf = pd.concat([s]*len(tmpset), ignore_index=True)
#     #         else:
#     #             tmpdf = pd.concat([tmpdf, pd.concat([s]*len(tmpset), ignore_index=True)], ignore_index=True)
#             
#             tmpdf = pd.concat([s]*len(tmpset), ignore_index=True)    
#             if len(tmpset) == len(tmpdict):
#                 #dfcopo.loc[index, 'ConsignPO Item'] = 'DELETE'
#                 rowindex.append(index)
#                 irow = 0            
#                 for k, v in sorted(tmpdict.iteritems(), key=lambda (k,v): (v,k)):           
#                     print k, v[0], v[1]                
#                     tmpdf.loc[irow, 'ConsignPO Item'] = k
#                     tmpdf.loc[irow, 'ConsignPO'] = v[0]
#                     tmpdf.loc[irow, 'Qty'] = v[1]
#                     irow += 1
#                     
#                 #another way: Creating a MultiIndex
#                 processdf = pd.concat([processdf, tmpdf], ignore_index=True)
#                 #print tmpdf
#             else:
#                 raise "Error"
#         else:
#             for k, v in tmpdict.items():
#                 dfcopo.set_value(index,'ConsignPO', v[0])
#                 dfcopo.loc[index, 'ConsignPO Item'] = k
#                 break
#     
#     # print tmpdf.loc[[1],:]
#     #print dfcopo
#     
#     print dfrest
#     
#     if not processdf.empty:
#         dfcopo = dfcopo.drop(rowindex)
#         processdf = processdf.sort_values(['SKU', 'ConsignPO', 'ConsignPO Item'], ascending=[1, 1, 1])
#         #print processdf
#         dfcopo = pd.concat([dfrest, dfcopo, processdf, dfrestcopo])
#     else:
#         print ("ProcessDF is Empty")
#         module_logger.info('---ProcessDF is EMPTY---')
#         dfcopo = pd.concat([dfrest, dfcopo, dfrestcopo])
#     
#     dfcopo = dfcopo.sort_values(['SKU', 'ConsignPO', 'ConsignPO Item', 'TransDate', 'TransTime'], ascending=[1, 1, 1, 1, 1])
#     #print dfcopo  
#/********END************

    dfsum = df2['Qty'].groupby([df2['SKU'], df2['TransCode']]).sum().unstack()
    print dfsum
    module_logger.info(dfsum)
    #print isinstance(dfsum, pd.DataFrame)
    #print dfsum.columns
    #print cmp(2, 8), cmp(8, 2), cmp(3, 3)
    #TransCode  DNGI        #if this dfsum, should determine if column 'COPO' exists
    #SKU            
    #HKCQ2PA/A   -10
    #HKEB2PA/A   -10


    dfcoporslt = pd.DataFrame(columns=df2.columns)
    
    dngisubtotal = 0
    for index, row in dfsum.iterrows():        
        #print row.isnull()['COPO'], row.isnull()['DNGI']
        if 'COPO' not in dfsum.columns or row.isnull()['COPO']:
            print 'COPO is null, Ignore!', index
            module_logger.info('%s COPO is null, Ignore!' % index)
        elif 'DNGI' not in dfsum.columns or row.isnull()['DNGI']:
            print 'DNGI is null', index
            module_logger.info('%s DNGI is null, Ignore!' % index)
            
            dfcopo = df2[(df2.TransCode == 'COPO') & (df2.SKU == index)].sort_values(by=['SKU','TransDate', 'TransTime'], ascending=[True, True, True])
            dfcoporslt = pd.concat([dfcoporslt, dfcopo])
        else:
            print "\n \n Index:", index, "row:", row['COPO'], row['DNGI']
            module_logger.info("Index: %s, COPO: %s, DNGI: %s " % (index, row['COPO'], row['DNGI']))
            
            cmprslt = cmp(row['COPO'], row['DNGI'])
            #HHXT2ZM/A -10.0 -14.0    1
            #HHXU2ZM/A  -7.0  -4.0    -1
            #HHXV2MM/A  -7.0  -7.0    0
            print cmprslt
            module_logger.info("COPO vs DNGI: %s" % cmprslt)
            dngisubtotal = row['DNGI']
              
            dfcopo = df2[(df2.TransCode == 'COPO') & (df2.SKU == index)].sort_values(by=['SKU','TransDate', 'TransTime'], ascending=[True, True, False])    
            dfdngi = df2[(df2.TransCode == 'DNGI') & (df2.SKU == index)].sort_values(by=['SKU','TransDate', 'TransTime', 'ConsignPO', 'ConsignPO Item'], ascending=[True, True, False, True, True])
            #print dfcopo
            #print dfdngi
            #print dfdngi.dtypes
            print dfcopo[['SKU', 'TransCode', 'TransDate', 'TransTime', 'Qty']]
            print dfdngi[['SKU', 'TransCode', 'TransDate', 'TransTime', 'Qty']]            
            module_logger.info(dfcopo[['SKU', 'TransCode', 'TransDate', 'TransTime', 'Qty']])
            module_logger.info(dfdngi[['SKU', 'TransCode', 'TransDate', 'TransTime', 'Qty']])
            #dfdngi['TransTime'] = pd.to_datetime(df['TransTime'], format='%Y-%m-%d')
            
#             def test_apply(x):
#                 try:
#                     print 'asdfasdf', type(x), type(datetime.datetime.strptime(x, "%Y-%m-%d"))
#                     return datetime.datetime.strptime(x, "%Y-%m-%d")
#                     #datetime.datetime.strptime(string_date, "%Y-%m-%d %H:%M:%S.%f")
#                 except ValueError:
#                     return None
#             def test_apply1(x):
#                 try:
#                     if isinstance(x, datetime.time):
#                         print type(x)
#                         return x
#                     else:
#                         print type(x), x
#                         print type(datetime.datetime.strptime(x, "%H:%M:%S").time()), 'mmm'
#                         return datetime.datetime.strptime(x, "%H:%M:%S").time()
#                     #datetime.datetime.strptime(string_date, "%Y-%m-%d %H:%M:%S.%f")
#                 except ValueError:
#                     return None
#                  
#             #dfdngi['TransDate'] = dfdngi['TransDate'].astype('datetime64[ns]')
#             dfdngi['TransDate'] = dfdngi['TransDate'].apply(test_apply)
#             #dfdngi['TransTime'] = dfdngi['TransTime'].astype('datetime64[ns]')
#             dfdngi['TransTime'] = dfdngi['TransTime'].apply(test_apply1).astype('datetime64[ns]')
#             print dfdngi
#             print dfdngi.dtypes
#             print 'DNGI Rows:', len(dfdngi.index)
            

            #print  dfcopo.set_index([dfcopo[col] for col in ['TransRef', 'TransRef Item']])
#use TransRef + TransRef Item as Index
#print dfcopo.set_index(dfcopo['TransRef'].astype(str) + '_' + dfcopo['TransRef Item'].astype(str))
            
            dngi_row_iterator = dfdngi.iterrows()
            
#             def annotate(gen):
#                 prev_i, prev_val = 0, gen.next()
#                 for i, val in enumerate(gen, start=1):
#                     yield prev_i, prev_val
#                     prev_i, prev_val = i, val
#                 yield '-1', prev_val            
#             for i, val in annotate(dngi_row_iterator):
#                 print i, val
                
            iremainder = 0
            tmpremainder = 0
            ssku = ''
            processdf = pd.DataFrame(columns=dfcopo.columns)
            rowindex = []
            
            ###    match logic here
            dngiqty = 0 ###To stop Iteration when cmprslt = -1
            dngipos = 0
            lastdngiitem = False
            dngirowcount = len(dfdngi.index)
            
            #tmpdict = {} unsorted
            tmpdict = collections.OrderedDict() #is a dictionary subclass that remembers the order in which its contents are added.
            bFirst = True
            for index, row in dfcopo.iterrows():
                #if (cmprslt == -1) and (dngiqty == dngisubtotal):
                #    break
                
                print '--COPO--', index, row['SKU'], row['Qty'], iremainder
                module_logger.info('--COPO-- %s, %s, %s, %s' % (index, row['SKU'], row['Qty'], iremainder))
                
                tmpset = collections.OrderedDict()  #changed to OrderedDisct(),  set()  ###used to record 'ConsignPO Item' for COPO split
                #if lastdngiitem:    #dngipos == len(dfdngi.index) or dngiqty == dngisubtotal:
                    #detect the last item while iterating over dfdngi.iterrows() 
                #    break
                #else:
                #    tmpdict = {}
                    
                if bFirst:         #if iremainder == 0: as COPO/DNGI DF are filtered for specific SKU, no need to use remainder
                    bFirst = False
                    #if (cmprslt == -1):  #and (dngiqty == dngisubtotal):
                        ###for the situation that dngi split for multi COPO rows
                    #    pass
                    #else:                    
                    idx, datarow = dngi_row_iterator.next()
                    dngipos += 1
                    if dngipos == dngirowcount:
                        lastdngiitem = True
                        module_logger.info('\t  !detect the last iterrows Pos: %s ' % dngipos)
                                
                    print '\t 1st(DNGI)', idx, datarow['SKU'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']
                    module_logger.info('\t 1st(DNGI) %s, %s, %s, %s, %s' % (idx, datarow['SKU'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']))
                    
                    iremainder = datarow['Qty']
                    dngiqty = datarow['Qty']                    
                    ssku = datarow['SKU']
                    
                    #tmpset.add(datarow['ConsignPO Item'])
                    tmpset[datarow['ConsignPO'] + '_' + datarow['ConsignPO Item']] = [datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']]
                    tmpdict[datarow['TransRef'] + '_' + datarow['TransRef Item']] = [datarow['TransRef'] + '_' +  datarow['TransRef Item'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty'], 1, datarow['Qty']]
                    #tmpdict[datarow['ConsignPO Item']] = (datarow['ConsignPO'], datarow['Qty'])
                    #if datarow['ConsignPO Item'] not in tmpset:
                    #    tmpset.add(datarow['ConsignPO Item'])
                    #    tmpdict[datarow['ConsignPO Item']] = (datarow['ConsignPO'], datarow['Qty'])
                    #    print "dict", tmpdict
                else:
                    dngiqty_remain = 0
                    if bool(tmpdict):
                        if len(tmpdict) > 1:
                            pass                   
                        remainv = tmpdict.values()[0]                        
                        iremainder = remainv[-1]    #if tmpdict is notnull, last element saved the unused DNGI Qty from last loop
                        dngiqty_remain = remainv[-1]
                        #tmpset.add(remainv[-4])
                        tmpset[remainv[-5] + '_' + remainv[-4]] = [remainv[-5], remainv[-4], remainv[-1]]     
                    else:
                        iremainder = 0
                    
                    if not lastdngiitem and (row['Qty'] < iremainder):
                        idx, datarow = dngi_row_iterator.next()
                        dngipos += 1
                        if dngipos == dngirowcount:
                            lastdngiitem = True
                            module_logger.info('\t  !detect the last iterrows Pos: %s ' % dngipos)
                    
                        iremainder += datarow['Qty']
                        dngiqty = dngiqty_remain + datarow['Qty']                    
                        ssku = datarow['SKU']
                    
                        #if datarow['ConsignPO Item'] not in tmpset:
                            #tmpset.add(datarow['ConsignPO Item'])
                        if (datarow['ConsignPO'] + '_' + datarow['ConsignPO Item']) not in tmpset:                            
                            tmpset[datarow['ConsignPO'] + '_' + datarow['ConsignPO Item']] = [datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']]                    
                        tmpdict[datarow['TransRef'] + '_' + datarow['TransRef Item']] = [datarow['TransRef'] + '_' + datarow['TransRef Item'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty'], 1, datarow['Qty']]
                         
                        
                icounter = 0                        #icounter used to avoid dead loop 
                while (row['Qty'] < iremainder) and (icounter < maxcounter):  #(row['SKU'] == ssku) and 
                    icounter += 1
                    
                    print '\t [WHILE], SKU, cmprslt, dngisubtotal, dngiqty:', ssku, cmprslt, dngisubtotal, dngiqty
                    module_logger.info('\t [WHILE], SKU: %s, cmprslt: %s, dngisubtotal: %s, dngiqty %s:' % (ssku, cmprslt, dngisubtotal, dngiqty))
                    if lastdngiitem:        #if (cmprslt == -1) and (dngiqty == dngisubtotal):
                        break  ###To stop Iteration when cmprslt = -1
                    
                    idx, datarow = dngi_row_iterator.next()
                    dngipos += 1
                    if dngipos == dngirowcount:
                        lastdngiitem = True
                        module_logger.info('\t  !detect the last iterrows Pos: %s ' % dngipos)
                        
                    dngiqty += datarow['Qty']
                    
                    print '\t --(DNGI)', idx, datarow['SKU'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']
                    module_logger.info('\t --(DNGI) %s, %s, %s, %s, %s' % (idx, datarow['SKU'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']))
                    
                    #if datarow['ConsignPO Item'] not in tmpset:
                    #    tmpset.add(datarow['ConsignPO Item'])
                    if (datarow['ConsignPO'] + '_' + datarow['ConsignPO Item']) not in tmpset:                            
                        tmpset[datarow['ConsignPO'] + '_' + datarow['ConsignPO Item']] = [datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty']]
                        #tmpdict[datarow['ConsignPO Item']] = (datarow['ConsignPO'], datarow['Qty'])                         
                        # for further DB save, use index as Key, tmpdict[datarow['ConsignPO Item']] = (datarow['ConsignPO'], datarow['Qty'])
                        print "\t tmpdict:", tmpdict
                    tmpdict[datarow['TransRef'] + '_' + datarow['TransRef Item']] = [datarow['TransRef'] + '_' + datarow['TransRef Item'], datarow['ConsignPO'], datarow['ConsignPO Item'], datarow['Qty'], 1, datarow['Qty']]
                    iremainder += datarow['Qty']
                
                if icounter == maxcounter:
                    pass
                
                print '\t  Length of tmpset:', len(tmpset)
                module_logger.info('\t  Length of tmpset: %s' % len(tmpset))
                print '\t * remainder:', iremainder
                module_logger.info('\t * remainder: %s' % iremainder)
#                iremainder = iremainder - row['Qty']
#                print '\t * new remainder:', iremainder
#                module_logger.info('\t * new remainder: %s' % iremainder)
                                   
                ### update dataframe
                #rewrite the update logic on 2/22/2017
                print len(tmpset)              
                if len(tmpset) > 1:
                    print "---Processing Split COPO here---"
                    module_logger.info("---Processing Split COPO here---")
                    #print tmpdict
                    #need to split current COPO
                    tmpremainder = iremainder - row['Qty']
                    
                    tmpdf = pd.DataFrame(columns=dfcopo.columns)     
                    s = dfcopo.loc[[index]]
                     
                    tmpdf = pd.concat([s]*len(tmpset), ignore_index=True) 
                    
                    #dfcopo.loc[index, 'ConsignPO Item'] = 'DELETE'
                    rowindex.append(index)
                    
                    irow = 0
                    ilen = len(tmpset)
                    
                    #for k, v in sorted(tmpdict.iteritems(), key=lambda (k,v): (v,k)):   ###To sort a dictionary by its value (only for Python 2, " Python3 key=lambda k,v: v,k")
                    
                    print "tmpset: ", tmpset
                    for k, v in tmpset.items():  
                        #print k, v
                        tmpdf.loc[irow, 'ConsignPO'] = v[0]            
                        tmpdf.loc[irow, 'ConsignPO Item'] = v[1]
                        print "Remarinder: ", iremainder
                        if cmprslt <> -1:                            
                            tmpdf.loc[irow, 'Mode'] = 'Split done (%s)' % row['Qty']
                        else:
                            if iremainder <= row['Qty']:
                                tmpdf.loc[irow, 'Mode'] = 'Split done (%s)' % row['Qty']
                            else:
                                tmpdf.loc[irow, 'Mode'] = 'Split partially done (%s)' % row['Qty']
                                
                        print row['Qty']
                        
                        myiter = tmpdict.iteritems()
                        splitqty = 0
                        #eachiterqty = 0
                        mylist = []
                        while True:
                            try:
                                x, y = next(myiter)
                                # loop body
                                #print "Y: ", len(tmpdict), y
                                if y[-5] == v[0] and y[-4] == v[1]:                                    
                                    splitqty += y[-1]
                                    mylist.append(y)
                                    print "Y: ", len(tmpdict), splitqty, y
                            except StopIteration:
                                break
                        
                        print "split row:", irow, splitqty, mylist
                        
                        #if irow == ilen - 1:
                        #    print "last~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
                        tmpdf.loc[irow, 'Qty'] = splitqty
                        tmpdf.loc[irow, 'Aging'] = str(mylist)
                        
                        irow += 1
                    
                    #after tmpdict iterated over done, use the same logic as BELOW ELSE
                    #tmpremainder = iremainder - row['Qty']  #if 0, identical else has remainder                 
                    lastk, lastv = tmpdict.popitem()
                    if cmprslt == -1 and lastdngiitem and iremainder > row['Qty']:      ###TO-DO, if COPO has positive number
                        lastv[-1] = lastv[-1]
                    else:
                        lastv[-1] = lastv[-1] - tmpremainder
                        #To update last record qty and list
                        #print irow
                        tmpdf.loc[irow-1, 'Qty'] = lastv[-1]        #Changed on 3/1/2017, logic error
                        tmpdf.loc[irow-1, 'Qty'] = splitqty - tmpremainder
                        mylist.pop()
                        mylist.append(lastv)
                        tmpdf.loc[irow-1, 'Aging'] = str(mylist)
                    
                    print lastv
                    tmpdict = collections.OrderedDict()
                    if tmpremainder <> 0:
                        #to be used for next iteration
                        if tmpremainder < 0:
                            lastv[-1] = tmpremainder
                            tmpdict[lastk] = lastv
                            print lastv    
                             
                    #another way: Creating a MultiIndex
                    processdf = pd.concat([processdf, tmpdf], ignore_index=True)
                    #print tmpdf
                    module_logger.info('>>>Split COPO<<< [%s](%s), %s link with <DNGI>: ' % (row['SKU'], row['Qty'], datarow['TransRef'] + '_' + datarow['TransRef Item']))
                    module_logger.info('\n\t '.join(map(str, mylist)))

                elif len(tmpset) == 1:
                    mylist = []
                    lastk, lastv = tmpdict.popitem()
                    dfcopo.set_value(index, 'ConsignPO', lastv[1])
                    dfcopo.loc[index, 'ConsignPO Item'] = lastv[2]
                    
                    tmpremainder = iremainder - row['Qty']  #if 0, identical else has remainder
                    if cmprslt <> -1: #and not lastdngiitem:
                        #make sure that COPO qty fully matched
                        dfcopo.loc[index, 'Mode'] = 'Done'
                    else:
                        if iremainder <= row['Qty']:                ###TO-DO, if COPO has positive number
                            dfcopo.loc[index, 'Mode'] = 'Done'
                        else:
                            dfcopo.loc[index, 'Mode'] = 'Partially Done (%s)' % row['Qty']
                            dfcopo.loc[index, 'Qty'] = iremainder
                            
                    #update used qty, default lastv[-3] = lastv[-1]
                    if cmprslt == -1 and lastdngiitem and iremainder > row['Qty']:      ###TO-DO, if COPO has positive number
                        lastv[-1] = lastv[-1]
                    else:
                        lastv[-1] = lastv[-1] - tmpremainder
                    
                    if bool(tmpdict):                            
                        for k, v in tmpdict.items():
                            mylist.append(v)                    
                    mylist.append(lastv)                      
                    #print mylist
                    dfcopo.loc[index, 'Aging'] = str(mylist)
                    
                    tmpdict = collections.OrderedDict()
                    if tmpremainder <> 0:
                        #to be used for next iteration
                        if tmpremainder < 0:
                            lastv[-1] = tmpremainder
                            tmpdict[lastk] = lastv
                            print tmpdict
                    
                    module_logger.info('>>>Normal COPO<<< [%s](%s), %s link with <DNGI>: ' % (row['SKU'], row['Qty'], datarow['TransRef'] + '_' + datarow['TransRef Item']))
                    module_logger.info('\n\t '.join(map(str, mylist)))
                    
                    #if cmprslt == -1 and lastdngiitem and tmpremainder <> 0:                        
                    #    break
                    #if tmpremainder <= 0:
                    #    continue
                    
                    
#                 if (cmprslt <> -1) and len(tmpset) > 1: ###len(tmpset) > 1 means COPO need to be splitted
#                     tmpdf = pd.DataFrame(columns=dfcopo.columns)
#                     #print tmpset      
#                     #tmpdf.loc[dfcopo.index[0]] = dfcopo.loc[index]        
#                     s = dfcopo.loc[[index]]
#                     
#                     tmpdf = pd.concat([s]*len(tmpset), ignore_index=True)    
#                     if len(tmpset) == len(tmpdict):
#                         #dfcopo.loc[index, 'ConsignPO Item'] = 'DELETE'
#                         rowindex.append(index)
#                         irow = 0            
#                         for k, v in sorted(tmpdict.iteritems(), key=lambda (k,v): (v,k)):   ###To sort a dictionary by its value (only for Python 2, " Python3 key=lambda k,v: v,k")       
#                             print k, v[0], v[1]                
#                             tmpdf.loc[irow, 'ConsignPO Item'] = k
#                             tmpdf.loc[irow, 'ConsignPO'] = v[0]
#                             tmpdf.loc[irow, 'Qty'] = v[1]
#                             irow += 1
#                             
#                         #another way: Creating a MultiIndex
#                         processdf = pd.concat([processdf, tmpdf], ignore_index=True)
#                         #print tmpdf
#                     else:
#                         raise "Error"
#                 elif (cmprslt == -1) and (dngiqty == dngisubtotal):
#                     ##to handle DNGI < COPO, DNGI run out, sometimes need to split COPO Qty
#                     if len(tmpset) > 1:
#                         pass
#                     else:
#                         for k, v in tmpdict.items():
#                             dfcopo.set_value(index,'ConsignPO', v[0])
#                             dfcopo.loc[index, 'ConsignPO Item'] = k
#                             dfcopo.loc[index, 'Qty'] = v[1] ###dngiqty 
#                             break
#                         print "-------->", dfcopo
#                         continue                        
#                 else:
#                     for k, v in tmpdict.items():
#                         dfcopo.set_value(index,'ConsignPO', v[0])
#                         dfcopo.loc[index, 'ConsignPO Item'] = k
#                         break                

                    
            if not processdf.empty:
                dfcopo = dfcopo.drop(rowindex)
                processdf = processdf.sort_values(['SKU', 'ConsignPO', 'ConsignPO Item'], ascending=[1, 1, 1])
                #print processdf
                #dfcopo = pd.concat([dfrest, dfcopo, processdf, dfrestcopo])
                dfcopo = pd.concat([dfcopo, processdf])
                
                criterion11 = dfcopo['ConsignPO'].isnull() & dfcopo['ConsignPO Item'].isnull() #added on 3/1/2017
                dfcopo = dfcopo.loc[~(criterion11), :]
            else:
                print ("ProcessDF is Empty")
                module_logger.info('---ProcessDF is EMPTY---')
                #dfcopo = pd.concat([dfrest, dfcopo, dfrestcopo])
                #print dfcopo                
                criterion11 = dfcopo['ConsignPO'].isnull() & dfcopo['ConsignPO Item'].isnull()                
                dfcopo = dfcopo.loc[~(criterion11), :]
                #print dfcopo          
            
            dfcoporslt = pd.concat([dfcoporslt, dfcopo])
            #print dfcoporslt
            
    dfcoporslt = pd.concat([dfrest, dfcoporslt, dfrestcopo])      
    dfcoporslt = dfcoporslt.sort_values(['SKU', 'ConsignPO', 'ConsignPO Item', 'TransDate', 'TransTime'], ascending=[1, 1, 1, 1, 1])

    
    ###Excel
    fnPrefix = '' if ProdEnvironment else 'Output_'
    
    outfile = os.path.join(fold_fileout_tw, fnPrefix+os.path.basename(infile))
    print outfile
    module_logger.info('Writing to output file: %s' % os.path.basename(infile))
    writer = pd.ExcelWriter(outfile)
    dfcoporslt.to_excel(writer, sht, index=False)
    writer.save()

def send_email(outputpath):
    module_logger = logging.getLogger("wmos_rpt.email") 
    
    mymail = DHLMail(USER, PWD)
   
    for excelfile in glob.glob(r'%s/*.xlsx' % outputpath):
        module_logger.info("sending file: %s" % excelfile)
        
        fn = os.path.splitext(os.path.basename(excelfile))[0].strip()        
        vendorcode = fn[28:32]
        sTo = ''
        sMsg = ''
        recipient = EmailReceiver.get(vendorcode, ())
        if recipient:            
            sTo = recipient[0]
            sCc = recipient[1]
            print recipient, sTo, sCc
        else:
            sTo = 'ben.wangs@dhl.com,da.li@dhl.com'
            sCc = 'sam.ma@dhl.com,Ming.Li2@dhl.com'
            sMsg = 'Rpt recipient is not found.\r\n'        
                
        mymail.email(sTo, sCc, 'Report: ' if ProdEnvironment else 'TEST: ', sMsg, excelfile)


def purgefiles(dstfolder):
    module_logger = logging.getLogger("wmos_rpt.delfiles")
    
    #Delete all files in directory
    for the_file in os.listdir(dstfolder):
        file_path = os.path.join(dstfolder, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
                module_logger.info('(%s) deleted', the_file)
            #elif os.path.isdir(file_path): shutil.rmtree(file_path)
        except Exception as e:
            print(e) 
            
            
def DeleteOutputTmpFiles():
    if datetime.date.today().isoweekday() <> 7: #Mon=1, Sat=6, Sun=7
        return
    
    purgefiles(fold_fileout_tw)

def DeleteArchivedEmails():
    if datetime.date.today().isoweekday() <> 7: #Mon=1, Sat=6, Sun=7
        return
    
    module_logger = logging.getLogger("wmos_rpt.delemails")

    #operate mail box
    try:    
        box = imapy.connect(
            host = HOST,
            username = USER,
            password = PWD,
            ssl=True,
        )
        module_logger.info('LOGIN completed, user: %s', USER)
    except Exception as e:
        print e.message
        module_logger.error(e.message)
        return
        
#     q = Q()
    emails = box.folder('Archive.SCi').emails() 

    if len(emails):
        for em in emails:
            #print em.uid, em['subject']
            em.mark(['deleted'])            
    
    box.logout()
    module_logger.info('Logout.')
     
                        
def main():
    #TO-DO, Log, and think about how to put them all into DB
    log_init()
    
    module_logger = logging.getLogger("wmos_rpt.main")
    thislogstart = '========START @ %s=======' % time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time()))
    module_logger.info(thislogstart)
    
    if ProdEnvironment:
        purgefiles(fold_fileout_tw)
        purgefiles(fold_tw_rp1282)    
    
    #delete all files under output folder: out/outszx
#    DeleteOutputTmpFiles()
    
    #TO-DO, save uid, filename into DB
    mailUID = RptUID
    
    ###fetch attachments from specific mailbox
    if ProdEnvironment:
        fetchemails(set(RP1282), fold_tw_rp1282, RptSender, RptSubject, mailUID)
        #pass
    
    ###process excel with Pandas
    process_xlsx(fold_tw_rp1282)    

    if ProdEnvironment:
        send_email(fold_fileout_tw)
        DeleteArchivedEmails()    
    
    module_logger.info('========END @ %s=======' % time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())))
    
    if ProdEnvironment:    
        sbody = ''
        HeadMark = False
        if thislogstart <> '':
            with open(os.path.join(fold_LOG, 'wmosreport.log')) as f:
                f = f.readlines() 
            print f 
            for line in f:
                if not HeadMark and thislogstart in line:
                    HeadMark = True
                
                if HeadMark:
                    sbody = sbody + line 
            print sbody
            
            mymail = DHLMail(USER, PWD)
            mymail.email('ben.wangs@dhl.com', 'sam.ma@dhl.com,da.li@dhl.com,Ming.Li2@dhl.com', 'RP1282 Process Summary %s' % datetime.datetime.now().strftime("%Y%m%d"), sbody)
            #mymail.email('sam.ma@dhl.com', '', 'RP1282 Process Summary %s' % datetime.datetime.now().strftime("%Y%m%d"), sbody)
    
if __name__ == '__main__':
    main()
