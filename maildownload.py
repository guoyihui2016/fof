# -*- coding: utf-8-*-

import os
import re
import time
import email
import poplib
import imaplib
import cStringIO
import shutil
import sys
import nfzbutl1
import  dbutl
import sendmail

reload(sys)
sys.setdefaultencoding('gbk')

DEFAULT_PORT = {
    "pop3": {False: 110, True: 995},
    "imap4": {False: 143, True: 993},
}

cfgfile=u'F:\\fof\\fof_cfg.xlsx'
logfile=u'F:\\fof\\fof_log.txt'
smode='1'


def writelog(errmsg):
    print errmsg
    return True

def writelog1(errmsg):
    writelog(errmsg)
    return True

def logonmailbox(cfg):
    '''获取邮箱设置信息，目前重excel配置文件获取'''
    mailinfo= nfzbutl1.getexceldata(cfg,'mailinfo','mail','nfzbfof',1,0,4)   #
    protocal=mailinfo['protocal']
    host=mailinfo['server']
    port=int(mailinfo['port'])
    usr=mailinfo['mailaddr']
    pwd=mailinfo['password']
    # print host,port,usr,pwd
    try:
        conn = imaplib.IMAP4_SSL(host, port)
        conn.login(usr, pwd)
        writelog('logon success!')

    except BaseException as e:
        writelog("Connect to {0}:{1} failed".format(host, port)+" ({0})".format(e))
    return conn

def scanmailbox(conn):
    list_pattern = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')
    type_, folders = conn.list()

    for folder in folders:
        flags, delimiter, folder_name = list_pattern.match(folder).groups()
        if folder_name!="INBOX" :
            continue

        conn.select(folder_name, readonly=True)
        type_, data = conn.search(None, "ALL")
        msg_id_list = [int(i) for i in data[0].split()]
        msg_num = len(msg_id_list)

        for i in msg_id_list:
            writelog('mailid: {0}'.format(i))
            type_, data = conn.fetch(i, "(RFC822)")
            msg = email.message_from_string(data[0][1])
            parsemail(msg)
            if False:
                print msg
                break
    dbinfo = nfzbutl1.getexceldata(cfgfile,'db','db','nfzbdsdb',1,1,4)
    dbcfg=dbinfo['usr']+'/'+dbinfo['pwd']+'@'+dbinfo['db']
    sqltext=nfzbutl1.getexceldata(cfgfile,'sql','id','lwdate',1,1,4)['sqltext']
    lwdate=dbutl.excute_dbquery(sqltext,dbcfg)[0][0]

    checkmaildownload(cfgfile,lwdate)
    return True

def checkmaildownload(cfgfile,lwdate):
    mfilter=nfzbutl1.getexceldata2(cfgfile,'mailfilter')
    dbinfo = nfzbutl1.getexceldata(cfgfile,'db','db','nfzbdsdb',1,1,4)
    dbcfg=dbinfo['usr']+'/'+dbinfo['pwd']+'@'+dbinfo['db']
    sqltext=nfzbutl1.getexceldata(cfgfile,'sql','id','filedownqry',1,1,4)['sqltext'].format(lwdate)
    print 'sqltext1:',sqltext
    mdl=dbutl.excute_dbquery(sqltext,dbcfg)

    a=[b[0] for b in mdl]
    mailtext='各位好,前一交易日（{0}）FOF系统还有如下子基金外部估值表文件未收到：'.format(lwdate)+'\n\n'
    textlist=''
    # for mf in mfilter:
    #     if mf['prod_code'] not in a:
    #         print mf['prod_code']
    if mfilter.__len__() >0:
        for mf in mfilter:
            if mf['prod_code'] not in a:
                textlist=textlist+mf['prod_code']+ ':'+mf['fundname'] + '\n'
        mailtext= mailtext.decode('utf-8')+textlist
        to_mail=nfzbutl1.getexceldata(cfgfile,'sendmail','mailtype','outfilecheck',1,1,4)['to_mail']
        sendmail.sendMail(mailtext,'123',to_mail)
    return True

def getcharset(str):
    if  str.find("?")>=0 and str.find("?",str.find("?")+1)>= 0 :
        return str[str.find("?")+1:str.find("?",str.find("?")+1)]
    else:
        return 'gbk'

def parsemail(mailmsg):
    subject = mailmsg.get("subject")
    print 'subject5:',subject
    charset=getcharset(subject)
    mailsubject = email.Header.decode_header( subject)[0][0].decode(charset)
    print 'mailsubject:',mailsubject
    sendfrom = email.utils.parseaddr(mailmsg.get("from"))[1]
    mailfilter1=nfzbutl1.getexceldata2(cfgfile,'mailfilter')
    for mfilter in mailfilter1:
        if mailsubject.find(mfilter['mailsubject']) >= 0:
            print 'matched msubject:', mfilter['mailsubject']
            parse_attachment(mailmsg,mfilter)
    return True

def parse_attachment(mailmsg,mfilter):
    for part in mailmsg.walk():
        filename=part.get_filename()
        if filename and not part.is_multipart():
            if filename=='@':
                continue
            charset=getcharset(filename)
            print 'fcharset:', charset
            ctype=part.get("Content-Type")
            print 'ctype:', ctype
            if  ctype.find('application/octet-stream') < 0 and ctype.find('application/vnd.ms-excel') < 0 and ctype.find('application/msexcel') < 0:
                continue
            print 'ctype:', ctype
            filename=email.Header.decode_header(email.Header.Header(filename))[0][0]
            filename=filename.decode(charset)
            print 'filename2:',filename
            name = str(unicode(filename))
            filedate=getfiledate(filename)
            if smode=='1' and   fileexists(filedate,mfilter):
                record(mfilter,filedate)
                continue
            exten=filename[filename.find('.'):]
            newfilename=mfilter['filename']
            filepath=mfilter['dir']
            if filedate and exten and newfilename:
                if int(mfilter['mode'])==1:
                    newfile=filedate+newfilename+exten
                elif  int(mfilter['mode'])==2:
                    if filename.find(mfilter['attachment'].decode())>=0:
                        newfile=filedate+newfilename+exten
                    else:
                        continue

                print 'newfile:',newfile
                newfile=os.path.join(filepath, newfile)
                try :
                    content = part.get_payload(decode = True)
                    with open(newfile, "wb") as f:
                        f.write(content)
                    record(mfilter,filedate)
                except BaseException as e:
                    print("[-] Write file of email failed: {0}".format(e))
    return  True

def record(mfilter,lwdate):
    dbinfo = nfzbutl1.getexceldata(cfgfile,'db','db','nfzbdsdb',1,1,4)
    dbcfg=dbinfo['usr']+'/'+dbinfo['pwd']+'@'+dbinfo['db']
    print dbcfg
    sqltext=nfzbutl1.getexceldata(cfgfile,'sql','id','filedownrec',1,1,4)['sqltext'].format(mfilter['prod_code'], lwdate)
    # print 'sqltext:',sqltext
    # sqltext='''select * from dim_fund_relation_20170221 a where a.vc_fundcode like '202A7_' '''
    rlt=dbutl.excute_sql(sqltext,dbcfg)
    print rlt
    return True


def getfiledate(filename):
    year_date = '2017'
    if filename.find(year_date) >= 0:
        print 'filename3:',filename
        nian='年'.decode('utf-8')
        yue='月'.decode('utf-8')
        ri='日'.decode('utf-8')
        length = filename.find(year_date)
        print length
        temp = filename[length+4]
        #日期的三种格式：一，20151030；二，2015_10_30；三，2015年10月30
        if temp.isdigit():
            print 'cond1:'
            date = filename[length:length+8]
            print date
        elif temp == '_' or temp == '-':
            year = filename[length:length+4]
            month = filename[length+5:length+7]
            day = filename[length+8:length+10]
            date = year + month + day
            print date
        else :
            if filename.find(nian) >= 0 :
                print 'niannian:'
                year = filename[length:length+4]
                name=filename.decode()
                month=filename[filename.find(nian,filename.find('2017'),filename.find(ri))+1:filename.find(yue,filename.find('2017'),filename.find(ri))]
                day=filename[filename.find(yue,filename.find('2017'),filename.find(ri))+1:filename.find(ri,filename.find(yue))]
                print name
                print name.find(nian,name.find('2017'),name.find(ri))
                print name.find(yue,name.find('2017'),name.find(ri))
                print name.find(yue,name.find('2017'),name.find(ri))
                if month.__len__()==1:
                    month='0'+month
                if day.__len__()==1:
                    day='0'+month
                date = year + month + day
                print date
        if date:
            return date
        else :
            return  False
    else:
        return False

def downloadmail():
    mailconn=logonmailbox(cfgfile)
    scanmailbox(mailconn)
    return True

def fileexists(filedate,mfilter):
    files=os.listdir(mfilter['dir'])
    dlist=[file[0:8] for file in files]
    if filedate in dlist:
        return True
    else:
        return False


# checkmaildownload(cfgfile,'20170219')
# rltcol=nfzbutl1.getexceldata2(cfgfile,'mailfilter')
# for x in rltcol :
#     print 'x:',x
downloadmail()

