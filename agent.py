# encoding:utf-8
 __author__ ="zhangzhe <zhangzhe0707@gmail.com>"
 __version__ =1.0
 __date__ ="12-10-29"
  
  
 import os
 import sys
 import win32com.client
 import ConfigParser
 import MySQLdb
 import logging
 import logging.handlers
 import ftplib
 import datetime
 import time
 import zipfile
 import socket
 import threading
  
  
 configP =ConfigParser.RawConfigParser()
 def read_config(config_file_path,field,key):
     configP =ConfigParser.ConfigParser()
     try:
         configP.read(config_file_path)
         result =configP.get(field,key)
     except:
         log_error('log',"read Config Error:%s" % config_file_path + " "+field)
         # sys.exit(1)
         result =False
     return result
  
  
 def write_config(config_file_path,field,key,value):
     configP =ConfigParser.ConfigParser()
     result=''
     try:
         configP.read(config_file_path)
         configP.set(field,key,value)
         configP.write(open(config_file_path,'w'))
         result =True
     except:
         log_error('log',"write Config Error:%s" % config_file_path + " "+field)     
         # sys.exit(1)
         result =False
     return result
  
 #��ȡ�ű��ļ��ĵ�ǰ·��
 #Get Code File Local Path
 def cur_file_dir():
      #Get Code Path
      path = sys.path[0]
      #�ж�Ϊ�ű��ļ�����py2exe�������ļ�������ǽű��ļ����򷵻ص��ǽű���Ŀ¼�������py2exe�������ļ����򷵻ص��Ǳ������ļ�·��
      if os.path.isdir(path):
         return '\\'.join(path.split('\\')[:-2])
      elif os.path.isfile(path):
         return '\\'.join((os.path.dirname(path)).split('\\')[:-1])
  
  
 localIP =socket.gethostbyname(socket.gethostname())                         #Get local IP
  
 #strFilePath ='\\'.join(os.getcwd().split('\\')[:-2])                         #Get local file path
 strFilePath =cur_file_dir()
  
 strLoggerPath =strFilePath + "\Agent"                                         #Set local Agent logger path 
 strFrameWorkPath =strFilePath + "\QTPScripts"
 strConfigurationPath =strFrameWorkPath +"\\bin\\Configuration.ini"
  
 CYCLETIME =read_config(strConfigurationPath,'TIME', 'cycle_time')                                                 #��ȡѭ��ִ�м��ʱ��
 write_config(strConfigurationPath,'FTPSERVER', 'ftpserver_userdefaultpath',strFilePath + '\\Agent\\FTPUSER')         #���û�Ĭ��·��д�������ļ�
  
 DB_HOST =read_config(strConfigurationPath,'DB', 'db_host')
 DB_PORT =read_config(strConfigurationPath,'DB', 'db_port')
 DB_USER =read_config(strConfigurationPath,'DB', 'db_user')
 DB_PWD =read_config(strConfigurationPath,'DB', 'db_pwd')
 DB_NAME =read_config(strConfigurationPath,'DB', 'db_name')
  
  
 #Get FTP server address
 FTPSERVER_HOST =read_config(strConfigurationPath,'FTPSERVER', 'ftpserver_host')
 #Get FTP server port                                     
 FTPSERVER_PORT =read_config(strConfigurationPath,'FTPSERVER', 'ftpserver_port') 
 #Get FTP access account
 FTPSERVER_USER =read_config(strConfigurationPath,'FTPSERVER', 'ftpserver_user')
 #Get FTP access password
 FTPSERVER_PWD =read_config(strConfigurationPath,'FTPSERVER', 'ftpserver_pwd') 
 #Get FTP account access
 FTPSERVER_USERPERM =read_config(strConfigurationPath,'FTPSERVER', 'ftpserver_userperm')
 #Get FTP default access directory
 FTPSERVER_USERDEFAULTPATH =read_config(strConfigurationPath,'FTPSERVER', 'ftpserver_userdefaultpath')
 #Get FTP anonymous user access directory
 FTPSERVER_ANONYMOUSPATH =read_config(strConfigurationPath,'FTPSERVER', 'ftpserver_anonymouspath')
  
  
 '''
 Logging Cofnig:
 '''
 try:
     from config.config import conf
 except ImportError:
     conf ={}
     conf["logger"] ={}
     conf["logger"]["path"] ="%s%slogs" % (strLoggerPath, os.sep)                                                                            #Log�ļ����·��
     # conf["logger"]["format"] ="%(asctime)s>>(FUNCNAME:%(funcName)s)-%(levelname)s-%(message)s"                                            #Log�ļ����ݸ�ʽ
     conf["logger"]["format"] ="%(asctime)s>>-%(levelname)s-%(message)s"                                                                     #Log�ļ����ݸ�ʽ
     conf["logger"]["backupcount"] =7                                                                                                        #Log������ʱָ�������ı����ļ��ĸ���
     conf["logger"]["level"] =logging.DEBUG                                                                                                  #����Log�������
 except :
     raise
  
  
  
 class logger():
     """
     """
     Ins =None
  
     @staticmethod
     def getLogger(modname):
         """
         """
         if None == logger.Ins:
             logger.Ins =InitLogging(modname)
         return logger.Ins
  
 def InitLogging(modname):
     """
     """
     logger =logging.getLogger(modname)
     format =logging.Formatter(conf["logger"]["format"])
  
     logpath =conf["logger"]["path"]
     if os.path.exists(logpath)==False:
         os.makedirs(logpath)
     logfile =os.path.join(logpath, "%s.log" % modname)
     handler =logging.handlers.TimedRotatingFileHandler(filename =logfile
                                                         , when ="D"
                                                         , backupCount =conf["logger"]["backupcount"])
     handler.setFormatter(format)
     logger.addHandler(handler)
     logger.setLevel(conf["logger"]["level"])
     return logger
  
 def log(LogName,level, msg, *args, **kwargs):
     """
     """
     # exst =traceback.extract_stack()
     # filepath =exst[-2][0]
     # file =os.path.split(filepath)[-1]
     # moduleName =os.path.splitext(file)[0]
  
     logger.getLogger(LogName).log(level, msg, args, kwargs)
  
 def log_info(LogName,*msg):
     """
     """
     logger.getLogger(LogName).info(msg)
  
 def log_debug(LogName,*msg):
     """
     """
     logger.getLogger(LogName).debug(msg)
  
 def log_warning(LogName,*msg):
     """
     """
     logger.getLogger(LogName).warning(msg)
  
 def log_error(LogName,*msg):
     """
     """
     logger.getLogger(LogName).error(msg)
  
 def log_exception(LogName,*msg):
     """
     """
     logger.getLogger(LogName).exception(msg)
  
 def log_critical(LogName,*msg):
     """
     """
     logger.getLogger(LogName).critical(msg)
  
          
 class SQLExample():
     #��ʼ��DB����
     def __init__(self):
         try:
             self.conn =MySQLdb.connect(host=DB_HOST,user=DB_USER,passwd=DB_PWD,db=DB_NAME)
             self.cursor =self.conn.cursor()
         except TypeError,err:
             log_error('log',"Mysql Connction Error:%s" % err)
  
  
     #Execute the SQL command
     def ExecuteSQL(self,sql):
         self.cursor.execute(sql)
         self.results =self.cursor.fetchall()
         return self.cursor.rowcount
  
     #Back to SQL execution results
     def Result(self):
         return self.results
  
     #Cancellation of DB connection
     def __Del__(self):
         self.cursor.close()
         self.conn.close()
  
  
 class FTP_Client(): 
     bolIsDir =False
     def __init__(self):
         try:
             self.ftp =ftplib.FTP()
             self.ftp.connect(FTPSERVER_HOST)
             self.ftp.login(FTPSERVER_USER,FTPSERVER_PWD)
             log_info('log',self.ftp.getwelcome())
         except Exception, err:
             log_error('log',"FTP Connection or Login Error:%s" % err)
  
     def ftpQuit(self):
         self.ftp.quit()
         log_error('log',"FTP Quit OK")
  
     def  DownLoadFile(self,LocalFile,RemoteFile):
         try:
             #self.ftp.cwd(FTPSERVER_USERDEFAULTPATH)                 #ѡ�����Ŀ¼
             #bufsize=1024                                            #���û�����С
             file_handler =open(LocalFile,'wb').write                 #��д��ģʽ�ڱ��ش��ļ�
             self.ftp.retrbinary('RETR %s' % (RemoteFile),file_handler,1024) #���շ��������ļ���д�뱾���ļ�
             file_handler.close()
             return True
         except Exception, err:
             log_error('log',"FTP DownLoad File Error:%s" % err)
              
     def UpLoadFile(self,LocalFile,RemoteFile):
         try:
             if os.path.isfile(LocalFile) == False:
                 log_error('log','LocalFile not exist:%s'% LocalFile)
                 return False
             file_handler=open(LocalFile,"rb")
             self.ftp.storbinary('STOR %s' % RemoteFile,file_handler,4094)
             file_handler.close()
             return True
         except Exception, err:
             log_error('log',"FTP UpLoad File Error:%s" % err)
  
  
 class QtpApp():
     def __init__(self):
         pass
  
     def Pad(self,strText,intLen):
         if len(strText) >= intLen:
             strText=strText
         else:
             while len(strText) != intLen:
                 strText ="0" + strText
         return strText
  
     def GetTimeStamp(self):
         nowTime =time.localtime()
         strYear =str(nowTime[0])
         strMonth =self.Pad(str(nowTime[1]),2)
         strDay =self.Pad(str(nowTime[2]),2)
         strTime =str(nowTime[3]) + str(nowTime[4]) + str(nowTime[5])
         tempTimeStamp =strYear + strMonth + strDay + "_" + strTime
         return tempTimeStamp
  
  
     def RunQTPtoTest(self):
         blnQtpVisible =True                    #define the qtp is visible or not
         blnQtpDisableSI =True                  #define Qtp Disable Smart Identification
         blnQtpRunMode ='FASE'
         blnQtpViewResults =False
         arrAddins =("Web","SAP")
         errorDescription =''
  
         strQTPResultPath =strFrameWorkPath + "\TestResult\Log\QTPResult_" + self.GetTimeStamp()
  
         QTPApplication =win32com.client.DispatchEx('QuickTest.Application')
  
  
         QTPApplication.Launch()
         QTPApplication.visible =blnQtpVisible
  
         QTPApplication.Open(strFrameWorkPath + '\\TestScript\\MainTest',False)
         QTPApplication.Test.Save()
  
         #set run settings for the test
         qtTest =QTPApplication.Test
         #Create the Run Results Options object
         qtResultsOpt =win32com.client.DispatchEx('QuickTest.RunResultsOptions')
  
         #Set the results locationF
         qtResultsOpt.ResultsLocation =strQTPResultPath
  
         qtTest.Run(qtResultsOpt)
         # time.sleep(3)
  
         # while QtpApp.Test.IsRunning == False:
         #     time.sleep(3)
  
  
 def zip_dir(dirname,zipfilename):
     filelist =[]
     if os.path.isfile(dirname):
         filelist.append(dirname)
     else :
         for root, dirs, files in os.walk(dirname):
             for name in files:
                 filelist.append(os.path.join(root, name))
          
     zf =zipfile.ZipFile(zipfilename, "w", zipfile.zlib.DEFLATED)
     for tar in filelist:
         arcname =tar[len(dirname):]
         zf.write(tar,arcname)
     zf.close()
  
 def zip_To_uploadReport():
     strReportName =read_config(strConfigurationPath,'OUTPUTREPORT','reportname')
     strReportPath =strFrameWorkPath + "\\TestResult\\" + strReportName + "\\"
     zipReportPath =strFrameWorkPath + "\\TestResult\\"
      
     #zip Report
     zip_dir(strReportPath,zipReportPath+ strReportName + '.zip')
  
     #Ftp UpLoad Report
     FTPClient=FTP_Client()
     FTPClient.UpLoadFile(zipReportPath + strReportName + '.zip',strReportName + '.zip')
  
  
 def StringToDate(DateTime):
     y,m,d,H,M,S =DateTime[0:6]                                     #�ѵõ���ʱ��Ԫ��ǰ����Ԫ�ظ�ֵ����������(Ҳ����������ʱ����)
     tmpDateTime =datetime.datetime(y,m,d,H,M,S)                     #���ʹ��datetime�Ѹղŵõ���ʱ�����תΪ��ʽ��ʱ���ʽ����
     return tmpDateTime                                              #���ظ�ʱ���ʽ����
  
 #��鲢��֤��ǰϵͳֻ����һ��Agent����
 def check_exsit(process_name):
     #����Windows��winmgmts����
     WMI =win32com.client.GetObject('winmgmts:')
     #����ִ�еĽ��������Ƿ����
     processCodeCov =WMI.ExecQuery('select * from Win32_Process where Name="%s"' % process_name)
     #����������ƴ���1��
     if len(processCodeCov) > 1:
         #�رյ�ǰ����
         sys.exit(1)
  
  
  
 def wait(WaiTime,ScenarioID,Curtime,SessionID):
     threading._sleep(WaiTime)
     log_info('log',ScenarioID,'Wait Time:%s' % WaiTime)
     write_config(strConfigurationPath,'FRAMEWORKRUN','sessionid',SessionID)
     RunScenario(ScenarioID,Curtime)
  
  
 def RunScenario(ScenarioID,Curtime):
     log_info(log,'LocalTime:'+str(Curtime)+' Start QTP.',strFrameWorkPath)
     qtpApp =QtpApp()
     qtpApp.RunQTPtoTest()
     threading._sleep(10)
     zip_To_uploadReport()
  
  
  
 def main():
     check_exsit('Agent.exe')
     while True:
         SQLDB =SQLExample()
         strSQL ='Call pr_non_execute("'+ localIP +'")'
         SQLDB.ExecuteSQL(strSQL)
         for ExeResult in SQLDB.Result():
             RunSessionID =ExeResult[0]                              #��ȡ���ݿ�SessionID
             RunScenarioID =ExeResult[1]                             #��ȡ���ݿ�ScenarioID            
             Planningtime =ExeResult[2]                              #��ȡ���ݿ�StartTime
             if str(Planningtime) == "None":
                 break
             else:
                 log_info('log',RunScenarioID,'Start Run: %s' % RunScenarioID)
                 Planningtime_Seconds =time.mktime(time.strptime(Planningtime,'%Y%m%d %H:%M:%S'))                #���ƻ�ʱ��ת��Ϊ��
                 Curtime =StringToDate(time.localtime())         #��ȡ��ǰʱ��
                 Curtime_Seconds=time.time()                     #��ȡ��ǰʱ��Ϊ�뵥λ
  
                 if Planningtime_Seconds-Curtime_Seconds>0:                                     #���ƻ�ʱ���뵱ǰʱ����ת�����жԱ�Ϊ��
                     WaitTime =Planningtime_Seconds-Curtime_Seconds
                     log_info('log',RunScenarioID,'Wait Time��%s' % WaitTime)
                     wait(WaitTime,RunScenarioID,Curtime,RunSessionID)
                 else:
                     WaitTime =1
                     log_info('log',RunScenarioID,'Wait Time��%s' % WaitTime)
                     wait(WaitTime,RunScenarioID,Curtime,RunSessionID)                
  
         time.sleep(int(CYCLETIME))
  
 if __name__== '__main__':