'''
/******************************************************************************/
/*  Copyright (c) CommVault Systems                                           */
/*  All Rights Reserved                                                       */
/*                                                                            */
/*  THIS IS UNPUBLISHED PROPRIETARY SOURCE CODE OF CommVault Systems          */
/*  The copyright notice above does not evidence any                          */
/*  actual or intended publication of such source code.                       */
/******************************************************************************/
Created on Mar 21, 2012
Revision: $Id: pruning.py,v 1.2.4.3 2013/05/23 16:17:45 sgupta Exp $
$Date: 2013/05/23 16:17:45 $
@author: Sumit Gupta
'''
from Database import Database
from Common import *

from datetime import datetime

import _winreg
import os, sys, shutil
from xml.dom.minidom import parse, parseString
import Logger
import logging
import sys, traceback
    
def parsePruning():		
        Printon = 0
        try:
            ProcessXMLQuery = "EXEC PruneSurveyResults %s" % (Printon)
            print __name__ + " :: Executing query '%s'" % ProcessXMLQuery
            Log.info("Executing query '%s'" % ProcessXMLQuery)
            if _db.Execute(ProcessXMLQuery) == False:
                print __name__ + " :: Query '%s' failed with error: '%s'" % (ProcessXMLQuery, self._db.GetErrorErrStr())
                return False, -1, self._db.GetErrorErrStr()
            else:
               row = _db.m_cursor.fetchone()
               if row:
                   if row.ErrorString and row.ErrorCode == 0:
                       print row.ErrorString                      
                       return True, row.ErrorCode, row.ErrorString
                   else:
                   	   print row.ErrorString
                   	   return False, row.ErrorCode, row.ErrorString
                   
        except:
            print __name__ + " :: Caught Exception: "
            Log.exception("Caught Exception")
            print sys.exc_info()[0]
            Log.exception(sys.exc_info()[0])
            traceback.print_exc()
            return False
    
def PostPrcessFile(iFileName, success=True):
    _bOpenConnection = False
    _sz_DbName = ''
    Log = ''

def DeleteOldArchiveFiles():
    SimpanaCloudRegPath = _db.SimpanaInstallPath + "\\cloud"
    aReg = _winreg.ConnectRegistry(None,_winreg.HKEY_LOCAL_MACHINE)
    aKey = _winreg.OpenKey(aReg, SimpanaCloudRegPath)
    try:
        value, Success = _winreg.QueryValueEx(aKey, r"nXMLPATH")
        Log.info("regKey nXMLPATH found")
        PathToUpload = value
    except:
        Log.info("regKey nXMLPATH not found")

    try:
        value, Success = _winreg.QueryValueEx(aKey, r"nFILEPRUNEDAYS")
        Log.info("regKey nFILEPRUNEDAYS found and deleting files older than %s days"%value)
        prune_days = value
    except WindowsError:
        Log.info("regKey nFILEPRUNEDAYS not found going with default number of days")
        prune_days = 90     
    PathToArchiveFiles = os.path.join(PathToUpload, "Archive")
    listOfFiles = os.listdir(PathToArchiveFiles)
    for ArchiveFile in listOfFiles:
        fullfilename = os.path.join(PathToArchiveFiles, ArchiveFile)
        createTimeEpoch = os.path.getctime(fullfilename)
        createTime = datetime.fromtimestamp(createTimeEpoch)
        timeDifference = datetime.now() - createTime
        daysSinceCreate = timeDifference.days
        if( daysSinceCreate >= prune_days):
            os.remove(fullfilename)
            print "Deleting file %s"%ArchiveFile
    return True

if __name__ == '__main__':
    #initialize logger
    Log = Logger.InitLogger()
    #initialize global database class
    _db = Database()
    Log.info(" Started to prune old archived files ")
    retVal = DeleteOldArchiveFiles()
    if retVal == True:
        Log.info(" Successfully pruned old files ")
    else:
        Log.info(" Files pruning failed")
    #pass
    GetSuccess, processName = isDriverProcessRunning(_db.SimpanaInstallPath, "Base\Cloud\pruning.py")
    if GetSuccess == True:
        if processName == "cvd":
            print "Please check whether cvd is running."
            Log.error("Please check whether cvd is running.")
        else:
            Log.error("Another Instance of Base\Cloud\pruning.py already running.")
        sys.exit()
        
    print __name__ + " :: Start Pruning....."
    Log.info("Start Pruning.....")
    SimpanaCloudDBRegPath = _db.SimpanaInstallPath + "\\Database"
    aReg = _winreg.ConnectRegistry(None,_winreg.HKEY_LOCAL_MACHINE)
    aKey = _winreg.OpenKey(aReg, SimpanaCloudDBRegPath)
    value, type = _winreg.QueryValueEx(aKey, r"sCLOUDDBNAME")
    sz_DbName = value
    value, type = _winreg.QueryValueEx(aKey, r"sCLOUDINSTANCE")
    sz_DbServer = value
    value, type = _winreg.QueryValueEx(aKey, r"sCLOUDCONNECTION")
    sz_DbDSN = value
     
    _bOpenConnection = False
    try:
        if _db.OpenDSN(sz_DbDSN, sz_DbName) == True:
            _bOpenConnection = True
    finally:
        if _bOpenConnection == False:
            print __name__ + " :: Failed to open DB '%s', with error '%s'" %(sz_DbName, _db.GetErrorErrStr())
            Log.error("Failed to open DB '%s', with error %s" % (sz_DbName, _db.GetErrorErrStr()))
                
    retVal = False
    szErrorString = ''
    ErrorCode = 0
    retVal, ErrorCode, szErrorString = parsePruning()
    if retVal == True:
        print __name__ + " :: Successfully completed pruning ['%s']" % (szErrorString)
        Log.info("Successfully completed pruning ['%s']" % (szErrorString))
    else:
       print __name__ + " :: Failed in pruning ErrorCode: ['%d' : '%s']" % (ErrorCode, szErrorString)
       Log.info("Failed in pruning ErrorCode: ['%d' : '%s']" % (ErrorCode, szErrorString))
       
