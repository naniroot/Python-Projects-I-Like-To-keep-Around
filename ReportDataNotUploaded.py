from Database import Database
from Common import *
from Table import *

import sys
reload(sys)
sys.setdefaultencoding("utf-8")

import smtplib, os
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders

emailFormat = """<html>
<head>
<title>#subject#</title>

<style>
/* Font Definitions */
@font-face
    {font-family:Calibri;
    panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
    {font-family:Tahoma;
    panose-1:2 11 6 4 3 5 4 4 2 4;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
    {margin:0in;
    margin-bottom:.0001pt;
    font-size:12.0pt;
    font-family:'Times New Roman','serif';}
a:link, span.MsoHyperlink
    {mso-style-priority:99;
    color:blue;
    text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
    {mso-style-priority:99;
    color:purple;
    text-decoration:underline;}
p
    {mso-style-priority:99;
    mso-margin-top-alt:auto;
    margin-right:0in;
    mso-margin-bottom-alt:auto;
    margin-left:0in;
    font-size:12.0pt;
    font-family:'Times New Roman','serif';}
p.MsoAcetate, li.MsoAcetate, div.MsoAcetate
    {mso-style-priority:99;
    mso-style-link:'Balloon Text Char';
    margin:0in;
    margin-bottom:.0001pt;
    font-size:8.0pt;
    font-family:'Tahoma','sans-serif';}
p.content1, li.content1, div.content1
    {mso-style-name:content1;
    mso-style-priority:99;
    mso-margin-top-alt:auto;
    margin-right:0in;
    mso-margin-bottom-alt:auto;
    margin-left:0in;
    font-size:12.0pt;
    font-family:'Times New Roman','serif';}
p.content2, li.content2, div.content2
    {mso-style-name:content2;
    mso-style-priority:99;
    mso-margin-top-alt:auto;
    margin-right:0in;
    mso-margin-bottom-alt:auto;
    margin-left:0in;
    background:white;
    font-size:12.0pt;
    font-family:'Calibri','sans-serif';
    color:#666666;}
span.EmailStyle20
    {mso-style-type:personal;
    font-family:'Calibri','sans-serif';
    color:#1F497D;}
span.BalloonTextChar
    {mso-style-name:'Balloon Text Char';
    mso-style-priority:99;
    mso-style-link:'Balloon Text';
    font-family:'Tahoma','sans-serif';}
span.EmailStyle23
    {mso-style-type:personal-reply;
    font-family:'Calibri','sans-serif';
    color:#1F497D;}
.MsoChpDefault
    {mso-style-type:export-only;
    font-size:10.0pt;}
@page WordSection1
    {size:8.5in 11.0in;
    margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection1
    {page:WordSection1;}
</style>

</head>
<body bgcolor=white background-repeat='repeat' lang=EN-US link=blue vlink=purple>
    <div align=center>
        <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=800 style='width:600.0pt;border-collapse:collapse'>
            <tr><td style='padding:15.0pt .75pt .75pt .75pt'>
                <p class=MsoNormal><img border=0 width=800 height=150 src='http://cloudtest:8080/CloudSurvey/images/cloud_mail_header.png' alt='Simpana Cloud'><o:p></o:p>
                </p>
                </td>
                <td style='padding:15.0pt .75pt .75pt .75pt'></td>
            </tr>
        </table>
    </div>

    <div align=center>
    <table class=MsoNormalTable border=0 cellpadding=0 width=800 style='width:600.0pt;background:white'>
        <tr><td style='padding:18.75pt 37.5pt 18.75pt 37.5pt'><p><span style='font-family:Calibri,sans-serif;color:#666666'>
        Hi,<o:p></o:p></span></p>

        <span style='font-family:Calibri,sans-serif;color:#666666'>

        <p>Overview of number of CommCells beaming data</p>

        #OverviewCommCell#

        <p><h4>See attachment for information on all commcells that did not upload any data between 7 and 90 Days.</h4></p>

        <p>If you have any questions, please <a href='mailto:cloudsurvey'><span style='color:#008ACD;text-decoration:none'>contact us</span></a>.<o:p></o:p></p>
        <p>Thank you,<o:p></o:p><br/>Engineering Reports<o:p></o:p></span></p>

        <div class=MsoNormal align=center style='text-align:center'><span style='font-family:Calibri,sans-serif;color:#666666'><hr size=1 width='100%' align=center></span></div>

        </td></tr>
    </table>
    </div>
</body>
</html>"""

attachFormat = """<html>
<head>
<title> List of Commcells Not beaming data</title>

<style>
/* Font Definitions */
@font-face
    {font-family:Calibri;
    panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
    {font-family:Tahoma;
    panose-1:2 11 6 4 3 5 4 4 2 4;}
/* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
    {margin:0in;
    margin-bottom:.0001pt;
    font-size:12.0pt;
    font-family:'Times New Roman','serif';}
a:link, span.MsoHyperlink
    {mso-style-priority:99;
    color:blue;
    text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
    {mso-style-priority:99;
    color:purple;
    text-decoration:underline;}
p
    {mso-style-priority:99;
    mso-margin-top-alt:auto;
    margin-right:0in;
    mso-margin-bottom-alt:auto;
    margin-left:0in;
    font-size:12.0pt;
    font-family:'Times New Roman','serif';}
p.MsoAcetate, li.MsoAcetate, div.MsoAcetate
    {mso-style-priority:99;
    mso-style-link:'Balloon Text Char';
    margin:0in;
    margin-bottom:.0001pt;
    font-size:8.0pt;
    font-family:'Tahoma','sans-serif';}
p.content1, li.content1, div.content1
    {mso-style-name:content1;
    mso-style-priority:99;
    mso-margin-top-alt:auto;
    margin-right:0in;
    mso-margin-bottom-alt:auto;
    margin-left:0in;
    font-size:12.0pt;
    font-family:'Times New Roman','serif';}
p.content2, li.content2, div.content2
    {mso-style-name:content2;
    mso-style-priority:99;
    mso-margin-top-alt:auto;
    margin-right:0in;
    mso-margin-bottom-alt:auto;
    margin-left:0in;
    background:white;
    font-size:12.0pt;
    font-family:'Calibri','sans-serif';
    color:#666666;}
span.EmailStyle20
    {mso-style-type:personal;
    font-family:'Calibri','sans-serif';
    color:#1F497D;}
span.BalloonTextChar
    {mso-style-name:'Balloon Text Char';
    mso-style-priority:99;
    mso-style-link:'Balloon Text';
    font-family:'Tahoma','sans-serif';}
span.EmailStyle23
    {mso-style-type:personal-reply;
    font-family:'Calibri','sans-serif';
    color:#1F497D;}
.MsoChpDefault
    {mso-style-type:export-only;
    font-size:10.0pt;}
@page WordSection1
    {size:8.5in 11.0in;
    margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection1
    {page:WordSection1;}
</style>

</head>
<body bgcolor=white background-repeat='repeat' lang=EN-US link=blue vlink=purple>
    <div align=center>
        <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=800 style='width:600.0pt;border-collapse:collapse'>
            <tr><td style='padding:15.0pt .75pt .75pt .75pt'>
                <p class=MsoNormal><img border=0 width=800 height=150 src='http://cloudtest:8080/CloudSurvey/images/cloud_mail_header.png' alt='Simpana Cloud'><o:p></o:p>
                </p>
                </td>
                <td style='padding:15.0pt .75pt .75pt .75pt'></td>
            </tr>
        </table>
    </div>

    <div align=center>
    <table class=MsoNormalTable border=0 cellpadding=0 width=800 style='width:600.0pt;background:white'>
        <tr><td style='padding:18.75pt 37.5pt 18.75pt 37.5pt'><p><span style='font-family:Calibri,sans-serif;color:#666666'>
        Hi,<o:p></o:p></span></p>

        <span style='font-family:Calibri,sans-serif;color:#666666'>

        <p>Information on all commcells that did not upload any data between 7 and 90 Days</p>

		#IndiviualCommCellInformation#

        <p>If you have any questions, please <a href='mailto:cloudsurvey@commvault.com'><span style='color:#008ACD;text-decoration:none'>contact us</span></a>.<o:p></o:p></p>
        <p>Thank you,<o:p></o:p><br/>Engineering Reports<o:p></o:p></span></p>

        <div class=MsoNormal align=center style='text-align:center'><span style='font-family:Calibri,sans-serif;color:#666666'><hr size=1 width='100%' align=center></span></div>

        </td></tr>
    </table>
    </div>
</body>
</html>"""


def send_mail(server, send_from, send_to, subject, text, files=[]):
    assert type(send_to) == list
    assert type(files) == list

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text, 'html'))

    for f in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(f, "rb").read())
        Encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
        msg.attach(part)

    smtp = smtplib.SMTP(server)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()

class CSInfo:
    '''
    classdocs
    '''
    CCId = 0
    CCUId = 0
    CustomerName = ''
    Version = ''
    LastUploadDate = ''
    NoOfDaysSinceUpload = 0
    HealthCheck = ''

class CSUploadInterval:
    '''
    classdocs
    '''
    _smtpServer = 'smtp.commvault.com'
    _bOpenConnectionSurvey = False
    _dbSurvey = Database()

    #Database holding survey table
    sz_DbName = 'cvcloud'
    sz_DbServer = '172.20.35.50\commvault'
    sz_DbUser = 'devlogin'
    sz_DbPw = 'commvault!12'
    Success = False

    def __init__(self):
        '''
        Constructor
        '''
        CnnStrSurvey = "Driver={SQL Server};Server=%s;Database=%s;" % (self.sz_DbServer, self.sz_DbName)
        self._bOpenConnectionSurvey = False
        try:
            if self._dbSurvey.Open(self.sz_DbUser, self.sz_DbPw, CnnStrSurvey) == True:
                self._bOpenConnectionSurvey = True
        finally:
            if self._bOpenConnectionSurvey == False:
                print __name__ + " :: Failed to open DB '%s', with error %s" % (self.sz_DbName, self._dbSurvey.GetErrorErrStr())

    def GetCSUploadInterval(self):
        print __name__ + "Start"

        sqlQuery = "select COUNT(*) AS 'TotalCSCount' from cf_SurveyResultFunc(2,0,1,NULL,NULL)"

        TotalCSCount = 0
        countTable = Table()
        if self._bOpenConnectionSurvey == True and self._dbSurvey.Execute(sqlQuery, countTable) == True:
            countTable.MoveFirst()
            if countTable.GetErrorErrStr() != "Success":
                print 'Failed to Get TotalCSCount'
            else:
                GetSucces, TotalCSCount = countTable.Get('TotalCSCount')
                if GetSucces == False:
                    print __name__ + "Failed to get Total Commcell Count"


        sqlQuery = "select COUNT(*) AS 'InSevenDayCSCount' from cf_SurveyResultFunc(2,0,1,NULL,NULL)WHERE DATEDIFF(day,Logdate,getutcdate()) <= 7"

        InSevenDayCSCount = 0
        countTable = Table()
        if self._bOpenConnectionSurvey == True and self._dbSurvey.Execute(sqlQuery, countTable) == True:
            countTable.MoveFirst()
            if countTable.GetErrorErrStr() != "Success":
                print 'Failed to Get InSevenDayCSCount'
            else:
                GetSucces, InSevenDayCSCount = countTable.Get('InSevenDayCSCount')
                if GetSucces == False:
                    print __name__ + "Failed to get Commcell count which beamed data in last seven day Count"

        sqlQuery = "select COUNT(*) AS 'WithinMonthCSCount' from cf_SurveyResultFunc(2,0,1,NULL,NULL)WHERE DATEDIFF(day,Logdate,getutcdate()) BETWEEN 8 AND 30"

        WithinMonthCSCount = 0
        countTable = Table()
        if self._bOpenConnectionSurvey == True and self._dbSurvey.Execute(sqlQuery, countTable) == True:
            countTable.MoveFirst()
            if countTable.GetErrorErrStr() != "Success":
                print 'Failed to Get WithinMonthCSCount'
            else:
                GetSucces, WithinMonthCSCount = countTable.Get('WithinMonthCSCount')
                if GetSucces == False:
                    print __name__ + "Failed to get Commcell count which beamed data in within the last month but greater than seven days"

        sqlQuery = "select COUNT(*) AS 'GreaterThanMonthCSCount' from cf_SurveyResultFunc(2,0,1,NULL,NULL)WHERE DATEDIFF(day,Logdate,getutcdate()) BETWEEN 30 AND 90"

        GreaterThanMonthCSCount = 0
        countTable = Table()
        if self._bOpenConnectionSurvey == True and self._dbSurvey.Execute(sqlQuery, countTable) == True:
            countTable.MoveFirst()
            if countTable.GetErrorErrStr() != "Success":
                print 'Failed to Get GreaterThanMonthCSCount'
            else:
                GetSucces, GreaterThanMonthCSCount = countTable.Get('GreaterThanMonthCSCount')
                if GetSucces == False:
                    print __name__ + "Failed to get Commcell count which beamed data before a month"

        sqlQuery = "select COUNT(*) AS 'GreaterThanThreeMonthCSCount' from cf_SurveyResultFunc(2,0,1,NULL,NULL)WHERE DATEDIFF(day,Logdate,getutcdate()) > 90"

        GreaterThanThreeMonthCSCount = 0
        countTable = Table()
        if self._bOpenConnectionSurvey == True and self._dbSurvey.Execute(sqlQuery, countTable) == True:
            countTable.MoveFirst()
            if countTable.GetErrorErrStr() != "Success":
                print 'Failed to Get GreaterThanThreeMonthCSCount'
            else:
                GetSucces, GreaterThanThreeMonthCSCount = countTable.Get('GreaterThanThreeMonthCSCount')
                if GetSucces == False:
                    print __name__ + "Failed to get Commcell count which beamed data before a month"

                strComcellOverview = '<p>Below is the overview of the number of Commcells that Uploaded Data.</p>'
                strComcellOverview += '<p>Total Commcell Count is <strong>%d</h4></strong>'%TotalCSCount
                strComcellOverview += '<table  border="1"><tr><th>Days Since Last Upload</th><th>Count</th></tr>'
                strComcellOverview += '<tr><td>Within 7 days</td><td><font color="green"><strong>%d</strong></font></td></tr>'%InSevenDayCSCount
                strComcellOverview += '<tr><td>Within 30 days but over 7 days</td><td><font color="orange"></strong>%d<strong></font></td></tr>'%WithinMonthCSCount
                strComcellOverview += '<tr><td>Within 90 days but over 30 days</td><td><font color="red"><strong>%d</strong></font></td></tr>'%GreaterThanMonthCSCount
                strComcellOverview += '<tr><td>Over 90 days</td><td><font color="red"><strong>%d</strong></font></td></tr>'%GreaterThanThreeMonthCSCount
                strComcellOverview += '</table>'

        sqlQuery = "SELECT BaseTable.CommServUniqueId, \
					BaseTable.CommCellID, BaseTable.CustomerName, BaseTable.CommServVersion, CONVERT(varchar(50),BaseTable.logdate, 106) AS 'Last Upload Time', \
          DATEDIFF(day,BaseTable.Logdate,getutcdate()) AS 'Days Since No Upload', \
          case when HealthTable.CommServUniqueId IS NULL then  'No' else  'Yes' end as 'Health Check' \
				FROM cf_SurveyResultFunc(2,0,1,NULL,NULL) as BaseTable \
                     LEFT OUTER JOIN ( select * from cf_SurveyResultFunc(16,0,1,NULL,NULL) )HealthTable ON \
               BaseTable.CommServUniqueId = HealthTable.CommServUniqueId \
            WHERE DATEDIFF(day,BaseTable.Logdate,getutcdate()) BETWEEN 8 AND 90 ORDER BY 'Days Since No Upload'"

        strComcellUploadInfo = None
        resultTable = Table()
        if self._bOpenConnectionSurvey == True and self._dbSurvey.Execute(sqlQuery, resultTable) == True:
            count = 1
            while True:
                if count == 1:
                    resultTable.MoveFirst()
                    if resultTable.GetErrorErrStr() != "Success":
                        print 'No Commcell that did not upload data Before 7 days'
                        break
                    else:
                        strComcellUploadInfo = '<p>Below is the list of Commcells that beamed data before a week.</p><table  border="1"><tr><th>CommCell UniqueID</th><th>CommCell ID</th><th>Customer Name</th><th>Version</th><th>Last Upload Time</th><th>Days Since No Upload</th></tr>'
                else:
                    resultTable.MoveNext()
                    if resultTable.GetErrorErrStr() != "Success":
                        print 'Done with all Installed CS'
                        break

                strComcellUploadInfo += '<tr>'
                cs = CSInfo()
                 
                GetSuccess, cs.CCId = resultTable.Get('CommCellID')
                if GetSuccess == False:
                    print __name__ + "Failed to get CS CommCell ID"
                    continue
                else:
                    strComcellUploadInfo += '<td>'
                    strComcellUploadInfo += str(cs.CCId)
                    strComcellUploadInfo += '</td>'

                GetSuccess, cs.CustomerName = resultTable.Get('CustomerName')
                GetSuccess, cs.CCUId = resultTable.Get('CommServUniqueId')
                if GetSuccess == False:
                    print __name__ + "Failed to get CommCell Customer Name"
                    continue
                else:
                    strComcellUploadInfo += '<td><a href="http://clouddriver/webconsole/survey/reports/commcellmonitoring.jsp?commUniId='
                    strComcellUploadInfo += str(cs.CCUId)
                    strComcellUploadInfo += '"><span style="color:#008ACD;text-decoration:none">'
                    strComcellUploadInfo += cs.CustomerName
                    strComcellUploadInfo += '</span></a></td>'

                GetSucces, cs.Version = resultTable.Get('CommServVersion')
                if GetSucces == False:
                    print __name__ + "Failed to get Version"
                    continue
                else:
                    strComcellUploadInfo += '<td>'
                    strComcellUploadInfo += cs.Version
                    strComcellUploadInfo += '</td>'

                GetSucces, cs.LastUploadDate = resultTable.Get('Last Upload Time')
                if GetSucces == False:
                    print __name__ + "Failed to get Last Upload date"
                    continue
                else:
                    strComcellUploadInfo += '<td>'
                    strComcellUploadInfo += cs.LastUploadDate
                    strComcellUploadInfo += '</td>'

                GetSucces, cs.NoOfDaysSinceUpload = resultTable.Get('Days Since No Upload')
                if GetSucces == False:
                    print __name__ + "Failed to get Number of days of upload"
                    continue
                else:
                    strComcellUploadInfo += '<td>'
                    strComcellUploadInfo += str(cs.NoOfDaysSinceUpload)
                    strComcellUploadInfo += '</td>'

                GetSuccess, cs.HealthCheck = resultTable.Get('Health Check')
                if GetSuccess == False:
                    print __name__ + "Failed to get Health Check"
                    continue
                else:
                    strComcellUploadInfo += '<td>'
                    strComcellUploadInfo += str(cs.HealthCheck)
                    strComcellUploadInfo += '</td>'

                strComcellUploadInfo += '</tr>'

                count = count+1
                cs = None

            if count >1:
                strComcellUploadInfo +='</table>'
            else:
                strComcellUploadInfo = 'There are No Commcells that have not beamed data in the last 7 days'

        sendEmailTo = []
        try:
            sendEmailTo = os.path.dirname(os.path.realpath(__file__)) + '\EmailList.txt'
            with open(sendEmailTo) as IPs:
                sendEmailTo = IPs.read().splitlines()
        except IOError:
            sendEmailTo = ['']
            print 'EmailList.txt... Procceeding with email ...'

        text = emailFormat.replace('#OverviewCommCell#', str(strComcellOverview))
        subject = 'Information about the Commcells Last Data Upload'
        text = text.replace('#subject#', subject)

        attachText = attachFormat.replace('#IndiviualCommCellInformation#', str(strComcellUploadInfo))

        with open("listOfComcells.html","w") as fs:
            fs.write(attachText)

        files = ['listOfComcells.html']

        send_mail(self._smtpServer, '', sendEmailTo, subject, text.encode('utf-8'), files)
        print __name__ + 'Email Sent'

        print 'End'
        return "Success"

if __name__ == '__main__':
    upgradedCS = CSUploadInterval()
    upgradedCS.GetCSUploadInterval()
