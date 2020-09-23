Attribute VB_Name = "modError"
Option Explicit

'
' Handle Errors. Send Service errors to log
'
 Public Sub HandleError(sDescription As String, Optional lNum As Long = 0, Optional SWhere As String = "", Optional lRating As Long = 1, Optional bDontLogError As Boolean = False)
    On Error Resume Next
    
        If bDontLogError Then
            'dont print error
        Else
            Debug.Print Now & " Error occured at:" & SWhere & vbCrLf & sDescription
        End If
        
        If bDontLogError Then
            'dont log
        Else
            LogErrorToFile App.Path & "\" & "errorlog.htm", "<B>" & Now & "</B>" & "&nbsp;&nbsp;<B>description</B>:" & sDescription & "&nbsp;&nbsp;<B>where:</B>" & SWhere & "&nbsp;&nbsp;<B>errnum:</B>:" & lNum & ")", CInt(lRating)
        End If
    
End Sub


'---------------------------------------------------------
' Appending an errormessage to a logfile
'  Importance can be 0,1,2
'  0 = General or unknown error, nothing wrong within the code
'  1 = Database error, Either Database is unavailable or results from it was handled uncorrectly
'  2 = Mission Critical, Internal error or some components missing
'---------------------------------------------------------
Public Sub LogErrorToFile(fileref As String, msg As String, Optional importance As Integer = 0)
On Error GoTo wempErrh
    
    Dim fnr
    On Error Resume Next
    
    'Does the file exist or does it need to be created first?
    If Dir(fileref) = "" Then
     WriteEmptyPage fileref
    End If
    
    fnr = FreeFile()
    Open fileref For Append As #fnr
    
    Select Case importance '0=minor error 1=db error 2=Critical error
     Case 1:
      Print #fnr, "<font color=""#CC9900""><b>&sect; </b></font>" & msg & "<BR>"
     Case 2:
      Print #fnr, "<font color=""#FF0000""><b>&sect; </b></font>" & msg & "<BR>"
     Case Else:
      Print #fnr, "<font color=""#BBBBBB""><b>&sect; </b></font>" & msg & "<BR>"
    End Select
    
    Close #fnr
    
    Exit Sub
wempErrh:
    Debug.Print "Error at LogErrorToFile() : " & Err.Description
    On Error Resume Next
    Close #fnr
End Sub


'---------------------------------------------------------
' Writes an empty html page on the disk
'---------------------------------------------------------
Public Sub WriteEmptyPage(fileref As String)
 On Error GoTo wempErrh
 Dim fnr
 
 On Error Resume Next
 
 fnr = FreeFile()
  
 Open fileref For Output As #fnr
  Print #fnr, "<HTML><HEAD><TITLE>Error Log</TITLE>" & vbCrLf & "<META content=""; text / html; charset=windows-1252; "" http-equiv=Content-Type></HEAD>" & _
               vbCrLf & "<style type=""text/css"">" & _
               "<!--body {  font-family: Arial, Helvetica, sans-serif; font-size: 10pt; font-style: normal}--></style>" & _
               "<BODY><P><H1>Error Log</H1>" & _
               "&nbsp;&nbsp;&nbsp;<font color=""#CCCCCC""><b>&sect;</b></font> = General or unknown error<BR>" & _
               "&nbsp;&nbsp;&nbsp;<font color=""#CC9900""><b>&sect;</b></font> = Database error, Either Database is unavailable or handled uncorrectly<BR>" & _
               "&nbsp;&nbsp;&nbsp;<font color=""#FF0000""><b>&sect;</b></font> = Mission Critical, Internal error or some components missing<BR>" & _
               "</P>"

 Close #fnr

Exit Sub

wempErrh:
 Debug.Print "Error at WriteEmptyPage() : " & Err.Description
 On Error Resume Next
 Close #fnr
 
End Sub


