Attribute VB_Name = "RA_Send_HoursWorkedSum_GH"
Sub GH_Send_Email_HoursWorkedSummary()
    
    '****************************************************************************************************************************************************************
    'February 10, 2021 - Created macro to mail merge through C:\Users\137504\OneDrive - American Airlines, Inc\Documents\Excel\Misc\RosterApps\HoursWorkedSummary_MM.xlsb
    '                Macro will create email and get TO and CC from worksheet "MailMerge", Attach filename, and send to listed recipients, using Signature block
    'February 17, 2021 - Automated Report Week Ending date based on text in A1
    'February 17, 2021 - Added "XXX" to send Master Hours Worked Summary Reports to Brian, Scott, Tiffany, Payroll.
    'April 20, 2021 - Modified MX version for Ground Handling.
    
    '****************************************************************************************************************************************************************
    

    Dim wb As Workbook: Set wb = ActiveWorkbook
    Dim ws As Worksheet: Set ws = wb.Sheets("GH MailMerge")
    Dim x As Long
    Dim strSubj As String
    Dim strBody As String
    Dim strSubjMaster As String
    Dim strBodyMaster As String
    
    Dim myAttachment As String
    Dim fn As String
    Dim RptEndDate As String
       
    TurnFuncOff
    
    RptEndDate = Format(ws.Cells(1, 5), "mmmm d, yyyy")
    
    strSubj = "Hours Worked Summary Report"
           
    strBody = "Attached is the Hours Worked Summary Report for week ending " & RptEndDate & "." & _
                    "<br><br>Thank you<br><br>"
                    
    strSubjMaster = "GH Hours Worked Summary Report Masters"
    
    strBodyMaster = "Attached are the GH Master Hours Worked Summary Reports for week ending " & RptEndDate & "." & _
                    "<br><br>Thank you<br><br>"
              
    
    Dim OutApp As Object
    Dim OutMail As Object
    Dim oAttachment As Object
    Dim LastRow As Long
    Dim files As Variant, file As Variant
    
    SigString = Environ("appdata") & _
                "\Microsoft\Signatures\KevinNoPicture.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
           
    x = 2
    
    Do While ws.Cells(x, 1) <> ""
        
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        Set oAttachment = OutMail.Attachments
        
        fn = ws.Cells(x, 2)
        
        myAttachment = GHfp & fn
        
        With OutMail
            .To = ws.Cells(x, 3).Value
            .CC = ws.Cells(x, 4).Value
            If ws.Cells(x, 1) <> "XXX" Then
                .Subject = strSubj
                .HTMLBody = strBody & "<br>" & Signature
            Else
                .Subject = strSubjMaster
                .HTMLBody = strBodyMaster & "<br>" & Signature
            End If
            files = Split(ws.Cells(x, 2), "%")
            For Each file In files
                file = GHfp & file
                .Attachments.Add file
            Next file
            '.Display
            .Send
        End With
        
        x = x + 1
    
    Loop
    
    TurnFuncOn

    Set OutApp = Nothing
    Set OutMail = Nothing
    
End Sub
Function GetBoiler(ByVal sFile As String) As String

    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.ReadAll
    ts.Close
    
End Function


