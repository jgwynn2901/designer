
'
'FNS DESIGNER/FNS CLAIMCAPTURE   
'Client					:	CCE
'Object					:	Default.asp   
'Script Date: 01/19/2007		Script By: Narayan.Ramachandran
'Work Request/ILog #	:	MROU-2726
'Requirement			: Changing the date format as the new search criteria the in the start date and end date
' The End Date criteria has been inttroduced.
'*/
'/*--------------------------------------------------------------------------------------------------------------------*/



Option Explicit On 
Option Strict Off

Imports System.Data.OracleClient
Imports System.Web.Mail

Public Class WebForm1
    Inherits System.Web.UI.Page
    Protected WithEvents Panel1 As System.Web.UI.WebControls.Panel

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here

        With Response
            .Clear()
            .Buffer = False
            .BufferOutput = False
            .Write("<HTML>" & vbCrLf)
            .Write("<body bgColor=""sandybrown"">" & vbCrLf)
            .Write("<TABLE id=""T1"" style=""Z-INDEX: 102; LEFT: 10px; POSITION: absolute; TOP: 148px"" cellSpacing=""1"" cellPadding=""1"" width=""100%"" align=""center"" border=""0"">" & vbCrLf)
            .Write("<TR>" & vbCrLf)
            .Write("<TD align=""middle"">" & vbCrLf)
            .Write("<div id=""P1"" align=""Center"" style=""color:White;background-color:brown;border-width:6px;border-style:Outset;font-family:Andale Mono;font-size:Large;height:51px;width:610px;"">Running Report" & vbCrLf)
            .Write("</div></TD>" & vbCrLf)
            .Write("</TR>" & vbCrLf)
            .Write("</TABLE>" & vbCrLf)
            .Write("</body>" & vbCrLf)
            .Write("</HTML>" & vbCrLf)
            .Flush()
        End With
        If createSummary() Then
            With Response
                .Write("<HTML>" & vbCrLf)
                .Write("<body bgColor=""sandybrown"">" & vbCrLf)
                .Write("<TABLE id=""T1"" style=""Z-INDEX: 102; LEFT: 10px; POSITION: absolute; TOP: 148px"" cellSpacing=""1"" cellPadding=""1"" width=""100%"" align=""center"" border=""0"">" & vbCrLf)
                .Write("<TR>" & vbCrLf)
                .Write("<TD align=""middle"">" & vbCrLf)
                .Write("<div id=""P1"" align=""Center"" style=""color:White;background-color:brown;border-width:6px;border-style:Outset;font-family:Andale Mono;font-size:Large;height:51px;width:610px;"">Report(s) mailed." & vbCrLf)
                .Write("</div></TD>" & vbCrLf)
                .Write("</TR>" & vbCrLf)
                .Write("</TABLE>" & vbCrLf)
                .Write("</body>" & vbCrLf)
                .Write("</HTML>" & vbCrLf)
                .Flush()
            End With
        Else
            With Response
                .Write("<HTML>" & vbCrLf)
                .Write("<body bgColor=""sandybrown"">" & vbCrLf)
                .Write("<TABLE id=""T1"" style=""Z-INDEX: 102; LEFT: 10px; POSITION: absolute; TOP: 148px"" cellSpacing=""1"" cellPadding=""1"" width=""100%"" align=""center"" border=""0"">" & vbCrLf)
                .Write("<TR>" & vbCrLf)
                .Write("<TD align=""middle"">" & vbCrLf)
                .Write("<div id=""P1"" align=""Center"" style=""color:White;background-color:brown;border-width:6px;border-style:Outset;font-family:Andale Mono;font-size:Large;height:51px;width:610px;"">$0 Invoice." & vbCrLf)
                .Write("</div></TD>" & vbCrLf)
                .Write("</TR>" & vbCrLf)
                .Write("</TABLE>" & vbCrLf)
                .Write("</body>" & vbCrLf)
                .Write("</HTML>" & vbCrLf)
                .Flush()
            End With
        End If
    End Sub

    Function createSummary() As Boolean
        Dim cAHS, cStartDate, cEndDate As String
        Dim oConn As New OracleConnection()
        Dim oCmd As New OracleCommand()

        'cAHS = "23"
		'cAHS = "10757117"		 '   FNS Test
        'cAHS = "13908410"
        'cStartDate = "Nov2003"
        cAHS = Request.QueryString("AHS")
        '---------------------------------------
        ' ILOG issue MROU-2726
        'Modified By R.Narayan
        cStartDate = UCase(Request.QueryString("DATEFROM"))
        cEndDate = UCase(Request.QueryString("DATETo"))
        '
        If cAHS = "" Then
            Response.Write("<script language=vbscript>")
            Response.Write("document.all.p1.innerText = ""Error. Incorrect parameters.""")
            Response.Write("</script>")
            Response.End()
        End If
        'cStartDate = "1-" & Left(cStartDate, 3) & "-" & Right(cStartDate, 4)

        oConn.ConnectionString = "Data Source=fnsp; user id=fnsowner;password=ctown_prod;Persist Security info=False;"
        oConn.Open()
        oCmd.CommandText = "ALTER SESSION SET NLS_DATE_FORMAT = 'DD-MON-YYYY HH:MI:SS'"
        oCmd.Connection = oConn
        oCmd.ExecuteNonQuery()
        oCmd.Dispose()
        If cAHS = "23" Then     '   do ALL
            getAllAgents(cStartDate, cEndDate, oConn)
            createSummary = True
        Else
            createSummary = getOneAgent(cAHS, cStartDate, cEndDate, oConn)
        End If
        oConn.Close()
        oConn.Dispose()
    End Function

    Sub getAllAgents(ByVal cRepStartDate As String, ByVal cRepEndDate As String, _
                        ByRef oConn As OracleConnection)

        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader

        cSQL = "Select * From ACCOUNT_HIERARCHY_STEP Where PARENT_NODE_ID = 23"
        oCmd.Connection = oConn
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            Do
                getOneAgent(CStr(oReader.GetValue(oReader.GetOrdinal("ACCNT_HRCY_STEP_ID"))), cRepStartDate, cRepEndDate, oConn)
            Loop Until Not oReader.Read
        End If
        oReader.Close()
        oReader.Dispose()
        oCmd.Dispose()

    End Sub

    Function getOneAgent(ByVal cAHS As String, _
                    ByVal cRepDate As String, _
                    ByVal cRepEndDate As String, _
                    ByVal oConn As OracleConnection) As Boolean
        Dim cSQL As String
        Dim oCmd As New OracleCommand
        Dim oReader As OracleDataReader
        Dim cServType As String = ""
        Dim oMonthRpt As AgentM
        Dim oYearRpt As AgentY
        Dim cAgentName As String
        Dim cFileName As String
        Dim cPath As String
        Dim oDiskOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
        Dim oMsg As New MailMessage
        Dim cEmail As String
        Dim oAttch As MailAttachment
        Dim dCreatedDate As Date
        Dim lFirstMonth As Boolean
        Dim lNoReport As Boolean

        Randomize()
        cFileName = "AgtBill" & CStr(CInt(Int((99900 * Rnd()) + 1))) & ".pdf"
        If Len(Dir(cFileName)) <> 0 Then
            Kill(cFileName)
        End If
        cPath = Server.MapPath("download") & "/" & cFileName
        oCmd.Connection = oConn
        cSQL = "Select AGENT_BILLING_METHOD, NAME, EMAIL_ADDRESS, CREATED_DT From ACCOUNT_HIERARCHY_STEP Where ACCNT_HRCY_STEP_ID = " & cAHS
        oCmd.CommandText = cSQL
        oReader = oCmd.ExecuteReader()
        oReader.Read()
        If oReader.HasRows Then
            If Not IsDBNull(oReader.GetValue(oReader.GetOrdinal("AGENT_BILLING_METHOD"))) Then
                cServType = UCase(CStr(oReader.GetValue(oReader.GetOrdinal("AGENT_BILLING_METHOD"))))
            End If
            If Not IsDBNull(oReader.GetValue(oReader.GetOrdinal("NAME"))) Then
                cAgentName = CStr(oReader.GetValue(oReader.GetOrdinal("NAME")))
            End If
            If Not IsDBNull(oReader.GetValue(oReader.GetOrdinal("EMAIL_ADDRESS"))) Then
                cEmail = CStr(oReader.GetValue(oReader.GetOrdinal("EMAIL_ADDRESS")))
            End If
            If Not IsDBNull(oReader.GetValue(oReader.GetOrdinal("CREATED_DT"))) Then
                dCreatedDate = CDate(oReader.GetValue(oReader.GetOrdinal("CREATED_DT")))
            End If
        End If
        oReader.Close()
        oReader.Dispose()
        oCmd.Dispose()
        If Len(cServType) <> 0 And Len(cEmail) <> 0 Then
            oDiskOptions.DiskFileName = cPath
            If cServType = "Y" Then
                oYearRpt = New AgentY
                If Month(CDate(cRepDate)) = Month(dCreatedDate) And Year(CDate(cRepDate)) = Year(dCreatedDate) Then
                    lFirstMonth = True
                Else
                    lFirstMonth = False
                End If
                lNoReport = True
                If getY.planY(cAHS, cRepDate, cRepEndDate, oConn, cAgentName, oYearRpt, lFirstMonth) Then
                    With oYearRpt
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Portrait
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                    lNoReport = False
                End If
                oYearRpt = Nothing
            ElseIf cServType = "M" Then
                oMonthRpt = New AgentM
                lNoReport = True
                If getM.planM(cAHS, cRepDate, cRepEndDate, oConn, cAgentName, oMonthRpt) Then
                    With oMonthRpt
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Portrait
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                    lNoReport = False
                End If
                oMonthRpt = Nothing
            End If
            getOneAgent = Not lNoReport
            If Not lNoReport Then
                oMsg.From = "agentPgm@firstnotice.com"
                oMsg.Subject = "RE: Monthly Billing Summary for ClaimCapture Agent program, " & MonthName(Month(CDate(cRepDate))) & " " & Year(CDate(cRepDate))
                oMsg.Body = vbCrLf & "Thank you for participating in the First Notice Systems' ClaimCapture Agent program." & vbCrLf & _
                                    "Attached please find a copy of your monthly billing summary." & vbCrLf & vbCrLf & _
                                    "If paying by credit card, your card will be charged the amount shown in the attached summary report." & vbCrLf & vbCrLf & _
                                    "If paying by check, please write your invoice number and agency name on your check," & vbCrLf & _
                                    "and make all checks payable to: ""First Notice Systems.""" & vbCrLf & vbCrLf & _
                                    "Payments should be mailed to:  First Notice Systems, Inc." & vbCrLf & _
                                    "                               PO Box 972325" & vbCrLf & _
                                    "                               Dallas, TX 75397-2325" & vbCrLf & vbCrLf & _
                                    "Please note, if you are on the annual plan, your first calendar month of service" & vbCrLf & _
                                    "will have no minimums enforced. All following months will have the minimums applied." & vbCrLf & vbCrLf & _
                                    "Should you have issues or questions on this invoice, please send your inquiries to" & vbCrLf & _
                                    "                               AgentBilling@Concentra.com" & vbCrLf & vbCrLf & _
                                    "Sincerely," & vbCrLf & vbCrLf & _
                                    "First Notice Systems, Inc." & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                                    "Note: Please DO NOT reply to this Email, it was sent by an automated process." & vbCrLf & vbCrLf & _
                                    "Email address: " & cEmail & vbCrLf & vbCrLf

                '**************************************************
                cEmail = "Agent_Billing@firstnotice.com"

                'oMsg.Cc = "Agent_Billing@firstnotice.com,Kathleen_Mullery@firstnotice.com,Pat_Lee@firstnotice.com"
                '**************************************************

                oMsg.To = cEmail
                'oMsg.Cc = "Agent_Billing@firstnotice.com"
                oAttch = New MailAttachment(cPath, MailEncoding.Base64)
                oMsg.Attachments.Add(oAttch)
                SmtpMail.SmtpServer = "localhost"
                SmtpMail.Send(oMsg)
            End If
            oMsg = Nothing
            oAttch = Nothing
        End If
    End Function
End Class
