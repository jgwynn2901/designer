Option Explicit On 
Option Strict On
'--------------------------------------------------------------------------------------------------------------------*/
'WORK REQUEST – AMCF-0163 updated [TDS/SOW Document # if Exists :]
'FNS DESIGNER/
'Client			:	EMC
'Script Date: 15/07/2005	By: Narayan Ramachandran
'Work Request/ILog #	:	AMCF-0163 updated
'Work Request/ILog #	:	ERCH-0017 updated (Modified on 05/11/2005)
'Requirement		: 	Creating A Billing Summary Report for EMC client.  
'*/
'---------------------------------------------------------------------------------------------------------------------->

'--------------------------------------------------------------------------------------------------------------------*/
'WORK REQUEST – JPRI-0970 updated [TDS/SOW Document # if Exists :]
'FNS DESIGNER/
'Client			:	KMP
'Script Date: 11/07/2005	By: Subhankar Sarkar & Avra Banerjee
'Requirement		: 	KMP CAT REPORT
'*/
'---------------------------------------------------------------------------------------------------------------------->

'--------------------------------------------------------------------------------------------------------------------*/
'WORK REQUEST – ERCH-0017 updated [TDS/SOW Document # if Exists :]
'FNS DESIGNER/
'Client			:	ALM
'Script Date: 07/11/2005	By: Sutapa Majumdar & Subhankar Sarkar
'Requirement		: 	ALM REPORT
'*/
'---------------------------------------------------------------------------------------------------------------------->
'--------------------------------------------------------------------------------------------------------------------*/
'WORK REQUEST – NBAR-5176 updated [TDS/SOW Document # if Exists :]
'FNS DESIGNER/
'Client			:	NBIC
'Script Date: 08/20/2010	By: Sohail Iqbal
'Requirement		: 	New report for NBIC
'*/
'---------------------------------------------------------------------------------------------------------------------->
'--------------------------------------------------------------------------------------------------------------------*/
'WORK REQUEST – KFAB-6227 updated [TDS/SOW Document # if Exists :]
'FNS DESIGNER/
'Client			:	AFF
'Script Date: 12/07/2010	By: Syed Waqas Ahmed Shah
'Requirement		: 	New report for AFF
'*/
'---------------------------------------------------------------------------------------------------------------------->
'--------------------------------------------------------------------------------------------------------------------*/
'WORK REQUEST – PMAC-1892 updated [TDS/SOW Document # if Exists :]
'FNS DESIGNER/
'Client			:	SEA
'Script Date: 12/14/2011	By: Sohail Iqbal
'Requirement		: 	Need to add Email as valid caller type & update billing report
'*/
'---------------------------------------------------------------------------------------------------------------------->

'--------------------------------------------------------------------------------------------------------------------*/
'WORK REQUEST – TPAL-0146 updated [TDS/SOW Document # if Exists :]
'FNS DESIGNER/
'Client			:	TOW
'Script Date: 02/22/2012	By: Syed Waqas Ahmed Shah
'Requirement		: 	New to introduce a new Summary Report for TOWER ASP (800)
'                       and include only those calls which belongs to IFN (SITE ID = 1 )
'*/
'---------------------------------------------------------------------------------------------------------------------->


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
    Const FMTNo As Integer = 75
    Const FRENo As Integer = 71
    Const MARNo As Integer = 72
    Const GBSNo As Integer = 46
    Const AIGNo As Integer = 62
    Const MCDNo As Integer = 73
    Const KMPNo As Integer = 31
    Const CCENo As Integer = 11
    Const LACNo As Integer = 52
    Const MGCNo As Integer = 103
    Const WIGNo As Integer = 104
    Const FGNo As Integer = 105
    'Chub
    Const CHBNo As Integer = 157
    Const CNLNo As Integer = 150
    Const RTWNo As Integer = 325
    Const CSGNoCall As Integer = 15
    Const CSGNoOnline As Integer = 15
    Const CSGSummary As Integer = 15
    Const WMANo As Integer = 12159325 ' West MANEGEMENT GBS
    Const SHPRNo As Integer = 14504629   'Sunston Hotel Property GBS
    Const PrudNo As Integer = 13776892 'CCE Prudential Financial
    'Const CRWNameASP = "Crawford ASP Claims"
    Const CRWNoASP As Integer = 81
    ' Const CRWNameFNS = "Crawford After Hours"
    Const CRWNoFNS As Integer = 81

    'LPIC-0063 Updated
    'Const CNLNo As Integer = 150
    'HML Implementation
    Const HMLNo As Integer = 387
    Const ALMNo As Integer = 650

    'Added for JPRI-0941
    Const AMCNo As Integer = 222

    ''Added for EMC for Summary Report By R.Narayan
    'Dated:15/07/2005 Work Request No:AMCF-0163 updated
    Const EMCNo As Integer = 67
    Const SELNo As Integer = 277
    'Added for SRS Summary report 
    Const SRSNo As Integer = 600
    'Added for CSAA Summary Report
    Const CSAANo As Integer = 102
    'Added for CSAA Summary Report
    Const AAANo As Integer = 111
    'Added for COV Billing Summary Report
    'CommonWealth Of Virginia
    Const COVNo As Integer = 12801314
    'ESIS KFAB-0042
    Const ESISNo As Integer = 202
    'EVR MROU-3087
    Const EVRNo As Integer = 155
    'PEMCO Mutual Insurance Company 

    Const PEMNo As Integer = 615
    'MROU-3021
    Const TGCNo As Integer = 190
    'TPAL-0146
    Const TGCNoASP As Integer = 800
    'MROU-3549
    Const SAFNo As Integer = 228
    Const AMENo As Integer = 1000
    'NBAR-5176
    Const NBICNo As Integer = 444
    'KFAB-6227
    Const AFFNo As Integer = 550
	'PMAC-1892
    Const SEANo As Integer = 22


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' If session variables have been passed through from Classic ASP, 
        ' assign them to .Net ASP Session variables
        If Request.Form.Count > 0 Then
            For i As Integer = 0 To Request.Form.Count - 1
                Session(Request.Form.GetKey(i)) = Request.Form(i)
            Next
        End If

        ' Process as per normal
        With Response
            .Buffer = False
            .BufferOutput = False
            .Write("<HTML>" & vbCrLf)
            .Write("<body bgColor=""darksalmon"">" & vbCrLf)
            .Write("<TABLE id=""T1"" style=""Z-INDEX: 102; LEFT: 10px; POSITION: absolute; TOP: 148px"" cellSpacing=""1"" cellPadding=""1"" width=""100%"" align=""center"" border=""0"">" & vbCrLf)
            .Write("<TR>" & vbCrLf)
            .Write("<TD align=""middle"">" & vbCrLf)
            .Write("<div id=""P1"" align=""Center"" style=""color:White;background-color:menu;border-width:6px;border-style:Outset;height:51px;width:610px;""><font face=""Verdana"" size=""5"" color=""black"">Running Report</font>" & vbCrLf)
            .Write("</div></TD>" & vbCrLf)
            .Write("</TR>" & vbCrLf)
            .Write("</TABLE>" & vbCrLf)
            .Write("</body>" & vbCrLf)
            .Write("</HTML>" & vbCrLf)
            .Flush()
            createSummary()
        End With

    End Sub

    Sub createSummary()
        Dim cCustName As String
        Dim oRptFxd As BillingSummaryFixed
        Dim oRptFxdRTW As RTWBillingSummaryFixed
        Dim oRptCnl As CnlBillingSummary
        Dim oRptFxdKMP As KMPBillingSummaryFixed
        Dim oRptFxdKMPCAT As KMPBillingSummaryCATFixed  'JPRI-0970	7-NOV-2005
        Dim oRptTiered As BillingSummaryTiered
        'HML Implementation
        Dim oRptFxdHML As HMLBillingSummaryFixed
        'Chub
        Dim oRptFxdCHB As ChubBillingSummary

        ''Added for EMC for Summary Report By R.Narayan
        'Dated:15/07/2005 Work Request No:AMCF-0163 updated
        Dim oRptTieredEMC As EMCBillingSummaryTiered

        'KFAB-6227
        Dim oRptTieredAFF As AFFBillingSummaryTiered

        'Dated:05/11/2005 Work Request No:ERCH-0017 updated
        Dim oRptALM As ALMBillingSummary

        'Added for JPRI-0941
        Dim oRptAMC As AMCBillingSummaryFixed

        Dim oRptSEL As SELBillingSummaryFixed
        Dim oRptSRS As SRSBillingSummaryFixed
        'Added for CSAA Reports
        Dim oRptCSAA As CSAABillingSummaryFixed
		'Added for AAA Reports
        Dim oRptAAA As AAABillingSummaryFixed
        'Added for COV Billing Summary
        Dim oRptCOV As COVBillingSummaryFixed

        'LPIC 0063 Updated
        'Dim oRptCnl As CnlBillingSummary
        'Added for GBS Report 
        Dim pRptGBS As GBSBillingSummaryFixed
        'Dated:06/02/2007 Work Request No:JMAR-0381
        Dim oRptPEM As PEMBillingSummaryFixed
        'MROU-3021
        Dim oRptTGC As TGCBillingSummaryFixed
        'TPAL-0146
        Dim oRptTGCASP As TGCASPBillingSummaryFixed
        'KFAB-0042
        Dim oRptESIS As ESISBillingSummaryFixed
        'MROU-3087
        Dim oRptEVR As EVRBillingSummaryFixed
        'MROU-3549
        Dim oRptSAF As SAFBillingSummaryFixed
        Dim oRptAME As AMEBillingSummary
        Dim oRptNBIC As NBICBillingSummaryFixed
		'PMAC-1892
		Dim oRptSEA as SEABillingSummaryFixed
        '-----------------------------
        Dim oRptTieredCRAW As CRAWBillingSummaryTiered
        Dim oRptPrud As PrudenBillingSummary
        Dim oParamFld As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        Dim oDiskOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions
        Dim cFileName As String = "BillingSummary.pdf"
        Dim cPath As String
        Dim cAHS, cStartDate As String, cEndDate As String
        Dim lTiered As Boolean
        Dim lTieredCRAW As Boolean
        Dim lIsCCE As Boolean
        Dim strUserSite As String
        Dim oPageMargins As CrystalDecisions.Shared.PageMargins
        Dim cCustCode As String
        Dim confFlag As String
        Dim oKMPflf As Boolean
        Dim oKMPCATflf As Boolean  'JPRI-0970	7-NOV-2005
        oKMPflf = False
        oKMPCATflf = False    'JPRI-0970	7-NOV-2005
        cPath = Server.MapPath("download") & "/" & cFileName
        'cAHS = "1000" '**
        'cCustName = "AME"
        'Session("ConnectionString") = "DSN=AMEBA;UID=FNSOWNER;PWD=CTOWN_DESIGNER;Server=AMEBA"
        

        cAHS = Request.QueryString("AHS")
        cCustName = Request.QueryString("CUSTNAME")
        lIsCCE = Len(Request.QueryString("CCE")) <> 0


        If cAHS = "" Then
            Response.Write("<script language=vbscript>")
            Response.Write("document.all.p1.innerHTML = ""<font face=""""Verdana"""" size=""""5"""" color=""""black"""">Error. Incorrect parameters.</font>""")
            Response.Write("</script>")
            Response.End()
        End If

        cStartDate = UCase(Request.QueryString("DATEFROM"))
        'cStartDate = "01-01-2009"
        cEndDate = UCase(Request.QueryString("DATETO"))
        'cEndDate = "05-20-2009"

        If Len(Dir(cFileName)) <> 0 Then
            Kill(cFileName)
        End If

        Select Case CInt(cAHS)
            Case FRENo, FMTNo    '	Fremont
                oRptFxd = New BillingSummaryFixed
                getClaims(oRptFxd, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
            Case CRWNoASP, CRWNoFNS
                oRptTieredCRAW = New CRAWBillingSummaryTiered
                If cCustName = "Crawford ASP Claims" Then
                    strUserSite = "ASP"
                Else
                    strUserSite = "FNS"
                End If
                getClaimsTieredCRAW(oRptTieredCRAW, strUserSite, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = True
            Case FGNo
                oRptFxd = New BillingSummaryFixed
                getClaimsFRG(oRptFxd, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
            Case WIGNo
                oRptTiered = New BillingSummaryTiered
                getClaimsTieredPricing(oRptTiered, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = True

                'Added for EMC for Summary Report By R.Narayan
                'Dated:15/07/2005 Work Request No:AMCF-0163 updated
                '***********************************
            Case SELNo
                oRptSEL = New SELBillingSummaryFixed
                getSELClaimsFixedPricing(oRptSEL, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False

            Case COVNo
                oRptCOV = New COVBillingSummaryFixed
                getClaimsCOV(oRptCOV, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
            Case SRSNo
                'Changes for New Email Type addition in report : 13/09/2006
                oRptSRS = New SRSBillingSummaryFixed
                getSRSClaimsFixedPricing(oRptSRS, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False

            Case EMCNo
                oRptTieredEMC = New EMCBillingSummaryTiered
                getClaimsTieredPricingEMC(oRptTieredEMC, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = True
                '***********************************
                'KFAB-6227
            Case AFFNo
                oRptTieredAFF = New AFFBillingSummaryTiered
                getClaimsTieredPricingAFF(oRptTieredAFF, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = True
            Case CSAANo
                oRptCSAA = New CSAABillingSummaryFixed
                getCSAAClaimsFixedPricing(oRptCSAA, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
				'BCAB-0586
            Case AAANo
                oRptAAA = New AAABillingSummaryFixed
                getAAAClaimsFixedPricing(oRptAAA, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
                'KFAB-0042
            Case ESISNo
                oRptESIS = New ESISBillingSummaryFixed
                getESISClaimsFixedPricing(oRptESIS, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
            Case MGCNo
                oRptTiered = New BillingSummaryTiered
                If InStr(cCustName, "IRC", CompareMethod.Text) = 0 Then
                    getMGCClaimsWithConfirmation(oRptTiered, cAHS, cCustName, cStartDate, cEndDate)
                Else
                    getMGCClaimsNoConfirmation(oRptTiered, cAHS, cCustName, cStartDate, cEndDate)
                End If
                lTiered = True
            Case WMANo '12159325 ' West MANEGEMENT GBS
                oRptFxd = New BillingSummaryFixed
                getClaimsWMA(oRptFxd, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
            Case SHPRNo '14504629   'Sunston Hotel Property GBS
                oRptFxd = New BillingSummaryFixed
                'getClaimsSHPR(oRptFxd, cAHS, cCustName, cStartDate)
                lTiered = False
            Case GBSNo
                pRptGBS = New GBSBillingSummaryFixed
                getClaimsGBS(pRptGBS, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
            Case KMPNo
                If cCustName = "Kemper-CAT" Then
                    'JPRI-0970	7-NOV-2005
                    oRptFxdKMPCAT = New KMPBillingSummaryCATFixed
                    getClaimsKMPCAT(oRptFxdKMPCAT, cAHS, cCustName, cStartDate, cEndDate)
                    lTiered = False
                    oKMPCATflf = True
                    'JPRI-0970	7-NOV-2005
                Else
                    oRptFxdKMP = New KMPBillingSummaryFixed
                    getClaimsFixedPricingKMP(oRptFxdKMP, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                    oKMPflf = True
                    lTiered = False
                End If
            Case CSGNoCall, CSGNoOnline
                oRptFxd = New BillingSummaryFixed
                getCISGPricing(oRptFxd, cAHS, cCustName, cStartDate, cEndDate)
            Case PrudNo
                oRptPrud = New PrudenBillingSummary
                getClaimsFixedPrudent(oRptPrud, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
                'LPIC-0063
            Case CNLNo
                oRptCnl = New CnlBillingSummary
                getCnlBillingSummaryReport(oRptCnl, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
            Case RTWNo
                oRptFxdRTW = New RTWBillingSummaryFixed
                getRTWClaimsFixedPricing(oRptFxdRTW, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
            Case HMLNo
                oRptFxdHML = New HMLBillingSummaryFixed
                getHMLClaimsFixedPricing(oRptFxdHML, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
            Case ALMNo
                oRptALM = New ALMBillingSummary
                getALMBillingSummaryReport(oRptALM, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
                'Added for JPRI-0941
            Case AMCNo
                oRptAMC = New AMCBillingSummaryFixed
                getClaimsFixedPricingAMC(oRptAMC, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
            Case CHBNo
                oRptFxdCHB = New ChubBillingSummary
                getChubClaimsFixedPricing(oRptFxdCHB, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
            Case PEMNo
                'Changes for New Email Type addition in report : 13/09/2006
                oRptPEM = New PEMBillingSummaryFixed
                getClaimsPEM(oRptPEM, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
                'MROU-3021
            Case TGCNo
                'Changes for New Email Type addition in report : 13/09/2006
                oRptTGC = New TGCBillingSummaryFixed
                getClaimsTGC(oRptTGC, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
                'TPAL-0146
            Case TGCNoASP
                oRptTGCASP = New TGCASPBillingSummaryFixed
                getClaimsTGCASP(oRptTGCASP, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
            Case EVRNo
                oRptEVR = New EVRBillingSummaryFixed
                getClaimsEVR(oRptEVR, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
                'MROU-3549
            Case SAFNo
                oRptSAF = New SAFBillingSummaryFixed
                getClaimsSAF(oRptSAF, cAHS, cCustName, cStartDate, cEndDate)
                lTiered = False
            Case AMENo
                oRptAME = New AMEBillingSummary
                getAMEClaimsFixedPricing(oRptAME, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
            Case NBICNo
                oRptNBIC = New NBICBillingSummaryFixed
                getNBICClaimsFixedPricing(oRptNBIC, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
            Case SEANo
                'PMAC-1892
                oRptSEA = New SEABillingSummaryFixed
                getSEAClaimsFixedPricing(oRptSEA, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
            Case Else
                oRptFxd = New BillingSummaryFixed
                getClaimsFixedPricing(oRptFxd, cAHS, cCustName, cStartDate, cEndDate, lIsCCE)
                lTiered = False
        End Select

        'getClaimsFixedPricing(oRpt, cAHS, cCustName, cStartDate)
        'getClaimsTieredPricing(oRpt, cAHS, cCustName, cStartDate)
        'getClaims(oRpt, cAHS, cCustName, cStartDate)
        'CrystalReportViewer1.ReportSource = oRpt
        'CrystalReportViewer1.ShowFirstPage()

        oDiskOptions.DiskFileName = cPath
        If Not lTiered Then
            If CInt(cAHS) = 13776892 Then    'CCE Prudential finc
                With oRptPrud
                    oPageMargins = .PrintOptions.PageMargins
                    oPageMargins.leftMargin = 0
                    oPageMargins.topMargin = 0
                    .PrintOptions.ApplyPageMargins(oPageMargins)
                    .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                    .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                    .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                    .ExportOptions.DestinationOptions = oDiskOptions
                    .Export()
                End With
            Else
                If oKMPflf = True Then
                    With oRptFxdKMP
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                    'JPRI-0970	7-NOV-2005
                ElseIf oKMPCATflf = True Then
                    With oRptFxdKMPCAT
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                    'JPRI-0970	7-NOV-2005
                ElseIf (CInt(cAHS) = 150) Then
                    With oRptCnl
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf (CInt(cAHS) = 102) Then
                    With oRptCSAA
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf (CInt(cAHS) = 111) Then
                    With oRptAAA
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf (CInt(cAHS) = 202) Then
                    With oRptESIS
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf CInt(cAHS) = 12801314 Then
                    With oRptCOV
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf CInt(cAHS) = 277 Then
                    With oRptSEL
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf CInt(cAHS) = 600 Then
                    With oRptSRS
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf CInt(cAHS) = 46 Then
                    With pRptGBS
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf CInt(cAHS) = 325 Then
                    With oRptFxdRTW
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf CInt(cAHS) = 387 Then
                    With oRptFxdHML
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf (CInt(cAHS) = 650) Then
                    With oRptALM
                        'Added for JPRI-0941
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf (CInt(cAHS) = 222) Then
                    With oRptAMC
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf (CInt(cAHS) = 615) Then
                    With oRptPEM
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                    'MROU-3021
                ElseIf (CInt(cAHS) = 190) Then
                    With oRptTGC
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                    'TPAL-0146
                ElseIf (CInt(cAHS) = 800) Then
                    With oRptTGCASP
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf (CInt(cAHS) = 228) Then
                    With oRptSAF
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf (CInt(cAHS) = 155) Then
                    With oRptEVR
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf CInt(cAHS) = 157 Then
                    With oRptFxdCHB
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf (CInt(cAHS) = 1000) Then
                    With oRptAME
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                ElseIf (CInt(cAHS) = 444) Then
                    With oRptNBIC
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                    'PMAC-1892
                ElseIf (CInt(cAHS) = 22) Then
                    With oRptSEA
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With

                Else
                    With oRptFxd
                        oPageMargins = .PrintOptions.PageMargins
                        oPageMargins.leftMargin = 0
                        oPageMargins.topMargin = 0
                        .PrintOptions.ApplyPageMargins(oPageMargins)
                        .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                        .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                        .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                        .ExportOptions.DestinationOptions = oDiskOptions
                        .Export()
                    End With
                End If
            End If
        ElseIf CInt(cAHS) = 81 Then
            With oRptTieredCRAW
                oPageMargins = .PrintOptions.PageMargins
                oPageMargins.leftMargin = 0
                oPageMargins.topMargin = 0
                .PrintOptions.ApplyPageMargins(oPageMargins)
                .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                .ExportOptions.DestinationOptions = oDiskOptions
                .Export()
            End With
            ' AMCF-0163 updated
        ElseIf CInt(cAHS) = 67 Then
            With oRptTieredEMC
                oPageMargins = .PrintOptions.PageMargins
                oPageMargins.leftMargin = 0
                oPageMargins.topMargin = 0
                .PrintOptions.ApplyPageMargins(oPageMargins)
                .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                .ExportOptions.DestinationOptions = oDiskOptions
                .Export()
            End With
            ' KFAB-6227
        ElseIf CInt(cAHS) = 550 Then
            With oRptTieredAFF
                oPageMargins = .PrintOptions.PageMargins
                oPageMargins.leftMargin = 0
                oPageMargins.topMargin = 0
                .PrintOptions.ApplyPageMargins(oPageMargins)
                .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                .ExportOptions.DestinationOptions = oDiskOptions
                .Export()
            End With
        Else
            With oRptTiered
                oPageMargins = .PrintOptions.PageMargins
                oPageMargins.leftMargin = 0
                oPageMargins.topMargin = 0
                .PrintOptions.ApplyPageMargins(oPageMargins)
                .PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Landscape
                .ExportOptions.ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
                .ExportOptions.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
                .ExportOptions.DestinationOptions = oDiskOptions
                .Export()
            End With
        End If
        Response.Write("<meta http-equiv=""refresh"" content=""1;url=download/" & cFileName & """>" & vbCrLf)
        Response.End()
    End Sub
End Class

