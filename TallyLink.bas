Attribute VB_Name = "TallyLink"
'==============================================================
' Excel -> TallyPrime voucher poster
' Voucher types: Payment, Receipt, Contra
' Sheets: Entries, Settings, Ledgers
'==============================================================
Option Explicit

Public Const APP_NAME As String = "Postify Bank"
Public Const APP_VERSION As String = "2.3.1"
Private gProgOn As Boolean, gProgTick As Single, gProgStart As Single

Private Const FIRST_ROW As Long = 5
Private Const COL_DATE As Long = 1
Private Const COL_TYPE As Long = 2
Private Const COL_BANK As Long = 3
Private Const COL_PARTY As Long = 4
Private Const COL_OUT As Long = 5
Private Const COL_IN As Long = 6
Private Const COL_NARR As Long = 7
Private Const COL_REF As Long = 8
Private Const COL_TXN As Long = 9
Private Const COL_STATUS As Long = 10
Private Const COL_NOTE As Long = 11
Private Const COL_Q As Long = 12          ' hidden: match quality 3 green / 2 yellow / 1 red
Private Const SHEET_PW As String = "postify"
Private Const LAYOUT_TAG As String = "2.2.0"
Private Const BB_SHEET As String = "Bank Balances"
Private Const FORMS_VER As String = "2"
Private Const SUSP_MARK As String = "POSTIFY-SUSPENSE"
Private Const BB_FIRST As Long = 6
Private Const BB_NOTE As String = "* Statement balance is as per the bank statement provided by you (balance after the last transaction imported), " & _
    "and the Tally balance as per your Tally company on the same date. The statement you provided is the basis for this check."
Private Const HOME_NOTE As String = "* Bank balance check: based on the bank statement provided by you and your Tally data for the same date."

Private Const CLR_OK As Long = 13561798     ' light green
Private Const CLR_ERR As Long = 13551615    ' light red
Private Const CLR_INFO As Long = 10284031   ' light yellow
Private Const CLR_TALLY As Long = 15652797  ' light blue
Private Const FX_SHEET As String = "Change Ledger"
Private Const S_FIRST As Long = 6
Private Const S_DATE As Long = 1
Private Const S_TYPE As Long = 2
Private Const S_NO As Long = 3
Private Const S_OTHER As Long = 4
Private Const S_DR As Long = 5
Private Const S_CR As Long = 6
Private Const S_NARR As Long = 7
Private Const S_CUR As Long = 8
Private Const S_NEW As Long = 9
Private Const S_STATUS As Long = 10
Private Const S_GUID As Long = 11
Private Const CLR_TALLYONLY As Long = 11389944   ' orange: change this one in Tally
Private gLastDayBook As String
Private gNoAnim As Boolean

' Learned memory (fast in-memory copy of the Learned sheet)
Private Const LH_SIZE As Long = 65536
Private gLN As Long, gLLoaded As Boolean, gLDirty As Boolean, gLStamp As Long
Private gLSig() As String, gLLed() As String, gLSeen() As String, gLCnt() As Long, gLDate() As Double
Private gLKind() As Long, gLLegacy() As Boolean, gLTok() As Variant, gLMark() As Long
Private gLHKey() As String, gLHIdx() As Long, gTHKey() As String, gTHList() As String

Private Const WHATS_NEW As String = "Dark mode now covers every sheet, not just the dashboard|" & _
    "WhatsApp: choose App or Browser (uses your WhatsApp Web login)|" & _
    "Clearer, easier-to-read pop-ups for email, WhatsApp and the mail password|" & _
    "Minor bug fixes"

Private gNames() As String, gGroups() As String, gComp() As String, gTok() As Variant
Private gDig() As String, gIni() As String, gDFKey() As String, gDFCnt() As Long
Private Const DF_SIZE As Long = 8192
Private gCount As Long

Private Const NARR_STOP As String = " NEFT RTGS IMPS UPI IFT INB MOB TPFT SELFFT MB P2A P2M ACH CR DR NACH ECS AUTOBPAY NBSM BILLDESK BILL DESK CLG TRF TRANSFER TPARTY PAYMENT PAYMENTS PAYOUT ACCOUNT AC ACC ACCO FOR THE OF AND MONTH ONLINE COMMON LTD LIMITED LIMITE LIMI PVT PRIVATE CO REF NA SAVINGS AUTOPE IMPSP2AP TO FROM BY SENT USI PAY POS ECOM PUR KOTAKMAH HDFCBANK HDFCBANKLTD ICICIBANK KOTAKMAHINDRABANK PUNJABNATIONALBANK SBIN YBL OKHDFCBANK OKICICI OKSBI OKAXIS PAYTM YESB IBL AXL APL UTIB ICIC HDFC0 "
Private Const LEDGER_STOP As String = " LTD LIMITED PVT PRIVATE AC THE AND OF CO MS DEBTOR DEBTORS CREDITOR CREDITORS FROM "
Private Const GENERIC_W As String = " MUTUAL FUND BROKERAGE COMMISSION INCOME EXPENSE PAYABLE RECEIVABLE INSURANCE ASSET MANAGEMENT FINANCIAL SERVICES SECURITIES FINANCE INDIA INTERNATIONAL HEALTH LIFE GENERAL OTHER ACCOUNT BANK CHARGES TRADING TRADERS ENTERPRISES INDUSTRIES SOLUTIONS TECH RECD RECEIPT INTEREST LOAN SHARES SHARE NORTH SOUTH EAST WEST DELHI "
Private Const IN_WORDS As String = " BROKERAGE COMMISSION INCOME RECEIVABLE RECD RECEIVED RECEIPT DIVIDEND REFUND MATURITY REDEMPTION LTCG STCG "
Private Const OUT_WORDS As String = " EXPENSE PAYABLE PAID PREMIUM EMI CHARGES "
Private Const UTR_PFX As String = " HDFCH CITIN SCBLH AXOMB UTIBH ICICN SBINN IDFBH ICMS YESBN KKBKH BARBN PUNBN IMPSP "
Private Const BANK_NAMES As String = "STATE BANK|STANDARD CHARTERED|CITI BANK|ICICI BANK|HDFC BANK|AXIS BANK|KOTAK MAHINDRA BANK|IDFC FIRST|IDBI BANK|YES BANK|RBL BANK|UNION BANK|PUNJAB NATIONAL|BANK OF BARODA|BANK OF INDIA|CANARA BANK|INDUSIND|FEDERAL BANK|NSDL PAYMENTS|PAYTM PAYMENTS|AIRTEL PAYMENTS|INDIAN BANK|INDIAN OVERSEAS|CENTRAL BANK|UCO BANK|DBS BANK|DEUTSCHE|HSBC BANK|BANDHAN BANK|AU SMALL|UNITY SMALL"

Private Function SettingValue(ByVal nm As String) As String
    On Error Resume Next
    SettingValue = Trim$(CStr(ThisWorkbook.Names(nm).RefersToRange.Value))
End Function

Private Function TallyUrl() As String
    Dim u As String
    u = SettingValue("TallyURL")
    If u = "" Then u = "http://localhost:9000"
    If LCase$(Left$(u, 4)) <> "http" Then u = "http://" & u
    TallyUrl = u
End Function

Private Function EntriesSheet() As Worksheet
    Set EntriesSheet = ThisWorkbook.Worksheets("Entries")
End Function

Private Function LastDataRow(ws As Worksheet) As Long
    Dim r As Long, c As Long, m As Long
    m = FIRST_ROW - 1
    For c = COL_DATE To COL_TXN
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > m Then m = r
    Next c
    LastDataRow = m
End Function

'---------------- Helpers ----------------
Private Function XmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    XmlEsc = s
End Function

Private Function Amt(ByVal v As Variant) As Double
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then Amt = CDbl(v)
End Function

Private Function AmtStr(ByVal d As Double) As String
    AmtStr = Replace(Format$(d, "0.00"), ",", ".")
End Function

Private Function GetTag(ByVal txt As String, ByVal tag As String) As String
    Dim p1 As Long, p2 As Long
    p1 = InStr(1, txt, "<" & tag & ">", vbTextCompare)
    If p1 = 0 Then Exit Function
    p1 = p1 + Len(tag) + 2
    p2 = InStr(p1, txt, "</" & tag & ">", vbTextCompare)
    If p2 = 0 Then Exit Function
    GetTag = Mid$(txt, p1, p2 - p1)
End Function

Private Function AllTags(ByVal txt As String, ByVal tag As String) As String
    Dim p1 As Long, p2 As Long, res As String
    p1 = 1
    Do
        p1 = InStr(p1, txt, "<" & tag & ">", vbTextCompare)
        If p1 = 0 Then Exit Do
        p1 = p1 + Len(tag) + 2
        p2 = InStr(p1, txt, "</" & tag & ">", vbTextCompare)
        If p2 = 0 Then Exit Do
        If res <> "" Then res = res & " | "
        res = res & Mid$(txt, p1, p2 - p1)
        p1 = p2
    Loop
    AllTags = res
End Function

Private Function IsDone(ByVal status As String) As Boolean
    Dim s As String
    s = LCase$(Trim$(status))
    IsDone = (Left$(s, 6) = "posted") Or (Left$(s, 8) = "exported") Or (Left$(s, 8) = "in tally")
End Function

Private Sub SetStatus(ws As Worksheet, ByVal r As Long, ByVal msg As String, ByVal clr As Long)
    ws.Cells(r, COL_STATUS).Value = msg
    ws.Cells(r, COL_STATUS).Interior.Color = clr
    ws.Cells(r, COL_STATUS).Font.Color = RGB(30, 41, 59)
    FitRow ws, r
End Sub

'---------------- XML building ----------------
Private Function Envelope(ByVal inner As String) As String
    Dim co As String, sv As String
    co = SettingValue("CompanyName")
    If co <> "" Then sv = "<STATICVARIABLES><SVCURRENTCOMPANY>" & XmlEsc(co) & "</SVCURRENTCOMPANY></STATICVARIABLES>"
    Envelope = "<ENVELOPE><HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>" & _
        "<BODY><IMPORTDATA><REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME>" & sv & "</REQUESTDESC>" & _
        "<REQUESTDATA>" & inner & "</REQUESTDATA></IMPORTDATA></BODY></ENVELOPE>"
End Function

Private Function LedgerEntry(ByVal ledger As String, ByVal isDebit As Boolean, ByVal amount As Double, _
        Optional ByVal bankAlloc As String = "") As String
    Dim s As String
    s = "<ALLLEDGERENTRIES.LIST><LEDGERNAME>" & XmlEsc(ledger) & "</LEDGERNAME>"
    If isDebit Then
        s = s & "<ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE><AMOUNT>-" & AmtStr(amount) & "</AMOUNT>"
    Else
        s = s & "<ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE><AMOUNT>" & AmtStr(amount) & "</AMOUNT>"
    End If
    LedgerEntry = s & bankAlloc & "</ALLLEDGERENTRIES.LIST>"
End Function

' Bank details (Transaction Type / Inst No / Inst Date) for a bank ledger line
Private Function BankAllocXml(ByVal txnType As String, ByVal ds As String, ByVal instNo As String, _
        ByVal favouring As String, ByVal isDebit As Boolean, ByVal amount As Double, Optional ByVal bankDate As String = "") As String
    Dim s As String
    If txnType = "" Then Exit Function
    s = "<BANKALLOCATIONS.LIST><DATE>" & ds & "</DATE><INSTRUMENTDATE>" & ds & "</INSTRUMENTDATE>" & _
        "<TRANSACTIONTYPE>" & XmlEsc(txnType) & "</TRANSACTIONTYPE>" & _
        "<PAYMENTFAVOURING>" & XmlEsc(favouring) & "</PAYMENTFAVOURING>" & _
        "<BANKPARTYNAME>" & XmlEsc(favouring) & "</BANKPARTYNAME>"
    If instNo <> "" Then s = s & "<INSTRUMENTNUMBER>" & XmlEsc(instNo) & "</INSTRUMENTNUMBER>"
    s = s & "<PAYMENTMODE>Transacted</PAYMENTMODE>"
    If bankDate <> "" Then s = s & "<BANKERSDATE>" & bankDate & "</BANKERSDATE>"
    If isDebit Then
        s = s & "<AMOUNT>-" & AmtStr(amount) & "</AMOUNT>"
    Else
        s = s & "<AMOUNT>" & AmtStr(amount) & "</AMOUNT>"
    End If
    BankAllocXml = s & "</BANKALLOCATIONS.LIST>"
End Function

' Group of a ledger from the Ledgers sheet ("" if not loaded / not found)
Private Function LedgerGroup(ByVal ledger As String) As String
    Dim v As Variant
    On Error Resume Next
    v = Application.VLookup(ledger, ThisWorkbook.Worksheets("Ledgers").Range("A:B"), 2, False)
    If Not IsError(v) Then LedgerGroup = CStr(v)
End Function

Private Function IsBankGroup(ByVal grp As String) As Boolean
    Dim g As String
    g = LCase$(grp)
    IsBankGroup = (InStr(g, "bank accounts") > 0) Or (InStr(g, "bank od") > 0) Or (InStr(g, "bank occ") > 0)
End Function

' Bank column: treat as bank unless Ledgers sheet says otherwise (or it is Cash)
Private Function BankColIsBank(ByVal ledger As String) As Boolean
    Dim grp As String
    grp = LedgerGroup(ledger)
    If grp <> "" Then
        BankColIsBank = IsBankGroup(grp)
    Else
        BankColIsBank = (LCase$(Trim$(ledger)) <> "cash")
    End If
End Function

' Returns "" if the row is blank. Sets errMsg if the row is invalid.
Private Function VoucherXml(ws As Worksheet, ByVal r As Long, ByRef errMsg As String) As String
    Dim vDate As Variant, d As Date, vType As String, bank As String, party As String
    Dim outAmt As Double, inAmt As Double, amount As Double, isIn As Boolean
    Dim narr As String, ref As String, ds As String, x As String
    Dim txn As String, bankAlloc As String, partyAlloc As String

    errMsg = ""
    vDate = ws.Cells(r, COL_DATE).Value
    outAmt = Amt(ws.Cells(r, COL_OUT).Value)
    inAmt = Amt(ws.Cells(r, COL_IN).Value)
    party = Trim$(CStr(ws.Cells(r, COL_PARTY).Value))
    If IsEmpty(vDate) And outAmt = 0 And inAmt = 0 And party = "" Then Exit Function

    If IsDate(vDate) Then
        d = CDate(vDate)
    Else
        errMsg = "Invalid or missing date": Exit Function
    End If
    If outAmt < 0 Or inAmt < 0 Then errMsg = "Amounts must be positive": Exit Function
    If outAmt > 0 And inAmt > 0 Then errMsg = "Fill only one of Withdrawal / Deposit": Exit Function
    If outAmt = 0 And inAmt = 0 Then errMsg = "Amount missing (Withdrawal or Deposit)": Exit Function
    If party = "" Then errMsg = "Party / Other Ledger is empty": Exit Function

    bank = Trim$(CStr(ws.Cells(r, COL_BANK).Value))
    If bank = "" Then bank = SettingValue("DefaultBank")
    If bank = "" Then errMsg = "Bank / Cash Ledger empty and no default bank in Settings": Exit Function

    isIn = (inAmt > 0)
    If isIn Then amount = inAmt Else amount = outAmt

    vType = Trim$(CStr(ws.Cells(r, COL_TYPE).Value))
    If vType = "" Then
        If isIn Then vType = "Receipt" Else vType = "Payment"
    End If

    narr = Trim$(CStr(ws.Cells(r, COL_NARR).Value))
    ref = Trim$(CStr(ws.Cells(r, COL_REF).Value))
    If ref <> "" Then
        If narr = "" Then narr = "Ref: " & ref Else narr = narr & " (Ref: " & ref & ")"
    End If
    ds = Format$(d, "yyyymmdd")

    x = "<TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & _
        "<VOUCHER VCHTYPE=""" & XmlEsc(vType) & """ ACTION=""Create"" OBJVIEW=""Accounting Voucher View"">" & _
        "<DATE>" & ds & "</DATE><EFFECTIVEDATE>" & ds & "</EFFECTIVEDATE>" & _
        "<VOUCHERTYPENAME>" & XmlEsc(vType) & "</VOUCHERTYPENAME>" & _
        "<PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>"
    If ref <> "" Then x = x & "<REFERENCE>" & XmlEsc(ref) & "</REFERENCE>"
    If narr <> "" Then x = x & "<NARRATION>" & XmlEsc(narr) & "</NARRATION>"

    ' Transaction type: row value, else default from Settings, else none
    txn = Trim$(CStr(ws.Cells(r, COL_TXN).Value))
    If txn = "" Then txn = SettingValue("DefaultTxnType")
    If txn = "" And AutoReconcile() Then txn = "Others"
    If txn <> "" Then
        If BankColIsBank(bank) Then bankAlloc = BankAllocXml(txn, ds, ref, party, isIn, amount, IIf(AutoReconcile(), ds, ""))
        ' bank-to-bank transfer: other side also gets bank details if it is a bank ledger
        If IsBankGroup(LedgerGroup(party)) Then partyAlloc = BankAllocXml(txn, ds, ref, bank, Not isIn, amount)
    End If

    If isIn Then
        ' Money into bank: Bank Dr, Party Cr
        x = x & LedgerEntry(party, False, amount, partyAlloc) & LedgerEntry(bank, True, amount, bankAlloc)
    Else
        ' Money out of bank: Party Dr, Bank Cr
        x = x & LedgerEntry(party, True, amount, partyAlloc) & LedgerEntry(bank, False, amount, bankAlloc)
    End If
    VoucherXml = x & "</VOUCHER></TALLYMESSAGE>"
End Function

'---------------- HTTP ----------------
Private Function PostXml(ByVal body As String, ByRef resp As String) As Boolean
    Dim http As Object
    On Error GoTo fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 5000, 60000, 120000
    http.Open "POST", TallyUrl(), False
    http.setRequestHeader "Content-Type", "text/xml;charset=utf-8"
    http.send body
    resp = CStr(http.responseText)
    PostXml = True
    Exit Function
fail:
    resp = Err.Description
    PostXml = False
End Function

Private Sub ConnectionHelp(ByVal detail As String)
    ProgressHide
    SetApp "B2", "Not connected"
    MsgBox "Tally isn't answering at " & TallyUrl() & vbCrLf & vbCrLf & _
        "Please check:" & vbCrLf & _
        "1. TallyPrime is open and your company is loaded." & vbCrLf & _
        "2. In TallyPrime: F1 (Help) > Settings > Connectivity > Client/Server configuration:" & vbCrLf & _
        "    TallyPrime acts as = Both,  Port = 9000. Then restart TallyPrime." & vbCrLf & _
        "3. The Tally URL on the Settings sheet is correct." & vbCrLf & vbCrLf & _
        "Then try again.   (Details: " & detail & ")", vbExclamation, APP_NAME & " - Tally not reachable"
End Sub

'================ BUTTON: Post to Tally ================
Public Sub PostToTally()
    Dim ws As Worksheet, r As Long, lastR As Long, x As String, errMsg As String
    Dim resp As String, created As Long, altered As Long, nOk As Long, nErr As Long, nSkip As Long
    Dim lineErr As String, nTally As Long, info As String, why As String

    Set ws = EntriesSheet()
    lastR = LastDataRow(ws)
    If lastR < FIRST_ROW Then MsgBox "No entries found on the Entries sheet.", vbInformation: Exit Sub

    If MsgBox("Post all new rows (" & FIRST_ROW & " to " & lastR & ") to Tally?" & vbCrLf & _
              "Company: " & IIf(SettingValue("CompanyName") = "", "(currently open company)", SettingValue("CompanyName")) & vbCrLf & _
              "Rows already marked Posted / Exported are skipped.", vbYesNo + vbQuestion, "Post to Tally") <> vbYes Then Exit Sub

    ProgressShow "Posting to Tally", "Checking which entries are already in Tally..."
    If LCase$(SettingValue("CheckBeforePost")) <> "no" Then
        nTally = CheckRowsAgainstTally(True, info)
        If nTally < 0 Then
            ProgressHide
            If MsgBox("Could not check which entries are already in Tally:" & vbCrLf & info & vbCrLf & vbCrLf & _
                "Post anyway? (Risk of duplicates if some entries were made manually.)", vbYesNo + vbExclamation, APP_NAME) <> vbYes Then Exit Sub
            nTally = 0
            ProgressShow "Posting to Tally", "Posting..."
        End If
    End If
    LoadLedgerCache
    LearnLoad
    Application.ScreenUpdating = False
    For r = FIRST_ROW To lastR
        ProgressUpdate (r - FIRST_ROW + 1) / (lastR - FIRST_ROW + 1), "Posting row " & (r - FIRST_ROW + 1) & " of " & (lastR - FIRST_ROW + 1) & "   (" & nOk & " posted)"
        If IsDone(CStr(ws.Cells(r, COL_STATUS).Value)) Then
            nSkip = nSkip + 1
        Else
            x = VoucherXml(ws, r, errMsg)
            If errMsg <> "" Then
                SetStatus ws, r, "Error: " & errMsg, CLR_ERR: nErr = nErr + 1
            ElseIf x <> "" Then
                Application.StatusBar = "Posting row " & r & " of " & lastR & "..."
                If Not PostXml(Envelope(x), resp) Then
                    LearnSave
                    Application.ScreenUpdating = True: Application.StatusBar = False
                    ProgressHide
                    ConnectionHelp resp
                    Exit Sub
                End If
                created = Val(GetTag(resp, "CREATED"))
                altered = Val(GetTag(resp, "ALTERED"))
                lineErr = AllTags(resp, "LINEERROR")
                If created + altered > 0 And lineErr = "" Then
                    SetStatus ws, r, "Posted " & Format$(Now, "dd-mmm-yy hh:nn"), CLR_OK: nOk = nOk + 1
                    If HasRowDate(ws, r) Then
                        LearnAdd CellText(ws.Cells(r, COL_NARR).Value), CellText(ws.Cells(r, COL_PARTY).Value), CellText(ws.Cells(r, COL_BANK).Value), CDbl(RowDate(ws, r)), why
                    Else
                        LearnAdd CellText(ws.Cells(r, COL_NARR).Value), CellText(ws.Cells(r, COL_PARTY).Value), CellText(ws.Cells(r, COL_BANK).Value), CDbl(Now), why
                    End If
                Else
                    If lineErr = "" Then lineErr = Left$(Replace(Replace(resp, vbCr, " "), vbLf, " "), 250)
                    SetStatus ws, r, "Tally error: " & lineErr, CLR_ERR: nErr = nErr + 1
                End If
            End If
        End If
    Next r
    LearnSave
    Application.StatusBar = False
    Application.ScreenUpdating = True
    ProgressDone "Posting finished"
    LogAct "Post to Tally", nOk & " posted, " & nErr & " errors, " & nSkip & " skipped", nOk
    Notify "Posting finished" & vbCrLf & "Posted: " & nOk & "    Errors: " & nErr & vbCrLf & _
           "Skipped (posted earlier / already in Tally): " & nSkip & IIf(nTally > 0, " (" & nTally & " found in Tally now)", ""), IIf(nErr > 0, "warn", "ok")
End Sub

'================ BUTTON: Save XML file (for cloud / manual import) ================
Public Sub SaveXmlForImport()
    Dim ws As Worksheet, r As Long, lastR As Long, x As String, errMsg As String
    Dim inner As String, nOk As Long, nErr As Long, fName As Variant, okRows As Collection, it As Variant
    Dim stm As Object

    Set ws = EntriesSheet()
    lastR = LastDataRow(ws)
    Set okRows = New Collection
    For r = FIRST_ROW To lastR
        If Not IsDone(CStr(ws.Cells(r, COL_STATUS).Value)) Then
            x = VoucherXml(ws, r, errMsg)
            If errMsg <> "" Then
                SetStatus ws, r, "Error: " & errMsg, CLR_ERR: nErr = nErr + 1
            ElseIf x <> "" Then
                inner = inner & x: okRows.Add r: nOk = nOk + 1
            End If
        End If
    Next r
    If nOk = 0 Then MsgBox "No valid new rows to export." & IIf(nErr > 0, vbCrLf & nErr & " row(s) have errors - see Status column.", ""), vbExclamation: Exit Sub

    fName = Application.GetSaveAsFilename("Tally_Vouchers_" & Format$(Now, "yyyymmdd_hhnn") & ".xml", "XML Files (*.xml), *.xml", , "Save XML for Tally import")
    If VarType(fName) = vbBoolean Then Exit Sub

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "unicode"
    stm.Open
    stm.WriteText Envelope(inner)
    stm.SaveToFile CStr(fName), 2
    stm.Close

    For Each it In okRows
        SetStatus ws, CLng(it), "Exported to XML " & Format$(Now, "dd-mmm-yy hh:nn"), CLR_INFO
    Next it
    MsgBox nOk & " voucher(s) saved to:" & vbCrLf & fName & vbCrLf & vbCrLf & _
        IIf(nErr > 0, nErr & " row(s) skipped with errors (see Status)." & vbCrLf & vbCrLf, "") & _
        "In TallyPrime: Gateway of Tally > Import > Transactions > choose this XML file.", vbInformation, "XML saved"
End Sub

'================ BUTTON: Test connection ================
Public Sub TestConnection()
    Dim http As Object, txt As String
    On Error GoTo fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 5000, 10000, 10000
    http.Open "GET", TallyUrl(), False
    http.send
    txt = CStr(http.responseText)
    MsgBox "Connected to Tally at " & TallyUrl() & vbCrLf & vbCrLf & "Tally says: " & txt, vbInformation, "Connection OK"
    Exit Sub
fail:
    ConnectionHelp Err.Description
End Sub

'================ BUTTON: Load ledgers from Tally ================
Private Function LedgerRequest(ByVal sv As String, ByVal withPrimary As Boolean) As String
    Dim extra As String
    If withPrimary Then extra = "<COMPUTE>PRIMGRP:$_PrimaryGroup</COMPUTE>"
    LedgerRequest = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>XLLedgerList</ID></HEADER>" & _
          "<BODY><DESC><STATICVARIABLES>" & sv & "</STATICVARIABLES><TDL><TDLMESSAGE>" & _
          "<COLLECTION NAME=""XLLedgerList"" ISMODIFY=""No""><TYPE>Ledger</TYPE><FETCH>NAME,PARENT</FETCH>" & extra & "</COLLECTION>" & _
          "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
End Function

Private Function LoadLedgersCore(ByVal silent As Boolean) As Long
    Dim req As String, resp As String, co As String, sv As String
    Dim re As Object, doc As Object, nodes As Object, n As Object, p As Object
    Dim ws As Worksheet, i As Long

    co = SettingValue("CompanyName")
    sv = "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    If co <> "" Then sv = sv & "<SVCURRENTCOMPANY>" & XmlEsc(co) & "</SVCURRENTCOMPANY>"
    req = LedgerRequest(sv, True)
    If Not PostXml(req, resp) Then ConnectionHelp resp: Exit Function
    If InStr(1, resp, "<LEDGER ", vbTextCompare) = 0 Then
        ' older Tally may not accept the primary-group formula -> ask again without it
        If Not PostXml(LedgerRequest(sv, False), resp) Then ConnectionHelp resp: Exit Function
    End If

    ' Remove control characters Tally sometimes sends (e.g. &#4;)
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "&#([0-8]|1[124-9]|2[0-9]|3[01]);"
    resp = re.Replace(resp, "")

    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    If Not doc.LoadXML(resp) Then
        MsgBox "Tally replied, but the ledger list could not be read." & vbCrLf & Left$(resp, 300), vbExclamation
        Exit Function
    End If
    Set nodes = doc.SelectNodes("//LEDGER")
    If nodes.Length = 0 Then
        MsgBox "No ledgers returned. Is the company '" & co & "' open in Tally? (Leave Company Name blank on Settings to use the currently open company.)", vbExclamation
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets("Ledgers")
    ws.Range("A2:C" & ws.Rows.Count).ClearContents
    i = 2
    For Each n In nodes
        ws.Cells(i, 1).Value = n.getAttribute("NAME")
        Set p = n.SelectSingleNode("PARENT")
        If Not p Is Nothing Then ws.Cells(i, 2).Value = p.Text
        Set p = n.SelectSingleNode("PRIMGRP")
        If Not p Is Nothing Then ws.Cells(i, 3).Value = p.Text
        i = i + 1
    Next n
    ws.Range("A1:C" & i - 1).Sort Key1:=ws.Range("A2"), Order1:=xlAscending, Header:=xlYes
    GridBorders ws.Range("A1:C" & i - 1)
    LoadLedgersCore = i - 2
    If Not silent Then MsgBox (i - 2) & " ledgers loaded from Tally." & vbCrLf & _
        "The Bank and Party columns on the Entries sheet now show them as a dropdown.", vbInformation, "Ledgers loaded"
End Function

Public Sub LoadLedgers()
    Dim n As Long, info As String
    ProgressShow "Loading ledgers from Tally", "Reading the ledger list..."
    n = LoadLedgersCore(True)
    If n = 0 Then ProgressHide: Exit Sub
    ProgressUpdate 0.6, n & " ledgers loaded - refreshing the Change Ledger sheet..."
    info = LoadSuspenseCore(True)
    ProgressDone "Ledgers loaded"
    RefreshTallyStatus
    LogAct "Load ledgers", n & " ledgers", n
    Notify n & " ledgers loaded from Tally." & vbCrLf & info, "ok"
End Sub

'================ BUTTON: Clear status of error rows ================
Public Sub ClearErrorStatus()
    Dim ws As Worksheet, r As Long, s As String
    Set ws = EntriesSheet()
    For r = FIRST_ROW To LastDataRow(ws)
        s = LCase$(CStr(ws.Cells(r, COL_STATUS).Value))
        If Left$(s, 5) = "error" Or Left$(s, 11) = "tally error" Then
            ws.Cells(r, COL_STATUS).ClearContents
            ws.Cells(r, COL_STATUS).Interior.ColorIndex = xlNone
        End If
    Next r
End Sub

'==============================================================
'  BANK STATEMENT IMPORT  +  AUTO LEDGER MATCHING
'==============================================================
Private Function CellText(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    CellText = Trim$(Replace(Replace(Replace(CStr(v), vbCr, " "), vbLf, " "), Chr(160), " "))
End Function

' Upper-case, non letters/digits -> single spaces
Private Function SpacedUpper(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, c As Long
    s = UCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1): c = Asc(ch)
        If (c >= 65 And c <= 90) Or (c >= 48 And c <= 57) Then out = out & ch Else out = out & " "
    Next i
    Do While InStr(out, "  ") > 0: out = Replace(out, "  ", " "): Loop
    SpacedUpper = Trim$(out)
End Function

Private Function HasDigit(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then HasDigit = True: Exit Function
    Next i
End Function

' Meaningful tokens of a narration (isLedger=False) or a ledger name (isLedger=True)
Private Function Levenshtein(ByVal a As String, ByVal b As String) As Long
    Dim la As Long, lb As Long, i As Long, j As Long, cost As Long
    Dim d() As Long
    la = Len(a): lb = Len(b)
    If la = 0 Then Levenshtein = lb: Exit Function
    If lb = 0 Then Levenshtein = la: Exit Function
    ReDim d(0 To la, 0 To lb)
    For i = 0 To la: d(i, 0) = i: Next i
    For j = 0 To lb: d(0, j) = j: Next j
    For i = 1 To la
        For j = 1 To lb
            If Mid$(a, i, 1) = Mid$(b, j, 1) Then cost = 0 Else cost = 1
            d(i, j) = d(i - 1, j) + 1
            If d(i, j - 1) + 1 < d(i, j) Then d(i, j) = d(i, j - 1) + 1
            If d(i - 1, j - 1) + cost < d(i, j) Then d(i, j) = d(i - 1, j - 1) + cost
        Next j
    Next i
    Levenshtein = d(la, lb)
End Function

Private Function Similarity(ByVal a As String, ByVal b As String) As Double
    Dim m As Long
    m = Len(a): If Len(b) > m Then m = Len(b)
    If m = 0 Then Exit Function
    Similarity = 1 - Levenshtein(a, b) / m
End Function

Private Function IsSubseq(ByVal a As String, ByVal b As String) As Boolean
    Dim i As Long, p As Long
    p = 1
    For i = 1 To Len(a)
        p = InStr(p, b, Mid$(a, i, 1))
        If p = 0 Then Exit Function
        p = p + 1
    Next i
    IsSubseq = True
End Function

Private Function SortedSig(ByVal tk As Variant) As String
    Dim i As Long, j As Long, tmp As String, a() As String
    If UBound(tk) < 0 Then Exit Function
    ReDim a(0 To UBound(tk))
    For i = 0 To UBound(tk): a(i) = tk(i): Next i
    For i = 0 To UBound(a) - 1
        For j = i + 1 To UBound(a)
            If a(j) < a(i) Then tmp = a(i): a(i) = a(j): a(j) = tmp
        Next j
    Next i
    SortedSig = Join(a, " ")
End Function

Private Function Jaccard(ByVal a As Variant, ByVal b As Variant) As Double
    Dim x As Variant, y As Variant, common As Long, longHit As Boolean, uni As Long
    If UBound(a) < 0 Or UBound(b) < 0 Then Exit Function
    For Each x In a
        For Each y In b
            If x = y Then
                common = common + 1
                If Len(x) >= 4 Then longHit = True
            End If
        Next y
    Next x
    uni = (UBound(a) + 1) + (UBound(b) + 1) - common
    If longHit And uni > 0 Then Jaccard = common / uni
End Function

'---------------- Tokenising ----------------
Private Function InList(ByVal lst As String, ByVal t As String) As Boolean
    InList = (InStr(lst, " " & t & " ") > 0)
End Function

Private Function IsAlphaStr(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        c = Asc(Mid$(s, i, 1))
        If c < 65 Or c > 90 Then Exit Function
    Next i
    IsAlphaStr = True
End Function

Private Function IsDigitStr(ByVal s As String) As Boolean
    Dim i As Long
    If s = "" Then Exit Function
    For i = 1 To Len(s)
        If Not (Mid$(s, i, 1) Like "#") Then Exit Function
    Next i
    IsDigitStr = True
End Function

Private Sub AddTok(ByRef out As String, ByVal t As String)
    If t = "" Then Exit Sub
    out = out & t & " "
End Sub

' Split into tokens; letters+digits codes keep only long letter runs; join single letters (K M M F -> KMMF)
Private Function BasicTokens(ByVal s As String, ByVal isNarr As Boolean) As String
    Dim i As Long, ch As String, c As Long, raw As String, out As String, parts As Variant, t As Variant
    Dim run As String, j As Long, buf As String, res As String
    s = UCase$(Replace(s, "&", " AND "))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1): c = Asc(ch)
        If (c >= 65 And c <= 90) Or (c >= 48 And c <= 57) Then raw = raw & ch Else raw = raw & " "
    Next i
    parts = Split(Trim$(raw), " ")
    For Each t In parts
        If t <> "" Then
            If HasDigit(CStr(t)) And Not IsDigitStr(CStr(t)) Then
                run = ""
                For j = 1 To Len(t) + 1
                    If j <= Len(t) Then ch = Mid$(t, j, 1) Else ch = "0"
                    If ch Like "[A-Z]" Then
                        run = run & ch
                    Else
                        If Len(run) >= 5 And Not InList(UTR_PFX, run) Then AddTok out, run
                        run = ""
                    End If
                Next j
            Else
                AddTok out, CStr(t)
            End If
        End If
    Next t
    If Not isNarr Then BasicTokens = Trim$(out): Exit Function
    For Each t In Split(Trim$(out), " ")
        If Len(t) = 1 And IsAlphaStr(CStr(t)) Then
            buf = buf & t
        Else
            If Len(buf) >= 3 Then res = res & buf & " "
            buf = ""
            If t <> "" Then res = res & t & " "
        End If
    Next t
    If Len(buf) >= 3 Then res = res & buf & " "
    BasicTokens = Trim$(res)
End Function

Private Function AbbrOf(ByVal t As String) As String
    Select Case t
        Case "MF": AbbrOf = "MUTUAL FUND"
        Case "BKG", "BRK", "BROK", "BROKER", "BRKG": AbbrOf = "BROKERAGE"
        Case "INS": AbbrOf = "INSURANCE"
        Case "COMMISIO", "COMMISION", "COMM": AbbrOf = "COMMISSION"
        Case "PRU": AbbrOf = "PRUDENTIAL"
        Case "EXP", "EXPENSES", "EXPS": AbbrOf = "EXPENSE"
        Case Else: AbbrOf = t
    End Select
End Function

' Expand abbreviations and glued MF (SBIMF -> SBI MUTUAL FUND, AXISMFBRK -> AXIS MUTUAL FUND BROKERAGE)
Private Function ExpandTokens(ByVal toks As String) As String
    Dim t As Variant, out As String, p As Long
    For Each t In Split(toks, " ")
        If t <> "" Then
            p = 0
            If Len(t) >= 5 And IsAlphaStr(CStr(t)) Then p = InStr(4, t, "MF")
            If AbbrOf(CStr(t)) <> t Then
                out = out & AbbrOf(CStr(t)) & " "
            ElseIf p > 0 Then
                out = out & Left$(t, p - 1) & " MUTUAL FUND "
                If p + 2 <= Len(t) Then out = out & AbbrOf(Mid$(t, p + 2)) & " "
            Else
                out = out & t & " "
            End If
        End If
    Next t
    ExpandTokens = Trim$(out)
End Function

Private Function UniqueFilter(ByVal toks As String, ByVal stopL As String) As Variant
    Dim t As Variant, out As String
    For Each t In Split(toks, " ")
        If Len(t) >= 2 And Not IsDigitStr(CStr(t)) And Not InList(stopL, CStr(t)) And Replace(CStr(t), "X", "") <> "" Then
            If InStr(" " & out, " " & t & " ") = 0 Then out = out & t & " "
        End If
    Next t
    out = Trim$(out)
    If out = "" Then UniqueFilter = Array() Else UniqueFilter = Split(out, " ")
End Function

Private Function LedgerTokens(ByVal ledName As String) As Variant
    LedgerTokens = UniqueFilter(ExpandTokens(BasicTokens(ledName, False)), LEDGER_STOP)
End Function

Private Function OwnNames() As Variant
    Dim s As String
    s = SettingValue("CompanyName")
    If SettingValue("OwnNames") <> "" Then s = s & "|" & SettingValue("OwnNames")
    OwnNames = Split(UCase$(s), "|")
End Function

Private Function IsBankSegment(ByVal sg As String) As Boolean
    Dim b As Variant
    sg = Trim$(sg)
    If InStr(" " & SpacedUpper(sg) & " ", " BANK ") > 0 Then IsBankSegment = True: Exit Function
    For Each b In Split(BANK_NAMES, "|")
        If b <> "" Then
            If Left$(sg, Len(b)) = b Then IsBankSegment = True: Exit Function
        End If
    Next b
    If Len(sg) = 11 Then
        If IsAlphaStr(Left$(sg, 4)) And Mid$(sg, 5, 1) = "0" Then IsBankSegment = True
    End If
End Function

' Meaningful tokens of a bank narration
Private Function NarrTokens(ByVal narr As String) As Variant
    Dim u As String, o As Variant, segs As Variant, i As Long, lastSeg As Long, keep As String, sg As String
    u = UCase$(narr)
    For Each o In OwnNames()
        If Trim$(o) <> "" Then
            u = Replace(u, Trim$(o), " ")
            u = Replace(u, Replace(Trim$(o), " ", ""), " ")
        End If
    Next o
    If Left$(u, 7) = "ACH-CR-" Or Left$(u, 7) = "ACH-DR-" Then u = Mid$(u, 8)
    segs = Split(u, "/")
    lastSeg = UBound(segs)
    If lastSeg >= 0 Then
        If (Trim$(segs(0)) = "IMPS" Or Trim$(segs(0)) = "UPI") And lastSeg > 3 Then lastSeg = 3
    End If
    For i = 0 To lastSeg
        sg = Trim$(segs(i))
        If sg <> "" Then
            If Not IsBankSegment(sg) Then keep = keep & " " & sg
        End If
    Next i
    NarrTokens = UniqueFilter(ExpandTokens(BasicTokens(keep, True)), NARR_STOP)
End Function

' Digit runs of the narration (spaces removed, so wrapped numbers join up)
Private Function NarrDigits(ByVal narr As String) As String
    Dim s As String, i As Long, ch As String, out As String
    s = Replace(narr, " ", "")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            out = out & ch
        ElseIf Right$(out, 1) <> " " Then
            out = out & " "
        End If
    Next i
    NarrDigits = " " & Trim$(out) & " "
End Function

Private Function LedgerDigits(ByVal ledName As String) As String
    Dim s As String, i As Long, ch As String, run As String, out As String
    s = ledName & " "
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Then
            run = run & ch
        Else
            If Len(run) >= 4 Then out = out & run & " "
            run = ""
        End If
    Next i
    LedgerDigits = Trim$(out)
End Function

Private Function LedgerInitials(ByVal ledName As String) As String
    Dim w As Variant, s As String, t As Variant, first As Boolean
    w = Split(ExpandTokens(BasicTokens(ledName, False)), " ")
    first = True
    For Each t In w
        If t <> "" And Not IsDigitStr(CStr(t)) And Not InList(LEDGER_STOP, CStr(t)) Then
            If first And InList(GENERIC_W, CStr(t)) Then Exit Function
            first = False
            s = s & Left$(t, 1)
        End If
    Next t
    If Len(s) >= 3 Then LedgerInitials = s
End Function

Private Function Squeeze(ByVal t As String) As String
    Dim i As Long, out As String, ch As String
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch <> Right$(out, 1) Then out = out & ch
    Next i
    Squeeze = out
End Function

'---------------- Ledger cache ----------------
Private Sub LoadLedgerCache()
    Dim ws As Worksheet, lastR As Long, i As Long, v As Variant, t As Variant, k As String, n As Long
    Dim comp As String, bank As String
    Set ws = ThisWorkbook.Worksheets("Ledgers")
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    gCount = 0
    ReDim gDFKey(0 To DF_SIZE - 1): ReDim gDFCnt(0 To DF_SIZE - 1)
    If lastR < 2 Then Exit Sub
    v = ws.Range("A2:C" & lastR).Value
    gCount = UBound(v, 1)
    ReDim gNames(1 To gCount): ReDim gGroups(1 To gCount): ReDim gComp(1 To gCount): ReDim gTok(1 To gCount)
    ReDim gDig(1 To gCount): ReDim gIni(1 To gCount)
    For i = 1 To gCount
        gNames(i) = CellText(v(i, 1))
        gGroups(i) = CellText(v(i, 3))
        If gGroups(i) = "" Then gGroups(i) = CellText(v(i, 2))
        gTok(i) = LedgerTokens(gNames(i))
        gDig(i) = LedgerDigits(gNames(i))
        gIni(i) = LedgerInitials(gNames(i))
        comp = ""
        If UBound(gTok(i)) >= 0 Then
            For Each t In gTok(i)
                If Not InList(GENERIC_W, CStr(t)) Then comp = comp & t
                DFAdd CStr(t)
            Next t
        End If
        gComp(i) = comp
    Next i
End Sub

Private Function DFSlot(ByVal k As String) As Long
    Dim h As Long, i As Long
    For i = 1 To Len(k)
        h = (h * 31 + Asc(Mid$(k, i, 1))) Mod DF_SIZE
    Next i
    Do While gDFKey(h) <> "" And gDFKey(h) <> k
        h = (h + 1) Mod DF_SIZE
    Loop
    DFSlot = h
End Function

Private Function DFGet(ByVal k As String) As Long
    DFGet = gDFCnt(DFSlot(k))
End Function

Private Sub DFAdd(ByVal k As String)
    Dim h As Long
    h = DFSlot(k)
    gDFKey(h) = k
    gDFCnt(h) = gDFCnt(h) + 1
End Sub

Private Function TokW(ByVal t As String) As Double
    Dim n As Long
    If InList(GENERIC_W, t) Then TokW = 0.25: Exit Function
    n = DFGet(t)
    If n < 1 Then n = 1
    TokW = Log(1 + gCount / n)
End Function

Private Function GroupOf(ByVal ledger As String) As String
    Dim i As Long
    For i = 1 To gCount
        If StrComp(gNames(i), ledger, vbTextCompare) = 0 Then GroupOf = gGroups(i): Exit Function
    Next i
End Function

Private Function IsCashOrBank(ByVal ledger As String) As Boolean
    Dim g As String
    g = GroupOf(ledger)
    If g = "" Then
        IsCashOrBank = (LCase$(ledger) = "cash")
    Else
        IsCashOrBank = IsBankGroup(g) Or (InStr(LCase$(g), "cash-in-hand") > 0) Or (InStr(LCase$(g), "cash in hand") > 0)
    End If
End Function

' 1 = party (debtor for money in, creditor for money out), 2 = receivable/payable, 3 = everything else
Private Function TierOf(ByVal i As Long, ByVal isIn As Boolean) As Long
    Dim g As String, n As String
    g = LCase$(gGroups(i)): n = UCase$(gNames(i))
    TierOf = 3
    If isIn Then
        If InStr(g, "sundry debtors") > 0 Then
            TierOf = 1
        ElseIf InStr(g, "current assets") > 0 Or InStr(n, "RECEIVABLE") > 0 Then
            TierOf = 2
        End If
    Else
        If InStr(g, "sundry creditors") > 0 Then
            TierOf = 1
        ElseIf InStr(g, "current liabilities") > 0 Or InStr(n, "PAYABLE") > 0 Then
            TierOf = 2
        End If
    End If
End Function

Private Function IsExcluded(ByVal i As Long, ByVal bank As String) As Boolean
    Dim o As Variant
    If StrComp(gNames(i), bank, vbTextCompare) = 0 Then IsExcluded = True: Exit Function
    If LCase$(gGroups(i)) = "primary" Then IsExcluded = True: Exit Function
    For Each o In OwnNames()
        If Trim$(o) <> "" Then
            If StrComp(Trim$(gNames(i)), Trim$(o), vbTextCompare) = 0 Then IsExcluded = True: Exit Function
        End If
    Next o
End Function

' "Interest Paid|Interest on Od" -> first ledger that exists (exact names in order, then by tier among partial names)
Private Function ResolveLedger(ByVal cands As String, ByVal excludeName As String, ByVal isIn As Boolean) As String
    Dim c As Variant, i As Long, cs As String, bestT As Long, bestI As Long, tr As Long
    For Each c In Split(cands, "|")
        cs = Trim$(CStr(c))
        If cs <> "" Then
            For i = 1 To gCount
                If StrComp(gNames(i), cs, vbTextCompare) = 0 And StrComp(gNames(i), excludeName, vbTextCompare) <> 0 Then ResolveLedger = gNames(i): Exit Function
            Next i
        End If
    Next c
    For Each c In Split(cands, "|")
        cs = Trim$(CStr(c))
        If Len(cs) >= 3 Then
            bestT = 99: bestI = 0
            For i = 1 To gCount
                If StrComp(Left$(gNames(i), Len(cs)), cs, vbTextCompare) = 0 And StrComp(gNames(i), excludeName, vbTextCompare) <> 0 Then
                    tr = TierOf(i, isIn)
                    If tr < bestT Then bestT = tr: bestI = i
                End If
            Next i
            If bestI > 0 Then ResolveLedger = gNames(bestI): Exit Function
        End If
    Next c
End Function

Private Function TMatch(ByVal t As String, ByVal nt As Variant, ByVal nc As String) As Double
    Dim n As Variant, best As Double, r As Double, a As String, b As String
    If UBound(nt) >= 0 Then
        For Each n In nt
            If n = t Then TMatch = 1: Exit Function
            If Len(t) >= 4 And Len(n) >= 4 Then
                If Left$(n, Len(t)) = t Or Left$(t, Len(n)) = n Then
                    If best < 0.9 Then best = 0.9
                ElseIf Squeeze(CStr(n)) = Squeeze(t) Then
                    If best < 0.95 Then best = 0.95
                ElseIf Len(n) < Len(t) And 2 * Len(n) >= Len(t) And Left$(t, 1) = Left$(n, 1) And IsSubseq(CStr(n), t) Then
                    If best < 0.7 Then best = 0.7
                ElseIf Len(t) >= 5 And Len(n) >= 5 And Left$(t, 1) = Left$(n, 1) Then
                    a = Squeeze(t): b = Squeeze(CStr(n))
                    r = Similarity(a, b)
                    If r >= 0.8 Then
                        If best < r * 0.85 Then best = r * 0.85
                    End If
                End If
            End If
        Next n
    End If
    If best = 0 And Len(t) >= 5 Then
        If InStr(nc, t) > 0 Then best = 0.85
    End If
    TMatch = best
End Function

' Score 0-100 for ledger i. hits = "|TOKEN|TOKEN|" of distinctive tokens that matched
Private Function ScoreLedger(ByVal i As Long, ByVal nt As Variant, ByVal nc As String, ByVal digs As String, _
        ByVal isIn As Boolean, ByRef hits As String, ByRef why As String, ByVal npre As String) As Double
    Dim lt As Variant, t As Variant, d As Variant, n As Variant, tot As Double, got As Double, w As Double, m As Double
    Dim dist As Boolean, s As Double, words As String, x As Variant, iniHit As Boolean, quick As Boolean
    hits = "|": why = ""
    lt = gTok(i)
    ' account / card number in the ledger name
    If gDig(i) <> "" Then
        For Each d In Split(gDig(i), " ")
            If InStr(digs, " " & d & " ") > 0 Or (Len(d) >= 6 And InStr(digs, d & " ") > 0) Then
                hits = "|#" & d & "|": why = "account no": ScoreLedger = 100: Exit Function
            End If
        Next d
    End If
    If UBound(lt) < 0 Then Exit Function
    ' quick skip: no ledger word shares its first 3 letters with the narration
    If UBound(nt) < 0 Then Exit Function
    quick = False
    For Each t In lt
        If Len(t) >= 3 Then
            If InStr(npre, " " & Left$(t, 3)) > 0 Or InStr(nc, Left$(t, 4)) > 0 Then quick = True: Exit For
        End If
    Next t
    If Not quick And gIni(i) <> "" Then quick = (InStr(" " & Join(nt, " ") & " ", " " & gIni(i) & " ") > 0)
    If Not quick Then Exit Function
    iniHit = False
    If gIni(i) <> "" And UBound(nt) >= 0 Then
        For Each n In nt
            If n = gIni(i) Then iniHit = True
        Next n
    End If
    For Each t In lt
        w = TokW(CStr(t)): tot = tot + w
        m = TMatch(CStr(t), nt, nc)
        got = got + w * m
        If m > 0 And Len(t) >= 3 And Not InList(GENERIC_W, CStr(t)) Then dist = True: hits = hits & t & "|"
    Next t
    If Len(gComp(i)) >= 6 And InStr(nc, gComp(i)) > 0 Then
        s = 100: why = "full name"
    Else
        s = 0
        If UBound(nt) >= 0 And UBound(lt) >= 1 Then
            For Each n In nt
                If Len(n) >= 6 And Len(n) >= Len(lt(0)) + 2 And Left$(gComp(i), Len(n)) = n Then s = 95: why = "short name"
            Next n
        End If
        If s = 0 And tot > 0 And dist Then s = 100 * got / tot: why = "name"
    End If
    If iniHit And s < 88 Then
        ' initials, e.g. KMMF = Kotak Mahindra Mutual Fund, TMF = Tata Mutual Fund
        s = 88: why = "initials"
        For Each t In lt
            If Not InList(GENERIC_W, CStr(t)) And InStr(hits, "|" & t & "|") = 0 Then hits = hits & t & "|"
        Next t
    End If
    If s = 0 Then Exit Function
    ' money direction vs ledger meaning (e.g. Premium/Expense ledgers rarely receive money)
    For Each x In lt
        If Not isIn And InList(IN_WORDS, CStr(x)) Then s = s * 0.6: Exit For
        If isIn And InList(OUT_WORDS, CStr(x)) Then s = s * 0.6: Exit For
    Next x
    ScoreLedger = s
End Function

Private Function HitsCover(ByVal big As String, ByVal small As String) As Boolean
    Dim t As Variant
    If small = "|" Or small = "" Then Exit Function
    For Each t In Split(small, "|")
        If t <> "" Then
            If InStr(big, "|" & t & "|") = 0 Then Exit Function
        End If
    Next t
    HitsCover = True
End Function

'================ LEARNING ENGINE ================
' Learned sheet: A key, B ledger, C times used, D latest transaction date, E ledgers seen (|A|B|)
' Keys: sorted narration words, "#AC <account no>" (full number) or "#AX <digits>" (masked, e.g. XXXX3508)
Private Function LHash(ByVal k As String) As Long
    Dim h As Long, i As Long
    For i = 1 To Len(k)
        h = (h * 31 + Asc(Mid$(k, i, 1))) Mod LH_SIZE
    Next i
    LHash = h
End Function

Private Function LSlotSig(ByVal k As String) As Long
    Dim h As Long
    h = LHash(k)
    Do While gLHKey(h) <> "" And gLHKey(h) <> k
        h = (h + 1) Mod LH_SIZE
    Loop
    LSlotSig = h
End Function

Private Function LSlotTok(ByVal k As String) As Long
    Dim h As Long
    h = LHash(k)
    Do While gTHKey(h) <> "" And gTHKey(h) <> k
        h = (h + 1) Mod LH_SIZE
    Loop
    LSlotTok = h
End Function

Private Sub LGrow(ByVal need As Long)
    Dim cap As Long
    On Error Resume Next
    cap = -1
    cap = UBound(gLSig)
    On Error GoTo 0
    If need <= cap Then Exit Sub
    If cap < 1000 Then cap = 1000
    Do While cap < need: cap = cap * 2: Loop
    ReDim Preserve gLSig(1 To cap): ReDim Preserve gLLed(1 To cap): ReDim Preserve gLSeen(1 To cap)
    ReDim Preserve gLCnt(1 To cap): ReDim Preserve gLDate(1 To cap): ReDim Preserve gLKind(1 To cap)
    ReDim Preserve gLLegacy(1 To cap): ReDim Preserve gLTok(1 To cap): ReDim Preserve gLMark(1 To cap)
End Sub

Private Function LAppend(ByVal sig As String, ByVal led As String, ByVal cnt As Long, ByVal dt As Double, _
        ByVal seen As String, ByVal legacy As Boolean) As Long
    Dim t As Variant, h As Long
    gLN = gLN + 1
    LGrow gLN
    gLSig(gLN) = sig: gLLed(gLN) = led: gLCnt(gLN) = cnt: gLDate(gLN) = dt
    gLSeen(gLN) = seen: gLLegacy(gLN) = legacy: gLMark(gLN) = 0
    If Left$(sig, 4) = "#AC " Then
        gLKind(gLN) = 1: gLTok(gLN) = Array()
    ElseIf Left$(sig, 4) = "#AX " Then
        gLKind(gLN) = 2: gLTok(gLN) = Array()
    Else
        gLKind(gLN) = 0
        gLTok(gLN) = Split(sig, " ")
        For Each t In gLTok(gLN)
            If Len(t) >= 4 Then
                h = LSlotTok(CStr(t))
                gTHKey(h) = CStr(t)
                gTHList(h) = gTHList(h) & gLN & ","
            End If
        Next t
    End If
    h = LSlotSig(sig)
    gLHKey(h) = sig: gLHIdx(h) = gLN
    LAppend = gLN
End Function

' Read the Learned sheet into memory
Private Sub LearnLoad()
    Dim ws As Worksheet, lastR As Long, v As Variant, r As Long, sig As String, led As String, dt As Double, seen As String
    gLN = 0: gLStamp = 0
    ReDim gLHKey(0 To LH_SIZE - 1): ReDim gLHIdx(0 To LH_SIZE - 1)
    ReDim gTHKey(0 To LH_SIZE - 1): ReDim gTHList(0 To LH_SIZE - 1)
    LGrow 1000
    gLLoaded = True: gLDirty = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Learned")
    If ws Is Nothing Then Exit Sub
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then Exit Sub
    v = ws.Range("A2:E" & lastR).Value
    On Error GoTo 0
    If Not IsArray(v) Then Exit Sub
    For r = 1 To UBound(v, 1)
        sig = CellText(v(r, 1)): led = CellText(v(r, 2))
        If sig <> "" And led <> "" Then
            If gLHKey(LSlotSig(sig)) = "" And gLN < 45000 Then
                dt = 0
                If IsDate(v(r, 4)) Then dt = CDbl(CDate(v(r, 4)))
                seen = CellText(v(r, 5))
                If seen = "" Then
                    LAppend sig, led, CLng(Val(CellText(v(r, 3)))), dt, "|" & led & "|", True
                Else
                    LAppend sig, led, CLng(Val(CellText(v(r, 3)))), dt, seen, False
                End If
            Else
                gLDirty = True          ' duplicate row - will be tidied on save
            End If
        End If
    Next r
End Sub

' Write memory back to the Learned sheet (one write)
Private Sub LearnSave()
    Dim ws As Worksheet, i As Long, outv() As Variant, lastR As Long
    If Not gLLoaded Or Not gLDirty Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Learned")
    If ws Is Nothing Or gLN = 0 Then Exit Sub
    ReDim outv(1 To gLN, 1 To 5)
    For i = 1 To gLN
        outv(i, 1) = gLSig(i): outv(i, 2) = gLLed(i): outv(i, 3) = gLCnt(i)
        If gLDate(i) > 0 Then outv(i, 4) = CDate(gLDate(i))
        outv(i, 5) = gLSeen(i)
    Next i
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If CellText(ws.Range("E1").Value) = "" Then
        ws.Range("E1").Value = "Ledgers seen"
        ws.Range("E1").Interior.Color = ws.Range("D1").Interior.Color
        ws.Range("E1").Font.Color = ws.Range("D1").Font.Color
        ws.Range("E1").Font.Bold = ws.Range("D1").Font.Bold
        ws.Range("E1").Font.Name = ws.Range("D1").Font.Name
        ws.Columns("E").ColumnWidth = 40
    End If
    If CellText(ws.Range("D1").Value) = "Last used" Then ws.Range("D1").Value = "Latest entry"
    ws.Range("A2").Resize(gLN, 1).NumberFormat = "@"
    ws.Range("D2").Resize(gLN, 1).NumberFormat = "dd-mm-yyyy"
    ws.Range("A2").Resize(gLN, 5).Value = outv
    GridBorders ws.Range("A1").Resize(gLN + 1, 5)
    If lastR > gLN + 1 Then ws.Range("A" & (gLN + 2) & ":E" & lastR).ClearContents
    gLDirty = False
End Sub

Private Function SeenCount(ByVal seen As String) As Long
    SeenCount = Len(seen) - Len(Replace(seen, "|", "")) - 1
End Function

' Store one key -> ledger. The entry with the latest transaction date wins.
Private Sub LearnPut(ByVal key As String, ByVal led As String, ByVal dt As Double)
    Dim h As Long, i As Long
    h = LSlotSig(key)
    If gLHKey(h) = "" Then
        If gLN >= 45000 Then Exit Sub
        LAppend key, led, 1, dt, "|" & led & "|", False
    Else
        i = gLHIdx(h)
        gLCnt(i) = gLCnt(i) + 1
        If gLLegacy(i) Or dt >= gLDate(i) Then
            gLLed(i) = led: gLDate(i) = dt: gLLegacy(i) = False
        End If
        If InStr(1, gLSeen(i), "|" & led & "|", vbTextCompare) = 0 And Len(gLSeen(i)) < 240 Then gLSeen(i) = gLSeen(i) & led & "|"
    End If
    gLDirty = True
End Sub

' Account numbers in a narration -> "|#AC 920010074344192|#AX 3508|"
Private Function AcctKeys(ByVal narr As String, ByVal bank As String) As String
    Dim u As String, i As Long, j As Long, ch As String, run As String, st As Long, out As String, k As String
    Dim isRef As Boolean, bd As Variant, b As Variant, nm As Long, lt As String, pre As String, acPre As Boolean, nx As String
    narr = JoinWrapped(narr)
    u = UCase$(narr) & " "
    pre = " " & SpacedUpper(narr) & " "
    isRef = (InStr(pre, " IMPS ") > 0) Or (InStr(pre, " UPI ") > 0)
    bd = Split(LedgerDigits(bank), " ")
    out = "|"
    For i = 1 To Len(u)
        ch = Mid$(u, i, 1)
        If ch Like "#" Then
            If run = "" Then st = i
            run = run & ch
        ElseIf run <> "" Then
            k = ""
            ' letters right before the number
            lt = "": j = st - 1
            Do While j >= 1
                If Mid$(u, j, 1) Like "[A-Z*]" Then lt = Mid$(u, j, 1) & lt: j = j - 1 Else Exit Do
            Loop
            nm = Len(lt) - Len(Replace(Replace(lt, "X", ""), "*", ""))
            pre = " " & SpacedUpper(Left$(narr, st - 1))
            acPre = (Right$(pre, 3) = " AC" Or Right$(pre, 4) = " A C" Or Right$(pre, 5) = " ACNO" Or Right$(pre, 6) = " AC NO" Or _
                     Right$(pre, 7) = " A C NO" Or Right$(pre, 5) = " ACCT" Or Right$(pre, 8) = " ACCOUNT" Or Right$(pre, 11) = " ACCOUNT NO" Or _
                     Right$(pre, 8) = " ACCT NO" Or Right$(lt, 2) = "AC" Or Right$(lt, 4) = "ACNO")
            nx = ch
            If Replace(run, "0", "") <> "" And Not (nx Like "[A-Z]") Then
                If nm >= 2 And nm = Len(lt) Then
                    If Len(run) >= 4 Then k = "#AX " & run                    ' XXXXXX3508
                ElseIf lt = "" Or acPre Then
                    If Len(run) >= 9 And Len(run) <= 18 Then
                        If Not (isRef And Len(run) = 12 And Not acPre) Then k = "#AC " & run
                    ElseIf acPre And Len(run) >= 4 Then
                        k = "#AX " & run                                       ' A/C 3508
                    End If
                End If
            End If
            ' never learn this statement's own account number
            If k <> "" Then
                For Each b In bd
                    If Len(b) >= 4 Then
                        If Right$(run, Len(b)) = b Or Right$(b, Len(run)) = run Then k = ""
                    End If
                Next b
            End If
            If k <> "" And InStr(out, "|" & k & "|") = 0 Then out = out & k & "|"
            run = ""
        End If
    Next i
    If out = "|" Then out = ""
    AcctKeys = out
End Function

' "924010021783 508" (number wrapped onto two lines in the statement) -> "924010021783508"
Private Function JoinWrapped(ByVal s As String) As String
    Dim i As Long, n1 As Long, j As Long, n2 As Long
    i = 2
    Do While i < Len(s)
        If Mid$(s, i, 1) = " " And Mid$(s, i - 1, 1) Like "#" And Mid$(s, i + 1, 1) Like "#" Then
            n1 = 0: j = i - 1
            Do While j >= 1
                If Mid$(s, j, 1) Like "#" Then n1 = n1 + 1: j = j - 1 Else Exit Do
            Loop
            n2 = 0: j = i + 1
            Do While j <= Len(s)
                If Mid$(s, j, 1) Like "#" Then n2 = n2 + 1: j = j + 1 Else Exit Do
            Loop
            If n1 >= 9 And n2 <= 6 And n1 + n2 <= 18 Then s = Left$(s, i - 1) & Mid$(s, i + 1)
        End If
        i = i + 1
    Loop
    JoinWrapped = s
End Function

Private Function AcctSame(ByVal k1 As Long, ByVal d1 As String, ByVal k2 As Long, ByVal d2 As String) As Boolean
    Dim sh As String, lg As String
    If k1 = 1 And k2 = 1 Then AcctSame = (d1 = d2): Exit Function
    If Len(d1) <= Len(d2) Then sh = d1: lg = d2 Else sh = d2: lg = d1
    If Len(sh) < 4 Then Exit Function
    AcctSame = (Right$(lg, Len(sh)) = sh)
End Function

' Learn one narration -> ledger. Returns number of keys stored (0 = skipped, why = "suspense" / "nowords" / "blank")
Private Function LearnAdd(ByVal narr As String, ByVal led As String, ByVal bank As String, ByVal dt As Double, ByRef why As String) As Long
    Dim sig As String, k As Variant, n As Long
    why = ""
    led = Trim$(led): narr = Trim$(narr)
    If narr = "" Or led = "" Then why = "blank": Exit Function
    If InStr(1, led, "suspense", vbTextCompare) > 0 Then why = "suspense": Exit Function
    If bank <> "" Then
        If StrComp(led, bank, vbTextCompare) = 0 Then why = "blank": Exit Function
    End If
    If Not gLLoaded Then LearnLoad
    If dt <= 0 Then dt = CDbl(Now)
    sig = SortedSig(NarrTokens(narr))
    If sig <> "" Then LearnPut sig, led, dt: n = 1
    For Each k In Split(AcctKeys(narr, bank), "|")
        If k <> "" Then LearnPut CStr(k), led, dt: n = n + 1
    Next k
    If n = 0 Then why = "nowords"
    LearnAdd = n
End Function

' Remember narration -> ledger (single row; used by Change Ledger)
Private Sub LearnFromRow(ByVal narr As String, ByVal ledger As String, Optional ByVal bank As String = "", Optional ByVal dt As Double = 0)
    Dim why As String
    On Error Resume Next
    LearnLoad
    LearnAdd narr, ledger, bank, dt, why
    LearnSave
End Sub

' Learn from the vouchers of one bank ledger that are already in Tally
Private Function LearnFromTally(ByVal bank As String, ByVal dFrom As Date, ByVal dTo As Date, ByRef nV As Long) As Boolean
    Dim vc As Collection, v As Variant, info As String, why As String
    On Error GoTo fin
    Set vc = FetchBankVouchers(bank, dFrom, dTo, True, info)
    If vc Is Nothing Then Exit Function
    If Not gLLoaded Then LearnLoad
    For Each v In vc
        If Not CBool(v(10)) And CStr(v(5)) <> "" And CStr(v(6)) <> "" Then
            If LearnAdd(CStr(v(6)), CStr(v(5)), bank, CDbl(v(0)), why) > 0 Then nV = nV + 1
        End If
    Next v
    LearnFromTally = True
fin:
End Function

' Ledger to use from learned row i on a statement of 'bank' (never the bank itself:
' "PRABHAT SINGHA" is the savings a/c on the OD statement and the OD a/c on the savings statement)
Private Function LedForBank(ByVal i As Long, ByVal bank As String) As String
    Dim p As Variant, k As Long
    If bank = "" Or StrComp(gLLed(i), bank, vbTextCompare) <> 0 Then LedForBank = gLLed(i): Exit Function
    p = Split(gLSeen(i), "|")
    For k = UBound(p) To 0 Step -1
        If p(k) <> "" And StrComp(CStr(p(k)), bank, vbTextCompare) <> 0 Then LedForBank = p(k): Exit Function
    Next k
End Function

' Learned account number in this narration? (latest entry wins)
Private Function AcctMatch(ByVal narr As String, ByVal bank As String, ByRef led As String, ByRef keyShown As String, ByRef conflict As Boolean) As Boolean
    Dim k As Variant, i As Long, kk As Long, kd As String, best As Long, cand As String, bestLed As String
    conflict = False
    For Each k In Split(AcctKeys(narr, bank), "|")
        If k <> "" Then
            If Left$(k, 4) = "#AC " Then kk = 1 Else kk = 2
            kd = Mid$(k, 5)
            For i = 1 To gLN
                If gLKind(i) > 0 Then
                    If AcctSame(kk, kd, gLKind(i), Mid$(gLSig(i), 5)) Then
                        cand = LedForBank(i, bank)
                        If SeenCount(gLSeen(i)) < 3 And cand <> "" Then
                            If GroupOf(cand) <> "" Or gCount = 0 Then
                                If best = 0 Then
                                    best = i: bestLed = cand
                                Else
                                    If StrComp(cand, bestLed, vbTextCompare) <> 0 Then conflict = True
                                    If gLDate(i) > gLDate(best) Then best = i: bestLed = cand
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    Next k
    If best > 0 Then
        led = bestLed: keyShown = Mid$(gLSig(best), 5)
        AcctMatch = True
    End If
End Function

' Closest learned narration (shares a long word; ties -> latest entry)
Private Function LearnedBest(ByVal nt As Variant, ByVal bank As String, ByRef bestL As String, ByRef bestSeen As Long) As Double
    Dim kw As Variant, h As Long, ix As Variant, i As Long, j As Double, bestJ As Double, bestD As Double, d As Double, cand As String
    If UBound(nt) < 0 Or gLN = 0 Then Exit Function
    gLStamp = gLStamp + 1
    For Each kw In nt
        If Len(kw) >= 4 Then
            h = LSlotTok(CStr(kw))
            If gTHKey(h) <> "" Then
                For Each ix In Split(gTHList(h), ",")
                    If ix <> "" Then
                        i = CLng(ix)
                        If gLMark(i) <> gLStamp Then
                            gLMark(i) = gLStamp
                            j = Jaccard(nt, gLTok(i))
                            If gLLegacy(i) Then d = 0 Else d = gLDate(i)
                            If j > bestJ Or (j = bestJ And j > 0 And d > bestD) Then
                                cand = LedForBank(i, bank)
                                If cand <> "" Then bestJ = j: bestL = cand: bestD = d: bestSeen = SeenCount(gLSeen(i))
                            End If
                        End If
                    End If
                Next ix
            End If
        End If
    Next kw
    LearnedBest = bestJ
End Function

Private Function KeywordHit(ByVal u As String, ByVal cu As String, ByVal kw As String) As Boolean
    Dim ks As String, kc As String
    ks = SpacedUpper(kw)
    If ks = "" Then Exit Function
    If InStr(u, " " & ks & " ") > 0 Then KeywordHit = True: Exit Function
    kc = Replace(ks, " ", "")
    If Len(kc) >= 5 Then KeywordHit = (InStr(cu, kc) > 0)
End Function

' Decide ledger + voucher type. quality: 3=green 2=yellow 1=red
Private Sub MatchNarration(ByVal narr As String, ByVal isIn As Boolean, ByVal bank As String, _
        ByRef ledger As String, ByRef vType As String, ByRef note As String, ByRef quality As Long)
    Dim ws As Worksheet, lastR As Long, r As Long, kw As Variant, u As String, cu As String, dirn As String
    Dim nt As Variant, nc As String, digs As String, i As Long, lv As Variant, bestJ As Double, j As Double
    Dim bestL As String, sigTok As Variant, npre As String
    Dim sc() As Double, tr() As Long, hs() As String, wy() As String, h As String, w As String
    Dim bestSc As Double, bestI As Long, pickI As Long, t As Long, k As Long, alt As String, nAlt As Long
    Dim used() As Boolean, altIdx As Long, altScore As Double, g As String, kwConflict As Boolean, nSeen As Long

    ledger = "": vType = "": note = "": quality = 1
    u = " " & SpacedUpper(narr) & " "
    cu = Replace(u, " ", "")
    nt = NarrTokens(narr)
    nc = ""
    If UBound(nt) >= 0 Then nc = Join(nt, "")
    digs = NarrDigits(narr)
    npre = " "
    If UBound(nt) >= 0 Then
        For Each kw In nt
            npre = npre & Left$(kw, 3) & " "
        Next kw
    End If

    ' 0) Your own earlier postings win: same account number, or exactly the same narration words
    If Not gLLoaded Then LearnLoad
    ' 0a) same account number as before
    If AcctMatch(narr, bank, bestL, alt, kwConflict) Then
        ledger = bestL
        note = "Learned: account ..." & Right$(alt, 4) & " was posted to this ledger before"
        If kwConflict Then
            quality = 2: note = note & " - CHECK, this number was used with other ledgers too"
        Else
            quality = 3
        End If
        alt = ""
        GoTo finish
    End If
    ' 0b) exactly the same narration as before (and always the same ledger)
    bestL = ""
    bestJ = LearnedBest(nt, bank, bestL, nSeen)
    If UBound(nt) >= 0 And gLN > 0 Then          ' short words (e.g. "TML DIV") - exact same words only
        k = gLHIdx(LSlotSig(SortedSig(nt)))
        If gLHKey(LSlotSig(SortedSig(nt))) <> "" And k > 0 Then
            If LedForBank(k, bank) <> "" Then bestJ = 1: bestL = LedForBank(k, bank): nSeen = SeenCount(gLSeen(k))
        End If
    End If
    If bestJ >= 0.999 And bestL <> "" And nSeen <= 1 Then
        If GroupOf(bestL) <> "" Or gCount = 0 Then
            ledger = bestL
            note = "Learned (100% same as before)": quality = 3
            GoTo finish
        End If
    End If

    ' 1) Rules sheet (your keywords win over name matching)
    Set ws = ThisWorkbook.Worksheets("Rules")
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        dirn = LCase$(CellText(ws.Cells(r, 3).Value))
        If dirn = "" Or (dirn = "in" And isIn) Or (dirn = "out" And Not isIn) Then
            For Each kw In Split(CellText(ws.Cells(r, 1).Value), "|")
                If Trim$(kw) <> "" Then
                    If KeywordHit(u, cu, CStr(kw)) Then
                        ledger = ResolveLedger(CellText(ws.Cells(r, 2).Value), bank, isIn)
                        If ledger <> "" Then
                            vType = CellText(ws.Cells(r, 4).Value)
                            note = "Rule: " & Trim$(kw): quality = 3
                            GoTo finish
                        End If
                    End If
                End If
            Next kw
        End If
    Next r

    ' 2) similar narration posted before
    If bestJ >= 0.6 And bestL <> "" Then
        If GroupOf(bestL) <> "" Or gCount = 0 Then
            ledger = bestL
            note = "Learned (" & Format$(bestJ * 100, "0") & "% same as before)"
            If bestJ >= 0.8 Then quality = 3 Else quality = 2
            GoTo finish
        End If
    End If

    ' 3) Name matching with accounting sense
    If gCount > 0 Then
        ReDim sc(1 To gCount): ReDim tr(1 To gCount): ReDim hs(1 To gCount): ReDim wy(1 To gCount): ReDim used(1 To gCount)
        For i = 1 To gCount
            If Not IsExcluded(i, bank) Then
                sc(i) = ScoreLedger(i, nt, nc, digs, isIn, h, w, npre)
                hs(i) = h: wy(i) = w: tr(i) = TierOf(i, isIn)
                If sc(i) > bestSc Then bestSc = sc(i): bestI = i
            End If
        Next i
        If bestSc >= 70 Then
            ' party first: debtor (money in) / creditor (money out), then receivable/payable, if it is the same name
            For t = 1 To 3
                pickI = 0
                For i = 1 To gCount
                    If tr(i) = t And sc(i) >= 40 Then
                        If sc(i) >= bestSc - 30 Or HitsCover(hs(i), hs(bestI)) Then
                            If pickI = 0 Then
                                pickI = i
                            ElseIf sc(i) > sc(pickI) Then
                                pickI = i
                            End If
                        End If
                    End If
                Next i
                If pickI > 0 Then Exit For
            Next t
            If pickI = 0 Then pickI = bestI
            ledger = gNames(pickI)
            If sc(pickI) >= 85 Then quality = 3 Else quality = 2
            note = "Match " & Format$(sc(pickI), "0") & "% (" & wy(pickI) & ")"
            ' another ledger of the same kind is just as close -> check
            For i = 1 To gCount
                If i <> pickI And tr(i) = tr(pickI) And sc(i) >= 40 And sc(i) >= sc(pickI) - 10 Then quality = 2
            Next i
            ' unusual direction: money paid to a debtor / received from a creditor
            g = LCase$(gGroups(pickI))
            If (Not isIn And InStr(g, "sundry debtors") > 0) Or (isIn And InStr(g, "sundry creditors") > 0) Then
                quality = 2: note = note & " - unusual: " & IIf(isIn, "receipt from a creditor", "payment to a debtor")
            End If
            ' top alternatives
            used(pickI) = True
            For k = 1 To 2
                altIdx = 0: altScore = 0
                For i = 1 To gCount
                    If Not used(i) And sc(i) >= 50 And sc(i) > altScore Then altScore = sc(i): altIdx = i
                Next i
                If altIdx > 0 Then
                    used(altIdx) = True
                    alt = alt & IIf(alt = "", "", " | ") & gNames(altIdx) & " " & Format$(sc(altIdx), "0") & "%"
                End If
            Next k
            If quality < 3 And alt <> "" Then note = note & " - CHECK, also: " & alt
            GoTo finish
        End If
    End If

    ' 4) Nothing found -> Suspense (if the company has it) or blank
    ledger = ResolveLedger("Suspense A/c|Suspense Account|Suspense", bank, isIn)
    If ledger <> "" Then note = "No match - put in Suspense" Else note = "No match - fill manually"
    quality = 1

finish:
    If vType = "" Then
        If ledger <> "" And IsCashOrBank(ledger) Then
            vType = "Contra"
        ElseIf isIn Then
            vType = "Receipt"
        Else
            vType = "Payment"
        End If
    End If
End Sub

Private Function GuessTxnType(ByVal narr As String) As String
    Dim u As String
    u = " " & SpacedUpper(narr) & " "
    If InStr(u, " NEFT ") + InStr(u, " RTGS ") + InStr(u, " IMPS ") + InStr(u, " UPI ") + InStr(u, " INB ") + _
       InStr(u, " ACH ") + InStr(u, " NACH ") + InStr(u, " ECS ") + InStr(u, " NETBANKING ") > 0 Then
        GuessTxnType = "e-Fund Transfer"
    ElseIf InStr(u, " CHQ ") + InStr(u, " CHEQUE ") + InStr(u, " CLG ") + InStr(u, " CLEARING ") > 0 Then
        GuessTxnType = "Cheque"
    End If
End Function

Private Sub FillMatch(ws As Worksheet, ByVal r As Long, ByVal bank As String)
    Dim narr As String, isIn As Boolean, led As String, vt As String, note As String, q As Long, clr As Long
    narr = CellText(ws.Cells(r, COL_NARR).Value)
    isIn = (Amt(ws.Cells(r, COL_IN).Value) > 0)
    MatchNarration narr, isIn, bank, led, vt, note, q
    ws.Cells(r, COL_PARTY).Value = led
    ws.Cells(r, COL_TYPE).Value = vt
    If CellText(ws.Cells(r, COL_TXN).Value) = "" Then ws.Cells(r, COL_TXN).Value = GuessTxnType(narr)
    Select Case q
        Case 3: clr = CLR_OK
        Case 2: clr = CLR_INFO
        Case Else: clr = CLR_ERR
    End Select
    ws.Cells(r, COL_NOTE).Value = note
    ws.Cells(r, COL_NOTE).Interior.Color = clr
    ws.Cells(r, COL_PARTY).Interior.Color = clr
    ws.Cells(r, COL_NOTE).Font.Color = RGB(30, 41, 59)
    ws.Cells(r, COL_PARTY).Font.Color = RGB(30, 41, 59)
    ws.Cells(r, COL_Q).Value = q
    FitRow ws, r
End Sub

' Row grows to show wrapped narration / status / note
Private Sub FitRow(ws As Object, ByVal r As Long)
    On Error Resume Next
    ws.Rows(r).AutoFit
    If ws.Rows(r).RowHeight < 16 Then ws.Rows(r).RowHeight = 16
End Sub

'---------------- Statement parsing ----------------
Private Function ParseAmount(ByVal v As Variant, ByRef side As String) As Double
    Dim s As String, u As String, neg As Boolean, i As Long, ch As String, clean As String
    side = ""
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbSingle Then
        ParseAmount = CDbl(v): Exit Function
    End If
    s = CellText(v): u = UCase$(s)
    If s = "" Then Exit Function
    If Right$(u, 2) = "CR" Or Right$(u, 3) = "CR." Then side = "C"
    If Right$(u, 2) = "DR" Or Right$(u, 3) = "DR." Then side = "D"
    If InStr(s, "(") > 0 And InStr(s, ")") > 0 Then neg = True
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "#" Or ch = "." Then
            clean = clean & ch
        ElseIf ch = "-" And clean = "" Then
            neg = True
        End If
    Next i
    If clean = "" Or clean = "." Then Exit Function
    If neg Then ParseAmount = -Val(clean) Else ParseAmount = Val(clean)
End Function

Private Function MonthNum(ByVal s As String) As Long
    Dim p As Long
    s = UCase$(Left$(s, 3))
    p = InStr("JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC", s)
    If p > 0 And (p - 1) Mod 3 = 0 And Len(s) = 3 Then MonthNum = (p - 1) \ 3 + 1
End Function

' Indian format (day first). Returns Empty if not a date.
Private Function ParseDateCell(ByVal v As Variant) As Variant
    Dim s As String, parts As Variant, d As Long, m As Long, y As Long, tmp As Long
    ParseDateCell = Empty
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then ParseDateCell = DateValue(v): Exit Function
    If IsNumeric(v) And Not VarType(v) = vbString Then
        If v > 20000 And v < 80000 Then ParseDateCell = DateValue(CDate(v))
        Exit Function
    End If
    s = UCase$(CellText(v))
    If Len(s) < 6 Or Len(s) > 25 Then Exit Function
    s = Replace(Replace(Replace(Replace(s, "/", " "), "-", " "), ".", " "), ",", " ")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    parts = Split(Trim$(s), " ")
    If UBound(parts) < 2 Then Exit Function
    If Len(parts(0)) = 4 And IsNumeric(parts(0)) Then        ' yyyy mm dd
        tmp = 0: y = Val(parts(0)): m = Val(parts(1)): d = Val(parts(2))
    Else
        If Not IsNumeric(parts(0)) Then Exit Function
        d = Val(parts(0))
        If IsNumeric(parts(1)) Then m = Val(parts(1)) Else m = MonthNum(CStr(parts(1)))
        If Not IsNumeric(Left$(parts(2), 2)) Then Exit Function
        y = Val(Left$(parts(2), 4))
    End If
    If y < 100 Then y = y + 2000
    If d < 1 Or d > 31 Or m < 1 Or m > 12 Or y < 2000 Or y > 2100 Then Exit Function
    Dim dt As Date
    On Error GoTo bad
    dt = DateSerial(y, m, d)
    If Month(dt) = m And Day(dt) = d Then ParseDateCell = dt
bad:
End Function

Private Function HdrMatch(ByVal h As String, ByVal words As String) As Boolean
    Dim w As Variant
    For Each w In Split(words, "|")
        If InStr(h, CStr(w)) > 0 Then HdrMatch = True: Exit Function
    Next w
End Function

' Returns a Collection of Array(date, narration, withdrawal, deposit, ref)
Private Function ParseStatement(ByVal arr As Variant, ByRef info As String) As Collection
    Dim res As New Collection, nR As Long, nC As Long, i As Long, j As Long, h As String
    Dim hr As Long, cDate_ As Long, cVal As Long, cNarr As Long, cWd As Long, cDep As Long
    Dim cAmt As Long, cDrCr As Long, cRef As Long, dv As Variant, narr As String
    Dim w As Double, dp As Double, a As Double, sd As String, sd2 As String, ref As String
    Dim last As Variant, lowNarr As String, filled As Long, cBal As Long, bal As Double, hasBal As Boolean, sb As String

    Set ParseStatement = res
    nR = UBound(arr, 1): nC = UBound(arr, 2)

    For i = LBound(arr, 1) To Application.Min(nR, 80)
        cDate_ = 0: cVal = 0: cNarr = 0: cWd = 0: cDep = 0: cAmt = 0: cDrCr = 0: cRef = 0: cBal = 0
        For j = LBound(arr, 2) To nC
            h = LCase$(CellText(arr(i, j)))
            h = Replace(h, ".", "")
            If h <> "" And Len(h) < 45 Then
                If HdrMatch(h, "dr/cr|cr/dr|dr / cr|debit/credit|type") Then
                    If cDrCr = 0 Then cDrCr = j
                ElseIf HdrMatch(h, "date") Or h = "dt" Or Right$(h, 3) = " dt" Then
                    If InStr(h, "value") > 0 Then
                        If cVal = 0 Then cVal = j
                    ElseIf cDate_ = 0 Then
                        cDate_ = j
                    End If
                ElseIf HdrMatch(h, "balance") Then
                    If cBal = 0 And InStr(h, "opening") = 0 Then cBal = j
                ElseIf HdrMatch(h, "withdraw|debit|paid out|money out") Or h = "dr" Or Left$(h, 3) = "dr " Or Left$(h, 3) = "dr(" Then
                    If cWd = 0 Then cWd = j
                ElseIf HdrMatch(h, "deposit|credit|paid in|money in") Or h = "cr" Or Left$(h, 3) = "cr " Or Left$(h, 3) = "cr(" Then
                    If cDep = 0 Then cDep = j
                ElseIf HdrMatch(h, "chq|cheque|ref|utr|instrument") Then
                    If cRef = 0 Then cRef = j
                ElseIf HdrMatch(h, "amount|amt") Then
                    If cAmt = 0 Then cAmt = j
                ElseIf HdrMatch(h, "narration|naration|description|particular|remark|details|transaction") Then
                    If cNarr = 0 Then cNarr = j
                End If
            End If
        Next j
        If cDate_ = 0 Then cDate_ = cVal
        If cDate_ > 0 And cNarr > 0 And ((cWd > 0 And cDep > 0) Or cAmt > 0) Then hr = i: Exit For
    Next i

    If hr = 0 Then
        info = "Could not find the header row (Date / Narration / Withdrawal / Deposit columns) in the statement."
        Exit Function
    End If

    For i = hr + 1 To nR
        dv = ParseDateCell(arr(i, cDate_))
        narr = CellText(arr(i, cNarr))
        lowNarr = LCase$(narr)
        If cWd > 0 And cDep > 0 Then
            w = Abs(ParseAmount(arr(i, cWd), sd)): dp = Abs(ParseAmount(arr(i, cDep), sd))
        Else
            a = ParseAmount(arr(i, cAmt), sd): sd2 = ""
            If cDrCr > 0 Then
                sd2 = UCase$(CellText(arr(i, cDrCr)))
                If Left$(sd2, 1) = "D" Then
                    sd = "D"
                ElseIf Left$(sd2, 1) = "C" Then
                    sd = "C"
                End If
            End If
            If sd = "" Then
                If a < 0 Then sd = "D" Else sd = "C"
            End If
            w = 0: dp = 0
            If sd = "D" Then w = Abs(a) Else dp = Abs(a)
        End If
        ref = ""
        If cRef > 0 Then ref = CellText(arr(i, cRef))
        bal = 0: hasBal = False
        If cBal > 0 Then
            If CellText(arr(i, cBal)) <> "" Then
                bal = ParseAmount(arr(i, cBal), sb)
                If sb = "D" Then bal = -Abs(bal)
                If sb = "C" Then bal = Abs(bal)
                hasBal = True
            End If
        End If

        If IsEmpty(dv) Then
            ' continuation of previous narration (multi-line PDFs)
            If narr <> "" And w = 0 And dp = 0 And res.Count > 0 And CellText(arr(i, cDate_)) = "" Then
                last = res(res.Count)
                res.Remove res.Count
                last(1) = Trim$(last(1) & " " & narr)
                res.Add last
            End If
        ElseIf (w > 0 Or dp > 0) And InStr(lowNarr, "opening balance") = 0 And InStr(lowNarr, "closing balance") = 0 Then
            res.Add Array(dv, narr, w, dp, ref, bal, hasBal)
        End If
    Next i
    If res.Count = 0 Then info = "Header row found (row " & hr & ") but no transactions could be read below it."
End Function

Private Function ExcelToArray(ByVal f As String) As Variant
    Dim wbS As Workbook, sh As Worksheet, best As Worksheet, mx As Double
    On Error GoTo fail
    Application.DisplayAlerts = False
    Set wbS = Workbooks.Open(Filename:=f, UpdateLinks:=0, ReadOnly:=True)
    For Each sh In wbS.Worksheets
        If sh.UsedRange.Cells.Count > mx Then mx = sh.UsedRange.Cells.Count: Set best = sh
    Next sh
    If Not best Is Nothing Then ExcelToArray = best.UsedRange.Value
    wbS.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Exit Function
fail:
    Application.DisplayAlerts = True
    MsgBox "Could not open the statement file:" & vbCrLf & Err.Description, vbExclamation
End Function

' Reads all tables of a PDF using Excel's Power Query (Excel 2021 / Microsoft 365)
Private Function PdfToArray(ByVal f As String) As Variant
    Dim qName As String, m As String, tmp As Worksheet, lo As ListObject
    On Error GoTo fail
    qName = "BankPdf" & Format$(Now, "hhnnss")
    m = "let Source = Pdf.Tables(File.Contents(""" & Replace(f, """", """""") & """))," & _
        " T = Table.SelectRows(Source, each [Kind] = ""Table"")," & _
        " U = if Table.RowCount(T) > 0 then T else Table.SelectRows(Source, each [Kind] = ""Page"")," & _
        " Data = Table.Combine(U[Data]) in Data"
    ThisWorkbook.Queries.Add Name:=qName, Formula:=m
    Application.DisplayAlerts = False
    Set tmp = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    Set lo = tmp.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qName & ";Extended Properties=""""", Destination:=tmp.Range("A1"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & qName & "]")
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
    End With
    PdfToArray = lo.Range.Value
cleanup:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Delete
    ThisWorkbook.Queries(qName).Delete
    ThisWorkbook.Connections("Query - " & qName).Delete
    Application.DisplayAlerts = True
    QuietActivate ThisWorkbook.Worksheets("Entries")
    Exit Function
fail:
    MsgBox "Could not read this PDF." & vbCrLf & vbCrLf & _
        "- PDF reading needs Excel 2021 or Microsoft 365." & vbCrLf & _
        "- Password-protected PDF? Open it in Chrome, press Ctrl+P, choose 'Save as PDF' to make an unlocked copy." & vbCrLf & _
        "- Scanned (photo) PDFs cannot be read - download the Excel statement from net banking instead." & vbCrLf & vbCrLf & _
        "Details: " & Err.Description, vbExclamation, "PDF import"
    PdfToArray = Empty
    Resume cleanup
End Function

'================ BUTTON: Import Bank Statement ================
Public Sub ImportBankStatement()
    Dim f As Variant, bank As Variant, arr As Variant, recs As Collection, info As String
    Dim ws As Worksheet, r As Long, rec As Variant, key As String, keys As Object, i As Long
    Dim nNew As Long, nDup As Long, nHi As Long, nMid As Long, nLow As Long, clr As Long, dupHit As Boolean
    Dim nTally As Long, tinfo As String, dMin As Date, nHist As Long
    Dim dMax As Date, cBalv As Double, haveBal As Boolean, ascending As Boolean, k As Long, bmsg As String

    Set ws = EntriesSheet()
    If ThisWorkbook.Worksheets("Ledgers").Range("A2").Value = "" Then
        If MsgBox("Ledgers are not loaded yet. Load them from Tally now? (Needed for auto-matching.)", vbYesNo + vbQuestion) = vbYes Then
            If LoadLedgersCore(True) = 0 Then Exit Sub
        End If
    End If

    f = Application.GetOpenFilename("Bank statement (*.xls;*.xlsx;*.xlsm;*.csv;*.pdf),*.xls;*.xlsx;*.xlsm;*.csv;*.pdf", , "Select bank statement")
    If VarType(f) = vbBoolean Then Exit Sub

    bank = Application.InputBox("Bank ledger in Tally for this statement:", "Bank Ledger", SettingValue("DefaultBank"), Type:=2)
    If VarType(bank) = vbBoolean Then Exit Sub
    bank = Trim$(CStr(bank))
    If bank = "" Then MsgBox "Bank ledger is required.", vbExclamation: Exit Sub
    LoadLedgerCache
    If gCount > 0 And GroupOf(CStr(bank)) = "" Then
        If MsgBox("'" & bank & "' is not in the loaded ledger list. Continue anyway?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
    End If

    QuietActivate ws
    ProgressShow "Importing bank statement", "Reading " & Mid$(CStr(f), InStrRev(CStr(f), "\") + 1) & "..."
    Application.ScreenUpdating = False
    If LCase$(Right$(CStr(f), 4)) = ".pdf" Then arr = PdfToArray(CStr(f)) Else arr = ExcelToArray(CStr(f))
    If Not IsArray(arr) Then Application.StatusBar = False: Application.ScreenUpdating = True: ProgressHide: Exit Sub

    ProgressUpdate 0.15, "Finding transactions in the statement..."
    Set recs = ParseStatement(arr, info)
    If recs.Count = 0 Then
        Application.StatusBar = False: Application.ScreenUpdating = True
        ProgressHide
        MsgBox info, vbExclamation, "Import statement": Exit Sub
    End If

    ' learn from this bank's entries already in Tally (picks up corrections made in Tally)
    ProgressUpdate 0.17, "Learning from your earlier entries in Tally..."
    LearnLoad
    dMin = Date
    For Each rec In recs
        If IsDate(rec(0)) Then
            If CDate(rec(0)) < dMin Then dMin = CDate(rec(0))
        End If
    Next rec
    LearnFromTally CStr(bank), DateAdd("m", -13, dMin), Date, nHist

    ' existing rows -> skip duplicates
    Set keys = CreateObject("Scripting.Dictionary")
    For i = FIRST_ROW To LastDataRow(ws)
        If IsDate(ws.Cells(i, COL_DATE).Value) Then
            key = Format$(ws.Cells(i, COL_DATE).Value, "yyyymmdd") & "|" & AmtStr(Amt(ws.Cells(i, COL_OUT).Value)) & "|" & _
                  AmtStr(Amt(ws.Cells(i, COL_IN).Value)) & "|" & Left$(CellText(ws.Cells(i, COL_NARR).Value), 40)
            keys(key) = keys(key) + 1
        End If
    Next i

    r = LastDataRow(ws) + 1
    i = 0
    For Each rec In recs
        i = i + 1
        ProgressUpdate 0.2 + 0.7 * i / recs.Count, "Matching ledgers...  " & i & " of " & recs.Count
        key = Format$(rec(0), "yyyymmdd") & "|" & AmtStr(CDbl(rec(2))) & "|" & AmtStr(CDbl(rec(3))) & "|" & Left$(CStr(rec(1)), 40)
        If keys.Exists(key) Then
            If keys(key) > 0 Then dupHit = True Else dupHit = False
        Else
            dupHit = False
        End If
        If dupHit Then
            keys(key) = keys(key) - 1     ' this exact row was already on the sheet from an earlier import
            nDup = nDup + 1
        Else
            Application.StatusBar = "Matching ledgers... row " & r
            ws.Cells(r, COL_DATE).Value = rec(0)
            ws.Cells(r, COL_BANK).Value = bank
            If rec(2) > 0 Then ws.Cells(r, COL_OUT).Value = rec(2)
            If rec(3) > 0 Then ws.Cells(r, COL_IN).Value = rec(3)
            ws.Cells(r, COL_NARR).Value = rec(1)
            If rec(4) <> "" Then ws.Cells(r, COL_REF).NumberFormat = "@": ws.Cells(r, COL_REF).Value = rec(4)
            FillMatch ws, r, CStr(bank)
            clr = ws.Cells(r, COL_NOTE).Interior.Color
            If clr = CLR_OK Then
                nHi = nHi + 1
            ElseIf clr = CLR_INFO Then
                nMid = nMid + 1
            Else
                nLow = nLow + 1
            End If
            nNew = nNew + 1
            r = r + 1
        End If
    Next rec
    LearnSave
    Application.StatusBar = False
    ProgressUpdate 0.93, "Checking Tally for entries that already exist..."
    nTally = CheckRowsAgainstTally(True, tinfo)
    If nTally < 0 Then tinfo = "Could not check Tally for existing entries (" & tinfo & ") - it will be checked again when you post." Else tinfo = nTally & " row(s) already exist in Tally (blue) and will be skipped."
    ' bank balance check (statement balance as per the document provided by the user)
    ProgressUpdate 0.97, "Comparing the bank balance with Tally..."
    rec = recs(1): dMax = CDate(rec(0))
    rec = recs(recs.Count): ascending = (dMax <= CDate(rec(0)))
    For k = 1 To recs.Count
        rec = recs(k)
        If CDate(rec(0)) > dMax Then dMax = CDate(rec(0))
    Next k
    If ascending Then
        For k = recs.Count To 1 Step -1
            rec = recs(k)
            If rec(6) Then cBalv = rec(5): haveBal = True: Exit For
        Next k
    Else
        For k = 1 To recs.Count
            rec = recs(k)
            If rec(6) Then cBalv = rec(5): haveBal = True: Exit For
        Next k
    End If
    bmsg = BankStmtBalance(CStr(bank), dMax, cBalv, haveBal, Mid$(CStr(f), InStrRev(CStr(f), "\") + 1))
    Application.ScreenUpdating = True
    ProgressDone "Import finished"
    QuietActivate ws
    LogAct "Import statement", Mid$(CStr(f), InStrRev(CStr(f), "\") + 1) & " -> " & nNew & " rows (" & nHi & " green, " & nMid & " yellow, " & nLow & " red)", nNew
    Notify nNew & " transactions imported" & IIf(nDup > 0, "  (" & nDup & " already on the sheet)", "") & vbCrLf & _
        "Green " & nHi & "   -   Yellow " & nMid & "   -   Red " & nLow & vbCrLf & tinfo & vbCrLf & _
        "Review yellow and red rows, then Post to Tally." & IIf(bmsg <> "", vbCrLf & bmsg, ""), "ok"
End Sub

'================ BUTTON: Re-match ledgers ================
Public Sub RematchLedgers()
    Dim ws As Worksheet, r As Long, bank As String, n As Long, banks As String, b As Variant, nHist As Long
    Set ws = EntriesSheet()
    If MsgBox("Re-run auto-matching for all rows not yet Posted/Exported?" & vbCrLf & _
              "This overwrites Party / Other Ledger and Voucher Type on those rows.", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    ProgressShow "Re-matching ledgers", "Learning from your entries in Tally..."
    LoadLedgerCache
    LearnLoad
    banks = "|"
    For r = FIRST_ROW To LastDataRow(ws)
        If Not IsDone(CStr(ws.Cells(r, COL_STATUS).Value)) Then
            bank = CellText(ws.Cells(r, COL_BANK).Value)
            If bank = "" Then bank = SettingValue("DefaultBank")
            If bank <> "" And InStr(1, banks, "|" & bank & "|", vbTextCompare) = 0 Then banks = banks & bank & "|"
        End If
    Next r
    For Each b In Split(banks, "|")
        If b <> "" Then LearnFromTally CStr(b), DateAdd("m", -13, Date), Date, nHist
    Next b
    LearnSave
    Application.ScreenUpdating = False
    For r = FIRST_ROW To LastDataRow(ws)
        ProgressUpdate (r - FIRST_ROW + 1) / (LastDataRow(ws) - FIRST_ROW + 1), "Matching row " & (r - FIRST_ROW + 1) & "..."
        If Not IsDone(CStr(ws.Cells(r, COL_STATUS).Value)) And CellText(ws.Cells(r, COL_NARR).Value) <> "" Then
            bank = CellText(ws.Cells(r, COL_BANK).Value)
            If bank = "" Then bank = SettingValue("DefaultBank")
            FillMatch ws, r, bank
            n = n + 1
        End If
    Next r
    Application.ScreenUpdating = True
    ProgressDone "Matching finished"
    Notify n & " rows re-matched.", "ok"
End Sub

'==============================================================
'  DUPLICATE CHECK AGAINST TALLY
'==============================================================
Private Function HasRowDate(ws As Worksheet, ByVal r As Long) As Boolean
    HasRowDate = Not IsEmpty(ParseDateCell(ws.Cells(r, COL_DATE).Value))
End Function

Private Function RowDate(ws As Worksheet, ByVal r As Long) As Date
    RowDate = ParseDateCell(ws.Cells(r, COL_DATE).Value)
End Function

Private Function TallyDateStr(ByVal d As Date) As String
    TallyDateStr = Day(d) & "-" & Mid$("JanFebMarAprMayJunJulAugSepOctNovDec", (Month(d) - 1) * 3 + 1, 3) & "-" & Year(d)
End Function

Private Function ParseTallyDate(ByVal s As String) As Variant
    s = Trim$(s)
    ParseTallyDate = Empty
    If Len(s) = 8 And IsDigitStr(s) Then
        ParseTallyDate = DateSerial(Val(Left$(s, 4)), Val(Mid$(s, 5, 2)), Val(Right$(s, 2)))
    Else
        ParseTallyDate = ParseDateCell(s)
    End If
End Function

Private Function DupDays() As Long
    DupDays = 3
    If SettingValue("DupDays") <> "" Then DupDays = Val(SettingValue("DupDays"))
    If DupDays < 0 Then DupDays = 0
End Function

' Vouchers of one bank ledger from Tally. Each item: Array(date, isIn, amount, vchType, vchNo, otherLedger, narration, used)
' Returns Nothing if Tally could not be read (message already shown unless silent).
Private Function FetchBankVouchers(ByVal bank As String, ByVal dFrom As Date, ByVal dTo As Date, _
        ByVal silent As Boolean, ByRef info As String) As Collection
    Dim req As String, resp As String, co As String, sv As String, re As Object, doc As Object
    Dim vn As Object, ln As Object, lines As Object, res As New Collection, dv As Variant
    Dim lname As String, lamt As Double, bankAmt As Double, hasBank As Boolean, other As String
    Dim nV As Long, nNoLine As Long, nd As Object, vtype As String, vno As String, vnarr As String, vguid As String, nLn As Long
    Dim multi As Boolean

    co = SettingValue("CompanyName")
    sv = "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    If co <> "" Then sv = sv & "<SVCURRENTCOMPANY>" & XmlEsc(co) & "</SVCURRENTCOMPANY>"
    sv = sv & "<SVFROMDATE>" & TallyDateStr(dFrom) & "</SVFROMDATE><SVTODATE>" & TallyDateStr(dTo) & "</SVTODATE>"
    req = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>XLBankVch</ID></HEADER>" & _
          "<BODY><DESC><STATICVARIABLES>" & sv & "</STATICVARIABLES><TDL><TDLMESSAGE>" & _
          "<COLLECTION NAME=""XLBankVch"" ISMODIFY=""No""><TYPE>Vouchers : Ledger</TYPE><CHILDOF>" & XmlEsc(bank) & "</CHILDOF>" & _
          "<BELONGSTO>Yes</BELONGSTO><FETCH>Date,VoucherTypeName,VoucherNumber,Narration,Guid,MasterId,AllLedgerEntries.LedgerName,AllLedgerEntries.Amount," & _
          "LedgerEntries.LedgerName,LedgerEntries.Amount</FETCH></COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"

    If Not PostXml(req, resp) Then
        If Not silent Then ConnectionHelp resp
        info = "Tally not reachable": Exit Function
    End If
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "&#([0-8]|1[124-9]|2[0-9]|3[01]);"
    resp = re.Replace(resp, "")
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    If Not doc.LoadXML(resp) Then
        info = "Tally's reply could not be read"
        SaveDebug "Voucher list for " & bank, resp
        If Not silent Then MsgBox "Tally replied, but the voucher list could not be read. The reply was saved on the 'TallyDebug' sheet - please send me a screenshot of it.", vbExclamation
        Exit Function
    End If

    For Each vn In doc.SelectNodes("//VOUCHER")
        nV = nV + 1
        Set nd = vn.SelectSingleNode("DATE"): dv = Empty
        If Not nd Is Nothing Then dv = ParseTallyDate(nd.Text)
        vtype = "": vno = "": vnarr = ""
        Set nd = vn.SelectSingleNode("VOUCHERTYPENAME"): If Not nd Is Nothing Then vtype = nd.Text
        If vtype = "" Then vtype = CStr(vn.getAttribute("VCHTYPE") & "")
        Set nd = vn.SelectSingleNode("VOUCHERNUMBER"): If Not nd Is Nothing Then vno = nd.Text
        Set nd = vn.SelectSingleNode("NARRATION"): If Not nd Is Nothing Then vnarr = nd.Text
        vguid = ""
        Set nd = vn.SelectSingleNode("GUID"): If Not nd Is Nothing Then vguid = Trim$(nd.Text)
        If vguid = "" Then vguid = CStr(vn.getAttribute("REMOTEID") & "")
        hasBank = False: other = "": bankAmt = 0: nLn = 0: multi = False
        Set lines = vn.SelectNodes("ALLLEDGERENTRIES.LIST")
        If lines.Length = 0 Then Set lines = vn.SelectNodes("LEDGERENTRIES.LIST")
        For Each ln In lines
            lname = "": lamt = 0
            Set nd = ln.SelectSingleNode("LEDGERNAME"): If Not nd Is Nothing Then lname = Trim$(nd.Text)
            Set nd = ln.SelectSingleNode("AMOUNT"): If Not nd Is Nothing Then lamt = Val(Replace(nd.Text, ",", ""))
            If StrComp(lname, bank, vbTextCompare) = 0 Then
                hasBank = True: bankAmt = bankAmt + lamt: nLn = nLn + 1
            ElseIf other = "" And lname <> "" Then
                other = lname
            ElseIf lname <> "" And StrComp(lname, other, vbTextCompare) <> 0 Then
                multi = True
            End If
        Next ln
        If hasBank And Not IsEmpty(dv) And Abs(bankAmt) > 0.004 Then
            ' Tally: negative amount = debit to bank = money IN
            res.Add Array(dv, (bankAmt < 0), Abs(bankAmt), vtype, vno, other, vnarr, False, vguid, nLn, multi)
        Else
            nNoLine = nNoLine + 1
        End If
    Next vn

    If nV > 0 And res.Count = 0 Then
        info = "Tally sent " & nV & " vouchers but without the bank lines"
        SaveDebug "Voucher list for " & bank, resp
        If Not silent Then MsgBox info & ". The reply was saved on the 'TallyDebug' sheet - please send me a screenshot of it.", vbExclamation
        Exit Function
    End If
    info = res.Count & " vouchers of '" & bank & "' read from Tally (" & TallyDateStr(dFrom) & " to " & TallyDateStr(dTo) & ")"
    Set FetchBankVouchers = res
End Function

Private Sub SaveDebug(ByVal title As String, ByVal txt As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("TallyDebug")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "TallyDebug"
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = title & "  (" & Format$(Now, "dd-mmm-yy hh:nn") & ")"
    ws.Range("A2").Value = Left$(txt, 32000)
End Sub

' Best unused voucher for (date, direction, amount) within the day window. Returns index or 0.
Private Function FindVoucher(ByVal vc As Collection, ByVal d As Date, ByVal isIn As Boolean, ByVal amount As Double, _
        ByVal window As Long, ByVal narr As String) As Long
    Dim k As Long, v As Variant, diff As Long, best As Long, bestDiff As Long, nm As Boolean
    bestDiff = 9999
    For k = 1 To vc.Count
        v = vc(k)
        If Not v(7) And v(1) = isIn And Abs(v(2) - amount) < 0.005 Then
            diff = Abs(DateDiff("d", v(0), d))
            If diff <= window Then
                nm = (narr <> "" And InStr(1, v(6), Left$(narr, 30), vbTextCompare) > 0)
                If nm Then diff = diff - 100        ' same narration -> surely this voucher
                If diff < bestDiff Then bestDiff = diff: best = k
            End If
        End If
    Next k
    FindVoucher = best
End Function

' Voucher named in an "In Tally already: Contra No. 58 dt 20-08-26 ..." status
Private Function VoucherFromStatus(ByVal vc As Collection, ByVal st As String, ByVal isIn As Boolean, ByVal amount As Double) As Long
    Dim p1 As Long, p2 As Long, vno As String, ds As String, dv As Variant, k As Long, v As Variant
    p1 = InStr(st, " No. "): p2 = InStr(st, " dt ")
    If p2 = 0 Then Exit Function
    If p1 > 0 And p1 < p2 Then vno = Trim$(Mid$(st, p1 + 5, p2 - p1 - 5))
    ds = Mid$(st, p2 + 4, 8)
    dv = ParseDateCell(ds)
    If IsEmpty(dv) Then Exit Function
    For k = 1 To vc.Count
        v = vc(k)
        If Not v(7) And v(1) = isIn And Abs(v(2) - amount) < 0.005 And v(0) = dv Then
            If vno = "" Or CStr(v(4)) = vno Then VoucherFromStatus = k: Exit Function
        End If
    Next k
End Function

Private Sub MarkUsed(ByVal vc As Collection, ByVal k As Long)
    Dim v As Variant
    v = vc(k): v(7) = True
    vc.Remove k
    If k > vc.Count Then
        vc.Add v
    Else
        vc.Add v, Before:=k
    End If
End Sub

' Marks statement rows that already exist in Tally. Returns number marked, or -1 if Tally could not be checked.
Private Function CheckRowsAgainstTally(ByVal silent As Boolean, ByRef info As String) As Long
    Dim ws As Worksheet, lastR As Long, r As Long, banks As String, b As Variant, bank As String
    Dim dMin As Date, dMax As Date, have As Boolean, vc As Collection, k As Long, v As Variant
    Dim isIn As Boolean, amount As Double, st As String, nMark As Long, inf As String, win As Long
    Dim nP As Long, pS() As Double, pR() As Long, pK() As Long, rowDone() As Boolean, vchDone() As Boolean
    Dim bestP As Long, i As Long, diff As Long, narr As String, party As String, partyKnown As Boolean
    Dim narrHit As Boolean, ledHit As Boolean

    Set ws = EntriesSheet()
    lastR = LastDataRow(ws)
    win = DupDays()
    banks = "|"
    For r = FIRST_ROW To lastR
        If HasRowDate(ws, r) Then
            bank = CellText(ws.Cells(r, COL_BANK).Value)
            If bank = "" Then bank = SettingValue("DefaultBank")
            If bank <> "" And InStr(1, banks, "|" & bank & "|", vbTextCompare) = 0 Then banks = banks & bank & "|"
        End If
    Next r

    For Each b In Split(banks, "|")
        If b <> "" Then
            have = False
            For r = FIRST_ROW To lastR
                If HasRowDate(ws, r) Then
                    bank = CellText(ws.Cells(r, COL_BANK).Value)
                    If bank = "" Then bank = SettingValue("DefaultBank")
                    If StrComp(bank, b, vbTextCompare) = 0 Then
                        If Not have Then
                            dMin = RowDate(ws, r): dMax = dMin: have = True
                        Else
                            If RowDate(ws, r) < dMin Then dMin = RowDate(ws, r)
                            If RowDate(ws, r) > dMax Then dMax = RowDate(ws, r)
                        End If
                    End If
                End If
            Next r
            If have Then
                Application.StatusBar = "Reading '" & b & "' vouchers from Tally..."
                Set vc = FetchBankVouchers(CStr(b), dMin - win, dMax + win, silent, inf)
                info = info & IIf(info = "", "", vbCrLf) & inf
                If vc Is Nothing Then CheckRowsAgainstTally = -1: Application.StatusBar = False: Exit Function
                ' pass 1: vouchers already claimed - rows posted by this file, and rows already marked blue
                For r = FIRST_ROW To lastR
                    st = LCase$(CellText(ws.Cells(r, COL_STATUS).Value))
                    If (Left$(st, 6) = "posted" Or Left$(st, 8) = "in tally") And HasRowDate(ws, r) Then
                        bank = CellText(ws.Cells(r, COL_BANK).Value): If bank = "" Then bank = SettingValue("DefaultBank")
                        If StrComp(bank, b, vbTextCompare) = 0 Then
                            isIn = (Amt(ws.Cells(r, COL_IN).Value) > 0)
                            amount = IIf(isIn, Amt(ws.Cells(r, COL_IN).Value), Amt(ws.Cells(r, COL_OUT).Value))
                            k = 0
                            If Left$(st, 8) = "in tally" Then k = VoucherFromStatus(vc, CellText(ws.Cells(r, COL_STATUS).Value), isIn, amount)
                            If k = 0 Then k = FindVoucher(vc, RowDate(ws, r), isIn, amount, IIf(Left$(st, 6) = "posted", 0, win), CellText(ws.Cells(r, COL_NARR).Value))
                            If k > 0 Then MarkUsed vc, k
                        End If
                    End If
                Next r
                ' pass 2: rows not yet posted -> already in Tally?  Best pairs first (same narration, same ledger, closest date)
                nP = 0
                For r = FIRST_ROW To lastR
                    st = CellText(ws.Cells(r, COL_STATUS).Value)
                    If Not IsDone(st) And HasRowDate(ws, r) Then
                        bank = CellText(ws.Cells(r, COL_BANK).Value): If bank = "" Then bank = SettingValue("DefaultBank")
                        If StrComp(bank, b, vbTextCompare) = 0 Then
                            isIn = (Amt(ws.Cells(r, COL_IN).Value) > 0)
                            amount = IIf(isIn, Amt(ws.Cells(r, COL_IN).Value), Amt(ws.Cells(r, COL_OUT).Value))
                            narr = CellText(ws.Cells(r, COL_NARR).Value)
                            party = CellText(ws.Cells(r, COL_PARTY).Value)
                            partyKnown = (party <> "" And InStr(1, party, "suspense", vbTextCompare) = 0)
                            If amount > 0 Then
                                For k = 1 To vc.Count
                                    v = vc(k)
                                    If Not v(7) And v(1) = isIn And Abs(v(2) - amount) < 0.005 Then
                                        diff = Abs(DateDiff("d", v(0), RowDate(ws, r)))
                                        If diff <= win Then
                                            narrHit = (narr <> "" And CStr(v(6)) <> "" And (InStr(1, v(6), Left$(narr, 30), vbTextCompare) > 0 Or InStr(1, narr, Left$(CStr(v(6)), 30), vbTextCompare) > 0))
                                            ledHit = (party <> "" And StrComp(party, CStr(v(5)), vbTextCompare) = 0)
                                            ' a different day is accepted only if narration or ledger also agrees
                                            If diff = 0 Or narrHit Or ledHit Or Not partyKnown Then
                                                nP = nP + 1
                                                ReDim Preserve pS(1 To nP): ReDim Preserve pR(1 To nP): ReDim Preserve pK(1 To nP)
                                                pS(nP) = diff * 10 - IIf(narrHit, 1000, 0) - IIf(ledHit, 500, 0)
                                                pR(nP) = r: pK(nP) = k
                                            End If
                                        End If
                                    End If
                                Next k
                            End If
                        End If
                    End If
                Next r
                If nP > 0 Then
                    ReDim rowDone(FIRST_ROW To lastR): ReDim vchDone(1 To vc.Count)
                    Do
                        bestP = 0
                        For i = 1 To nP
                            If pR(i) > 0 Then
                                If Not rowDone(pR(i)) And Not vchDone(pK(i)) Then
                                    If bestP = 0 Then
                                        bestP = i
                                    ElseIf pS(i) < pS(bestP) Then
                                        bestP = i
                                    End If
                                End If
                            End If
                        Next i
                        If bestP = 0 Then Exit Do
                        r = pR(bestP): k = pK(bestP)
                        rowDone(r) = True: vchDone(k) = True: pR(bestP) = 0
                        v = vc(k)
                        SetStatus ws, r, "In Tally already: " & v(3) & IIf(v(4) <> "", " No. " & v(4), "") & " dt " & _
                            Format$(v(0), "dd-mm-yy") & IIf(v(5) <> "", " (" & v(5) & ")", "") & " - skipped", CLR_TALLY
                        nMark = nMark + 1
                    Loop
                End If
            End If
        End If
    Next b
    Application.StatusBar = False
    CheckRowsAgainstTally = nMark
End Function

'================ BUTTON: Check against Tally ================
Public Sub CheckAgainstTally()
    Dim n As Long, info As String
    ProgressShow "Checking against Tally", "Reading vouchers from Tally..."
    n = CheckRowsAgainstTally(False, info)
    If n < 0 Then ProgressHide: Exit Sub
    ProgressDone "Check finished"
    LogAct "Check against Tally", n & " row(s) already in Tally", n
    Notify n & " row(s) are already in Tally (blue) and will be skipped." & vbCrLf & _
        "Wrong match? Clear that row's Status cell.", "ok"
End Sub

'================ BUTTON: Find duplicates already in Tally ================
Public Sub FindDuplicatesInTally()
    Dim ws As Worksheet, rp As Worksheet, bank As Variant, dFrom As Variant, dTo As Variant, vc As Collection, info As String
    Dim i As Long, j As Long, vi As Variant, vj As Variant, grp() As Long, nG As Long, rr As Long, win As Long
    Dim r As Long, lastR As Long, mine As Boolean, nStmt As Long, gi As Long, members As String, m As Variant
    Dim nVch As Long, nMine As Long, extra As Long, d0 As Date, isIn As Boolean, amount As Double, sug As String
    Dim narrPosted As String, dRow As Date, nDel As Long, bws As Worksheet, br As Long, tb As Double, okB As Boolean
    Dim balNote As String, balOK As Boolean, rb As String

    Set ws = EntriesSheet()
    lastR = LastDataRow(ws)
    win = DupDays()
    bank = Application.InputBox("Bank ledger to check in Tally:", "Find duplicates", SettingValue("DefaultBank"), Type:=2)
    If VarType(bank) = vbBoolean Then Exit Sub
    dFrom = Empty: dTo = Empty
    For r = FIRST_ROW To lastR
        If HasRowDate(ws, r) Then
            If IsEmpty(dFrom) Then dFrom = RowDate(ws, r): dTo = dFrom
            If RowDate(ws, r) < dFrom Then dFrom = RowDate(ws, r)
            If RowDate(ws, r) > dTo Then dTo = RowDate(ws, r)
        End If
    Next r
    If IsEmpty(dFrom) Then dFrom = DateSerial(Year(Date), 4, 1): dTo = Date
    dFrom = Application.InputBox("From date (dd-mm-yyyy):", "Find duplicates", Format$(dFrom, "dd-mm-yyyy"), Type:=2)
    If VarType(dFrom) = vbBoolean Then Exit Sub
    dTo = Application.InputBox("To date (dd-mm-yyyy):", "Find duplicates", Format$(dTo, "dd-mm-yyyy"), Type:=2)
    If VarType(dTo) = vbBoolean Then Exit Sub
    dFrom = ParseDateCell(dFrom): dTo = ParseDateCell(dTo)
    If IsEmpty(dFrom) Or IsEmpty(dTo) Then MsgBox "Please enter dates like 01-04-2026.", vbExclamation: Exit Sub

    ' 1) balance first: if Tally matches the statement, same-amount entries are real
    Set bws = BankSheet()
    br = BBRow(bws, CStr(bank), False)
    If br > 0 Then
        If IsNumeric(bws.Cells(br, 3).Value) And CellText(bws.Cells(br, 3).Value) <> "" Then
            ProgressShow "Finding duplicates", "Comparing the balance with your statement..."
            BBCheckRow bws, br
            ProgressHide
            RefreshBankSummary
            If Left$(CellText(bws.Cells(br, 6).Value), 1) = ChrW(&H2714) Then
                balOK = True
                balNote = "Tally balance matches the statement on " & Format$(bws.Cells(br, 2).Value, "dd-mm-yyyy") & " *"
                If MsgBox("The balance of '" & bank & "' in Tally matches your bank statement on " & Format$(bws.Cells(br, 2).Value, "dd-mm-yyyy") & "." & vbCrLf & vbCrLf & _
                    "So entries with the same amount are real transactions - there is nothing to delete." & vbCrLf & vbCrLf & _
                    "Show same-amount entries anyway (for information only)?" & vbCrLf & vbCrLf & _
                    "* as per the bank statement provided by you", vbYesNo + 256 + vbInformation, "Find duplicates") <> vbYes Then
                    LogAct "Find duplicates", "'" & bank & "': balance matches the statement - nothing to delete", 0
                    Exit Sub
                End If
            ElseIf CellText(bws.Cells(br, 5).Value) <> "" Then
                balNote = "Tally differs from the statement by " & Format$(bws.Cells(br, 5).Value, "#,##0.00") & " on " & Format$(bws.Cells(br, 2).Value, "dd-mm-yyyy") & " *"
            End If
        End If
    End If
    If balNote = "" Then balNote = "Balance not compared (no statement balance for this bank yet)"

    ProgressShow "Finding duplicates", "Reading '" & bank & "' vouchers from Tally..."
    Set vc = FetchBankVouchers(CStr(bank), CDate(dFrom), CDate(dTo), False, info)
    If vc Is Nothing Then ProgressHide: Exit Sub
    ProgressDone "Vouchers read - looking for duplicates"

    ' group vouchers: same direction + amount, dates within the window
    ReDim grp(1 To vc.Count + 1)
    For i = 1 To vc.Count
        If grp(i) = 0 Then
            vi = vc(i)
            For j = i + 1 To vc.Count
                If grp(j) = 0 Then
                    vj = vc(j)
                    If vj(1) = vi(1) And Abs(vj(2) - vi(2)) < 0.005 And Abs(DateDiff("d", vj(0), vi(0))) <= win Then
                        If grp(i) = 0 Then nG = nG + 1: grp(i) = nG
                        grp(j) = grp(i)
                    End If
                End If
            Next j
        End If
    Next i

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("Tally Duplicates").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set rp = ThisWorkbook.Worksheets.Add(After:=ws)
    rp.Name = "Tally Duplicates"
    DarkIfNeeded rp
    rp.Range("A1").Value = "Possible duplicate vouchers in '" & bank & "'  (" & info & ")"
    rp.Range("A1").Font.Bold = True
    rp.Range("A2").Value = "Same amount and direction within " & win & " day(s), more copies in Tally than in the statement. 'Made by this file' = narration matches a row posted from the Entries sheet. Delete the suggested ones in Tally: Day Book > select voucher > Alt+D."
    rp.Range("A3").Value = "Bank balance: " & balNote & IIf(InStr(balNote, "*") > 0, "    (* based on the bank statement provided by you)", "")
    rp.Range("A3").Font.Italic = True
    rp.Range("A3").Font.Color = IIf(balOK, RGB(5, 150, 105), RGB(180, 83, 9))
    rp.Range("A4:J4").Value = Array("Group", "Date", "In/Out", "Amount", "Voucher Type", "Voucher No", "Other ledger", "Narration", "Made by this file?", "Suggestion")
    rp.Range("A4:J4").Font.Bold = True
    rp.Range("A4:J4").Interior.Color = 15652797
    rr = 5

    For gi = 1 To nG
        ' how many statement rows exist for this amount/direction around this date?
        members = ""
        For i = 1 To vc.Count
            If grp(i) = gi Then members = members & i & " "
        Next i
        vi = vc(CLng(Split(Trim$(members), " ")(0)))
        isIn = vi(1): amount = vi(2)
        nStmt = 0
        For r = FIRST_ROW To lastR
            rb = CellText(ws.Cells(r, COL_BANK).Value)
            If rb = "" Then rb = SettingValue("DefaultBank")
            If HasRowDate(ws, r) And StrComp(rb, CStr(bank), vbTextCompare) = 0 Then
                If Abs(DateDiff("d", RowDate(ws, r), vi(0))) <= win Then
                    If isIn And Abs(Amt(ws.Cells(r, COL_IN).Value) - amount) < 0.005 Then nStmt = nStmt + 1
                    If Not isIn And Abs(Amt(ws.Cells(r, COL_OUT).Value) - amount) < 0.005 Then nStmt = nStmt + 1
                End If
            End If
        Next r
        nVch = UBound(Split(Trim$(members), " ")) + 1
        If nStmt = 0 Then nStmt = 1           ' statement not on the sheet: assume one real transaction
        extra = nVch - nStmt
        If extra > 0 Then
            nDel = 0
            For Each m In Split(Trim$(members), " ")
                vi = vc(CLng(m))
                mine = False
                For r = FIRST_ROW To lastR
                    If Left$(LCase$(CellText(ws.Cells(r, COL_STATUS).Value)), 6) = "posted" Then
                        narrPosted = CellText(ws.Cells(r, COL_NARR).Value)
                        If narrPosted <> "" And InStr(1, vi(6), Left$(narrPosted, 30), vbTextCompare) > 0 Then mine = True: Exit For
                    End If
                Next r
                sug = ""
                If balOK Then
                    sug = "Balance matches - probably a real transaction"
                ElseIf mine And nDel < extra Then
                    sug = "DELETE (duplicate made by this file)": nDel = nDel + 1
                End If
                rp.Cells(rr, 1).Value = gi
                rp.Cells(rr, 2).Value = vi(0): rp.Cells(rr, 2).NumberFormat = "dd-mm-yyyy"
                rp.Cells(rr, 3).Value = IIf(vi(1), "In", "Out")
                rp.Cells(rr, 4).Value = vi(2): SetRupee rp.Cells(rr, 4)
                rp.Cells(rr, 5).Value = vi(3)
                rp.Cells(rr, 6).NumberFormat = "@": rp.Cells(rr, 6).Value = vi(4)
                rp.Cells(rr, 7).Value = vi(5)
                rp.Cells(rr, 8).Value = vi(6)
                rp.Cells(rr, 9).Value = IIf(mine, "Yes", "No")
                rp.Cells(rr, 10).Value = sug
                If Left$(sug, 6) = "DELETE" Then rp.Range(rp.Cells(rr, 1), rp.Cells(rr, 10)).Interior.Color = CLR_ERR
                rr = rr + 1
            Next m
            If nDel < extra And Not balOK Then
                rp.Cells(rr - 1, 10).Value = rp.Cells(rr - 1, 10).Value & IIf(rp.Cells(rr - 1, 10).Value = "", "", " / ") & _
                    "Group has " & extra - nDel & " more extra voucher(s) - check which one to delete"
            End If
            rr = rr + 1
        End If
    Next gi
    On Error Resume Next
    If rr > 5 Then GridBorders rp.Range(rp.Cells(4, 1), rp.Cells(rr - 1, 10))
    rp.Columns("A:G").AutoFit
    rp.Columns("H").ColumnWidth = 60
    rp.Columns("I:J").AutoFit
    QuietActivate rp
    On Error GoTo 0
    If rr = 5 Then
        rp.Range("A5").Value = "No duplicates found."
        MsgBox "No duplicate vouchers found in '" & bank & "' for this period.", vbInformation
    Else
        MsgBox "Possible duplicates are listed on the 'Tally Duplicates' sheet." & vbCrLf & _
            "Red rows = extra vouchers created by this file - delete those in Tally (Day Book > select > Alt+D).", vbExclamation
    End If
End Sub

'================ BUTTON: Learn from posted rows ================
' Learns every "Posted" row on the Entries sheet (rows further down = later = win), then optionally Tally itself
Public Sub LearnFromPostedRows()
    Dim ws As Worksheet, r As Long, lastR As Long, n As Long, nAc As Long, nSus As Long, nNo As Long
    Dim led As String, narr As String, bank As String, why As String, dt As Double, msg As String
    Set ws = EntriesSheet()
    lastR = LastDataRow(ws)
    ProgressShow "Learning from posted rows", "Reading rows..."
    LoadLedgerCache
    LearnLoad
    Application.ScreenUpdating = False
    For r = FIRST_ROW To lastR
        ProgressUpdate (r - FIRST_ROW + 1) / (lastR - FIRST_ROW + 1), "Learning row " & (r - FIRST_ROW + 1) & " of " & (lastR - FIRST_ROW + 1)
        If Left$(LCase$(CellText(ws.Cells(r, COL_STATUS).Value)), 6) = "posted" Then
            led = CellText(ws.Cells(r, COL_PARTY).Value)
            narr = CellText(ws.Cells(r, COL_NARR).Value)
            bank = CellText(ws.Cells(r, COL_BANK).Value)
            If HasRowDate(ws, r) Then dt = CDbl(RowDate(ws, r)) Else dt = CDbl(Now)
            If LearnAdd(narr, led, bank, dt, why) > 0 Then
                n = n + 1
                If AcctKeys(narr, bank) <> "" Then nAc = nAc + 1
            ElseIf why = "suspense" Then
                nSus = nSus + 1
            ElseIf why = "nowords" Then
                nNo = nNo + 1
            End If
        End If
    Next r
    LearnSave
    Application.ScreenUpdating = True
    ProgressDone "Learned " & n & " rows"
    LogAct "Learn from posted rows", n & " learned (" & nAc & " with account no.), " & nSus & " Suspense skipped, " & nNo & " without name/number", n
    msg = n & " posted row(s) learned  (" & nAc & " of them also by account number)."
    If nSus > 0 Then msg = msg & vbCrLf & nSus & " skipped - posted to Suspense (correct them with Change Ledger and they are learned)."
    If nNo > 0 Then msg = msg & vbCrLf & nNo & " skipped - the narration has no name or account number to learn."
    If MsgBox(msg & vbCrLf & vbCrLf & "Also learn from the entries already in Tally?" & vbCrLf & _
        "(All bank ledgers, last 13 months - this picks up corrections you made directly in Tally.)", _
        vbYesNo + vbQuestion, APP_NAME) = vbYes Then
        LearnFromTallyHistory
    Else
        Notify msg, "ok"
    End If
End Sub

' Learn from every bank ledger's vouchers in Tally (last 13 months)
Public Sub LearnFromTallyHistory()
    Dim i As Long, banks As String, b As Variant, nB As Long, nV As Long, nAll As Long, k As Long, nFail As Long
    LoadLedgerCache
    For i = 1 To gCount
        If IsBankGroup(gGroups(i)) Then banks = banks & gNames(i) & "|"
    Next i
    If banks = "" Then MsgBox "No bank ledgers found on the Ledgers sheet - click Load ledgers first.", vbExclamation, APP_NAME: Exit Sub
    nAll = UBound(Split(banks, "|"))
    ProgressShow "Learning from Tally", "Connecting to Tally..."
    LearnLoad
    For Each b In Split(banks, "|")
        If b <> "" Then
            k = k + 1
            ProgressUpdate k / nAll, "Reading " & b & "  (" & k & " of " & nAll & ")"
            If LearnFromTally(CStr(b), DateAdd("m", -13, Date), Date, nV) Then
                nB = nB + 1
            Else
                nFail = nFail + 1
                If nB = 0 And nFail >= 2 Then Exit For        ' Tally not answering
            End If
        End If
    Next b
    LearnSave
    ProgressDone "Learning finished"
    If nB = 0 Then
        ConnectionHelp "Could not read vouchers from Tally"
        Exit Sub
    End If
    LogAct "Learn from Tally", nV & " vouchers from " & nB & " bank ledger(s)", nV
    Notify nV & " entries learned from Tally (" & nB & " bank ledger" & IIf(nB = 1, "", "s") & ")." & vbCrLf & _
        "Re-match ledgers to apply it to rows not yet posted.", "ok"
End Sub

'==============================================================
'  CHANGE LEDGER  (alter the same voucher in Tally)
'==============================================================
' "Change Ledger" sheet: B2 ledger, E2 from date, G2 to date; data from row 6

Private Function FxSheet() As Worksheet
    Set FxSheet = ThisWorkbook.Worksheets(FX_SHEET)
End Function

Private Function SuspenseLedgerName() As String
    Dim ws As Worksheet, lastR As Long, r As Long, nm As String, g As String
    Set ws = ThisWorkbook.Worksheets("Ledgers")
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastR
        nm = CellText(ws.Cells(r, 1).Value)
        g = LCase$(CellText(ws.Cells(r, 2).Value) & "|" & CellText(ws.Cells(r, 3).Value))
        If nm <> "" And (InStr(g, "suspense") > 0 Or InStr(1, nm, "suspense", vbTextCompare) > 0) Then SuspenseLedgerName = nm: Exit Function
    Next r
End Function

Private Function FyStart() As Date
    Dim v As Variant
    v = ParseDateCell(SettingValue("SuspFrom"))
    If IsEmpty(v) Then
        If Month(Date) >= 4 Then v = DateSerial(Year(Date), 4, 1) Else v = DateSerial(Year(Date) - 1, 4, 1)
    End If
    FyStart = v
End Function

Private Function LedgerExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet, v As Variant
    Set ws = ThisWorkbook.Worksheets("Ledgers")
    v = Application.Match(nm, ws.Range("A:A"), 0)
    LedgerExists = Not IsError(v)
End Function

Private Function LedgerGroupFull(ByVal nm As String) As String
    Dim ws As Worksheet, v As Variant
    Set ws = ThisWorkbook.Worksheets("Ledgers")
    v = Application.Match(nm, ws.Range("A:A"), 0)
    If Not IsError(v) Then LedgerGroupFull = LCase$(CellText(ws.Cells(v, 2).Value) & "|" & CellText(ws.Cells(v, 3).Value))
End Function

Private Function IsBankOrCashGroup(ByVal g As String) As Boolean
    IsBankOrCashGroup = IsBankGroup(g) Or InStr(g, "cash-in-hand") > 0 Or InStr(g, "cash in hand") > 0
End Function

' Loads all vouchers of the ledger chosen at the top of the sheet. Returns a one-line summary.
Private Function LoadFixCore(ByVal silent As Boolean) As String
    Dim ws As Worksheet, led As String, dFrom As Variant, dTo As Variant, vc As Collection, info As String
    Dim k As Long, v As Variant, keep As String, r As Long, lastR As Long, g As String, n As Long, nMulti As Long
    Dim p1 As Long, p2 As Long

    Set ws = FxSheet()
    led = CellText(ws.Range("B2").Value)
    If led = "" Then
        led = SuspenseLedgerName()
        If led = "" Then LoadFixCore = "Choose a ledger at the top of the Change Ledger sheet.": GoTo done
        ws.Range("B2").Value = led
    End If
    If Not LedgerExists(led) Then LoadFixCore = "'" & led & "' is not in the ledger list - pick it from the dropdown (or click Load Ledgers).": GoTo done
    If IsBankOrCashGroup(LedgerGroupFull(led)) Then LoadFixCore = "'" & led & "' is a Bank/Cash ledger - those are not changed here. Choose the other ledger (party, expense, Suspense...).": GoTo done
    dFrom = ParseDateCell(ws.Range("E2").Value)
    If IsEmpty(dFrom) Then dFrom = FyStart(): ws.Range("E2").Value = dFrom
    dTo = ParseDateCell(ws.Range("G2").Value)
    If IsEmpty(dTo) Then dTo = Date: ws.Range("G2").Value = dTo
    If dTo < dFrom Then LoadFixCore = "'To' date is before 'From' date.": GoTo done

    ' remember ledgers already picked (by voucher id)
    keep = Chr(1)
    lastR = ws.Cells(ws.Rows.Count, S_DATE).End(xlUp).Row
    For r = S_FIRST To lastR
        g = CellText(ws.Cells(r, S_GUID).Value)
        If g <> "" And CellText(ws.Cells(r, S_NEW).Value) <> "" Then keep = keep & g & Chr(2) & CellText(ws.Cells(r, S_NEW).Value) & Chr(1)
    Next r

    Application.StatusBar = "Reading '" & led & "' from Tally..."
    Set vc = FetchBankVouchers(led, CDate(dFrom), CDate(dTo), silent, info)
    Application.StatusBar = False
    If vc Is Nothing Then LoadFixCore = "Entries not loaded: " & info: GoTo done

    Application.ScreenUpdating = False
    If lastR >= S_FIRST Then
        ws.Range(ws.Cells(S_FIRST, 1), ws.Cells(lastR, S_GUID)).ClearContents
        ws.Range(ws.Cells(S_FIRST, 1), ws.Cells(lastR, S_STATUS)).Interior.ColorIndex = xlNone
        ws.Range(ws.Cells(S_FIRST, S_NEW), ws.Cells(lastR, S_NEW)).Interior.Color = 13431551
    End If
    r = S_FIRST
    For k = 1 To vc.Count
        v = vc(k)
        ws.Cells(r, S_DATE).Value = v(0): ws.Cells(r, S_DATE).NumberFormat = "dd-mm-yyyy"
        ws.Cells(r, S_TYPE).Value = v(3)
        ws.Cells(r, S_NO).NumberFormat = "@": ws.Cells(r, S_NO).Value = v(4)
        ws.Cells(r, S_OTHER).Value = v(5)
        If v(1) Then ws.Cells(r, S_DR).Value = v(2) Else ws.Cells(r, S_CR).Value = v(2)
        ws.Cells(r, S_NARR).Value = v(6)
        ws.Cells(r, S_CUR).Value = led
        ws.Cells(r, S_GUID).Value = v(8)
        p1 = InStr(keep, Chr(1) & v(8) & Chr(2))
        If p1 > 0 And CStr(v(8)) <> "" Then
            p1 = p1 + Len(v(8)) + 2
            p2 = InStr(p1, keep, Chr(1))
            ws.Cells(r, S_NEW).Value = Mid$(keep, p1, p2 - p1)
        End If
        If v(9) > 1 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, S_STATUS)).Interior.Color = CLR_TALLYONLY: ws.Range(ws.Cells(r, 1), ws.Cells(r, S_STATUS)).Font.Color = RGB(30, 41, 59)
            ws.Cells(r, S_STATUS).Value = "This voucher has " & v(9) & " lines of '" & led & "' - change it in Tally"
            nMulti = nMulti + 1
        End If
        FitRow ws, r
        r = r + 1: n = n + 1
    Next k
    If lastR >= r Then ws.Range(ws.Rows(r), ws.Rows(lastR)).RowHeight = 16
    Application.ScreenUpdating = True
    LoadFixCore = n & " entr" & IIf(n = 1, "y", "ies") & " of '" & led & "' from " & Format$(dFrom, "dd-mm-yyyy") & " to " & Format$(dTo, "dd-mm-yyyy") & _
        IIf(nMulti > 0, " (" & nMulti & " orange = change in Tally)", "") & "."
done:
    Application.ScreenUpdating = True
    If Not silent Then Notify LoadFixCore, IIf(InStr(LoadFixCore, "entr") > 0, "ok", "warn")
End Function

' Called by Load Ledgers: refresh the sheet for the chosen ledger (Suspense by default)
Private Function LoadSuspenseCore(ByVal silent As Boolean) As String
    LoadSuspenseCore = LoadFixCore(silent)
End Function

'================ BUTTONS (Change Ledger sheet) ================
Public Sub LoadLedgerEntries()
    Dim msg As String
    QuietActivate FxSheet()
    ProgressShow "Loading entries", "Reading '" & CellText(FxSheet().Range("B2").Value) & "' from Tally..."
    msg = LoadFixCore(True)
    ProgressDone "Entries loaded"
    If InStr(msg, " entr") > 0 And InStr(msg, " from ") > 0 Then
        Notify msg & vbCrLf & "Pick the right ledger in the yellow column, then Update in Tally.", "ok"
    Else
        Notify msg, "warn"
    End If
    QuietActivate FxSheet()
End Sub

Public Sub LoadSuspenseEntries()
    Dim s As String, msg As String
    s = SuspenseLedgerName()
    If s = "" Then MsgBox "No Suspense ledger found in the ledger list (click Load Ledgers).", vbExclamation: Exit Sub
    FxSheet().Range("B2").Value = s
    QuietActivate FxSheet()
    ProgressShow "Loading Suspense entries", "Reading '" & s & "' from Tally..."
    msg = LoadFixCore(True)
    ProgressDone "Entries loaded"
    Notify msg, "ok"
    QuietActivate FxSheet()
End Sub

' Full vouchers of one day from Tally (Day Book export, the same format Tally itself imports)
Private Function DayBookDoc(ByVal d As Date, ByRef errTxt As String) As Object
    Dim req As String, resp As String, co As String, sv As String, re As Object, doc As Object
    co = SettingValue("CompanyName")
    sv = "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    If co <> "" Then sv = sv & "<SVCURRENTCOMPANY>" & XmlEsc(co) & "</SVCURRENTCOMPANY>"
    sv = sv & "<SVFROMDATE>" & TallyDateStr(d) & "</SVFROMDATE><SVTODATE>" & TallyDateStr(d) & "</SVTODATE>" & _
              "<SVCURRENTDATE>" & TallyDateStr(d) & "</SVCURRENTDATE>"
    req = "<ENVELOPE><HEADER><TALLYREQUEST>Export Data</TALLYREQUEST></HEADER><BODY><EXPORTDATA><REQUESTDESC>" & _
          "<REPORTNAME>Day Book</REPORTNAME><STATICVARIABLES>" & sv & "</STATICVARIABLES></REQUESTDESC></EXPORTDATA></BODY></ENVELOPE>"
    If Not PostXml(req, resp) Then errTxt = "Tally not reachable: " & resp: Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "&#([0-8]|1[124-9]|2[0-9]|3[01]);"
    resp = re.Replace(resp, "")
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.preserveWhiteSpace = True
    If Not doc.LoadXML(resp) Then
        errTxt = "Day Book reply could not be read"
        SaveDebug "Day Book " & Format$(d, "dd-mm-yyyy"), resp
        Exit Function
    End If
    gLastDayBook = resp
    Set DayBookDoc = doc
End Function

Private Function NodeText(ByVal n As Object, ByVal tag As String) As String
    Dim nd As Object
    Set nd = n.SelectSingleNode(tag)
    If Not nd Is Nothing Then NodeText = Trim$(nd.Text)
End Function

' Finds the voucher: by its Tally id (GUID), else by voucher number + type + date + amount of this ledger line
Private Function FindVoucherForRow(ByVal doc As Object, ByVal guid As String, ByVal vno As String, ByVal vtype As String, _
        ByVal d As Date, ByVal ledger As String, ByVal amount As Double) As Object
    Dim vn As Object, g As String, dv As Variant, ln As Object, a As Double, ok As Boolean
    For Each vn In doc.SelectNodes("//VOUCHER")
        g = NodeText(vn, "GUID")
        If g = "" Then g = CStr(vn.getAttribute("REMOTEID") & "")
        If guid <> "" And StrComp(g, guid, vbTextCompare) = 0 Then Set FindVoucherForRow = vn: Exit Function
    Next vn
    If vno = "" Then Exit Function
    For Each vn In doc.SelectNodes("//VOUCHER")
        If NodeText(vn, "VOUCHERNUMBER") = vno And StrComp(NodeText(vn, "VOUCHERTYPENAME"), vtype, vbTextCompare) = 0 Then
            dv = ParseTallyDate(NodeText(vn, "DATE"))
            If Not IsEmpty(dv) Then
                If dv = d Then
                    ok = False
                    For Each ln In vn.SelectNodes("ALLLEDGERENTRIES.LIST|LEDGERENTRIES.LIST")
                        If StrComp(NodeText(ln, "LEDGERNAME"), ledger, vbTextCompare) = 0 Then
                            a = Abs(Val(Replace(NodeText(ln, "AMOUNT"), ",", "")))
                            If Abs(a - amount) < 0.005 Then ok = True
                        End If
                    Next ln
                    If ok Then Set FindVoucherForRow = vn: Exit Function
                End If
            End If
        End If
    Next vn
End Function

' Also correct the matching posted row on the Entries sheet
Private Sub FixEntriesRow(ByVal d As Date, ByVal amount As Double, ByVal oldL As String, ByVal newL As String)
    Dim ws As Worksheet, r As Long, a As Double
    Set ws = EntriesSheet()
    For r = FIRST_ROW To LastDataRow(ws)
        If HasRowDate(ws, r) Then
            If RowDate(ws, r) = d And StrComp(CellText(ws.Cells(r, COL_PARTY).Value), oldL, vbTextCompare) = 0 Then
                a = Amt(ws.Cells(r, COL_OUT).Value) + Amt(ws.Cells(r, COL_IN).Value)
                If Abs(a - amount) < 0.005 Then
                    ws.Cells(r, COL_PARTY).Value = newL
                    ws.Cells(r, COL_PARTY).Interior.Color = CLR_OK
                    ws.Cells(r, COL_NOTE).Value = "Changed from '" & oldL & "' " & Format$(Now, "dd-mmm-yy")
                    Exit Sub
                End If
            End If
        End If
    Next r
End Sub

Private Sub SetFxStatus(ws As Worksheet, ByVal r As Long, ByVal msg As String, ByVal clr As Long)
    ws.Cells(r, S_STATUS).Value = msg
    ws.Cells(r, S_STATUS).Interior.Color = clr
    ws.Cells(r, S_STATUS).Font.Color = RGB(30, 41, 59)
    FitRow ws, r
End Sub

'================ BUTTON (Change Ledger sheet): Update in Tally ================
Public Sub UpdateLedgerInTally()
    Dim ws As Worksheet, lastR As Long, r As Long, nTodo As Long, newL As String, oldL As String, guid As String
    Dim d As Date, doc As Object, errTxt As String, vn As Object, ln As Object, lines As Object, hit As Object, hitLine As Object
    Dim nHit2 As Long, hit2 As Object, hitLine2 As Object
    Dim nHit As Long, nd As Object, body As String, resp As String, created As Long, altered As Long
    Dim lineErr As String, nOk As Long, nErr As Long, co As String, sv As String, g As String, amount As Double
    Dim cacheDay As Date, haveCache As Boolean, st As String, agst As Boolean, nVch As Long, vno As String

    Set ws = FxSheet()
    lastR = ws.Cells(ws.Rows.Count, S_DATE).End(xlUp).Row
    For r = S_FIRST To lastR
        st = LCase$(CellText(ws.Cells(r, S_STATUS).Value))
        If CellText(ws.Cells(r, S_NEW).Value) <> "" And Left$(st, 7) <> "updated" Then nTodo = nTodo + 1
    Next r
    If nTodo = 0 Then MsgBox "Pick the correct ledger in the yellow 'Correct Ledger' column for the entries you want to change, then click Update in Tally.", vbInformation: Exit Sub
    If MsgBox(nTodo & " entr" & IIf(nTodo = 1, "y", "ies") & " will be changed in Tally." & vbCrLf & vbCrLf & _
        "The SAME voucher is altered (same number, date, amount, bank details) - only the ledger is replaced." & vbCrLf & _
        "Tip: take a Tally backup before the first use." & vbCrLf & vbCrLf & "Continue?", vbYesNo + vbQuestion, "Update in Tally") <> vbYes Then Exit Sub

    co = SettingValue("CompanyName")
    If co <> "" Then sv = "<STATICVARIABLES><SVCURRENTCOMPANY>" & XmlEsc(co) & "</SVCURRENTCOMPANY></STATICVARIABLES>"
    ProgressShow "Updating Tally", "Changing " & nTodo & " entr" & IIf(nTodo = 1, "y", "ies") & "..."

    For r = S_FIRST To lastR
        ProgressUpdate (r - S_FIRST + 1) / (lastR - S_FIRST + 1), "Updating... (" & nOk & " done)"
        newL = CellText(ws.Cells(r, S_NEW).Value)
        st = LCase$(CellText(ws.Cells(r, S_STATUS).Value))
        If newL <> "" And Left$(st, 7) <> "updated" Then
            oldL = CellText(ws.Cells(r, S_CUR).Value)
            guid = CellText(ws.Cells(r, S_GUID).Value)
            If IsEmpty(ParseDateCell(ws.Cells(r, S_DATE).Value)) Then GoTo nextRow
            d = ParseDateCell(ws.Cells(r, S_DATE).Value)
            amount = Amt(ws.Cells(r, S_DR).Value) + Amt(ws.Cells(r, S_CR).Value)
            Application.StatusBar = "Updating row " & r & "..."
            g = LedgerGroupFull(newL)
            If InStr(st, "lines of") > 0 Then
                SetFxStatus ws, r, CellText(ws.Cells(r, S_STATUS).Value), CLR_TALLYONLY: nErr = nErr + 1
            ElseIf Not LedgerExists(newL) Then
                SetFxStatus ws, r, "Ledger '" & newL & "' not in the ledger list - pick from the dropdown", CLR_ERR: nErr = nErr + 1
            ElseIf StrComp(newL, oldL, vbTextCompare) = 0 Then
                SetFxStatus ws, r, "Correct Ledger is the same as the current one", CLR_ERR: nErr = nErr + 1
            ElseIf IsBankOrCashGroup(g) Then
                SetFxStatus ws, r, "Bank/Cash ledger needs a Contra voucher - change this one in Tally", CLR_TALLYONLY: nErr = nErr + 1
            ElseIf guid = "" Then
                SetFxStatus ws, r, "Voucher id missing - click Load Entries again", CLR_ERR: nErr = nErr + 1
            Else
                If Not haveCache Or cacheDay <> d Then
                    Set doc = DayBookDoc(d, errTxt)
                    If doc Is Nothing Then
                        Application.StatusBar = False
                        If InStr(errTxt, "not reachable") > 0 Then ProgressHide: ConnectionHelp errTxt: Exit Sub
                        SetFxStatus ws, r, errTxt & " (see TallyDebug sheet)", CLR_ERR: nErr = nErr + 1
                        GoTo nextRow
                    End If
                    cacheDay = d: haveCache = True
                End If
                Set vn = FindVoucherForRow(doc, guid, CellText(ws.Cells(r, S_NO).Value), CellText(ws.Cells(r, S_TYPE).Value), d, oldL, amount)
                If vn Is Nothing Then
                    nVch = doc.SelectNodes("//VOUCHER").Length
                    SaveDebug "Day Book " & Format$(d, "dd-mm-yyyy") & " - looking for " & CellText(ws.Cells(r, S_TYPE).Value) & " No. " & _
                        CellText(ws.Cells(r, S_NO).Value) & " id " & guid, gLastDayBook
                    SetFxStatus ws, r, "Voucher not found: Tally's Day Book returned " & nVch & " voucher(s) for " & Format$(d, "dd-mm-yyyy") & _
                        " - reply saved on TallyDebug sheet, please send a screenshot", CLR_ERR: nErr = nErr + 1
                    GoTo nextRow
                End If
                ' invoices with stock items are not changed from here
                If Not vn.SelectSingleNode("ALLINVENTORYENTRIES.LIST/STOCKITEMNAME|INVENTORYENTRIES.LIST/STOCKITEMNAME") Is Nothing Then
                    SetFxStatus ws, r, "Invoice with stock items - change it in Tally", CLR_TALLYONLY: nErr = nErr + 1
                    GoTo nextRow
                End If
                ' invoice-style vouchers carry the same lines in two lists - count each list on its own
                nHit = 0: nHit2 = 0: Set hit2 = Nothing: Set hitLine2 = Nothing
                For Each ln In vn.SelectNodes("ALLLEDGERENTRIES.LIST")
                    Set nd = ln.SelectSingleNode("LEDGERNAME")
                    If Not nd Is Nothing Then
                        If StrComp(Trim$(nd.Text), oldL, vbTextCompare) = 0 Then nHit = nHit + 1: Set hit = nd: Set hitLine = ln
                    End If
                Next ln
                For Each ln In vn.SelectNodes("LEDGERENTRIES.LIST")
                    Set nd = ln.SelectSingleNode("LEDGERNAME")
                    If Not nd Is Nothing Then
                        If StrComp(Trim$(nd.Text), oldL, vbTextCompare) = 0 Then nHit2 = nHit2 + 1: Set hit2 = nd: Set hitLine2 = ln
                    End If
                Next ln
                If nHit = 0 And nHit2 > 0 Then nHit = nHit2: Set hit = hit2: Set hitLine = hitLine2: nHit2 = 0: Set hit2 = Nothing: Set hitLine2 = Nothing
                If nHit2 > nHit Then nHit = nHit2
                If nHit <> 1 Then
                    ws.Range(ws.Cells(r, 1), ws.Cells(r, S_STATUS)).Interior.Color = CLR_TALLYONLY: ws.Range(ws.Cells(r, 1), ws.Cells(r, S_STATUS)).Font.Color = RGB(30, 41, 59)
                    SetFxStatus ws, r, "This voucher has " & nHit & " lines of '" & oldL & "' - change it in Tally", CLR_TALLYONLY: nErr = nErr + 1
                    GoTo nextRow
                End If
                ' bill-wise "Agst Ref" belongs to the old party - do not move it
                agst = False
                For Each nd In hitLine.SelectNodes("BILLALLOCATIONS.LIST/BILLTYPE")
                    If InStr(1, nd.Text, "Agst", vbTextCompare) > 0 Then agst = True
                Next nd
                If Not hitLine2 Is Nothing Then
                    For Each nd In hitLine2.SelectNodes("BILLALLOCATIONS.LIST/BILLTYPE")
                        If InStr(1, nd.Text, "Agst", vbTextCompare) > 0 Then agst = True
                    Next nd
                End If
                If agst Then
                    SetFxStatus ws, r, "Settled against a bill (Agst Ref) of '" & oldL & "' - change it in Tally", CLR_TALLYONLY: nErr = nErr + 1
                    GoTo nextRow
                End If
                ' replace the ledger everywhere it names this party
                hit.Text = newL
                If Not hit2 Is Nothing Then hit2.Text = newL
                For Each nd In vn.SelectNodes(".//PARTYLEDGERNAME|.//PAYMENTFAVOURING|.//BANKPARTYNAME")
                    If StrComp(Trim$(nd.Text), oldL, vbTextCompare) = 0 Then nd.Text = newL
                Next nd
                ' Tally identifies the voucher to alter by type + number + date (TAGNAME / TAGVALUE)
                vno = NodeText(vn, "VOUCHERNUMBER")
                If vno = "" Then
                    SetFxStatus ws, r, "Voucher has no number - change it in Tally", CLR_TALLYONLY: nErr = nErr + 1
                    GoTo nextRow
                End If
                On Error Resume Next
                vn.removeAttribute "REMOTEID"
                vn.removeAttribute "VCHKEY"
                On Error GoTo 0
                vn.setAttribute "DATE", TallyDateStr(d)
                vn.setAttribute "TAGNAME", "Voucher Number"
                vn.setAttribute "TAGVALUE", vno
                vn.setAttribute "VCHTYPE", NodeText(vn, "VOUCHERTYPENAME")
                vn.setAttribute "ACTION", "Alter"
                body = "<ENVELOPE><HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER><BODY><IMPORTDATA><REQUESTDESC>" & _
                       "<REPORTNAME>Vouchers</REPORTNAME>" & sv & "</REQUESTDESC><REQUESTDATA>" & _
                       "<TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vn.xml & "</TALLYMESSAGE></REQUESTDATA></IMPORTDATA></BODY></ENVELOPE>"
                If Not PostXml(body, resp) Then Application.StatusBar = False: ProgressHide: ConnectionHelp resp: Exit Sub
                created = Val(GetTag(resp, "CREATED"))
                altered = Val(GetTag(resp, "ALTERED"))
                lineErr = AllTags(resp, "LINEERROR")
                haveCache = False
                If created > 0 Then
                    SetFxStatus ws, r, "STOPPED: Tally created a NEW voucher instead of changing this one. Open Day Book " & _
                        Format$(d, "dd-mm-yyyy") & " and delete the extra voucher, then send me a screenshot.", CLR_ERR
                    SaveDebug "Alter reply (created instead of altered)", resp
                    Application.StatusBar = False
                    ProgressHide
                    MsgBox "Tally created a new voucher instead of changing the existing one (row " & r & ")." & vbCrLf & _
                        "Updating was stopped so this cannot repeat. Please delete the extra voucher in Day Book " & Format$(d, "dd-mm-yyyy") & ".", vbCritical
                    Exit Sub
                ElseIf altered > 0 And lineErr = "" Then
                    SetFxStatus ws, r, "Updated in Tally " & Format$(Now, "dd-mmm-yy hh:nn") & " (was '" & oldL & "')", CLR_OK: nOk = nOk + 1
                    ws.Cells(r, S_CUR).Value = newL
                    If IsBankOrCashGroup(LedgerGroupFull(CellText(ws.Cells(r, S_OTHER).Value))) Then LearnFromRow CellText(ws.Cells(r, S_NARR).Value), newL, CellText(ws.Cells(r, S_OTHER).Value), CDbl(d)
                    FixEntriesRow d, amount, oldL, newL
                Else
                    If lineErr = "" Then lineErr = Left$(Replace(Replace(resp, vbCr, " "), vbLf, " "), 250)
                    SetFxStatus ws, r, "Tally error: " & lineErr, CLR_ERR: nErr = nErr + 1
                End If
            End If
        End If
nextRow:
    Next r
    Application.StatusBar = False
    ProgressDone "Tally updated"
    LogAct "Change ledger in Tally", nOk & " updated, " & nErr & " not updated", nOk
    Notify "Updated in Tally: " & nOk & vbCrLf & "Not updated: " & nErr & IIf(nErr > 0, "  (see Status; orange = change in Tally)", ""), IIf(nErr > 0, "warn", "ok")
End Sub

'==============================================================
'  POSTIFY UI LAYER  (navigation, ribbon, progress, toasts, log)
'==============================================================
Private Function AutoReconcile() As Boolean
    AutoReconcile = (LCase$(SettingValue("AutoRecon")) <> "no")
End Function

Public Function AppVersion() As String
    AppVersion = APP_VERSION
End Function

Private Function AppData() As Worksheet
    On Error Resume Next
    Set AppData = ThisWorkbook.Worksheets("AppData")
End Function

Private Sub SetApp(ByVal cellAddr As String, ByVal v As Variant)
    On Error Resume Next
    AppData().Range(cellAddr).Value = v
End Sub

' Runs when the file opens (called from ThisWorkbook) and right after an update
Public Sub AppStartup()
    Dim was As Boolean
    On Error Resume Next
    SetApp "B7", APP_VERSION
    was = gNoAnim: gNoAnim = True
    QuietActivate ThisWorkbook.Worksheets("Home")
    OnSheetActivate "Home"
    gNoAnim = was
    ThisWorkbook.Worksheets("Home").Range("A1").Select
    StartupWithSplash
    If CStr(AppData().Range("B8").Value) <> APP_VERSION Then
        SetApp "A11", "Last updated"
        SetApp "B11", Date
        Application.OnTime Now + TimeSerial(0, 0, 1), "TallyLink.ShowWhatsNew"
    End If
    Application.OnTime Now + TimeSerial(0, 0, 4), "Updater.AutoCheck"
    If Not FormExists("PfySend") Or CStr(AppData().Range("B18").Value) <> FORMS_VER Then
        Application.OnTime Now + TimeSerial(0, 0, 7), "TallyLink.BuildFormsQuiet"
    End If
End Sub

' Clean look on dashboard-style sheets + a soft fade-in when a sheet tab is clicked
Public Sub OnSheetActivate(ByVal shName As String)
    On Error Resume Next
    Select Case shName
        Case "Home", "Help", "Activity Log"
            ActiveWindow.DisplayHeadings = False
        Case Else
            ActiveWindow.DisplayHeadings = True
    End Select
    ActiveWindow.DisplayGridlines = False
    If shName = "Home" Then FitHome
    If gNoAnim Or gProgOn Then Exit Sub
    Select Case shName
        Case "Home", "Entries", "Change Ledger", "Activity Log", "Settings", "Ledgers", "Rules", "Learned", "Help", BB_SHEET, "Update History"
            SheetFadeIn ActiveSheet
    End Select
End Sub

' Activate a sheet from code without the tab animation
Private Sub QuietActivate(ws As Object)
    Dim was As Boolean
    was = gNoAnim
    gNoAnim = True
    On Error Resume Next
    ws.Activate
    gNoAnim = was
End Sub

' Home dashboard fills the window width on any screen
Private Sub FitHome()
    Dim ws As Worksheet, ev As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Home")
    If Not ActiveSheet Is ws Then Exit Sub
    ev = Application.EnableEvents
    Application.EnableEvents = False
    ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    ws.Range("A1:AS1").Select
    ActiveWindow.Zoom = True
    If ActiveWindow.Zoom > 175 Then ActiveWindow.Zoom = 175
    If ActiveWindow.Zoom < 70 Then ActiveWindow.Zoom = 70
    ActiveWindow.ScrollColumn = 1: ActiveWindow.ScrollRow = 1
    ws.Range("A1").Select
    Application.EnableEvents = ev
End Sub

' What the user really sees (sheet points), inside frozen panes, headings and scroll bars
Private Sub ViewRect(ByRef l As Single, ByRef t As Single, ByRef w As Single, ByRef h As Single)
    Dim vr As Object, z As Single, uw As Single, uh As Single, fw As Single, fh As Single
    On Error Resume Next
    With ActiveWindow
        Set vr = .Panes(.Panes.Count).VisibleRange
        If vr Is Nothing Then Set vr = .VisibleRange
        l = vr.Left: t = vr.Top: w = vr.Width: h = vr.Height
        z = .Zoom / 100
        If z <= 0 Then z = 1
        If .FreezePanes Or .Panes.Count > 1 Then
            If .SplitColumn > 0 Then fw = .Panes(1).VisibleRange.Width
            If .SplitRow > 0 Then fh = .Panes(1).VisibleRange.Height
        End If
        uw = .UsableWidth / z - fw - IIf(.DisplayHeadings, 28, 0) - 22
        uh = .UsableHeight / z - fh - IIf(.DisplayHeadings, 16, 0) - 22
        If uw > 120 And uw < w Then w = uw
        If uh > 120 And uh < h Then h = uh
    End With
End Sub

' Everything on screen incl. frozen rows/columns (for full-screen covers)
Private Sub CoverRect(ByRef l As Single, ByRef t As Single, ByRef w As Single, ByRef h As Single)
    Dim a As Object, z As Object
    On Error Resume Next
    With ActiveWindow
        Set a = .Panes(1).VisibleRange
        Set z = .Panes(.Panes.Count).VisibleRange
    End With
    If a Is Nothing Or z Is Nothing Then
        Set z = ActiveWindow.VisibleRange: Set a = z
    End If
    l = a.Left: t = a.Top
    w = z.Left + z.Width - l + 40
    h = z.Top + z.Height - t + 40
End Sub

' Sheet tab clicked: a thin teal line sweeps across the top (no cover, so no white flash)
Private Sub SheetFadeIn(ws As Object)
    Dim bar As Object, i As Long, l As Single, t As Single, w As Single, h As Single
    On Error GoTo fin
    CoverRect l, t, w, h
    w = w - 40
    Set bar = NewBox(ws, "trn_sweep", 1, l, t, 4, 3, RGB(20, 184, 166))
    Application.ScreenUpdating = True
    For i = 1 To 7
        bar.Width = w * i / 7
        TransitionPause 0.018
    Next i
    For i = 1 To 4
        bar.Fill.Transparency = i / 4
        TransitionPause 0.02
    Next i
fin:
    DeleteShapes "trn_"
End Sub

' Make a group of shapes (fill + text) more / less transparent
Private Sub ShapesAlpha(ByVal prefix As String, ws As Object, ByVal a As Single)
    Dim shp As Object
    On Error Resume Next
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(prefix)) = prefix Then
            If shp.Fill.Visible Then
                If shp.Name <> prefix & "bg" Then shp.Fill.Transparency = a
            End If
            If shp.TextFrame2.HasText Then shp.TextFrame2.TextRange.Font.Fill.Transparency = a
        End If
    Next shp
End Sub

' Opening animation on the dashboard; the start-up work runs while it plays
Private Sub StartupWithSplash()
    Dim ws As Object, shp As Object, bar As Object, stp As Object, i As Long, k As Long
    Dim l As Single, t As Single, w As Single, h As Single, vl As Single, vt As Single, vw As Single, vh As Single
    Dim cx As Single, cy As Single, tw As Single, msgs As Variant, t0 As Single
    On Error GoTo plain
    Set ws = ActiveSheet
    CoverRect l, t, w, h
    ViewRect vl, vt, vw, vh
    cx = vl + vw / 2: cy = vt + vh / 2 - 30
    Set shp = NewBox(ws, "spl_bg", 1, l, t, w, h, RGB(11, 31, 58))
    Set shp = NewBox(ws, "spl_logo", 5, cx - 30, cy - 104, 60, 60, RGB(20, 184, 166))
    shp.Adjustments.Item(1) = 0.22
    StyleText shp, "P", 30, RGB(255, 255, 255), True, 2
    shp.TextFrame2.VerticalAnchor = 3
    Set shp = NewText(ws, "spl_title", cx - 220, cy - 32, 440, 36)
    StyleText shp, APP_NAME, 26, RGB(255, 255, 255), True, 2
    Set shp = NewText(ws, "spl_tag", cx - 220, cy + 8, 440, 20)
    StyleText shp, "Bank statements to Tally, in minutes", 11, RGB(148, 180, 210), False, 2
    tw = 260
    Set shp = NewBox(ws, "spl_track", 5, cx - tw / 2, cy + 50, tw, 6, RGB(30, 58, 95))
    shp.Adjustments.Item(1) = 0.5
    Set bar = NewBox(ws, "spl_bar", 5, cx - tw / 2, cy + 50, 6, 6, RGB(20, 184, 166))
    bar.Adjustments.Item(1) = 0.5
    Set stp = NewText(ws, "spl_step", cx - 220, cy + 64, 440, 18)
    StyleText stp, "Starting up...", 9, RGB(148, 163, 184), False, 2
    Set shp = NewText(ws, "spl_ver", cx - 220, cy + 88, 440, 16)
    StyleText shp, "v" & APP_VERSION, 8, RGB(100, 116, 139), False, 2
    Application.ScreenUpdating = True
    t0 = Timer
    msgs = Array("Starting up...", "Tidying your sheets...", "Checking the connection to Tally...", "Preparing your dashboard...", "Ready")
    For k = 0 To 4
        stp.TextFrame2.TextRange.Text = msgs(k)
        Select Case k
            Case 1: ApplyLayout220
                    ApplyLayout221
                    ApplyLayout222
                    ApplyAmountFormats
                    ApplyLayout223
                    ApplyLayout230
                    ApplyLayout231
                    ProtectFormulaSheets
            Case 2: RefreshTallyStatus
            Case 3: EnsureHelpUpdateGuide
                    EnsureHistorySheet
        End Select
        For i = 1 To 5
            bar.Width = 6 + (tw - 6) * (k * 5 + i) / 25
            TransitionPause 0.03
        Next i
    Next k
    Do While Abs(Timer - t0) < 1.2
        DoEvents
    Loop
    For i = 1 To 8
        ShapesAlpha "spl_", ws, i / 8
        ws.Shapes("spl_bg").Fill.Transparency = i / 8
        TransitionPause 0.03
    Next i
    DeleteShapes "spl_"
    Exit Sub
plain:
    DeleteShapes "spl_"
    On Error Resume Next
    ApplyLayout220
    ApplyLayout221
    ApplyLayout222
    ApplyAmountFormats
    ApplyLayout223
    ApplyLayout230
    ApplyLayout231
    ProtectFormulaSheets
    RefreshTallyStatus
    EnsureHelpUpdateGuide
    EnsureHistorySheet
End Sub

' One-time sheet improvements of v2.2.0 (wrap + fit rows, colour keys, Needs review figure)
Private Sub ApplyLayout220()
    Dim ws As Worksheet, home As Worksheet, lastR As Long, r As Long, c As Long, q As Long
    On Error Resume Next
    If CStr(AppData().Range("B9").Value) = LAYOUT_TAG Then Exit Sub
    ' Entries
    Set ws = EntriesSheet()
    lastR = LastDataRow(ws)
    ws.Cells(4, COL_Q).Value = "q"
    ws.Columns(COL_Q).Hidden = True
    For r = FIRST_ROW To lastR
        If CellText(ws.Cells(r, COL_NOTE).Value) <> "" And CellText(ws.Cells(r, COL_Q).Value) = "" Then
            c = ws.Cells(r, COL_NOTE).Interior.Color
            If c = CLR_OK Then
                q = 3
            ElseIf c = CLR_INFO Then
                q = 2
            Else
                q = 1
            End If
            ws.Cells(r, COL_Q).Value = q
        End If
    Next r
    With ws.Range(ws.Cells(FIRST_ROW, COL_NARR), ws.Cells(FIRST_ROW + 2499, COL_NARR))
        .WrapText = True
    End With
    ws.Range(ws.Cells(FIRST_ROW, COL_STATUS), ws.Cells(FIRST_ROW + 2499, COL_NOTE)).WrapText = True
    ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + 2499, COL_NOTE)).VerticalAlignment = -4108
    If lastR >= FIRST_ROW Then ws.Range(ws.Rows(FIRST_ROW), ws.Rows(lastR)).AutoFit
    ColourKey ws.Range("A3"), "Colour key:", Array( _
        Array(RGB(34, 197, 94), "Green = sure match"), Array(RGB(234, 179, 8), "Yellow = please check"), _
        Array(RGB(239, 68, 68), "Red = no match / Suspense / error"), Array(RGB(96, 165, 250), "Blue = already in Tally (skipped)"), _
        Array(RGB(16, 185, 129), "Status green = posted"))
    ws.Rows(3).RowHeight = 20
    ' Change Ledger
    Set ws = FxSheet()
    ws.Range(ws.Cells(S_FIRST, S_NARR), ws.Cells(S_FIRST + 2999, S_NARR)).WrapText = True
    ws.Range(ws.Cells(S_FIRST, S_STATUS), ws.Cells(S_FIRST + 2999, S_STATUS)).WrapText = True
    ws.Range(ws.Cells(S_FIRST, 1), ws.Cells(S_FIRST + 2999, S_STATUS)).VerticalAlignment = -4108
    lastR = ws.Cells(ws.Rows.Count, S_DATE).End(xlUp).Row
    If lastR >= S_FIRST Then ws.Range(ws.Rows(S_FIRST), ws.Rows(lastR)).AutoFit
    ColourKey ws.Range("A3"), "Load entries > pick the right ledger in the yellow column > Update in Tally.   Colour key:", Array( _
        Array(RGB(250, 204, 21), "Yellow = choose here"), Array(RGB(249, 115, 22), "Orange = change it in Tally"), _
        Array(RGB(34, 197, 94), "Green = updated"), Array(RGB(239, 68, 68), "Red = error"))
    ws.Range("A3").WrapText = True
    ws.Rows(3).RowHeight = 22
    ' Home: Needs review = yellow + red rows not posted yet, plus rows with errors (red figure)
    Set home = ThisWorkbook.Worksheets("Home")
    home.Unprotect SHEET_PW
    With home.Range("K13")
        .Formula = "=SUMPRODUCT((Entries!$A$5:$A$5000<>"""")*(Entries!$J$5:$J$5000="""")*(Entries!$L$5:$L$5000>0)*(Entries!$L$5:$L$5000<3))" & _
                   "+SUMPRODUCT((Entries!$A$5:$A$5000<>"""")*ISNUMBER(SEARCH(""error"",Entries!$J$5:$J$5000)))"
        .Font.Color = RGB(220, 38, 38)
        .NumberFormat = "#,##0;-#,##0;[Color10]0"
    End With
    home.Range("K12:Q12").Interior.Color = RGB(220, 38, 38)
    home.Range("K15").Value = "yellow, red & error rows"
    SetApp "A9", "Layout version"
    SetApp "B9", LAYOUT_TAG
End Sub

' "Colour key:  [] Green = ...   [] Yellow = ..." with coloured squares
Private Sub ColourKey(cel As Range, ByVal lead As String, ByVal items As Variant)
    Dim txt As String, i As Long, pos() As Long, it As Variant
    On Error Resume Next
    ReDim pos(0 To UBound(items))
    txt = lead & "   "
    For i = 0 To UBound(items)
        pos(i) = Len(txt) + 1
        txt = txt & ChrW(&H25A0) & " " & items(i)(1) & "     "
    Next i
    cel.Value = txt
    cel.Font.Italic = False
    cel.Font.Size = 9
    cel.Font.Color = RGB(100, 116, 139)
    cel.Characters(1, Len(lead)).Font.Bold = True
    For i = 0 To UBound(items)
        cel.Characters(pos(i), 1).Font.Color = items(i)(0)
        cel.Characters(pos(i), 1).Font.Size = 12
    Next i
End Sub

' Formula cells are locked: the dashboard sheet is protected (macros can still update it)
Private Sub ProtectFormulaSheets()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Home")
    If ws Is Nothing Then Exit Sub
    ws.Unprotect SHEET_PW
    ws.Cells.Locked = True
    ws.Protect Password:=SHEET_PW, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ws.EnableSelection = 0
End Sub

' Quick check whether Tally answers (shown on Home)
Public Sub RefreshTallyStatus()
    Dim http As Object, ok As Boolean
    On Error GoTo fin
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 1500, 1500, 2500, 2500
    http.Open "GET", TallyUrl(), False
    http.send
    ok = (http.Status = 200)
fin:
    SetApp "B2", IIf(ok, "Connected", "Not connected")
    SetApp "B3", Now
End Sub

Public Sub TestConnectionUI()
    RefreshTallyStatus
    If AppData().Range("B2").Value = "Connected" Then
        Notify "Tally is connected and ready." & vbCrLf & TallyUrl(), "ok"
    Else
        ConnectionHelp "No answer from Tally"
    End If
End Sub

' Ribbon buttons: the button's tag is the macro to run
Public Sub RibbonAction(control As IRibbonControl)
    Application.Run "TallyLink." & control.Tag
End Sub

'---------------- navigation ----------------
Private Sub GoSheet(ByVal nm As String, Optional ByVal cellAddr As String = "A1")
    Dim src As Object, dst As Object
    On Error GoTo plain
    Set dst = ThisWorkbook.Worksheets(nm)
    Set src = ActiveSheet
    If src Is dst Then dst.Range(cellAddr).Select: Exit Sub
    PageTransition src, dst, cellAddr
    Exit Sub
plain:
    On Error Resume Next
    QuietActivate ThisWorkbook.Worksheets(nm)
    ThisWorkbook.Worksheets(nm).Range(cellAddr).Select
    DeleteShapes "trn_"
End Sub

' Smooth page change: a soft navy veil with an "Opening ..." pill, switch underneath, veil fades away
Private Sub PageTransition(src As Object, dst As Object, ByVal cellAddr As String)
    Dim cov As Object, pill As Object, i As Long, steps As Long, veil As Single
    Dim l As Single, t As Single, w As Single, h As Single, vl As Single, vt As Single, vw As Single, vh As Single
    On Error GoTo fin
    steps = 5: veil = 0.55
    Application.ScreenUpdating = True
    ' 1) dim the current page
    CoverRect l, t, w, h
    ViewRect vl, vt, vw, vh
    Set cov = NewBox(src, "trn_out", 1, l, t, w, h, RGB(11, 31, 58))
    cov.Fill.Transparency = 1
    Set pill = NewBox(src, "trn_pill", 5, vl + vw / 2 - 110, vt + vh / 2 - 22, 220, 40, RGB(11, 31, 58))
    pill.Adjustments.Item(1) = 0.5
    StyleText pill, "Opening " & dst.Name & "...", 11, RGB(255, 255, 255), True, 2
    pill.TextFrame2.VerticalAnchor = 3
    pill.Fill.Transparency = 1
    pill.TextFrame2.TextRange.Font.Fill.Transparency = 1
    For i = 1 To steps
        cov.Fill.Transparency = 1 - (1 - veil) * i / steps
        pill.Fill.Transparency = 1 - i / steps
        pill.TextFrame2.TextRange.Font.Fill.Transparency = 1 - i / steps
        TransitionPause 0.02
    Next i
    ' 2) switch page underneath the same veil
    Application.ScreenUpdating = False
    QuietActivate dst
    dst.Range(cellAddr).Select
    CoverRect l, t, w, h
    ViewRect vl, vt, vw, vh
    Set cov = NewBox(dst, "trn_in", 1, l, t, w, h, RGB(11, 31, 58))
    cov.Fill.Transparency = veil
    Set pill = NewBox(dst, "trn_badge", 5, vl + vw / 2 - 110, vt + vh / 2 - 22, 220, 40, RGB(11, 31, 58))
    pill.Adjustments.Item(1) = 0.5
    StyleText pill, "Opening " & dst.Name & "...", 11, RGB(255, 255, 255), True, 2
    pill.TextFrame2.VerticalAnchor = 3
    DeleteShapes "trn_out"
    DeleteShapes "trn_pill"
    Application.ScreenUpdating = True
    ' 3) lift the veil
    For i = 1 To 6
        cov.Fill.Transparency = veil + (1 - veil) * i / 6
        pill.Fill.Transparency = i / 6
        pill.TextFrame2.TextRange.Font.Fill.Transparency = i / 6
        TransitionPause 0.025
    Next i
fin:
    DeleteShapes "trn_"
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not ActiveSheet Is dst Then QuietActivate dst: dst.Range(cellAddr).Select
End Sub

Private Sub TransitionPause(ByVal secs As Single)
    Dim t0 As Single
    t0 = Timer
    DoEvents
    Do While Abs(Timer - t0) < secs
        DoEvents
    Loop
End Sub
Public Sub GoHome()
    GoSheet "Home"
End Sub
Public Sub GoEntries()
    GoSheet "Entries", "A5"
End Sub
Public Sub GoChangeLedger()
    GoSheet "Change Ledger", "B2"
End Sub
Public Sub GoSettings()
    GoSheet "Settings", "B4"
End Sub
Public Sub GoHelp()
    GoSheet "Help"
End Sub
Public Sub GoLog()
    GoSheet "Activity Log"
End Sub
Public Sub GoRules()
    GoSheet "Rules", "A2"
End Sub
Public Sub GoLearned()
    GoSheet "Learned", "A2"
End Sub
Public Sub GoLedgers()
    GoSheet "Ledgers", "A2"
End Sub

Public Sub ShowAbout()
    MsgBox APP_NAME & "   v" & APP_VERSION & vbCrLf & vbCrLf & _
        "Bank statements to Tally, in minutes." & vbCrLf & _
        "Import statements, smart ledger matching that learns, duplicate-safe posting," & vbCrLf & _
        "suspense and ledger correction without opening Tally." & vbCrLf & vbCrLf & _
        "Your data stays on this PC: " & APP_NAME & " only talks to Tally on your computer" & vbCrLf & _
        "(and checks the update page for new versions).", vbInformation, "About " & APP_NAME
End Sub

'---------------- shapes helpers ----------------
Private Sub StyleText(shp As Object, ByVal txt As String, ByVal sz As Single, ByVal clr As Long, _
        Optional ByVal bold As Boolean = False, Optional ByVal align As Long = 1)
    With shp.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = -1
        .TextRange.Text = txt
        .TextRange.Font.Name = "Segoe UI"
        .TextRange.Font.Size = sz
        .TextRange.Font.Bold = IIf(bold, -1, 0)
        .TextRange.Font.Fill.ForeColor.RGB = clr
        .TextRange.ParagraphFormat.Alignment = align
    End With
End Sub

Private Function NewBox(ws As Object, ByVal nm As String, ByVal shapeType As Long, ByVal l As Single, ByVal t As Single, _
        ByVal w As Single, ByVal h As Single, ByVal fillClr As Long) As Object
    Dim shp As Object
    Set shp = ws.Shapes.AddShape(shapeType, l, t, w, h)
    shp.Name = nm
    shp.Fill.ForeColor.RGB = fillClr
    shp.Fill.Solid
    shp.Line.Visible = 0
    shp.Placement = 3
    Set NewBox = shp
End Function

Private Function NewText(ws As Object, ByVal nm As String, ByVal l As Single, ByVal t As Single, ByVal w As Single, ByVal h As Single) As Object
    Dim shp As Object
    Set shp = ws.Shapes.AddTextbox(1, l, t, w, h)
    shp.Name = nm
    shp.Fill.Visible = 0
    shp.Line.Visible = 0
    shp.Placement = 3
    Set NewText = shp
End Function

Private Sub DeleteShapes(ByVal prefix As String)
    Dim ws As Object, i As Long
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For i = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(i).Name, Len(prefix)) = prefix Then ws.Shapes(i).Delete
        Next i
    Next ws
End Sub

'---------------- loading overlay ----------------
Public Sub ProgressShow(ByVal title As String, Optional ByVal stepText As String = "Starting...")
    Dim ws As Object, vr As Object, l As Single, t As Single, w As Single, h As Single, cl As Single, ct As Single
    Dim shp As Object
    On Error GoTo fin
    ProgressHide
    Set ws = ActiveSheet
    CoverRect l, t, w, h
    Set shp = NewBox(ws, "pfy_ov", 1, l, t, w, h, RGB(11, 31, 58))
    shp.Fill.Transparency = 0.45
    ViewRect l, t, w, h
    cl = l + w / 2 - 200: ct = t + h / 2 - 80
    If cl < l + 4 Then cl = l + 4
    If ct < t + 4 Then ct = t + 4
    Set shp = NewBox(ws, "pfy_card", 5, cl, ct, 400, 150, RGB(255, 255, 255))
    shp.Adjustments.Item(1) = 0.08
    shp.Shadow.Visible = -1: shp.Shadow.Blur = 18: shp.Shadow.OffsetY = 4: shp.Shadow.Transparency = 0.75
    Set shp = NewBox(ws, "pfy_accent", 1, cl + 24, ct + 22, 4, 26, RGB(20, 184, 166))
    Set shp = NewText(ws, "pfy_title", cl + 38, ct + 20, 340, 28)
    StyleText shp, title, 14, RGB(11, 31, 58), True
    Set shp = NewText(ws, "pfy_step", cl + 24, ct + 62, 352, 20)
    StyleText shp, stepText, 10, RGB(71, 85, 105)
    Set shp = NewBox(ws, "pfy_barbg", 5, cl + 24, ct + 96, 352, 10, RGB(226, 232, 240))
    shp.Adjustments.Item(1) = 0.5
    Set shp = NewBox(ws, "pfy_bar", 5, cl + 24, ct + 96, 10, 10, RGB(20, 184, 166))
    shp.Adjustments.Item(1) = 0.5
    Set shp = NewText(ws, "pfy_pct", cl + 24, ct + 114, 352, 18)
    StyleText shp, APP_NAME & "  -  please wait", 9, RGB(148, 163, 184), False, 1
    gProgOn = True: gProgTick = 0: gProgStart = Timer
    Application.ScreenUpdating = True
    DoEvents
fin:
End Sub

' pct 0..1 ; updates at most ~6 times a second
Public Sub ProgressUpdate(ByVal pct As Double, ByVal stepText As String)
    Dim ws As Object, su As Boolean
    If Not gProgOn Then Exit Sub
    If pct < 1 And Abs(Timer - gProgTick) < 0.35 Then Exit Sub
    gProgTick = Timer
    On Error GoTo fin
    If pct < 0.03 Then pct = 0.03
    If pct > 1 Then pct = 1
    For Each ws In ThisWorkbook.Worksheets
        If ShapeExists(ws, "pfy_bar") Then
            ws.Shapes("pfy_bar").Width = 352 * pct
            ws.Shapes("pfy_step").TextFrame2.TextRange.Text = stepText
            ws.Shapes("pfy_pct").TextFrame2.TextRange.Text = Format$(pct * 100, "0") & "%   -   " & APP_NAME
            Exit For
        End If
    Next ws
    su = Application.ScreenUpdating
    Application.ScreenUpdating = True
    DoEvents
    Application.ScreenUpdating = su
fin:
End Sub

Public Sub ProgressHide()
    gProgOn = False
    DeleteShapes "pfy_"
End Sub

' Finish nicely: show 100% / Done and keep the card readable for a moment
Public Sub ProgressDone(Optional ByVal doneText As String = "Done")
    Dim ws As Object, el As Single, waitFor As Single, t0 As Single
    If Not gProgOn Then Exit Sub
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ShapeExists(ws, "pfy_bar") Then
            ws.Shapes("pfy_bar").Width = 352
            ws.Shapes("pfy_step").TextFrame2.TextRange.Text = doneText
            ws.Shapes("pfy_pct").TextFrame2.TextRange.Text = "100%   -   " & APP_NAME
            Exit For
        End If
    Next ws
    Application.ScreenUpdating = True
    DoEvents
    el = Timer - gProgStart
    If el < 0 Then el = el + 86400
    waitFor = 2.5 - el
    If waitFor < 0.9 Then waitFor = 0.9
    t0 = Timer
    Do While Abs(Timer - t0) < waitFor
        DoEvents
    Loop
    ProgressHide
End Sub

Private Function ShapeExists(ws As Object, ByVal nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = ws.Shapes(nm)
    ShapeExists = Not s Is Nothing
End Function

'---------------- toast notifications ----------------
' kind: "ok" (teal), "warn" (amber), "err" (red)
Public Sub Notify(ByVal msg As String, Optional ByVal kind As String = "ok")
    Dim ws As Object, vr As Object, w As Single, h As Single, l As Single, t As Single, n As Long, clr As Long
    Dim shp As Object, secs As Long, vl As Single, vt As Single, vw As Single, vh As Single
    On Error GoTo fallback
    ToastHide
    Set ws = ActiveSheet
    ViewRect vl, vt, vw, vh
    n = UBound(Split(msg, vbCrLf)) + 1
    w = 360: h = 40 + n * 16
    If w > vw - 16 Then w = vw - 16
    If w < 200 Then w = 200
    l = vl + vw - w - 16
    t = vt + 12
    If l < vl + 4 Then l = vl + 4
    Select Case kind
        Case "warn": clr = RGB(245, 158, 11)
        Case "err": clr = RGB(239, 68, 68)
        Case Else: clr = RGB(20, 184, 166)
    End Select
    Set shp = NewBox(ws, "tst_card", 5, l, t, w, h, RGB(255, 255, 255))
    shp.Adjustments.Item(1) = 0.12
    shp.Shadow.Visible = -1: shp.Shadow.Blur = 14: shp.Shadow.OffsetY = 3: shp.Shadow.Transparency = 0.72
    shp.OnAction = "TallyLink.ToastHide"
    Set shp = NewBox(ws, "tst_bar", 5, l + 10, t + 12, 5, h - 24, clr)
    shp.OnAction = "TallyLink.ToastHide"
    Set shp = NewText(ws, "tst_title", l + 26, t + 10, w - 60, 18)
    StyleText shp, APP_NAME, 9, clr, True
    shp.OnAction = "TallyLink.ToastHide"
    Set shp = NewText(ws, "tst_msg", l + 26, t + 28, w - 40, h - 34)
    StyleText shp, msg, 10, RGB(30, 41, 59)
    shp.OnAction = "TallyLink.ToastHide"
    Set shp = NewText(ws, "tst_x", l + w - 26, t + 6, 20, 20)
    StyleText shp, ChrW(&H2715), 11, RGB(100, 116, 139), True, 2
    shp.OnAction = "TallyLink.ToastHide"
    Application.ScreenUpdating = True
    secs = 7 + 2 * n
    If secs > 20 Then secs = 20
    Application.OnTime Now + TimeSerial(0, 0, secs), "TallyLink.ToastHide"
    Exit Sub
fallback:
    MsgBox msg, IIf(kind = "err", vbExclamation, vbInformation), APP_NAME
End Sub

Public Sub ToastHide()
    DeleteShapes "tst_"
End Sub

'---------------- activity log ----------------
Public Sub LogAct(ByVal action As String, ByVal details As String, Optional ByVal n As Long = -1)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Activity Log")
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 1
    If r < 5 Then r = 5
    ws.Cells(r, 2).Value = Now
    ws.Cells(r, 2).NumberFormat = "dd-mmm-yy hh:mm"
    ws.Cells(r, 3).Value = action
    ws.Cells(r, 4).Value = details
    If n >= 0 Then ws.Cells(r, 5).Value = n
    ws.Cells(r, 6).Value = Environ$("USERNAME")
    GridBorders ws.Range(ws.Cells(r, 2), ws.Cells(r, 6))
End Sub

'---------------- undo a ledger change ----------------
Public Sub UndoLedgerChange()
    Dim ws As Worksheet, c As Range, r As Long, st As String, p1 As Long, p2 As Long, oldL As String, n As Long
    Set ws = FxSheet()
    If Not ActiveSheet Is ws Then QuietActivate ws: MsgBox "Select the green 'Updated in Tally' rows you want to undo, then click Undo again.", vbInformation, APP_NAME: Exit Sub
    For Each c In Selection.Rows
        r = c.Row
        If r >= S_FIRST Then
            st = CellText(ws.Cells(r, S_STATUS).Value)
            p1 = InStr(st, "(was '")
            If LCase$(Left$(st, 7)) = "updated" And p1 > 0 Then
                p1 = p1 + 6
                p2 = InStr(p1, st, "')")
                If p2 > p1 Then
                    oldL = Mid$(st, p1, p2 - p1)
                    ws.Cells(r, S_NEW).Value = oldL
                    ws.Cells(r, S_STATUS).Value = ""
                    ws.Cells(r, S_STATUS).Interior.ColorIndex = xlNone
                    n = n + 1
                End If
            End If
        End If
    Next c
    If n = 0 Then MsgBox "Select the green 'Updated in Tally' rows you want to undo, then click Undo again.", vbInformation, APP_NAME: Exit Sub
    LogAct "Undo ledger change", n & " row(s) set back to their previous ledger", n
    UpdateLedgerInTally
End Sub

'---------------- what's new (shown once after an update) ----------------
Public Sub ShowWhatsNew()
    Dim ws As Object, items As Variant, i As Long, txt As String, shp As Object, card As Object, lst As Object
    Dim l As Single, t As Single, w As Single, h As Single, cw As Single, cl As Single, ct As Single, y As Single
    On Error GoTo fin
    SetApp "B8", APP_VERSION
    WhatsNewHide
    ToastHide
    Set ws = ActiveSheet
    items = Split(WHATS_NEW, "|")
    CoverRect l, t, w, h
    Set shp = NewBox(ws, "wn_ov", 1, l, t, w, h, RGB(11, 31, 58))
    shp.Fill.Transparency = 0.45
    shp.OnAction = "TallyLink.WhatsNewGotIt"
    ViewRect l, t, w, h
    cw = 500
    If cw > w - 20 Then cw = w - 20
    If cw < 300 Then cw = 300
    cl = l + (w - cw) / 2: ct = t + 24
    If cl < l + 4 Then cl = l + 4
    Set card = NewBox(ws, "wn_card", 5, cl, ct, cw, 300, RGB(255, 255, 255))
    card.Adjustments.Item(1) = 0.04
    card.Shadow.Visible = -1: card.Shadow.Blur = 22: card.Shadow.OffsetY = 5: card.Shadow.Transparency = 0.7
    card.OnAction = "TallyLink.WhatsNewNoop"
    Set shp = NewBox(ws, "wn_badge", 5, cl + 28, ct + 24, 78, 20, RGB(230, 250, 246))
    shp.Adjustments.Item(1) = 0.5
    StyleText shp, "UPDATED", 8, RGB(15, 118, 110), True, 2
    shp.TextFrame2.VerticalAnchor = 3
    shp.OnAction = "TallyLink.WhatsNewNoop"
    Set shp = NewText(ws, "wn_title", cl + 28, ct + 52, cw - 56, 28)
    StyleText shp, "What's new in " & APP_NAME & " v" & APP_VERSION, 16, RGB(11, 31, 58), True
    shp.OnAction = "TallyLink.WhatsNewNoop"
    Set shp = NewText(ws, "wn_sub", cl + 28, ct + 82, cw - 56, 18)
    StyleText shp, "Your entries, ledgers, rules, learning and settings are unchanged.", 9.5, RGB(100, 116, 139)
    shp.OnAction = "TallyLink.WhatsNewNoop"
    Set shp = NewBox(ws, "wn_line", 1, cl + 28, ct + 110, cw - 56, 1, RGB(226, 232, 240))
    For i = 0 To UBound(items)
        txt = txt & IIf(i > 0, vbCr, "") & ChrW(&H2713) & "   " & items(i)
    Next i
    Set lst = NewText(ws, "wn_list", cl + 28, ct + 124, cw - 56, 40)
    StyleText lst, txt, 10.5, RGB(30, 41, 59)
    With lst.TextFrame2
        .TextRange.ParagraphFormat.SpaceAfter = 7
        .TextRange.ParagraphFormat.LeftIndent = 18
        .TextRange.ParagraphFormat.FirstLineIndent = -18
        For i = 1 To UBound(items) + 1
            .TextRange.Paragraphs(i).Characters(1, 1).Font.Fill.ForeColor.RGB = RGB(20, 184, 166)
            .TextRange.Paragraphs(i).Characters(1, 1).Font.Bold = -1
        Next i
        .AutoSize = 1
    End With
    lst.OnAction = "TallyLink.WhatsNewNoop"
    y = lst.Top + lst.Height + 18
    Set shp = NewBox(ws, "wn_ok", 5, cl + (cw - 140) / 2, y, 140, 34, RGB(20, 184, 166))
    shp.Adjustments.Item(1) = 0.25
    StyleText shp, "Got it", 11, RGB(255, 255, 255), True, 2
    shp.TextFrame2.VerticalAnchor = 3
    shp.OnAction = "TallyLink.WhatsNewGotIt"
    card.Height = y + 34 + 26 - ct
    Application.ScreenUpdating = True
    LogAct "Updated", APP_NAME & " v" & APP_VERSION & " installed"
    Exit Sub
fin:
    WhatsNewHide
    MsgBox APP_NAME & " v" & APP_VERSION & vbCrLf & vbCrLf & "What's new:" & vbCrLf & "  - " & Join(Split(WHATS_NEW, "|"), vbCrLf & "  - "), vbInformation, "What's new"
End Sub

Public Sub WhatsNewHide()
    DeleteShapes "wn_"
End Sub

Public Sub WhatsNewGotIt()
    WhatsNewHide
    If CStr(AppData().Range("B17").Value) <> "card" Then MailIntro
End Sub

Public Sub WhatsNewNoop()
End Sub

' Help sheet: "How to update" section (added once). Returns its row.
Private Function EnsureHelpUpdateGuide() As Long
    Dim ws As Worksheet, f As Range, r As Long, i As Long, steps As Variant
    On Error GoTo fin
    Set ws = ThisWorkbook.Worksheets("Help")
    Set f = ws.Columns(2).Find(What:="How to update Postify Bank", LookIn:=-4163, LookAt:=1)
    If Not f Is Nothing Then EnsureHelpUpdateGuide = f.Row: Exit Function
    steps = Array( _
        "1.  One-time setting: File > Options > Trust Center > Trust Center Settings > Macro Settings > tick 'Trust access to the VBA project object model' > OK. Without it the file cannot update itself.", _
        "2.  Save your work (Ctrl+S) and keep the internet on. The Update URL on the Settings sheet must be filled in (it comes pre-filled).", _
        "3.  Home > Check for updates (or the Postify Bank ribbon > Check for updates). When a new version is shown, click Yes / Update now.", _
        "4.  Wait a few seconds. Only the engine is replaced - your entries, ledgers, rules, learned narrations and settings stay exactly as they are, and the file saves itself.", _
        "5.  A 'What's new' window lists the changes. Click Got it. Home shows the new version number.", _
        "6.  If anything goes wrong, the previous version is put back automatically - nothing is lost.", _
        "7.  Manual update (no internet): you receive a TallyLink.bas file. Press Alt+F11, in the left panel open Modules, right-click TallyLink > Remove TallyLink > No. Then File > Import File > choose the new TallyLink.bas. Close the window and press Ctrl+S. Always import BEFORE you save.")
    r = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 3
    ws.Cells(r, 2).Value = "How to update Postify Bank"
    With ws.Cells(r, 2)
        .Font.Bold = True: .Font.Size = 13: .Font.Color = RGB(11, 31, 58): .VerticalAlignment = -4107
    End With
    ws.Rows(r).RowHeight = 26
    For i = 0 To UBound(steps)
        With ws.Cells(r + 1 + i, 2)
            .Value = steps(i)
            .WrapText = True
            .VerticalAlignment = -4160
            .IndentLevel = 1
            .Interior.Color = RGB(255, 255, 255)
            .Font.Color = RGB(30, 41, 59)
            .Font.Size = 10
        End With
        ws.Rows(r + 1 + i).RowHeight = 16 * (Len(steps(i)) \ 120 + 1) + 8
    Next i
    EnsureHelpUpdateGuide = r
fin:
End Function

' Button in What's new / Help
Public Sub ShowUpdateGuide()
    Dim r As Long
    WhatsNewHide
    r = EnsureHelpUpdateGuide()
    If r = 0 Then r = 1
    GoSheet "Help", "B" & r
    On Error Resume Next
    ActiveWindow.ScrollRow = IIf(r > 2, r - 1, 1)
End Sub

'==============================================================
'  BANK BALANCE CHECK  (Tally vs the statement provided by the user)
'==============================================================
Private Function FyStartOf(ByVal d As Date) As Date
    If Month(d) >= 4 Then FyStartOf = DateSerial(Year(d), 4, 1) Else FyStartOf = DateSerial(Year(d) - 1, 4, 1)
End Function

' Bank ledger balance in Tally at the end of day d. Positive = money in the account (Dr), minus = overdrawn.
Private Function TallyBankBalance(ByVal bank As String, ByVal d As Date, ByRef ok As Boolean) As Double
    Dim req As String, resp As String, co As String, sv As String, doc As Object, re As Object, ln As Object, nd As Object, nm As String
    ok = False
    On Error GoTo fin
    co = SettingValue("CompanyName")
    sv = "<SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
    If co <> "" Then sv = sv & "<SVCURRENTCOMPANY>" & XmlEsc(co) & "</SVCURRENTCOMPANY>"
    sv = sv & "<SVFROMDATE>" & TallyDateStr(FyStartOf(d)) & "</SVFROMDATE><SVTODATE>" & TallyDateStr(d) & "</SVTODATE>"
    req = "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>XLBankBal</ID></HEADER>" & _
          "<BODY><DESC><STATICVARIABLES>" & sv & "</STATICVARIABLES><TDL><TDLMESSAGE>" & _
          "<COLLECTION NAME=""XLBankBal"" ISMODIFY=""No""><TYPE>Ledger</TYPE><FETCH>NAME,CLOSINGBALANCE</FETCH></COLLECTION>" & _
          "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    If Not PostXml(req, resp) Then Exit Function
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = "&#([0-8]|1[124-9]|2[0-9]|3[01]);"
    resp = re.Replace(resp, "")
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    If Not doc.LoadXML(resp) Then Exit Function
    For Each ln In doc.SelectNodes("//LEDGER")
        nm = CStr(ln.getAttribute("NAME") & "")
        If nm = "" Then
            Set nd = ln.SelectSingleNode("NAME"): If Not nd Is Nothing Then nm = nd.Text
        End If
        If StrComp(Trim$(nm), Trim$(bank), vbTextCompare) = 0 Then
            Set nd = ln.SelectSingleNode("CLOSINGBALANCE")
            If Not nd Is Nothing Then TallyBankBalance = -Val(Replace(nd.Text, ",", ""))
            ok = True
            Exit Function
        End If
    Next ln
fin:
End Function

Private Function BankSheet() As Worksheet
    Dim ws As Worksheet, shp As Object, i As Long, hd As Variant, wd As Variant, back As Object, su As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(BB_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then Set BankSheet = ws: Exit Function
    Set back = ActiveSheet
    su = Application.ScreenUpdating
    Application.ScreenUpdating = False
    gNoAnim = True
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(FX_SHEET))
    ws.Name = BB_SHEET
    On Error Resume Next
    ws.Tab.Color = RGB(20, 184, 166)
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10
    wd = Array(36, 15, 20, 22, 16, 28, 18, 40)
    For i = 0 To 7: ws.Columns(i + 1).ColumnWidth = wd(i): Next i
    ws.Rows(1).RowHeight = 34
    With ws.Range("A1:H1")
        .Interior.Color = RGB(11, 31, 58)
    End With
    ws.Range("A1").Value = "  Bank Balances    Tally vs your bank statement, for every bank account"
    ws.Range("A1").Font.Color = RGB(148, 180, 210)
    ws.Range("A1").VerticalAlignment = -4108
    ws.Range("A1").Characters(1, 15).Font.Bold = True
    ws.Range("A1").Characters(1, 15).Font.Size = 15
    ws.Range("A1").Characters(1, 15).Font.Color = RGB(255, 255, 255)
    ws.Rows(2).RowHeight = 30
    ws.Range("A3:H3").Merge
    ws.Range("A3").Value = BB_NOTE
    ws.Range("A3").WrapText = True
    ws.Range("A3").Font.Italic = True
    ws.Range("A3").Font.Size = 9
    ws.Range("A3").Font.Color = RGB(180, 83, 9)
    ws.Range("A3").VerticalAlignment = -4160
    ws.Rows(3).RowHeight = 30
    ws.Range("A4").Value = "Positive = money in the account, minus = overdrawn. Updated on every statement import - click Refresh balances after posting or fixing entries."
    ws.Range("A4").Font.Size = 9: ws.Range("A4").Font.Color = RGB(100, 116, 139)
    hd = Array("Bank ledger (Tally)", "Statement up to", "Statement balance *", "Tally balance (same date)", "Difference", "Status", "Last checked", "Statement file")
    ws.Rows(5).RowHeight = 30
    For i = 0 To 7
        With ws.Cells(5, i + 1)
            .Value = hd(i)
            .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(19, 52, 94)
            .HorizontalAlignment = -4108: .VerticalAlignment = -4108: .WrapText = True
        End With
    Next i
    ws.Range("B6:B500").NumberFormat = "dd-mm-yyyy"
    SetRupee ws.Range("C6:E500")
    ws.Range("G6:G500").NumberFormat = "dd-mm-yy hh:mm"
    ' buttons
    Set shp = NewBox(ws, "bb_refresh", 5, ws.Range("A2").Left + 4, ws.Range("A2").Top + 4, 190, 22, RGB(20, 184, 166))
    StyleText shp, ChrW(&H21BB) & "  Refresh balances", 10, RGB(255, 255, 255), True, 2
    shp.TextFrame2.VerticalAnchor = 3: shp.OnAction = "TallyLink.RefreshBankBalances": shp.Placement = 2
    Set shp = NewBox(ws, "bb_home", 5, ws.Range("A2").Left + 204, ws.Range("A2").Top + 4, 110, 22, RGB(230, 250, 246))
    StyleText shp, ChrW(&H2302) & "  Home", 10, RGB(15, 118, 110), True, 2
    shp.TextFrame2.VerticalAnchor = 3: shp.OnAction = "TallyLink.GoHome": shp.Placement = 2
    ActiveWindow.DisplayGridlines = False
    DarkIfNeeded ws
    On Error Resume Next
    If Not back Is Nothing Then back.Activate
    gNoAnim = False
    Application.ScreenUpdating = su
    Set BankSheet = ws
End Function

Private Function BBRow(ws As Worksheet, ByVal bank As String, ByVal addIfMissing As Boolean) As Long
    Dim r As Long, lastR As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = BB_FIRST To lastR
        If StrComp(CellText(ws.Cells(r, 1).Value), bank, vbTextCompare) = 0 Then BBRow = r: Exit Function
    Next r
    If addIfMissing Then
        If lastR < BB_FIRST Then BBRow = BB_FIRST Else BBRow = lastR + 1
        ws.Cells(BBRow, 1).Value = bank
        GridBorders ws.Range(ws.Cells(5, 1), ws.Cells(BBRow, 8))
        SetRupee ws.Range(ws.Cells(BBRow, 3), ws.Cells(BBRow, 5))
    End If
End Function

' Compare one row with Tally and colour it
Private Sub BBCheckRow(ws As Worksheet, ByVal r As Long)
    Dim bank As String, d As Variant, tb As Double, ok As Boolean, st As String, clr As Long
    bank = CellText(ws.Cells(r, 1).Value)
    d = ParseDateCell(ws.Cells(r, 2).Value)
    If bank = "" Or IsEmpty(d) Then Exit Sub
    tb = TallyBankBalance(bank, CDate(d), ok)
    If ok Then ws.Cells(r, 4).Value = tb Else ws.Cells(r, 4).ClearContents
    If Not IsNumeric(ws.Cells(r, 3).Value) Or CellText(ws.Cells(r, 3).Value) = "" Then
        ws.Cells(r, 5).ClearContents
        st = "No balance in the statement": clr = RGB(241, 245, 249)
    ElseIf Not ok Then
        ws.Cells(r, 5).ClearContents
        st = "Tally not reachable - Refresh later": clr = RGB(254, 243, 199)
    Else
        ws.Cells(r, 5).Value = Round(CDbl(ws.Cells(r, 3).Value) - tb, 2)
        If Abs(CDbl(ws.Cells(r, 5).Value)) < 0.01 Then
            st = ChrW(&H2714) & "  Match": clr = CLR_OK
        Else
            st = ChrW(&H2716) & "  Does not match - check": clr = CLR_ERR
        End If
    End If
    ws.Cells(r, 6).Value = st
    ws.Range(ws.Cells(r, 5), ws.Cells(r, 6)).Interior.Color = clr
    ws.Range(ws.Cells(r, 5), ws.Cells(r, 6)).Font.Color = RGB(30, 41, 59)
    ws.Cells(r, 7).Value = Now
End Sub

' Called after an import: remember the statement balance (latest statement wins) and compare. Returns a short message.
Private Function BankStmtBalance(ByVal bank As String, ByVal dTo As Date, ByVal bal As Double, ByVal hasBal As Boolean, ByVal fileName As String) As String
    Dim ws As Worksheet, r As Long, old As Variant
    On Error GoTo fin
    Set ws = BankSheet()
    r = BBRow(ws, bank, True)
    old = ParseDateCell(ws.Cells(r, 2).Value)
    If Not IsEmpty(old) Then
        If CDate(old) > dTo And IsNumeric(ws.Cells(r, 3).Value) And CellText(ws.Cells(r, 3).Value) <> "" Then GoTo check
    End If
    ws.Cells(r, 2).Value = dTo
    If hasBal Then ws.Cells(r, 3).Value = bal Else ws.Cells(r, 3).Value = "not in statement"
    ws.Cells(r, 8).Value = fileName
check:
    BBCheckRow ws, r
    RefreshBankSummary
    If Left$(CellText(ws.Cells(r, 6).Value), 1) = ChrW(&H2714) Then
        BankStmtBalance = "Bank balance on " & Format$(ws.Cells(r, 2).Value, "dd-mm-yyyy") & ": Tally matches the statement *"
    ElseIf IsNumeric(ws.Cells(r, 5).Value) And CellText(ws.Cells(r, 5).Value) <> "" Then
        BankStmtBalance = "Bank balance on " & Format$(ws.Cells(r, 2).Value, "dd-mm-yyyy") & ": Tally differs from the statement by " & _
            Format$(ws.Cells(r, 5).Value, "#,##0.00") & " * (see Bank Balances)"
    End If
fin:
End Function

'================ BUTTON: Refresh balances ================
Public Sub RefreshBankBalances()
    Dim ws As Worksheet, r As Long, lastR As Long, n As Long, nOk As Long, nBad As Long
    Set ws = BankSheet()
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < BB_FIRST Then
        Notify "No bank statements imported yet." & vbCrLf & "Import a statement - its balance is checked automatically.", "warn"
        Exit Sub
    End If
    ProgressShow "Checking bank balances", "Reading balances from Tally..."
    For r = BB_FIRST To lastR
        ProgressUpdate (r - BB_FIRST + 1) / (lastR - BB_FIRST + 1), "Checking " & CellText(ws.Cells(r, 1).Value) & "..."
        BBCheckRow ws, r
        n = n + 1
        If Left$(CellText(ws.Cells(r, 6).Value), 1) = ChrW(&H2714) Then nOk = nOk + 1
        If Left$(CellText(ws.Cells(r, 6).Value), 1) = ChrW(&H2716) Then nBad = nBad + 1
    Next r
    ProgressDone "Balances checked"
    RefreshBankSummary
    LogAct "Refresh bank balances", nOk & " match, " & nBad & " do not match", n
    Notify nOk & " of " & n & " bank account(s) match Tally." & IIf(nBad > 0, vbCrLf & nBad & " do not match - see the red rows.", "") & _
        vbCrLf & "* as per the statements provided by you", IIf(nBad > 0, "warn", "ok")
End Sub

Public Sub GoBankBalances()
    BankSheet
    GoSheet BB_SHEET, "A6"
End Sub

' Home card "Bank balance check"
Private Sub RefreshBankSummary()
    Dim ws As Worksheet, home As Worksheet, r As Long, lastR As Long, nOk As Long, nBad As Long, firstBad As String, st As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(BB_SHEET)
    Set home = ThisWorkbook.Worksheets("Home")
    If home Is Nothing Then Exit Sub
    If Not ws Is Nothing Then
        lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For r = BB_FIRST To lastR
            st = CellText(ws.Cells(r, 6).Value)
            If Left$(st, 1) = ChrW(&H2714) Then nOk = nOk + 1
            If Left$(st, 1) = ChrW(&H2716) Then
                nBad = nBad + 1
                If firstBad = "" Then firstBad = CellText(ws.Cells(r, 1).Value)
            End If
        Next r
    End If
    home.Unprotect SHEET_PW
    home.Range("AL14").Value = "Bank balance check *"
    If nOk + nBad = 0 Then
        home.Range("AL13").Value = ChrW(&H2013)
        home.Range("AL13").Font.Color = RGB(148, 163, 184)
        home.Range("AL15").Value = "import a statement to check"
        home.Range("AL12:AR12").Interior.Color = RGB(203, 213, 225)
    ElseIf nBad = 0 Then
        home.Range("AL13").Value = nOk & " / " & nOk
        home.Range("AL13").Font.Color = RGB(5, 150, 105)
        home.Range("AL15").Value = "all accounts match Tally"
        home.Range("AL12:AR12").Interior.Color = RGB(52, 211, 153)
    Else
        home.Range("AL13").Value = nOk & " / " & (nOk + nBad)
        home.Range("AL13").Font.Color = RGB(220, 38, 38)
        home.Range("AL15").Value = "check " & Left$(firstBad, 14) & IIf(Len(firstBad) > 14, "..", "") & IIf(nBad > 1, " +" & (nBad - 1), "")
        home.Range("AL12:AR12").Interior.Color = RGB(220, 38, 38)
    End If
    home.Range("AL13").HorizontalAlignment = -4131
    home.Range("AR39").Value = APP_NAME & " v" & APP_VERSION
    home.Protect Password:=SHEET_PW, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    home.EnableSelection = 0
End Sub

' One-time changes of v2.2.1: Bank Balances sheet, Home card, footnote, Help section
Private Sub ApplyLayout221()
    Dim home As Worksheet, shp As Object, rg As Range, hl As Worksheet, f As Range, r As Long, i As Long, steps As Variant
    On Error Resume Next
    If CStr(AppData().Range("B10").Value) = "2.2.1" Then Exit Sub
    BankSheet
    Set home = ThisWorkbook.Worksheets("Home")
    home.Unprotect SHEET_PW
    ' 5th card (was Ledgers) becomes the bank balance check; click opens the Bank Balances sheet
    home.Range("AL15").Font.Color = RGB(100, 116, 139)
    For Each shp In home.Shapes
        If shp.Name = "kpi_bank" Then shp.Delete
    Next shp
    Set rg = home.Range("AL12:AR15")
    Set shp = home.Shapes.AddShape(1, rg.Left, rg.Top, rg.Width, rg.Height)
    shp.Name = "kpi_bank"
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255): shp.Fill.Transparency = 0.99
    shp.Line.Visible = 0
    shp.OnAction = "TallyLink.GoBankBalances"
    shp.Placement = 1
    ' footnote for the *
    home.Range("B40").Value = HOME_NOTE
    home.Range("B40").Font.Size = 8: home.Range("B40").Font.Italic = True: home.Range("B40").Font.Color = RGB(148, 163, 184)
    home.Rows(40).RowHeight = 16
    home.Protect Password:=SHEET_PW, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ' Help: short section
    Set hl = ThisWorkbook.Worksheets("Help")
    Set f = hl.Columns(2).Find(What:="Bank balance check", LookIn:=-4163, LookAt:=1)
    If f Is Nothing Then
        steps = Array( _
            "Every time you import a statement, Postify reads its closing balance and compares it with the bank ledger's balance in Tally on the same date. The result is on the Home card 'Bank balance check' and, for every bank account, on the Bank Balances sheet.", _
            "Green = Tally matches the statement. Red = there is a difference: entries missing in Tally, entered twice, or with a wrong amount. Click Refresh balances (Bank Balances sheet) after posting or fixing entries.", _
            "Find duplicates checks the balance first: when Tally matches the statement, entries of the same amount are real transactions and nothing is suggested for deletion.", _
            "* The statement balance is as per the bank statement provided by you, and the Tally balance as per your Tally data for the same date. The statement you provided is the basis for this check.")
        r = hl.Cells(hl.Rows.Count, 2).End(xlUp).Row + 3
        hl.Cells(r, 2).Value = "Bank balance check"
        With hl.Cells(r, 2)
            .Font.Bold = True: .Font.Size = 13: .Font.Color = RGB(11, 31, 58): .VerticalAlignment = -4107
        End With
        hl.Rows(r).RowHeight = 26
        For i = 0 To UBound(steps)
            With hl.Cells(r + 1 + i, 2)
                .Value = steps(i)
                .WrapText = True: .VerticalAlignment = -4160: .IndentLevel = 1
                .Interior.Color = RGB(255, 255, 255): .Font.Color = RGB(30, 41, 59): .Font.Size = 10
            End With
            hl.Rows(r + 1 + i).RowHeight = 16 * (Len(steps(i)) \ 120 + 1) + 8
        Next i
    End If
    SetApp "A10", "Layout 2.2.1"
    SetApp "B10", "2.2.1"
    RefreshBankSummary
End Sub

'==============================================================
'  v2.2.2: comma style amounts, borders, update history
'==============================================================
' Amounts in Excel's Comma Style (Alt+H+K): Indian commas as per Windows settings, 2 decimals, no currency symbol
Private Sub SetRupee(rg As Range)
    On Error Resume Next
    rg.NumberFormat = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * ""-""??_ ;_ @_ "
    If Err.Number <> 0 Then
        Err.Clear
        rg.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    End If
End Sub

' One-time: amount columns in Comma Style (also replaces the rupee format of an earlier 2.2.2 build)
Private Sub ApplyAmountFormats()
    Dim ws As Worksheet
    On Error Resume Next
    If CStr(AppData().Range("B13").Value) = "comma" Then Exit Sub
    Set ws = EntriesSheet()
    SetRupee ws.Range(ws.Cells(FIRST_ROW, COL_OUT), ws.Cells(FIRST_ROW + 2499, COL_IN))
    Set ws = FxSheet()
    SetRupee ws.Range(ws.Cells(S_FIRST, S_DR), ws.Cells(S_FIRST + 2999, S_CR))
    Set ws = BankSheet()
    SetRupee ws.Range("C6:E500")
    SetApp "A13", "Amount format"
    SetApp "B13", "comma"
End Sub

' Thin light-grey grid on a table range
Private Sub GridBorders(rg As Range)
    Dim b As Variant
    On Error Resume Next
    For Each b In Array(7, 8, 9, 10, 11, 12)
        With rg.Borders(b)
            .LineStyle = 1
            .Weight = 2
            .ColorIndex = -4105
        End With
    Next b
End Sub

Private Function LastUsedRow(ws As Worksheet, ByVal col As Long, ByVal minRow As Long) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
    If LastUsedRow < minRow Then LastUsedRow = minRow
End Function

Private Sub ApplyLayout222()
    Dim ws As Worksheet, home As Worksheet, hl As Worksheet, f As Range, rg As Range, shp As Object, lastR As Long
    On Error Resume Next
    If CStr(AppData().Range("B12").Value) = "2.2.2" Then Exit Sub
    ' Change Ledger: grid
    Set ws = FxSheet()
    GridBorders ws.Range(ws.Cells(S_FIRST - 1, 1), ws.Cells(S_FIRST + 2999, S_STATUS))
    ' Bank Balances
    Set ws = BankSheet()
    ws.Range("A3").Value = BB_NOTE
    ws.Range("A3").Font.Color = RGB(100, 116, 139)
    GridBorders ws.Range("A5:H" & LastUsedRow(ws, 1, BB_FIRST))
    ' other tables
    Set ws = ThisWorkbook.Worksheets("Activity Log")
    If Not ws Is Nothing Then GridBorders ws.Range("B4:F" & LastUsedRow(ws, 2, 5))
    Set ws = ThisWorkbook.Worksheets("Ledgers")
    If Not ws Is Nothing Then GridBorders ws.Range("A1:C" & LastUsedRow(ws, 1, 2))
    Set ws = ThisWorkbook.Worksheets("Rules")
    If Not ws Is Nothing Then GridBorders ws.Range("A1:D" & LastUsedRow(ws, 1, 2))
    Set ws = ThisWorkbook.Worksheets("Learned")
    If Not ws Is Nothing Then GridBorders ws.Range("A1:E" & LastUsedRow(ws, 1, 2))
    ' Home: last updated date, update history button, softer footnote
    Set home = ThisWorkbook.Worksheets("Home")
    home.Unprotect SHEET_PW
    home.Range("Y33").Value = "Last updated"
    home.Range("Y33").Font.Color = home.Range("Y32").Font.Color
    home.Range("AF33").Formula = "=IF(AppData!$B$11="""",""-"",TEXT(AppData!$B$11,""dd-mmm-yyyy""))"
    home.Range("AF33").Font.Bold = True
    home.Range("AF33").Font.Color = RGB(11, 31, 58)
    home.Range("B40").Value = HOME_NOTE
    Set rg = home.Range("Y37:AQ37")
    rg.UnMerge
    rg.Merge
    rg.Value = ChrW(&H2630) & "  Update history  -  features of the past updates"
    rg.HorizontalAlignment = -4108: rg.VerticalAlignment = -4108
    rg.Font.Bold = True: rg.Font.Color = RGB(15, 118, 110)
    rg.Interior.Color = RGB(230, 250, 246)
    GridBorders rg
    home.Rows(37).RowHeight = 24
    For Each shp In home.Shapes
        If shp.Name = "upd_hist" Then shp.Delete
    Next shp
    Set shp = home.Shapes.AddShape(1, rg.Left, rg.Top, rg.Width, rg.Height)
    shp.Name = "upd_hist"
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255): shp.Fill.Transparency = 0.99
    shp.Line.Visible = 0
    shp.OnAction = "TallyLink.ShowUpdateHistory"
    shp.Placement = 1
    home.Protect Password:=SHEET_PW, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    home.EnableSelection = 0
    ' Help: softer wording of the bank balance note
    Set hl = ThisWorkbook.Worksheets("Help")
    Set f = hl.Columns(2).Find(What:="accepts no liability", LookIn:=-4163, LookAt:=2)
    If Not f Is Nothing Then f.Value = "* The statement balance is as per the bank statement provided by you, and the Tally balance as per your Tally data for the same date. The statement you provided is the basis for this check."
    EnsureHistorySheet
    SetApp "A12", "Layout 2.2.2"
    SetApp "B12", "2.2.2"
End Sub

' Every version and what it brought (newest first): "version|date|item|item..." separated by "##"
Private Function HistoryText() As String
    Dim s As String
    s = "2.3.1|13-Nov-2026|Dark mode for every sheet|WhatsApp opens in App or Browser (WhatsApp Web) - your choice|" & _
        "Readable info cards for the suspense file, email and WhatsApp steps|Reminders: 2-Step Verification + App password for email, log in to WhatsApp first|" & _
        "Password window says your password stays on this PC|Tidier Updates & info box|Minor bug fixes"
    s = s & "##2.3.0|12-Nov-2026|Suspense file - Excel of all Suspense entries with a dropdown of every ledger and a Remarks column|" & _
        "Send it by Email (your own mailbox) or WhatsApp, or Save as|Import the filled file back into Change Ledger|" & _
        "Email & WhatsApp settings with saved recipient, subject and message|Minor bug fixes"
    s = s & "##2.2.3|11-Nov-2026|Dark mode - switch the dashboard between light and dark (button at the top right of Home)|" & _
        "All borders on every table (same as Alt+H+B+A)|Tidier Updates & info box and dashboard cards|Minor bug fixes"
    s = s & "##2.2.2|09-Nov-2026|Cleaner formatting - amounts with Indian commas and 2 decimals, and neat borders on every table|" & _
        "Last updated date shown on Home (Updates & info)|Update history - this page|Minor bugs fixed"
    s = s & "##2.2.1|09-Nov-2026|Bank balance check - Tally balance vs your bank statement for every bank account (Home card + Bank Balances sheet)|" & _
        "Find duplicates checks the balance first - nothing is suggested for deletion when Tally matches your statement|Minor bugs fixed"
    s = s & "##2.2.0|07-Nov-2026|Cells with formulas locked, so the dashboard cannot be broken by an accidental edit|Opening animation|" & _
        "Colour key on Entries and Change Ledger|Rows grow to fit long narrations|Smoother page changes|" & _
        "Needs review counts yellow and red rows and errors|Change Ledger orange rows fixed"
    s = s & "##2.1.1|06-Nov-2026|Home dashboard fills your screen"
    s = s & "##2.1.0|06-Nov-2026|Smarter learning - account numbers are remembered (e.g. MOB/SELFFT/9200.., XXXX3508)|" & _
        "Learns from the entries already in Tally - the latest entry wins|Learn from posted rows shows what was learned and skipped|" & _
        "Messages stay inside the screen|Smooth animation when switching sheets|What's new window and update instructions"
    s = s & "##2.0.0|05-Nov-2026|Postify Bank launch - dashboard, ribbon, loading screen, notifications, one-click updates|" & _
        "Bank reconciliation date filled automatically|Undo for Change Ledger|Activity Log"
    HistoryText = s
End Function

' "Update History" sheet, rebuilt whenever the version changes
Private Sub EnsureHistorySheet()
    Dim ws As Worksheet, back As Object, su As Boolean, vers As Variant, v As Variant, p As Variant, r As Long, i As Long, shp As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Update History")
    If Not ws Is Nothing Then
        If CStr(ws.Range("Z1").Value) = APP_VERSION Then Exit Sub
    End If
    Set back = ActiveSheet
    su = Application.ScreenUpdating
    Application.ScreenUpdating = False
    gNoAnim = True
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("Help"))
        ws.Name = "Update History"
        ws.Tab.Color = RGB(148, 163, 184)
    End If
    ws.Cells.Clear
    For Each shp In ws.Shapes: shp.Delete: Next shp
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10
    ws.Cells.Interior.Color = RGB(244, 247, 251)
    ws.Columns("A").ColumnWidth = 3: ws.Columns("B").ColumnWidth = 110: ws.Columns("C").ColumnWidth = 3
    ws.Rows(1).RowHeight = 34
    ws.Range("A1:C1").Interior.Color = RGB(11, 31, 58)
    ws.Range("B1").Value = "Update History    what changed in each version of " & APP_NAME
    ws.Range("B1").Font.Color = RGB(148, 180, 210)
    ws.Range("B1").VerticalAlignment = -4108
    ws.Range("B1").Characters(1, 14).Font.Bold = True
    ws.Range("B1").Characters(1, 14).Font.Size = 15
    ws.Range("B1").Characters(1, 14).Font.Color = RGB(255, 255, 255)
    ws.Range("Z1").Value = APP_VERSION
    ws.Range("Z1").Font.Color = RGB(11, 31, 58)
    ws.Rows(2).RowHeight = 30
    Set shp = NewBox(ws, "uh_home", 5, ws.Range("B2").Left, ws.Range("B2").Top + 4, 110, 22, RGB(230, 250, 246))
    StyleText shp, ChrW(&H2302) & "  Home", 10, RGB(15, 118, 110), True, 2
    shp.TextFrame2.VerticalAnchor = 3: shp.OnAction = "TallyLink.GoHome": shp.Placement = 2
    r = 4
    vers = Split(HistoryText(), "##")
    For Each v In vers
        p = Split(v, "|")
        ws.Cells(r, 2).Value = "v" & p(0) & "     " & p(1) & IIf(p(0) = APP_VERSION, "     (installed)", "")
        ws.Cells(r, 2).Font.Bold = True: ws.Cells(r, 2).Font.Size = 13: ws.Cells(r, 2).Font.Color = RGB(11, 31, 58)
        ws.Cells(r, 2).VerticalAlignment = -4107
        ws.Rows(r).RowHeight = 26
        r = r + 1
        For i = 2 To UBound(p)
            With ws.Cells(r, 2)
                .Value = ChrW(&H2713) & "   " & p(i)
                .WrapText = True: .IndentLevel = 1: .VerticalAlignment = -4108
                .Interior.Color = RGB(255, 255, 255): .Font.Color = RGB(30, 41, 59)
                .Characters(1, 1).Font.Color = RGB(20, 184, 166)
                .Characters(1, 1).Font.Bold = True
            End With
            ws.Rows(r).RowHeight = 20 * (Len(p(i)) \ 125 + 1)
            r = r + 1
        Next i
        GridBorders ws.Range(ws.Cells(r - UBound(p) + 1, 2), ws.Cells(r - 1, 2))
        r = r + 1
    Next v
    ActiveWindow.DisplayGridlines = False
    If IsDark() Then
        SwapPairs ws.Range("A1:AZ" & r + 5), True
        ws.Range("A" & r & ":E" & r + 300).Interior.Color = RGB(11, 18, 32)
        SheetBackground ws, True
    End If
    If Not back Is Nothing Then back.Activate
    gNoAnim = False
    Application.ScreenUpdating = su
End Sub

' Button on Home (Updates & info): Update history
Public Sub ShowUpdateHistory()
    EnsureHistorySheet
    GoSheet "Update History", "B4"
End Sub

'==============================================================
'  v2.2.3: all borders, tidier Home, light / dark dashboard
'==============================================================
Private Sub ApplyLayout223()
    Dim ws As Worksheet, home As Worksheet, fc As Object
    On Error Resume Next
    If CStr(AppData().Range("B15").Value) = "2.2.3" Then Exit Sub
    ' All borders (Alt+H+B+A) on every table
    Set ws = EntriesSheet()
    GridBorders ws.Range(ws.Cells(FIRST_ROW - 1, 1), ws.Cells(FIRST_ROW + 2499, COL_NOTE))
    Set ws = FxSheet()
    GridBorders ws.Range(ws.Cells(S_FIRST - 1, 1), ws.Cells(S_FIRST + 2999, S_STATUS))
    Set ws = BankSheet()
    GridBorders ws.Range("A5:H" & LastUsedRow(ws, 1, BB_FIRST))
    Set ws = ThisWorkbook.Worksheets("Activity Log")
    If Not ws Is Nothing Then GridBorders ws.Range("B4:F" & LastUsedRow(ws, 2, 5))
    Set ws = Nothing: Set ws = ThisWorkbook.Worksheets("Ledgers")
    If Not ws Is Nothing Then GridBorders ws.Range("A1:C" & LastUsedRow(ws, 1, 2))
    Set ws = Nothing: Set ws = ThisWorkbook.Worksheets("Rules")
    If Not ws Is Nothing Then GridBorders ws.Range("A1:D" & LastUsedRow(ws, 1, 2))
    Set ws = Nothing: Set ws = ThisWorkbook.Worksheets("Learned")
    If Not ws Is Nothing Then GridBorders ws.Range("A1:E" & LastUsedRow(ws, 1, 2))
    ThisWorkbook.Worksheets("Update History").Range("Z1").Value = ""     ' rebuild history with the new borders
    ' Home
    Set home = ThisWorkbook.Worksheets("Home")
    home.Unprotect SHEET_PW
    ' Updates & info: version / status / last updated move up one row, so they no longer touch the buttons
    If CellText(home.Range("Y30").Value) = "" And CellText(home.Range("Y31").Value) = "Installed version" Then
        home.Range("Y31:AQ33").Cut home.Range("Y30")
    End If
    ' Needs review: accent bar red only when something needs review (green when 0)
    home.Range("K15").Value = "yellow, red & error rows"
    With home.Range("K12:Q12")
        .FormatConditions.Delete
        .Interior.Color = RGB(52, 211, 153)
        Set fc = .FormatConditions.Add(Type:=2, Formula1:="=$K$13>0")
        fc.Interior.Color = RGB(220, 38, 38)
    End With
    home.Protect Password:=SHEET_PW, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    home.EnableSelection = 0
    RefreshBankSummary
    ThemeButton
    SetApp "A15", "Layout 2.2.3"
    SetApp "B15", "2.2.3"
End Sub

' Light / dark switch at the top right of the dashboard
Private Sub ThemeButton()
    Dim home As Worksheet, shp As Object, x As Single, dark As Boolean
    On Error Resume Next
    Set home = ThisWorkbook.Worksheets("Home")
    For Each shp In home.Shapes
        If shp.Name = "theme_btn" Then shp.Delete
    Next shp
    dark = (LCase$(CStr(AppData().Range("B14").Value)) = "dark")
    x = home.Range("AS1").Left - 128
    Set shp = NewBox(home, "theme_btn", 5, x, home.Range("A3").Top + 2, 112, 26, IIf(dark, RGB(226, 232, 240), RGB(11, 31, 58)))
    shp.Adjustments.Item(1) = 0.5
    shp.Line.Visible = -1: shp.Line.ForeColor.RGB = RGB(20, 184, 166): shp.Line.Weight = 1.25
    If dark Then
        StyleText shp, ChrW(&H2600) & "  Light mode", 9.5, RGB(11, 31, 58), True, 2
    Else
        StyleText shp, ChrW(&H263E) & "  Dark mode", 9.5, RGB(255, 255, 255), True, 2
    End If
    shp.TextFrame2.VerticalAnchor = 3
    shp.OnAction = "TallyLink.ToggleTheme"
    shp.Placement = 2
End Sub

' Swap one colour for another on the dashboard (fills, fonts and borders)
Private Sub SwapColour(rg As Range, ByVal c1 As Long, ByVal c2 As Long, ByVal doFill As Boolean, ByVal doFont As Boolean, ByVal doBorder As Boolean)
    Dim e As Variant
    On Error Resume Next
    If doFill Then
        Application.FindFormat.Clear: Application.ReplaceFormat.Clear
        Application.FindFormat.Interior.Color = c1
        Application.ReplaceFormat.Interior.Color = c2
        rg.Replace What:="", Replacement:="", LookAt:=2, SearchFormat:=True, ReplaceFormat:=True
    End If
    If doFont Then
        Application.FindFormat.Clear: Application.ReplaceFormat.Clear
        Application.FindFormat.Font.Color = c1
        Application.ReplaceFormat.Font.Color = c2
        rg.Replace What:="", Replacement:="", LookAt:=2, SearchFormat:=True, ReplaceFormat:=True
    End If
    If doBorder Then
        For Each e In Array(7, 8, 9, 10)
            Application.FindFormat.Clear: Application.ReplaceFormat.Clear
            Application.FindFormat.Borders(e).Color = c1
            Application.ReplaceFormat.Borders(e).Color = c2
            rg.Replace What:="", Replacement:="", LookAt:=2, SearchFormat:=True, ReplaceFormat:=True
        Next e
    End If
    Application.FindFormat.Clear: Application.ReplaceFormat.Clear
End Sub



'==============================================================
'  v2.3: SUSPENSE FILE  (make -> Email / WhatsApp / Save as -> import back)
'==============================================================
Private Sub SetSetting(ByVal nm As String, ByVal v As Variant)
    On Error Resume Next
    ThisWorkbook.Names(nm).RefersToRange.Value = v
End Sub

Private Function MailDomain(ByVal em As String) As String
    If InStr(em, "@") > 0 Then MailDomain = LCase$(Trim$(Mid$(em, InStr(em, "@") + 1)))
End Function

' Outlook.com / Hotmail / Microsoft 365: password sending is being switched off by Microsoft
Private Function IsMicrosoftMail(ByVal em As String, ByVal server As String) As Boolean
    Select Case MailDomain(em)
        Case "outlook.com", "outlook.in", "hotmail.com", "hotmail.co.in", "live.com", "live.in", "msn.com"
            IsMicrosoftMail = True
    End Select
    If InStr(1, server, "office365", vbTextCompare) > 0 Or InStr(1, server, "outlook", vbTextCompare) > 0 Then IsMicrosoftMail = True
End Function

' Known mail servers by address (others: user fills Settings)
Private Sub MailPreset(ByVal em As String, ByRef server As String, ByRef port As Long)
    server = "": port = 465
    Select Case MailDomain(em)
        Case "gmail.com", "googlemail.com": server = "smtp.gmail.com"
        Case "yahoo.com", "yahoo.co.in", "yahoo.in", "ymail.com", "rocketmail.com": server = "smtp.mail.yahoo.com"
        Case "zoho.com": server = "smtp.zoho.com"
        Case "zohomail.in", "zoho.in": server = "smtp.zoho.in"
        Case "rediffmail.com": server = "smtp.rediffmail.com"
        Case "outlook.com", "outlook.in", "hotmail.com", "hotmail.co.in", "live.com", "live.in", "msn.com"
            server = "smtp-mail.outlook.com": port = 587
    End Select
End Sub

Private Function Obf(ByVal txt As String, ByVal decode As Boolean) As String
    Dim k As String, i As Long, out As String, c As Long
    k = "Pf9#yB4nk!q2"
    On Error GoTo fin
    If decode Then
        For i = 1 To Len(txt) \ 4
            c = CLng("&H" & Mid$(txt, (i - 1) * 4 + 1, 4)) Xor AscW(Mid$(k, ((i - 1) Mod Len(k)) + 1, 1))
            out = out & ChrW(c)
        Next i
    Else
        For i = 1 To Len(txt)
            c = (AscW(Mid$(txt, i, 1)) And &HFFFF&) Xor AscW(Mid$(k, ((i - 1) Mod Len(k)) + 1, 1))
            out = out & Right$("0000" & Hex$(c), 4)
        Next i
    End If
    Obf = out
fin:
End Function

Private Function DocsFolder() As String
    Dim p As String
    On Error Resume Next
    p = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
    If p = "" Then p = Environ$("USERPROFILE") & "\Documents"
    p = p & "\Postify Suspense"
    If Dir(p, vbDirectory) = "" Then MkDir p
    DocsFolder = p
End Function

Private Function FileSafe(ByVal s As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        s = Replace(s, ch, "-")
    Next ch
    FileSafe = Trim$(s)
End Function

' Build the Excel for the person who knows the entries. Returns the saved path ("" if failed)
Private Function BuildSuspenseWorkbook(vc As Collection, ByVal sus As String, ByVal dFrom As Date, ByVal dTo As Date) As String
    Dim wb As Workbook, ws As Worksheet, wl As Worksheet, src As Worksheet, v As Variant, r As Long, k As Long
    Dim lastL As Long, lv As Variant, n As Long, co As String, path As String, hd As Variant, i As Long, nRows As Long
    On Error GoTo fail
    co = SettingValue("CompanyName")
    If co = "" Then co = "Company"
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add(-4167)
    Set ws = wb.Worksheets(1)
    ws.Name = "Suspense"
    Set wl = wb.Worksheets.Add(After:=ws)
    wl.Name = "Ledgers"
    ' --- ledger list as on today
    Set src = ThisWorkbook.Worksheets("Ledgers")
    lastL = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    wl.Cells.Font.Name = "Segoe UI": wl.Cells.Font.Size = 10
    wl.Range("A1").Value = "Ledgers in Tally as on " & Format$(Date, "dd-mm-yyyy")
    wl.Range("A1").Font.Bold = True: wl.Range("A1").Font.Size = 12
    wl.Range("A2:B2").Value = Array("Ledger", "Group")
    wl.Range("A2:B2").Font.Bold = True: wl.Range("A2:B2").Interior.Color = RGB(226, 232, 240)
    n = 0
    If lastL >= 2 Then
        lv = src.Range("A2:C" & lastL).Value
        For k = 1 To UBound(lv, 1)
            If CellText(lv(k, 1)) <> "" Then
                n = n + 1
                wl.Cells(2 + n, 1).Value = CellText(lv(k, 1))
                wl.Cells(2 + n, 2).Value = IIf(CellText(lv(k, 3)) <> "", CellText(lv(k, 3)), CellText(lv(k, 2)))
            End If
        Next k
    End If
    If n = 0 Then n = 1
    wb.Names.Add Name:="LedgerList", RefersTo:="=Ledgers!$A$3:$A$" & (n + 2)
    wl.Columns("A").ColumnWidth = 45: wl.Columns("B").ColumnWidth = 30
    GridBorders wl.Range("A2:B" & (n + 2))
    ' --- suspense sheet
    ws.Cells.Font.Name = "Segoe UI": ws.Cells.Font.Size = 10
    ws.Range("A1").Value = "Suspense entries - " & co
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 15: ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:J1").Interior.Color = RGB(11, 31, 58)
    ws.Rows(1).RowHeight = 30: ws.Range("A1").VerticalAlignment = -4108
    ws.Range("A2").Value = "Ledger: " & sus & "     Period: " & Format$(dFrom, "dd-mm-yyyy") & " to " & Format$(dTo, "dd-mm-yyyy") & _
        "     Entries: " & vc.Count & "     Extracted on: " & Format$(Date, "dd-mm-yyyy")
    ws.Range("A2").Font.Color = RGB(71, 85, 105)
    ws.Range("A3").Value = "Please choose the correct ledger for each entry in the yellow column (pick from the dropdown), add a remark if needed, save and send this file back."
    ws.Range("A3").Font.Bold = True: ws.Range("A3").Font.Color = RGB(180, 83, 9)
    hd = Array("Date", "Voucher Type", "Voucher No", "Bank / Other side", "Debit", "Credit", "Narration", "Current Ledger", "Correct Ledger", "Remarks", "Id")
    For i = 0 To 10
        ws.Cells(5, i + 1).Value = hd(i)
    Next i
    With ws.Range("A5:J5")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(19, 52, 94)
        .HorizontalAlignment = -4108: .VerticalAlignment = -4108: .WrapText = True
    End With
    ws.Rows(5).RowHeight = 30
    r = 6
    For Each v In vc
        ws.Cells(r, 1).Value = v(0)
        ws.Cells(r, 2).Value = v(3)
        ws.Cells(r, 3).NumberFormat = "@": ws.Cells(r, 3).Value = CStr(v(4))
        ws.Cells(r, 4).Value = v(5)
        If v(1) Then ws.Cells(r, 5).Value = v(2) Else ws.Cells(r, 6).Value = v(2)
        ws.Cells(r, 7).Value = v(6)
        ws.Cells(r, 8).Value = sus
        ws.Cells(r, 11).NumberFormat = "@": ws.Cells(r, 11).Value = CStr(v(8))
        r = r + 1
    Next v
    nRows = r - 6
    ws.Range("A6:A" & r - 1).NumberFormat = "dd-mm-yyyy"
    ws.Range("E6:F" & r - 1).NumberFormat = "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * ""-""??_ ;_ @_ "
    ws.Range("G6:G" & r - 1).WrapText = True
    ws.Range("A6:J" & r - 1).VerticalAlignment = -4108
    ws.Range("I6:I" & r - 1).Interior.Color = RGB(255, 248, 225)
    ws.Range("J6:J" & r - 1).Interior.Color = RGB(255, 255, 255)
    ws.Range("J6:J" & r - 1).WrapText = True
    With ws.Range("I6:I" & r - 1).Validation
        .Delete
        .Add Type:=3, AlertStyle:=2, Operator:=1, Formula1:="=LedgerList"
        .IgnoreBlank = True: .InCellDropdown = True
        .ErrorTitle = "Ledger": .ErrorMessage = "Please pick a ledger from the list (sheet Ledgers)."
    End With
    GridBorders ws.Range("A5:J" & r - 1)
    hd = Array(12, 13, 12, 26, 14, 14, 50, 22, 32, 30)
    For i = 0 To 9: ws.Columns(i + 1).ColumnWidth = hd(i): Next i
    ws.Range("A6:J" & r - 1).Rows.AutoFit
    ' hidden: voucher id + period, used when the file comes back
    ws.Columns("K:O").Hidden = True
    ws.Range("M1").Value = dFrom: ws.Range("N1").Value = dTo: ws.Range("O1").Value = SUSP_MARK
    ' footnote
    ws.Cells(r + 1, 1).Value = "Generated by " & APP_NAME & " v" & APP_VERSION & " on " & Format$(Now, "dd-mm-yyyy hh:nn")
    With ws.Cells(r + 1, 1).Font
        .Italic = True: .Size = 8: .Color = RGB(148, 163, 184)
    End With
    ws.PageSetup.CenterFooter = "&I&8Generated by " & APP_NAME
    ws.PageSetup.Orientation = 2
    ws.Activate
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    ' only the yellow and remarks columns can be typed in
    ws.Cells.Locked = True
    ws.Range("I6:J" & r - 1).Locked = False
    ws.Range("A5:J" & r - 1).AutoFilter
    ws.Protect DrawingObjects:=True, Contents:=True, AllowFiltering:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True
    path = DocsFolder() & "\" & FileSafe("Suspense - " & co & " - " & Format$(Date, "dd-mm-yyyy")) & ".xlsx"
    Application.DisplayAlerts = False
    wb.SaveAs fileName:=path, FileFormat:=51
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Application.ScreenUpdating = True
    BuildSuspenseWorkbook = path
    Exit Function
fail:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Application.ScreenUpdating = True
    MsgBox "The suspense file could not be made: " & Err.Description, vbExclamation, APP_NAME
End Function

'================ BUTTON (Change Ledger): Make suspense file ================
Public Sub MakeSuspenseFile()
    Dim sus As String, dFrom As Variant, dTo As Variant, vc As Collection, info As String, path As String
    sus = SuspenseLedgerName()
    If sus = "" Then MsgBox "No Suspense ledger found in the ledger list (click Load ledgers).", vbExclamation, APP_NAME: Exit Sub
    dFrom = ParseDateCell(FxSheet().Range("E2").Value)
    If IsEmpty(dFrom) Then dFrom = FyStart()
    dTo = ParseDateCell(FxSheet().Range("G2").Value)
    If IsEmpty(dTo) Then dTo = Date
    ProgressShow "Making the suspense file", "Reading '" & sus & "' from Tally..."
    Set vc = FetchBankVouchers(sus, CDate(dFrom), CDate(dTo), True, info)
    If vc Is Nothing Then ProgressHide: ConnectionHelp info: Exit Sub
    If vc.Count = 0 Then
        ProgressHide
        Notify "No Suspense entries from " & Format$(dFrom, "dd-mm-yyyy") & " to " & Format$(dTo, "dd-mm-yyyy") & "." & vbCrLf & _
            "Change the From / To dates on Change Ledger if needed.", "ok"
        Exit Sub
    End If
    ProgressUpdate 0.6, "Building the Excel file (" & vc.Count & " entries)..."
    path = BuildSuspenseWorkbook(vc, sus, CDate(dFrom), CDate(dTo))
    If path = "" Then ProgressHide: Exit Sub
    ProgressDone "File ready"
    LogAct "Suspense file", vc.Count & " entries -> " & Mid$(path, InStrRev(path, "\") + 1), vc.Count
    SendSuspenseDialog path, CDate(dFrom), CDate(dTo), vc.Count
End Sub

Private Function FillPlaceholders(ByVal t As String, ByVal dFrom As Date, ByVal dTo As Date, ByVal n As Long) As String
    Dim co As String
    co = SettingValue("CompanyName")
    t = Replace(t, "{company}", co)
    t = Replace(t, "{period}", Format$(dFrom, "dd-mm-yyyy") & " to " & Format$(dTo, "dd-mm-yyyy"))
    t = Replace(t, "{count}", CStr(n))
    FillPlaceholders = t
End Function

' Pop-up: Send Email / Send on WhatsApp / Save as. Saved details are pre-filled and can be edited for this send.
Private Sub SendSuspenseDialog(ByVal path As String, ByVal dFrom As Date, ByVal dTo As Date, ByVal n As Long)
    Dim f As Object, ch As String, sTo As String, sMob As String, sSub As String, sMsg As String
    Dim cTo As String, cMob As String, cSub As String, cMsg As String, done As Boolean, k As Variant, note As String
    sTo = SettingValue("SendToEmail"): sMob = SettingValue("SendToMobile")
    sSub = FillPlaceholders(SettingValue("SendSubject"), dFrom, dTo, n)
    sMsg = FillPlaceholders(Replace(ThisWorkbook.Names("SendMessage").RefersToRange.Value, vbLf, vbCrLf), dFrom, dTo, n)
    cTo = sTo: cMob = sMob: cSub = sSub: cMsg = sMsg
    If IsMicrosoftMail(SettingValue("MailFrom"), SettingValue("MailServer")) Then
        note = "Note: Outlook / Hotmail / Microsoft 365 mail may not send from here (Microsoft has switched off password sending). WhatsApp is recommended."
    End If
    If Not FormReady("PfySend") Then
        ' simple fallback without the pop-up window
        k = Application.InputBox("File ready: " & Mid$(path, InStrRev(path, "\") + 1) & vbCrLf & vbCrLf & _
            "1 = Send Email" & vbCrLf & "2 = Send on WhatsApp" & vbCrLf & "3 = Save as..." & IIf(note <> "", vbCrLf & vbCrLf & note, ""), APP_NAME, 1, Type:=1)
        If VarType(k) = vbBoolean Then Exit Sub
        Select Case CLng(k)
            Case 1
                k = Application.InputBox("Send to (email):", APP_NAME, sTo, Type:=2)
                If VarType(k) <> vbBoolean Then DoSendEmail path, CStr(k), cSub, cMsg
            Case 2
                k = Application.InputBox("WhatsApp mobile number:", APP_NAME, sMob, Type:=2)
                If VarType(k) <> vbBoolean Then DoWhatsApp path, CStr(k), cMsg
            Case 3: DoSaveAs path
        End Select
        Exit Sub
    End If
    Do
        Set f = VBA.UserForms.Add("PfySend")
        f.Controls("lblFile").Caption = "File ready:  " & Mid$(path, InStrRev(path, "\") + 1) & "   (" & n & " entries)"
        f.Controls("txtTo").Text = cTo: f.Controls("txtTo").Tag = sTo
        f.Controls("txtMob").Text = cMob: f.Controls("txtMob").Tag = sMob
        f.Controls("txtSub").Text = cSub: f.Controls("txtSub").Tag = sSub
        f.Controls("txtMsg").Text = cMsg: f.Controls("txtMsg").Tag = sMsg
        f.Controls("lblNote").Caption = note
        f.Controls("optWeb").Value = (LCase$(SettingValue("WaOpenIn")) <> "app")
        f.Controls("optApp").Value = (LCase$(SettingValue("WaOpenIn")) = "app")
        f.Show
        ch = f.Choice
        cTo = Trim$(f.Controls("txtTo").Text): cMob = Trim$(f.Controls("txtMob").Text)
        cSub = f.Controls("txtSub").Text: cMsg = f.Controls("txtMsg").Text
        If ch = "wa" Then SetSetting "WaOpenIn", IIf(f.Controls("optApp").Value, "App", "Browser")
        If f.Controls("chkDefault").Value And ch <> "" Then
            SetSetting "SendToEmail", cTo: SetSetting "SendToMobile", cMob
            sTo = cTo: sMob = cMob
        End If
        Unload f
        Set f = Nothing
        Select Case ch
            Case "email": done = DoSendEmail(path, cTo, cSub, cMsg)
            Case "wa": done = DoWhatsApp(path, cMob, cMsg)
            Case "save": done = DoSaveAs(path)
            Case Else
                Notify "The suspense file is saved in:" & vbCrLf & DocsFolder(), "ok"
                done = True
        End Select
    Loop Until done
End Sub

Private Function DoSaveAs(ByVal path As String) As Boolean
    Dim fn As Variant
    fn = Application.GetSaveAsFilename(InitialFileName:=Mid$(path, InStrRev(path, "\") + 1), FileFilter:="Excel Workbook (*.xlsx), *.xlsx", Title:="Save the suspense file")
    If VarType(fn) = vbBoolean Then Exit Function
    On Error GoTo fail
    If StrComp(CStr(fn), path, vbTextCompare) <> 0 Then FileCopy path, CStr(fn)
    Notify "Suspense file saved:" & vbCrLf & CStr(fn), "ok"
    LogAct "Suspense file", "Saved as " & CStr(fn)
    DoSaveAs = True
    Exit Function
fail:
    MsgBox "Could not save there: " & Err.Description, vbExclamation, APP_NAME
End Function

' UTF-8 URL encoding (for the WhatsApp message)
Private Function UrlEnc(ByVal s As String) As String
    Dim i As Long, c As Long, out As String, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1): c = AscW(ch) And &HFFFF&
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or ch = "-" Or ch = "." Or ch = "_" Or ch = "~" Then
            out = out & ch
        ElseIf c < 128 Then
            out = out & "%" & Right$("0" & Hex$(c), 2)
        ElseIf c < 2048 Then
            out = out & "%" & Hex$(192 + (c \ 64)) & "%" & Hex$(128 + (c And 63))
        Else
            out = out & "%" & Hex$(224 + (c \ 4096)) & "%" & Hex$(128 + ((c \ 64) And 63)) & "%" & Hex$(128 + (c And 63))
        End If
    Next i
    UrlEnc = out
End Function

Private Function DoWhatsApp(ByVal path As String, ByVal mob As String, ByVal msg As String) As Boolean
    Dim d As String, i As Long, ch As String, useApp As Boolean, txt As String
    For i = 1 To Len(mob)
        ch = Mid$(mob, i, 1)
        If ch Like "#" Then d = d & ch
    Next i
    If Len(d) = 11 And Left$(d, 1) = "0" Then d = Mid$(d, 2)
    If Len(d) = 10 Then d = "91" & d
    If Len(d) < 11 Then MsgBox "Please enter a valid mobile number (10 digits, or with country code).", vbExclamation, APP_NAME: Exit Function
    txt = UrlEnc(Replace(msg, vbCrLf, vbLf))
    useApp = (LCase$(SettingValue("WaOpenIn")) = "app")
    On Error Resume Next
    If useApp Then
        ThisWorkbook.FollowHyperlink "whatsapp://send?phone=" & d & "&text=" & txt
        If Err.Number <> 0 Then Err.Clear: useApp = False
    End If
    If Not useApp Then ThisWorkbook.FollowHyperlink "https://web.whatsapp.com/send?phone=" & d & "&text=" & txt
    Shell "explorer.exe /select,""" & path & """", vbNormalFocus
    On Error GoTo 0
    LogAct "Suspense file", "WhatsApp (" & IIf(useApp, "app", "browser") & ") opened for +" & d
    InfoCard "WhatsApp is opening", "Chat with +" & d & " - the message is already typed in.", Array( _
        "#Send the file", _
        "1.   The folder with your file is open: " & Mid$(path, InStrRev(path, "") + 1), _
        "2.   Drag the file into the WhatsApp chat (or click the paperclip / + and choose it).", _
        "3.   Press Send.", _
        "#If you see a QR code", _
        "WhatsApp is not logged in " & IIf(useApp, "in the app", "in this browser") & ". On your phone: WhatsApp > Linked devices > Link a device, scan the code, then make the file again and send.", _
        "Tip: the pop-up lets you choose Browser (WhatsApp Web) or the WhatsApp app."), _
        Array(Array("Got it", "TallyLink.InfoCardHide", "teal"))
    DoWhatsApp = True
End Function

' Ask for the mail password (dots, Show password, Remember)
Private Function PwdPrompt(ByVal em As String, ByRef remember As Boolean, ByRef cancelled As Boolean) As String
    Dim f As Object, v As Variant
    cancelled = False
    If FormReady("PfyPwd") Then
        Set f = VBA.UserForms.Add("PfyPwd")
        f.Controls("lblP").Caption = "Mail password for" & vbCrLf & em
        f.Controls("chkRem").Value = remember
        f.Show
        If f.Choice = "ok" Then
            PwdPrompt = f.Controls("txtPwd").Text
            remember = f.Controls("chkRem").Value
        Else
            cancelled = True
        End If
        Unload f
    Else
        v = Application.InputBox("Mail password for " & em & " (Gmail / Yahoo: the App password):", APP_NAME, "", Type:=2)
        If VarType(v) = vbBoolean Then cancelled = True Else PwdPrompt = CStr(v)
    End If
End Function

Private Function DoSendEmail(ByVal path As String, ByVal sendTo As String, ByVal subj As String, ByVal body As String) As Boolean
    Dim em As String, server As String, port As Long, ssl As Boolean, pwd As String, remember As Boolean, cancelled As Boolean
    Dim usedSaved As Boolean, m As Object, cfg As Object, pre As String, v As Variant, errTxt As String
    sendTo = Trim$(sendTo)
    If InStr(sendTo, "@") = 0 Then MsgBox "Please enter the email address to send to.", vbExclamation, APP_NAME: Exit Function
    em = SettingValue("MailFrom")
    If em = "" Then
        v = Application.InputBox("Your email address (the file is sent from this mailbox):", APP_NAME, "", Type:=2)
        If VarType(v) = vbBoolean Then Exit Function
        em = Trim$(CStr(v))
        If InStr(em, "@") = 0 Then Exit Function
        SetSetting "MailFrom", em
    End If
    server = SettingValue("MailServer")
    port = Val(SettingValue("MailPort"))
    If server = "" Then
        MailPreset em, server, port
        If server = "" Then
            MsgBox "Postify does not know the mail server for '" & MailDomain(em) & "'." & vbCrLf & vbCrLf & _
                "Please fill 'Mail server' and 'Port' on the Settings sheet (ask your mail provider - e.g. Google Workspace: smtp.gmail.com / 465, " & _
                "Zoho: smtp.zoho.in / 465). Then click Send Email again." & vbCrLf & vbCrLf & "Tip: WhatsApp needs no setup.", vbInformation, APP_NAME
            GoSheet "Settings", "B20"
            Exit Function
        End If
        SetSetting "MailServer", server: SetSetting "MailPort", port
    End If
    If port = 0 Then port = 465
    ssl = (LCase$(SettingValue("MailSSL")) <> "no")
    remember = (LCase$(SettingValue("MailRemember")) <> "no")
    If remember And CStr(AppData().Range("B16").Value) <> "" Then
        pwd = Obf(CStr(AppData().Range("B16").Value), True): usedSaved = True
    Else
        pwd = PwdPrompt(em, remember, cancelled)
        If cancelled Or pwd = "" Then Exit Function
    End If
    If InStr(1, server, "gmail", vbTextCompare) > 0 Or InStr(1, server, "yahoo", vbTextCompare) > 0 Then pwd = Replace(pwd, " ", "")
    ProgressShow "Sending email", "Connecting to " & server & "..."
    On Error GoTo fail
    Set cfg = CreateObject("CDO.Configuration")
    pre = "http://schemas.microsoft.com/cdo/configuration/"
    With cfg.Fields
        .Item(pre & "sendusing") = 2
        .Item(pre & "smtpserver") = server
        .Item(pre & "smtpserverport") = port
        .Item(pre & "smtpauthenticate") = 1
        .Item(pre & "sendusername") = em
        .Item(pre & "sendpassword") = pwd
        .Item(pre & "smtpusessl") = ssl
        .Item(pre & "smtpconnectiontimeout") = 40
        .Update
    End With
    Set m = CreateObject("CDO.Message")
    Set m.Configuration = cfg
    m.From = em
    m.To = sendTo
    m.Subject = subj
    m.TextBody = body
    m.AddAttachment path
    ProgressUpdate 0.6, "Sending to " & sendTo & "..."
    m.Send
    On Error GoTo 0
    ProgressDone "Email sent"
    If remember And Not usedSaved Then
        SetApp "B16", Obf(pwd, False)
        SetSetting "MailRemember", "Yes"
    ElseIf Not remember Then
        SetApp "B16", ""
    End If
    LogAct "Suspense file", "Emailed to " & sendTo & " from " & em
    Notify "Email sent to " & sendTo & vbCrLf & "with " & Mid$(path, InStrRev(path, "\") + 1), "ok"
    DoSendEmail = True
    Exit Function
fail:
    errTxt = Err.Description
    On Error Resume Next
    ProgressHide
    If usedSaved Then SetApp "B16", ""        ' a saved password that fails is forgotten - it will be asked next time
    MsgBox "The email could not be sent." & vbCrLf & vbCrLf & errTxt & vbCrLf & vbCrLf & _
        "Please check:" & vbCrLf & _
        "  - Gmail / Yahoo need an App password (not your normal password)." & vbCrLf & _
        "  - Mail server / Port on Settings (now " & server & " / " & port & ")." & vbCrLf & _
        "  - Internet connection." & vbCrLf & _
        IIf(IsMicrosoftMail(em, server), "  - Outlook / Hotmail may not allow this at all - use WhatsApp instead." & vbCrLf, "") & vbCrLf & _
        "The pop-up opens again so you can retry or choose WhatsApp / Save as.", vbExclamation, APP_NAME
End Function

'================ BUTTON (Change Ledger): Import filled suspense file ================
Public Sub ImportFilledSuspense()
    Dim fn As Variant, wb As Workbook, ws As Worksheet, fx As Worksheet, r As Long, lastR As Long, rr As Long, fxLast As Long
    Dim guid As String, newL As String, rem_ As String, nFill As Long, nUnknown As Long, nMiss As Long, found As Boolean
    Dim sus As String, dFrom As Variant, dTo As Variant, hdrRow As Long, cId As Long, cNew As Long, cRem As Long, c As Long, msg As String
    fn = Application.GetOpenFilename("Excel files (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm", , "Choose the filled suspense file")
    If VarType(fn) = vbBoolean Then Exit Sub
    ProgressShow "Importing filled suspense file", "Opening " & Mid$(CStr(fn), InStrRev(CStr(fn), "\") + 1) & "..."
    On Error GoTo fail
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(CStr(fn), UpdateLinks:=0, ReadOnly:=True)
    On Error Resume Next
    Set ws = wb.Worksheets("Suspense")
    On Error GoTo fail
    If ws Is Nothing Then Set ws = wb.Worksheets(1)
    For r = 1 To 15
        For c = 1 To 15
            If LCase$(CellText(ws.Cells(r, c).Value)) = "correct ledger" Then hdrRow = r: cNew = c
            If LCase$(CellText(ws.Cells(r, c).Value)) = "remarks" Then cRem = c
            If LCase$(CellText(ws.Cells(r, c).Value)) = "id" Then cId = c
        Next c
        If hdrRow > 0 Then Exit For
    Next r
    If hdrRow = 0 Then
        wb.Close SaveChanges:=False
        ThisWorkbook.Activate: Application.ScreenUpdating = True: ProgressHide
        MsgBox "This does not look like a Postify suspense file (no 'Correct Ledger' column).", vbExclamation, APP_NAME
        Exit Sub
    End If
    If cId = 0 Then cId = 11
    dFrom = ParseDateCell(ws.Range("M1").Value): dTo = ParseDateCell(ws.Range("N1").Value)
    ' make sure Change Ledger shows these Suspense entries
    Set fx = FxSheet()
    sus = SuspenseLedgerName()
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    fxLast = fx.Cells(fx.Rows.Count, S_DATE).End(xlUp).Row
    found = False
    For r = hdrRow + 1 To lastR
        guid = CellText(ws.Cells(r, cId).Value)
        If guid <> "" Then
            For rr = S_FIRST To fxLast
                If CellText(fx.Cells(rr, S_GUID).Value) = guid Then found = True: Exit For
            Next rr
        End If
        If found Then Exit For
    Next r
    If Not found And sus <> "" Then
        ProgressUpdate 0.3, "Loading the Suspense entries into Change Ledger..."
        fx.Range("B2").Value = sus
        If Not IsEmpty(dFrom) Then fx.Range("E2").Value = dFrom
        If Not IsEmpty(dTo) Then fx.Range("G2").Value = dTo
        msg = LoadFixCore(True)
        fxLast = fx.Cells(fx.Rows.Count, S_DATE).End(xlUp).Row
    End If
    ProgressUpdate 0.7, "Copying the chosen ledgers..."
    For r = hdrRow + 1 To lastR
        guid = CellText(ws.Cells(r, cId).Value)
        newL = CellText(ws.Cells(r, cNew).Value)
        rem_ = ""
        If cRem > 0 Then rem_ = CellText(ws.Cells(r, cRem).Value)
        If guid <> "" And (newL <> "" Or rem_ <> "") Then
            found = False
            For rr = S_FIRST To fxLast
                If CellText(fx.Cells(rr, S_GUID).Value) = guid Then
                    found = True
                    If newL <> "" And LCase$(Left$(CellText(fx.Cells(rr, S_STATUS).Value), 7)) <> "updated" Then
                        fx.Cells(rr, S_NEW).Value = newL
                        nFill = nFill + 1
                        If Not LedgerExists(newL) Then nUnknown = nUnknown + 1: fx.Cells(rr, S_NEW).Interior.Color = CLR_ERR
                    End If
                    If rem_ <> "" And CellText(fx.Cells(rr, S_STATUS).Value) = "" Then fx.Cells(rr, S_STATUS).Value = "Remark: " & rem_
                    FitRow fx, rr
                    Exit For
                End If
            Next rr
            If Not found Then nMiss = nMiss + 1
        End If
    Next r
    wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Application.ScreenUpdating = True
    ProgressDone "Imported"
    QuietActivate fx
    LogAct "Import filled suspense file", nFill & " ledgers from " & Mid$(CStr(fn), InStrRev(CStr(fn), "\") + 1), nFill
    msg = nFill & " correct ledger(s) copied into the yellow column." & vbCrLf
    If nUnknown > 0 Then msg = msg & nUnknown & " are not in your ledger list (red) - pick another ledger or create it in Tally and Load ledgers." & vbCrLf
    If nMiss > 0 Then msg = msg & nMiss & " entr" & IIf(nMiss = 1, "y was", "ies were") & " not found (already changed or deleted in Tally)." & vbCrLf
    Notify msg & "Check them, then click Update in Tally.", IIf(nUnknown + nMiss > 0, "warn", "ok")
    Exit Sub
fail:
    msg = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Application.ScreenUpdating = True
    ProgressHide
    MsgBox "The file could not be read: " & msg, vbExclamation, APP_NAME
End Sub

'---------------- pop-up windows (built once inside this file) ----------------
Private Function FormExists(ByVal nm As String) As Boolean
    Dim c As Object
    On Error Resume Next
    Set c = ThisWorkbook.VBProject.VBComponents(nm)
    FormExists = Not c Is Nothing
End Function

Private Function FormReady(ByVal nm As String) As Boolean
    If FormExists(nm) And CStr(AppData().Range("B18").Value) = FORMS_VER Then FormReady = True: Exit Function
    FormReady = BuildForms()
    If FormReady Then FormReady = FormExists(nm)
End Function

Public Sub BuildFormsQuiet()
    BuildForms
End Sub

Private Function AddCtl(frm As Object, ByVal progId As String, ByVal nm As String, ByVal l As Single, ByVal t As Single, _
        ByVal w As Single, ByVal h As Single, Optional ByVal cap As String = "") As Object
    Dim c As Object
    Set c = frm.Designer.Controls.Add(progId)
    c.Name = nm
    c.Left = l: c.Top = t: c.Width = w: c.Height = h
    On Error Resume Next
    c.Font.Name = "Segoe UI"
    c.Font.Size = 9
    If cap <> "" Then c.Caption = cap
    Set AddCtl = c
End Function

' Creates the two small windows (needs "Trust access to the VBA project object model", same as updates)
Private Function BuildForms() As Boolean
    Dim vbp As Object, comp As Object, c As Object, cnt As Long, code As String, nm As Variant
    On Error Resume Next
    Set vbp = ThisWorkbook.VBProject
    cnt = vbp.VBComponents.Count
    If Err.Number <> 0 Or cnt = 0 Then Exit Function
    Err.Clear
    ' remove old versions (renamed first, so the names are free straight away)
    For Each nm In Array("PfySend", "PfyPwd")
        Set comp = Nothing
        Set comp = vbp.VBComponents(nm)
        If Not comp Is Nothing Then
            comp.Name = nm & "_old" & Format$(Timer * 100, "0")
            vbp.VBComponents.Remove comp
        End If
    Next nm
    Err.Clear
    On Error GoTo fail
    ' --- PfySend
    Set comp = vbp.VBComponents.Add(3)
    comp.Name = "PfySend"
    comp.Properties("Caption").Value = APP_NAME & " - send the suspense file"
    comp.Properties("Width").Value = 500
    comp.Properties("Height").Value = 560
    comp.Properties("BackColor").Value = RGB(244, 247, 251)
    Set c = AddCtl(comp, "Forms.Label.1", "lblHead", 16, 12, 460, 22, "Send the suspense file")
    c.Font.Size = 14: c.Font.Bold = True: c.ForeColor = RGB(11, 31, 58)
    Set c = AddCtl(comp, "Forms.Label.1", "lblFile", 16, 38, 460, 18, "File ready")
    c.Font.Size = 10: c.ForeColor = RGB(100, 116, 139)
    Set c = AddCtl(comp, "Forms.Label.1", "lblTo", 16, 70, 120, 18, "Email to")
    c.Font.Size = 10.5: c.Font.Bold = True: c.ForeColor = RGB(30, 41, 59)
    Set c = AddCtl(comp, "Forms.TextBox.1", "txtTo", 140, 66, 336, 24)
    c.Font.Size = 10.5
    Set c = AddCtl(comp, "Forms.Label.1", "lblMob", 16, 104, 120, 18, "WhatsApp mobile")
    c.Font.Size = 10.5: c.Font.Bold = True: c.ForeColor = RGB(30, 41, 59)
    Set c = AddCtl(comp, "Forms.TextBox.1", "txtMob", 140, 100, 336, 24)
    c.Font.Size = 10.5
    Set c = AddCtl(comp, "Forms.Label.1", "lblWa", 16, 136, 120, 18, "WhatsApp opens in")
    c.Font.Size = 10.5: c.Font.Bold = True: c.ForeColor = RGB(30, 41, 59)
    Set c = AddCtl(comp, "Forms.OptionButton.1", "optWeb", 140, 132, 170, 22, "Browser (WhatsApp Web)")
    c.Font.Size = 10: c.GroupName = "wa"
    Set c = AddCtl(comp, "Forms.OptionButton.1", "optApp", 316, 132, 160, 22, "WhatsApp app")
    c.Font.Size = 10: c.GroupName = "wa"
    Set c = AddCtl(comp, "Forms.Label.1", "lblSub", 16, 170, 120, 18, "Subject")
    c.Font.Size = 10.5: c.Font.Bold = True: c.ForeColor = RGB(30, 41, 59)
    Set c = AddCtl(comp, "Forms.TextBox.1", "txtSub", 140, 166, 336, 24)
    c.Font.Size = 10.5
    Set c = AddCtl(comp, "Forms.Label.1", "lblMsg", 16, 204, 120, 18, "Message")
    c.Font.Size = 10.5: c.Font.Bold = True: c.ForeColor = RGB(30, 41, 59)
    Set c = AddCtl(comp, "Forms.TextBox.1", "txtMsg", 140, 200, 336, 150)
    c.MultiLine = True: c.EnterKeyBehavior = True: c.WordWrap = True: c.ScrollBars = 2: c.Font.Size = 10.5
    Set c = AddCtl(comp, "Forms.CommandButton.1", "btnReset", 140, 358, 150, 24, "Use saved details")
    c.Font.Size = 10
    Set c = AddCtl(comp, "Forms.CheckBox.1", "chkDefault", 300, 358, 180, 24, "Save these as my default")
    c.Font.Size = 10
    Set c = AddCtl(comp, "Forms.Label.1", "lblTips", 16, 392, 460, 58, _
        "Before you send:" & vbCrLf & _
        "   Email - 2-Step Verification must be ON and you need an App password (not your normal password)." & vbCrLf & _
        "   WhatsApp - log in to WhatsApp on this PC first (in the browser or the app you choose above).")
    c.Font.Size = 9.5: c.ForeColor = RGB(15, 118, 110): c.WordWrap = True
    Set c = AddCtl(comp, "Forms.Label.1", "lblNote", 16, 452, 460, 30, "")
    c.Font.Size = 9.5: c.ForeColor = RGB(180, 83, 9): c.WordWrap = True
    Set c = AddCtl(comp, "Forms.CommandButton.1", "btnEmail", 16, 486, 146, 34, "Send Email")
    c.BackColor = RGB(20, 184, 166): c.ForeColor = RGB(255, 255, 255): c.Font.Bold = True: c.Font.Size = 10.5
    Set c = AddCtl(comp, "Forms.CommandButton.1", "btnWa", 170, 486, 160, 34, "Send on WhatsApp")
    c.BackColor = RGB(37, 211, 102): c.ForeColor = RGB(255, 255, 255): c.Font.Bold = True: c.Font.Size = 10.5
    Set c = AddCtl(comp, "Forms.CommandButton.1", "btnSave", 338, 486, 138, 34, "Save as...")
    c.Font.Size = 10.5
    Set c = AddCtl(comp, "Forms.CommandButton.1", "btnClose", 388, 12, 88, 24, "Close")
    c.Font.Size = 9.5
    code = "Public Choice As String" & vbCrLf & _
        "Private Sub btnEmail_Click()" & vbCrLf & "    Choice = ""email"": Me.Hide" & vbCrLf & "End Sub" & vbCrLf & _
        "Private Sub btnWa_Click()" & vbCrLf & "    Choice = ""wa"": Me.Hide" & vbCrLf & "End Sub" & vbCrLf & _
        "Private Sub btnSave_Click()" & vbCrLf & "    Choice = ""save"": Me.Hide" & vbCrLf & "End Sub" & vbCrLf & _
        "Private Sub btnClose_Click()" & vbCrLf & "    Choice = """": Me.Hide" & vbCrLf & "End Sub" & vbCrLf & _
        "Private Sub btnReset_Click()" & vbCrLf & "    txtTo.Text = txtTo.Tag: txtMob.Text = txtMob.Tag: txtSub.Text = txtSub.Tag: txtMsg.Text = txtMsg.Tag" & vbCrLf & "End Sub" & vbCrLf & _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf & _
        "    If CloseMode = 0 Then Cancel = True: Choice = """": Me.Hide" & vbCrLf & "End Sub"
    comp.CodeModule.AddFromString code
    ' --- PfyPwd
    Set comp = vbp.VBComponents.Add(3)
    comp.Name = "PfyPwd"
    comp.Properties("Caption").Value = APP_NAME & " - mail password"
    comp.Properties("Width").Value = 400
    comp.Properties("Height").Value = 290
    comp.Properties("BackColor").Value = RGB(244, 247, 251)
    Set c = AddCtl(comp, "Forms.Label.1", "lblP", 16, 12, 360, 40, "Mail password")
    c.WordWrap = True: c.ForeColor = RGB(11, 31, 58): c.Font.Size = 10.5: c.Font.Bold = True
    Set c = AddCtl(comp, "Forms.TextBox.1", "txtPwd", 16, 58, 360, 26)
    c.PasswordChar = "*": c.Font.Size = 12
    Set c = AddCtl(comp, "Forms.CheckBox.1", "chkShow", 16, 90, 150, 22, "Show password")
    c.Font.Size = 10
    Set c = AddCtl(comp, "Forms.CheckBox.1", "chkRem", 176, 90, 200, 22, "Remember this password")
    c.Font.Size = 10
    Set c = AddCtl(comp, "Forms.Label.1", "lblHelp", 16, 118, 360, 36, _
        "Gmail / Yahoo: use the 16-letter App password (Google Account > Security > 2-Step Verification ON > App passwords) - not your normal password.")
    c.WordWrap = True: c.Font.Size = 9.5: c.ForeColor = RGB(15, 118, 110)
    Set c = AddCtl(comp, "Forms.Label.1", "lblSafe", 16, 158, 360, 36, _
        "Your password stays on this PC only. It is never shared with anyone - it goes only to your own mail server to send this email.")
    c.WordWrap = True: c.Font.Size = 9.5: c.ForeColor = RGB(71, 85, 105)
    Set c = AddCtl(comp, "Forms.CommandButton.1", "btnOK", 184, 204, 92, 30, "Send")
    c.BackColor = RGB(20, 184, 166): c.ForeColor = RGB(255, 255, 255): c.Font.Bold = True: c.Font.Size = 10.5: c.Default = True
    Set c = AddCtl(comp, "Forms.CommandButton.1", "btnCancel", 284, 204, 92, 30, "Cancel")
    c.Cancel = True: c.Font.Size = 10.5
    code = "Public Choice As String" & vbCrLf & _
        "Private Sub chkShow_Click()" & vbCrLf & "    If chkShow.Value Then txtPwd.PasswordChar = """" Else txtPwd.PasswordChar = ""*""" & vbCrLf & "End Sub" & vbCrLf & _
        "Private Sub btnOK_Click()" & vbCrLf & "    Choice = ""ok"": Me.Hide" & vbCrLf & "End Sub" & vbCrLf & _
        "Private Sub btnCancel_Click()" & vbCrLf & "    Choice = """": Me.Hide" & vbCrLf & "End Sub" & vbCrLf & _
        "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf & _
        "    If CloseMode = 0 Then Cancel = True: Choice = """": Me.Hide" & vbCrLf & "End Sub"
    comp.CodeModule.AddFromString code
    SetApp "A18", "Pop-up windows version"
    SetApp "B18", FORMS_VER
    BuildForms = True
    Exit Function
fail:
    BuildForms = False
End Function

'---------------- one-time setup of v2.3 ----------------
Private Sub ApplyLayout230()
    Dim st As Worksheet, fx As Worksheet, hl As Worksheet, shp As Object, rg As Range, r As Long, i As Long, items As Variant, f As Range
    On Error Resume Next
    If CStr(AppData().Range("B19").Value) = "2.3.0" Then Exit Sub
    ' Settings: Email & WhatsApp block
    Set st = ThisWorkbook.Worksheets("Settings")
    st.Range("A3:C3").Copy Destination:=st.Range("A18:C18")
    st.Range("A18").Value = "Email & WhatsApp": st.Range("B18").Value = "Your value": st.Range("C18").Value = "Used by Change Ledger > Make suspense file"
    items = Array( _
        Array("MailFrom", "Your email (sender)", "", "The suspense file is sent from this mailbox. Gmail, Yahoo, Zoho and Rediff are set up automatically."), _
        Array("MailServer", "Mail server (SMTP)", "", "Filled automatically for Gmail / Yahoo / Zoho / Rediff. Company mail: ask your mail provider (e.g. smtp.yourdomain.com)."), _
        Array("MailPort", "Port", "", "Usually 465 (secure). Filled automatically with the server."), _
        Array("MailSSL", "Secure connection (SSL)", "Yes", "Keep Yes."), _
        Array("MailRemember", "Remember mail password", "Yes", "Yes = typed once and kept in this file. No = asked at every send (safer on a shared file). Gmail / Yahoo need an App password."), _
        Array("SendToEmail", "Send suspense file to (email)", "", "Saved recipient; you can change it in the pop-up for each send."), _
        Array("SendToMobile", "Send on WhatsApp to (mobile)", "", "10-digit mobile (91 is added). You can change it in the pop-up."), _
        Array("SendSubject", "Email subject", "Suspense entries {period} - please confirm the ledgers", "{period}, {count} and {company} are filled in automatically."), _
        Array("SendMessage", "Message (email & WhatsApp)", "Dear Sir/Madam," & vbLf & vbLf & "Please find attached {count} bank entries kept in Suspense in our books for {period}." & vbLf & _
            "Kindly select the correct ledger for each entry in the yellow column (dropdown), add remarks where needed, and send the file back." & vbLf & vbLf & "Regards," & vbLf & "{company}", _
            "Text of the email / WhatsApp message."))
    For i = 0 To UBound(items)
        r = 19 + i
        st.Range("A13:C13").Copy Destination:=st.Range("A" & r & ":C" & r)
        st.Cells(r, 1).Value = items(i)(1)
        If CellText(st.Cells(r, 2).Value) = "" Or st.Cells(r, 2).Value = st.Range("B13").Value Then st.Cells(r, 2).Value = items(i)(2)
        st.Cells(r, 3).Value = items(i)(3)
        ThisWorkbook.Names.Add Name:=items(i)(0), RefersTo:="=Settings!$B$" & r
    Next i
    st.Range("B27").WrapText = True: st.Rows(27).RowHeight = 110
    st.Range("B26").WrapText = True
    For i = 22 To 23
        st.Cells(i, 2).Validation.Delete
        st.Cells(i, 2).Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="Yes,No"
    Next i
    ' Change Ledger: two new buttons
    Set fx = FxSheet()
    For Each shp In fx.Shapes
        If shp.Name = "cl_mkfile" Or shp.Name = "cl_import" Then shp.Delete
    Next shp
    Set rg = fx.Range("I4")
    Set shp = NewBox(fx, "cl_mkfile", 1, rg.Left + 1, rg.Top + 1, rg.Width - 2, rg.Height - 2, RGB(20, 184, 166))
    StyleText shp, ChrW(&H2709) & "  Make suspense file", 10, RGB(255, 255, 255), True, 2
    shp.TextFrame2.VerticalAnchor = 3: shp.OnAction = "TallyLink.MakeSuspenseFile": shp.Placement = 1
    Set rg = fx.Range("J4")
    Set shp = NewBox(fx, "cl_import", 1, rg.Left + 1, rg.Top + 1, rg.Width - 2, rg.Height - 2, RGB(230, 250, 246))
    StyleText shp, ChrW(&H21E9) & "  Import filled suspense file", 10, RGB(15, 118, 110), True, 2
    shp.TextFrame2.VerticalAnchor = 3: shp.OnAction = "TallyLink.ImportFilledSuspense": shp.Placement = 1
    ' Help section
    Set hl = ThisWorkbook.Worksheets("Help")
    Set f = hl.Columns(2).Find(What:="Suspense file by Email / WhatsApp", LookIn:=-4163, LookAt:=1)
    If f Is Nothing Then
        items = Array( _
            "Change Ledger > Make suspense file: an Excel of all Suspense entries (Change Ledger From / To dates) with a yellow 'Correct Ledger' dropdown of every ledger and a Remarks column. A pop-up then offers Send Email, Send on WhatsApp or Save as.", _
            "Send Email sends from your own mailbox (Settings > Email & WhatsApp). Gmail / Yahoo: turn on 2-Step Verification in your account and create an 'App password' (16 letters) - use that, not your normal password. Company mail: fill Mail server and Port.", _
            "Outlook / Hotmail / Microsoft 365 may not allow sending from here - use WhatsApp. WhatsApp opens with the message typed in and the file's folder open: drag the file into the chat and press Send.", _
            "When the file comes back: Change Ledger > Import filled suspense file. The chosen ledgers go into the yellow column - check them and click Update in Tally.")
        r = hl.Cells(hl.Rows.Count, 2).End(xlUp).Row + 3
        hl.Cells(r, 2).Value = "Suspense file by Email / WhatsApp"
        With hl.Cells(r, 2)
            .Font.Bold = True: .Font.Size = 13: .Font.Color = RGB(11, 31, 58): .VerticalAlignment = -4107
        End With
        hl.Rows(r).RowHeight = 26
        For i = 0 To UBound(items)
            With hl.Cells(r + 1 + i, 2)
                .Value = items(i)
                .WrapText = True: .VerticalAlignment = -4160: .IndentLevel = 1
                .Interior.Color = RGB(255, 255, 255): .Font.Color = RGB(30, 41, 59): .Font.Size = 10
            End With
            hl.Rows(r + 1 + i).RowHeight = 16 * (Len(items(i)) \ 120 + 1) + 8
        Next i
        GridBorders hl.Range(hl.Cells(r + 1, 2), hl.Cells(r + 1 + UBound(items), 2))
    End If
    SetApp "A19", "Layout 2.3.0"
    SetApp "B19", "2.3.0"
End Sub



'==============================================================
'  v2.3.1: info cards, WhatsApp app / browser, dark mode on every sheet
'==============================================================
' Readable card on the sheet. lines: "#Heading" or text. buttons: Array(caption, macro, "teal" / "white")
Public Sub InfoCard(ByVal title As String, ByVal subTitle As String, ByVal lines As Variant, ByVal buttons As Variant)
    Dim ws As Object, shp As Object, card As Object, lst As Object, i As Long, txt As String, fs As Single, tries As Long
    Dim l As Single, t As Single, w As Single, h As Single, cw As Single, cl As Single, ct As Single, y As Single, bx As Single, bw As Single
    On Error GoTo fin
    InfoCardHide
    ToastHide
    Set ws = ActiveSheet
    CoverRect l, t, w, h
    Set shp = NewBox(ws, "ic_ov", 1, l, t, w, h, RGB(11, 31, 58))
    shp.Fill.Transparency = 0.4
    shp.OnAction = "TallyLink.WhatsNewNoop"
    ViewRect l, t, w, h
    cw = 620
    If cw > w - 20 Then cw = w - 20
    If cw < 320 Then cw = 320
    cl = l + (w - cw) / 2: ct = t + 12
    If cl < l + 4 Then cl = l + 4
    Set card = NewBox(ws, "ic_card", 5, cl, ct, cw, 300, RGB(255, 255, 255))
    card.Adjustments.Item(1) = 0.03
    card.Shadow.Visible = -1: card.Shadow.Blur = 22: card.Shadow.OffsetY = 5: card.Shadow.Transparency = 0.7
    card.OnAction = "TallyLink.WhatsNewNoop"
    Set shp = NewBox(ws, "ic_bar", 1, cl, ct, cw, 6, RGB(20, 184, 166))
    Set shp = NewText(ws, "ic_title", cl + 30, ct + 24, cw - 60, 28)
    StyleText shp, title, 16, RGB(11, 31, 58), True
    shp.OnAction = "TallyLink.WhatsNewNoop"
    Set shp = NewText(ws, "ic_sub", cl + 30, ct + 54, cw - 60, 20)
    StyleText shp, subTitle, 10.5, RGB(100, 116, 139)
    shp.OnAction = "TallyLink.WhatsNewNoop"
    Set shp = NewBox(ws, "ic_line", 1, cl + 30, ct + 82, cw - 60, 1, RGB(226, 232, 240))
    For i = 0 To UBound(lines)
        If Left$(lines(i), 1) = "#" Then
            txt = txt & IIf(i > 0, vbCr, "") & Mid$(lines(i), 2)
        Else
            txt = txt & IIf(i > 0, vbCr, "") & lines(i)
        End If
    Next i
    fs = 11
    Set lst = NewText(ws, "ic_body", cl + 30, ct + 94, cw - 60, 40)
    For tries = 1 To 3
        StyleText lst, txt, fs, RGB(30, 41, 59)
        With lst.TextFrame2
            .TextRange.ParagraphFormat.SpaceAfter = fs * 0.55
            .TextRange.ParagraphFormat.LineRuleWithin = -1
            .TextRange.ParagraphFormat.SpaceWithin = 1.12
            For i = 0 To UBound(lines)
                If Left$(lines(i), 1) = "#" Then
                    With .TextRange.Paragraphs(i + 1)
                        .Font.Bold = -1
                        .Font.Size = fs + 1.5
                        .Font.Fill.ForeColor.RGB = RGB(15, 118, 110)
                        If i > 0 Then .ParagraphFormat.SpaceBefore = fs * 0.9
                    End With
                End If
            Next i
            .AutoSize = 1
        End With
        If lst.Top + lst.Height + 80 < t + h - 8 Or fs <= 9 Then Exit For
        fs = fs - 1
    Next tries
    lst.OnAction = "TallyLink.WhatsNewNoop"
    y = lst.Top + lst.Height + 18
    bx = cl + cw - 30
    For i = UBound(buttons) To 0 Step -1
        bw = 60 + Len(buttons(i)(0)) * 7.2
        bx = bx - bw
        Set shp = NewBox(ws, "ic_btn" & i, 5, bx, y, bw, 36, IIf(buttons(i)(2) = "teal", RGB(20, 184, 166), RGB(255, 255, 255)))
        shp.Adjustments.Item(1) = 0.25
        If buttons(i)(2) <> "teal" Then shp.Line.Visible = -1: shp.Line.ForeColor.RGB = RGB(20, 184, 166): shp.Line.Weight = 1.25
        StyleText shp, buttons(i)(0), 11, IIf(buttons(i)(2) = "teal", RGB(255, 255, 255), RGB(15, 118, 110)), True, 2
        shp.TextFrame2.VerticalAnchor = 3
        shp.OnAction = buttons(i)(1)
        bx = bx - 12
    Next i
    card.Height = y + 36 + 26 - ct
    Application.ScreenUpdating = True
    Exit Sub
fin:
    InfoCardHide
End Sub

Public Sub InfoCardHide()
    DeleteShapes "ic_"
End Sub

' Shown once after the update (after What's new > Got it)
Private Sub MailIntro()
    InfoCard "New: send your Suspense entries by Email or WhatsApp", "Change Ledger > Make suspense file  -  one click, ready to send", Array( _
        "#What it does", _
        "Makes an Excel of all your Suspense entries with a dropdown of every ledger. The person picks the right ledgers and sends the file back.", _
        "Send it by Email, on WhatsApp, or Save as. When it comes back: Change Ledger > Import filled suspense file.", _
        "#Before you send by Email", _
        "1.   Turn ON 2-Step Verification in your Google / Yahoo account.", _
        "2.   Create an App password: open myaccount.google.com/apppasswords, type a name (e.g. Postify) and click Create.", _
        "3.   Use those 16 letters as the mail password in Postify - not your normal Gmail password.", _
        "Company mail: enter your mail server in Settings > Email & WhatsApp.   Outlook / Hotmail may not work - use WhatsApp.", _
        "#Before you send on WhatsApp", _
        "Log in to WhatsApp on this PC first - WhatsApp Web in your browser, or the WhatsApp app. You choose which one in the pop-up.", _
        "#Your password", _
        "It stays on this PC only and is never shared with anyone. Should Postify remember it after you type it once?"), _
        Array(Array("Ask me every time", "TallyLink.MailIntroAsk", "white"), Array("Remember my password", "TallyLink.MailIntroRemember", "teal"))
End Sub

Public Sub MailIntroRemember()
    MailIntroDone "Yes"
End Sub

Public Sub MailIntroAsk()
    MailIntroDone "No"
End Sub

Private Sub MailIntroDone(ByVal remember As String)
    InfoCardHide
    SetSetting "MailRemember", remember
    SetApp "A17", "Mail intro shown"
    SetApp "B17", "card"
    If Not FormReady("PfySend") Then BuildForms
    GoSheet "Settings", "B19"
    Notify "Enter your email and the person's email / WhatsApp number here." & vbCrLf & _
        "Password: " & IIf(remember = "Yes", "remembered after the first send.", "asked at every send.") & " Change it any time in row 23.", "ok"
End Sub

Private Sub ApplyLayout231()
    Dim home As Worksheet, st As Worksheet, shp As Object, rg As Range, mv As Variant, i As Long, r1 As Long, r2 As Long
    On Error Resume Next
    If CStr(AppData().Range("B20").Value) = "2.3.1" Then Exit Sub
    ' Settings: where WhatsApp opens
    Set st = ThisWorkbook.Worksheets("Settings")
    st.Range("A13:C13").Copy Destination:=st.Range("A28:C28")
    st.Range("A28").Value = "WhatsApp opens in"
    st.Range("B28").Value = "Browser"
    st.Range("C28").Value = "Browser = WhatsApp Web (uses your browser login). App = the WhatsApp app. You can also choose in the pop-up."
    st.Range("B28").Validation.Delete
    st.Range("B28").Validation.Add Type:=3, AlertStyle:=1, Operator:=1, Formula1:="Browser,App"
    ThisWorkbook.Names.Add Name:="WaOpenIn", RefersTo:="=Settings!$B$28"
    ' Home: Updates & info buttons move up one row each (no empty row under Last updated)
    Set home = ThisWorkbook.Worksheets("Home")
    home.Unprotect SHEET_PW
    If CellText(home.Range("Y33").Value) = "" And InStr(CellText(home.Range("Y34").Value), "Check for updates") > 0 Then
        mv = Array(Array(34, 33), Array(36, 35), Array(37, 36))
        For i = 0 To 2
            r1 = mv(i)(0): r2 = mv(i)(1)
            home.Range("Y" & r1 & ":AQ" & r1).Cut home.Range("Y" & r2)
            home.Rows(r2).RowHeight = IIf(r2 = 36, 24, 26)
            For Each shp In home.Shapes
                If shp.TopLeftCell.Row = r1 And shp.TopLeftCell.Column >= 25 And shp.Name <> "theme_btn" Then
                    shp.Top = home.Rows(r2).Top
                    shp.Height = home.Rows(r2).Height
                End If
            Next shp
        Next i
        home.Rows(37).RowHeight = 14
    End If
    home.Protect Password:=SHEET_PW, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    home.EnableSelection = 0
    SetApp "A20", "Layout 2.3.1"
    SetApp "B20", "2.3.1"
End Sub

'---------------- dark mode for every sheet ----------------
Private Function IsDark() As Boolean
    IsDark = (LCase$(CStr(AppData().Range("B14").Value)) = "dark")
End Function

' 1x1 dark picture used as the sheet background (fills everything that has no colour of its own)
Private Function DarkPng() As String
    Dim hx As String, b() As Byte, i As Long, f As Integer, p As String
    p = Environ$("TEMP") & "\postify_dark_bg.png"
    If Dir(p) <> "" Then DarkPng = p: Exit Function
    hx = "89504E470D0A1A0A0000000D4948445200000001000000010802000000907753DE0000000C49444154789C63E0165200000069003E33E861380000000049454E44AE426082"
    ReDim b(0 To Len(hx) \ 2 - 1)
    For i = 0 To UBound(b)
        b(i) = CByte("&H" & Mid$(hx, i * 2 + 1, 2))
    Next i
    f = FreeFile
    Open p For Binary Access Write As #f
    Put #f, , b
    Close #f
    DarkPng = p
End Function

Private Sub SheetBackground(ws As Worksheet, ByVal dark As Boolean)
    On Error Resume Next
    If dark Then
        ws.SetBackgroundPicture DarkPng()
    Else
        ws.SetBackgroundPicture ""
    End If
End Sub

' New sheets made while dark mode is on (reports) get the dark background too
Private Sub DarkIfNeeded(ws As Worksheet)
    If IsDark() Then SheetBackground ws, True
End Sub

Private Sub SwapPairs(rg As Range, ByVal toDark As Boolean)
    Dim pairs As Variant, i As Long, a As Long, b As Long
    pairs = Array( _
        Array(RGB(244, 247, 251), RGB(11, 18, 32), True, True, False), _
        Array(RGB(255, 255, 255), RGB(21, 32, 51), True, False, False), _
        Array(RGB(230, 250, 246), RGB(15, 59, 58), True, False, False), _
        Array(RGB(11, 31, 58), RGB(226, 232, 240), False, True, False), _
        Array(RGB(30, 41, 59), RGB(203, 213, 225), False, True, False), _
        Array(RGB(100, 116, 139), RGB(148, 163, 184), False, True, False), _
        Array(RGB(15, 118, 110), RGB(94, 234, 212), False, True, False), _
        Array(RGB(226, 232, 240), RGB(51, 65, 85), False, False, True))
    For i = 0 To UBound(pairs)
        If toDark Then a = pairs(i)(0): b = pairs(i)(1) Else a = pairs(i)(1): b = pairs(i)(0)
        SwapColour rg, a, b, pairs(i)(2), pairs(i)(3), pairs(i)(4)
    Next i
End Sub

' Coloured cells keep dark text so they stay readable on the dark background
Private Sub DarkTextOnLightFills(ws As Worksheet)
    Dim fills As Variant, c As Variant
    On Error Resume Next
    fills = Array(RGB(255, 255, 255), RGB(255, 248, 225), RGB(255, 242, 204), RGB(230, 250, 246), RGB(244, 247, 251), RGB(226, 232, 240), _
        RGB(241, 245, 249), RGB(254, 243, 199), CLR_OK, CLR_INFO, CLR_ERR, CLR_TALLY, CLR_TALLYONLY, 13431551)
    For Each c In fills
        Application.FindFormat.Clear: Application.ReplaceFormat.Clear
        Application.FindFormat.Interior.Color = c
        Application.FindFormat.Font.ColorIndex = -4105
        Application.ReplaceFormat.Font.Color = RGB(30, 41, 59)
        ws.UsedRange.Replace What:="", Replacement:="", LookAt:=2, SearchFormat:=True, ReplaceFormat:=True
    Next c
    Application.FindFormat.Clear: Application.ReplaceFormat.Clear
End Sub

Private Sub GridColour(ws As Worksheet, ByVal toDark As Boolean)
    Dim e As Variant
    On Error Resume Next
    For Each e In Array(7, 8, 9, 10)
        Application.FindFormat.Clear: Application.ReplaceFormat.Clear
        If toDark Then
            Application.FindFormat.Borders(e).ColorIndex = -4105
            Application.ReplaceFormat.Borders(e).Color = RGB(71, 85, 105)
        Else
            Application.FindFormat.Borders(e).Color = RGB(71, 85, 105)
            Application.ReplaceFormat.Borders(e).ColorIndex = -4105
        End If
        ws.UsedRange.Replace What:="", Replacement:="", LookAt:=2, SearchFormat:=True, ReplaceFormat:=True
    Next e
    Application.FindFormat.Clear: Application.ReplaceFormat.Clear
End Sub

' Button: Dark mode / Light mode - now for the whole workbook
Public Sub ToggleTheme()
    Dim ws As Worksheet, toDark As Boolean, lastR As Long, n As Long, k As Long, page As Boolean
    On Error Resume Next
    toDark = Not IsDark()
    ProgressShow IIf(toDark, "Switching to dark mode", "Switching to light mode"), "Restyling the sheets..."
    Application.ScreenUpdating = False
    n = ThisWorkbook.Worksheets.Count
    For Each ws In ThisWorkbook.Worksheets
        k = k + 1
        ProgressUpdate k / n, "Restyling " & ws.Name & "..."
        If ws.Name <> "AppData" And ws.Name <> "TallyDebug" Then
            ws.Unprotect SHEET_PW
            page = (ws.Name = "Home" Or ws.Name = "Help" Or ws.Name = "Update History")
            If toDark And Not page Then DarkTextOnLightFills ws
            If page Then
                SwapPairs ws.Range("A1:AZ" & (ws.UsedRange.Row + ws.UsedRange.Rows.Count + 5)), toDark
                ' rows below the used area (they only have the column colour)
                lastR = ws.UsedRange.Row + ws.UsedRange.Rows.Count
                If ws.Name <> "Home" Then ws.Range("A" & lastR & ":E" & lastR + 300).Interior.Color = IIf(toDark, RGB(11, 18, 32), RGB(244, 247, 251))
            Else
                GridColour ws, toDark
                ' grey notes / descriptions: a lighter grey on the dark background
                If toDark Then
                    SwapColour ws.UsedRange, RGB(100, 116, 139), RGB(148, 163, 184), False, True, False
                Else
                    SwapColour ws.UsedRange, RGB(148, 163, 184), RGB(100, 116, 139), False, True, False
                End If
            End If
            SheetBackground ws, toDark
            If ws.Name = "Home" Then
                ws.Protect Password:=SHEET_PW, DrawingObjects:=False, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
                ws.EnableSelection = 0
            End If
        End If
    Next ws
    ' text without its own colour: light in dark mode, normal black in light mode
    If toDark Then
        ThisWorkbook.Styles("Normal").Font.Color = RGB(226, 232, 240)
    Else
        ThisWorkbook.Styles("Normal").Font.ColorIndex = -4105
    End If
    SetApp "A14", "Theme"
    SetApp "B14", IIf(toDark, "dark", "light")
    ThemeButton
    Application.ScreenUpdating = True
    ProgressDone IIf(toDark, "Dark mode on", "Light mode on")
    LogAct "Theme", IIf(toDark, "Dark mode", "Light mode")
End Sub
