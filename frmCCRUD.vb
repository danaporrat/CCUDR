Imports System
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Data.Odbc

Imports Microsoft.VisualBasic
Imports VBto

Public Class frmCCUDR
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdCRUD As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
    Friend WithEvents cmdCRUD_5 As System.Windows.Forms.Button
    Friend WithEvents cmdCRUD_4 As System.Windows.Forms.Button
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtCOD As System.Windows.Forms.TextBox
    Friend WithEvents grdA As AxMSHierarchicalFlexGridLib.AxMSHFlexGrid
    Friend WithEvents cmdCRUD_1 As System.Windows.Forms.Button
    Friend WithEvents cmdCRUD_2 As System.Windows.Forms.Button
    Friend WithEvents cmdCRUD_3 As System.Windows.Forms.Button
    Friend WithEvents cmdCRUD_0 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCCUDR))
        Me.cmdCRUD = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(Me.components)
        Me.cmdCRUD_5 = New System.Windows.Forms.Button()
        Me.cmdCRUD_4 = New System.Windows.Forms.Button()
        Me.cmdCRUD_1 = New System.Windows.Forms.Button()
        Me.cmdCRUD_2 = New System.Windows.Forms.Button()
        Me.cmdCRUD_3 = New System.Windows.Forms.Button()
        Me.cmdCRUD_0 = New System.Windows.Forms.Button()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.txtCOD = New System.Windows.Forms.TextBox()
        Me.grdA = New AxMSHierarchicalFlexGridLib.AxMSHFlexGrid()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        CType(Me.cmdCRUD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grdA, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdCRUD
        '
        '
        'cmdCRUD_5
        '
        Me.cmdCRUD_5.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCRUD_5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCRUD.SetIndex(Me.cmdCRUD_5, CType(5, Short))
        Me.cmdCRUD_5.Location = New System.Drawing.Point(251, 388)
        Me.cmdCRUD_5.Name = "cmdCRUD_5"
        Me.cmdCRUD_5.Size = New System.Drawing.Size(66, 50)
        Me.cmdCRUD_5.TabIndex = 11
        Me.cmdCRUD_5.Text = "Error"
        Me.cmdCRUD_5.UseVisualStyleBackColor = False
        '
        'cmdCRUD_4
        '
        Me.cmdCRUD_4.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCRUD_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCRUD.SetIndex(Me.cmdCRUD_4, CType(4, Short))
        Me.cmdCRUD_4.Location = New System.Drawing.Point(251, 324)
        Me.cmdCRUD_4.Name = "cmdCRUD_4"
        Me.cmdCRUD_4.Size = New System.Drawing.Size(66, 50)
        Me.cmdCRUD_4.TabIndex = 9
        Me.cmdCRUD_4.Text = "Read by Name"
        Me.cmdCRUD_4.UseVisualStyleBackColor = False
        '
        'cmdCRUD_1
        '
        Me.cmdCRUD_1.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCRUD_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCRUD.SetIndex(Me.cmdCRUD_1, CType(1, Short))
        Me.cmdCRUD_1.Location = New System.Drawing.Point(251, 129)
        Me.cmdCRUD_1.Name = "cmdCRUD_1"
        Me.cmdCRUD_1.Size = New System.Drawing.Size(66, 50)
        Me.cmdCRUD_1.TabIndex = 3
        Me.cmdCRUD_1.Text = "Read"
        Me.cmdCRUD_1.UseVisualStyleBackColor = False
        '
        'cmdCRUD_2
        '
        Me.cmdCRUD_2.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCRUD_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCRUD.SetIndex(Me.cmdCRUD_2, CType(2, Short))
        Me.cmdCRUD_2.Location = New System.Drawing.Point(251, 194)
        Me.cmdCRUD_2.Name = "cmdCRUD_2"
        Me.cmdCRUD_2.Size = New System.Drawing.Size(66, 50)
        Me.cmdCRUD_2.TabIndex = 2
        Me.cmdCRUD_2.Text = "Update"
        Me.cmdCRUD_2.UseVisualStyleBackColor = False
        '
        'cmdCRUD_3
        '
        Me.cmdCRUD_3.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCRUD_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCRUD.SetIndex(Me.cmdCRUD_3, CType(3, Short))
        Me.cmdCRUD_3.Location = New System.Drawing.Point(251, 259)
        Me.cmdCRUD_3.Name = "cmdCRUD_3"
        Me.cmdCRUD_3.Size = New System.Drawing.Size(66, 50)
        Me.cmdCRUD_3.TabIndex = 1
        Me.cmdCRUD_3.Text = "Delete"
        Me.cmdCRUD_3.UseVisualStyleBackColor = False
        '
        'cmdCRUD_0
        '
        Me.cmdCRUD_0.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCRUD_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCRUD.SetIndex(Me.cmdCRUD_0, CType(0, Short))
        Me.cmdCRUD_0.Location = New System.Drawing.Point(251, 65)
        Me.cmdCRUD_0.Name = "cmdCRUD_0"
        Me.cmdCRUD_0.Size = New System.Drawing.Size(66, 50)
        Me.cmdCRUD_0.TabIndex = 0
        Me.cmdCRUD_0.Text = "Create"
        Me.cmdCRUD_0.UseVisualStyleBackColor = False
        '
        'txtName
        '
        Me.txtName.BackColor = System.Drawing.SystemColors.Window
        Me.txtName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtName.Location = New System.Drawing.Point(339, 24)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(66, 20)
        Me.txtName.TabIndex = 8
        '
        'txtCOD
        '
        Me.txtCOD.BackColor = System.Drawing.SystemColors.Window
        Me.txtCOD.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCOD.Location = New System.Drawing.Point(235, 24)
        Me.txtCOD.Name = "txtCOD"
        Me.txtCOD.Size = New System.Drawing.Size(50, 20)
        Me.txtCOD.TabIndex = 6
        '
        'grdA
        '
        Me.grdA.CausesValidation = False
        Me.grdA.DataSource = Nothing
        Me.grdA.Location = New System.Drawing.Point(16, 49)
        Me.grdA.Name = "grdA"
        Me.grdA.OcxState = CType(resources.GetObject("grdA.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdA.Size = New System.Drawing.Size(171, 583)
        Me.grdA.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(16, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(171, 24)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Choose a line and then an action"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(291, 22)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(42, 24)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "NAME"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(193, 22)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(47, 24)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "CODE"
        '
        'frmCCUDR
        '
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(986, 817)
        Me.Controls.Add(Me.cmdCRUD_5)
        Me.Controls.Add(Me.cmdCRUD_4)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.txtCOD)
        Me.Controls.Add(Me.grdA)
        Me.Controls.Add(Me.cmdCRUD_1)
        Me.Controls.Add(Me.cmdCRUD_2)
        Me.Controls.Add(Me.cmdCRUD_3)
        Me.Controls.Add(Me.cmdCRUD_0)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Name = "frmCCUDR"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Connect Create  Read Update Delete "
        CType(Me.cmdCRUD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grdA, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

	'=========================================================

    Public Enum ARAZOTE
        ARAZOT_EREZ_ID
        ARAZOT_EREZ_DSC
        ARAZOT_EREZ_ISO3
        ARAZOT_EREZ_ENG
        ARAZOT_BESICUN
    End Enum

    Public glblConnection As New OleDb.OleDbConnection
    Public glblConnectionODBC As New Odbc.OdbcConnection
    Dim Rrset As New DataTable
    Dim adaptor As OleDbDataAdapter
    Dim strConnectGlbl As String
    Dim strFile As String = ""
    Dim bFox As Boolean
    Public glblArazotRset As New DataTable
    Dim strSql As String

    Private Sub cmdCRUD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCRUD.Click
        Dim Index As Short = cmdCRUD.GetIndex(sender)
        Dim strWhere As String
        Dim strAction As String
        Dim tempRset As New DataTable()
        Dim Command As OleDbCommand
        Dim currentRows() As DataRow
        Dim iAffected As Integer

        strWhere = " where erez_id = '" & grdA.get_TextMatrix(grdA.Row, 0) & "';"
        Select Case Index
            Case 0:
                'create
                If (testInput(False) = True) Then
                    strAction = "insert into ARAZOT (EREZ_ID, EREZ_DSC, EREZ_ISO3, EREZ_ENG, BESICUN)" & "values ('" & txtCOD.Text & "','" & txtName.Text & "','" & txtCOD.Text & Strings.Right(txtCOD.Text, 1) & "','" & txtName.Text & "', '0')"
                    Command = New OleDbCommand(strAction, glblConnection) : iAffected = Command.ExecuteNonQuery()
                    'glblConnection.ExecuteNonQuery(strAction)
                    'Rrset.Requery()
                    Rrset = Nothing
                    Rrset = New DataTable("ARAZOT")
                    adaptor.Fill(Rrset)
                    SetgrdA()
                    GetErezCheck()
                End If
            Case 1
                'read
                'strAction = " select * from ARAZOT " & strWhere
                If (txtCOD.Text <> "") Then
                    If (testInput(True) = True) Then
                        strAction = "erez_id=" & "'" & txtCOD.Text & "'"

                        'tempRset.Open(strAction, glblConnection) 
                        'If (Not tempRset.EOF) Then
                        '    txtCOD.Text = Trim(Convert.ToString(tempRset.Fields(ARAZOTE.ARAZOT_EREZ_ID).Value))
                        '    txtName.Text = Trim(Convert.ToString(tempRset.Fields(ARAZOTE.ARAZOT_EREZ_ENG).Value))
                        currentRows = Rrset.Select(strAction)
                        If (currentRows.Count > 0) Then
                            With currentRows(0)
                                txtCOD.Text = Trim(currentRows(0)(ARAZOTE.ARAZOT_EREZ_ID))
                                txtName.Text = Trim(currentRows(0)(ARAZOTE.ARAZOT_EREZ_ENG))
                            End With
                        Else
                            MsgBox("Not found")
                        End If
                    Else
                        MsgBox("ציין קוד ארץ")
                    End If
                End If
            Case 2
                'update
                If (testInput(True) = True) Then

                    strAction = "Update ARAZOT set EREZ_ENG = '" & txtName.Text & "'" & " where EREZ_ID = '" & txtCOD.Text & "'"
                    'glblConnection.ExecuteNonQuery(strAction)
                    'Rrset.Requery()
                    Command = New OleDbCommand(strAction, glblConnection) : iAffected = Command.ExecuteNonQuery()
                    Rrset = Nothing
                    Rrset = New DataTable("ARAZOT")
                    adaptor.Fill(Rrset)
                    SetgrdA()
                    GetErezCheck()
                End If
            Case 3
                'delete
                If (testInput(True) = True) Then

                    strAction = "Delete from ARAZOT" & " where EREZ_ID = '" & txtCOD.Text & "'"
                    Command = New OleDbCommand(strAction, glblConnection) : iAffected = Command.ExecuteNonQuery()
                    Rrset = Nothing
                    Rrset = New DataTable("ARAZOT")
                    adaptor.Fill(Rrset)
                    SetgrdA()
                    GetErezCheck()
                End If
            Case 4
                'read by name
                If (testInput(True) = True) Then

                    strAction = " EREZ_ID = '" & txtCOD.Text & "'"
                    'tempRset.Open(strAction, glblConnection)
                    'If (Not tempRset.EOF) Then
                    '    txtCOD.Text = Trim(Convert.ToString(tempRset.Fields("EREZ_ID").Value))
                    '    txtName.Text = Trim(Convert.ToString(tempRset.Fields("EREZ_ENG").Value))
                    currentRows = Rrset.Select(strAction)
                    If (currentRows.Count > 0) Then
                        With currentRows(0)
                            txtCOD.Text = Trim(currentRows(0)("EREZ_ID"))
                            txtName.Text = Trim(currentRows(0)("EREZ_ENG"))
                        End With
                    Else
                        MsgBox("Not found")
                    End If
                End If
            Case 5
                'deliberate error
                Try
                    strAction = " EREZ_ID1 = '" & txtCOD.Text & "'"
                    'tempRset.Open(strAction, glblConnection)
                    'If (Not tempRset.EOF) Then
                    '    txtCOD.Text = Trim(Convert.ToString(tempRset.Fields("EREZ_ID").Value))
                    '    txtName.Text = Trim(Convert.ToString(tempRset.Fields("EREZ_ENG").Value))
                    currentRows = Rrset.Select(strAction)
                    If (currentRows.Count > 0) Then
                        With currentRows(0)
                            txtCOD.Text = Trim(currentRows(0)("EREZ_ID"))
                            txtName.Text = Trim(currentRows(0)("EREZ_ENG"))
                        End With
                    Else
                        MsgBox("Not found")
                    End If
                Catch ex As OleDbException
                    OLEerrors(ex)
                Catch ex As Exception
                    ErrLogLine("", "", ex)
                End Try


        End Select
    End Sub

    Private Sub frmCCUDR_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bFox = True

        strConnectGlbl = "Provider=vfpoledb.1;" & "data Source=" & Application.StartupPath
        glblConnection.ConnectionString = strConnectGlbl
        glblConnection.Open()

        '- Rrset.CursorType = VBto.CursorTypeEnum.adOpenDynamic
        '- Rrset.CursorLocation = VBto.CursorLocationEnum.adUseClient
        '- Rrset.LockType = Stub_ADODB_LockTypeEnum_adLockOptimistic

        strFile = "ARAZOT"
        'Rrset.Open("select * from " & strFile & " order by EREZ_ID;", glblConnection)
        adaptor = New OleDbDataAdapter("Select * from ARAZOT order by EREZ_ID;", glblConnection)
        adaptor.Fill(Rrset)
        SetgrdA()
        GetErezCheck()
    End Sub

    Private Sub GetErezCheck()
        If (OpenReadRecordSet("ARAZOT", "", glblArazotRset, "*", " order by erez_id", 1, strSql) = False) Then
            MsgBox("הטבלה " & " ARAZOT " & "לא נפתח")
            Exit Sub
        End If
    End Sub

    Private Sub SetgrdA()
        Dim iRow As Integer

        With grdA
            .set_Cols(2)
            .Rows = IIf(Rrset.Rows.Count = 0, 2, Rrset.Rows.Count + 1)
            .set_TextMatrix(0, 0, "CODE")
            .set_TextMatrix(0, 1, "NAME   ")
        End With
        With Rrset
            'If Not .BOF Then .MoveFirst()
            If (.Rows.Count > 0) Then
                iRow = 1
                Dim reader As DataTableReader = .CreateDataReader

                Do While reader.Read
                    grdA.set_TextMatrix(iRow, 0, Trim(reader(ARAZOTE.ARAZOT_EREZ_ID)))
                    grdA.set_TextMatrix(iRow, 1, Trim(reader(ARAZOTE.ARAZOT_EREZ_ENG)))
                    iRow += 1
                Loop
                grdA.Row = 1
            End If
        End With
    End Sub

    Private Function testInput(bMustexist As Boolean) As Boolean
        testInput = True
        If (Len(txtCOD.Text) <> 2) Then
            MsgBox("code must be 2 letters")
            testInput = False
            Exit Function
        ElseIf (txtName.Text = "") Then
            MsgBox("name missing")
            testInput = False
            Exit Function
        End If
        Dim Myrow As DataRow = CheckErez(txtCOD.Text, False)
        If (IsNothing(Myrow)) Then
            If (bMustexist) Then
                MsgBox(txtCOD.Text & " does not exist")
                testInput = False
            End If
        Else
            If (bMustexist = False) Then
                MsgBox(txtCOD.Text & " already exists")
                testInput = False
            End If
        End If
    End Function

    Private Sub OLEerrors(e As OleDbException)
        Dim errorMessages As String
        Dim i As Integer

        For i = 0 To e.Errors.Count - 1
            errorMessages += "Index #" & i.ToString() & ControlChars.Cr _
                           & "Message: " & e.Errors(i).Message & ControlChars.Cr _
                           & "NativeError: " & e.Errors(i).NativeError & ControlChars.Cr _
                           & "Source: " & e.Errors(i).Source & ControlChars.Cr _
                           & "SQLState: " & e.Errors(i).SQLState & ControlChars.Cr
        Next i

        'Dim log As System.Diagnostics.EventLog = New System.Diagnostics.EventLog()
        'log.Source = "My Application"
        'log.WriteEntry(errorMessages)
        Console.WriteLine("An exception occurred. Please contact your system administrator.")
    End Sub

    Public Function ConcatenateComma(ByVal strLine As String, ByVal strAddition As String, Optional ByVal strDivider As String = ", ", Optional ByVal bLog As Boolean = True) As String
        ConcatenateComma = ""
        Try
            ConcatenateComma = strLine
            If (Len(strAddition) > 0) Then
                If (Len(ConcatenateComma) > 0) Then ConcatenateComma = ConcatenateComma & strDivider
                ConcatenateComma = ConcatenateComma & strAddition
            End If
            Exit Function
        Catch ex As Exception
        End Try
        'ErrLog("ConcatenateComma " & strAddition, "module1", True)
    End Function

    Public Sub ErrLogLine(ByVal funct As String, ByVal strSql As String, ByVal ex As Exception)
        Dim Result As String
        Dim hr As Integer = Runtime.InteropServices.Marshal.GetHRForException(ex)
        Dim strTemp As String

        Result = ConcatenateComma(Result, strSql, Environment.NewLine)
        strTemp = ex.GetType.ToString & "(0x" & hr.ToString("X8") & "): " & ex.Message & Environment.NewLine & ex.StackTrace & Environment.NewLine
        Result &= ConcatenateComma(Result, strTemp, Environment.NewLine)
        MsgBox(Result)
    End Sub

    Public Function CheckErez(ByRef strIn As String, ByVal bMsg As Boolean, Optional ByVal iARAZField As ARAZOTE = ARAZOTE.ARAZOT_EREZ_ID) As DataRow
        Dim row As DataRow
        CheckErez = Nothing
        Try
            With glblArazotRset
                For Each row In .Rows
                    If (UCase(Trim(strIn)) = UCase(Trim(row(iARAZField)))) Then
                        CheckErez = row
                        Exit Function
                    End If
                Next
            End With
            If (bMsg) Then
                MsgBox(" ארץ אינה מוכרת " & strIn, MsgBoxStyle.MsgBoxRight Or MsgBoxStyle.MsgBoxRtlReading Or MsgBoxStyle.Information, strIn)
            End If
            Exit Function
        Catch ex As Exception
            ErrLogLine(System.Reflection.MethodBase.GetCurrentMethod().Name, "", ex)
        End Try
    End Function

    Public Function OpenReadRecordSet(ByVal table1 As String, ByVal table2 As String, ByVal RecordSet As DataTable, ByRef strSelect As String, ByVal strWhere As String, ByVal iLockType As Integer, ByRef strSql As String, Optional ByVal adCursor As VBto.CursorLocationEnum = VBto.CursorLocationEnum.adUseClient, Optional ByVal bDisplayOnStatus As Boolean = True, Optional ByVal bDisplayOnLOG As Boolean = True, Optional ByVal bAllowColumns As Boolean = True, Optional ByVal bForceLog As Boolean = False, Optional ByVal desc As String = "") As Boolean
        OpenReadRecordSet = False
        Dim strRead As String
        Dim strREPORT_PRINTER As String = ""
        Dim sTable2 As String
        Dim adaptor As OleDbDataAdapter
        Dim adaptorODBC As OdbcDataAdapter

        'errlog is not used
        RecordSet.CaseSensitive = False
        'If (bMainSet And bDisplayOnStatus) Then
        '    strREPORT_PRINTER = MainFormPanelText(PanelsENUM.PANEL_REPORT_PRINTER)
        '    SetMainFormPanelText(PanelsENUM.PANEL_REPORT_PRINTER, table1 & " " & table2 & " שולף " & " " & desc)
        'End If
        'If (RecordSet.State <> ConnectionState.Closed) Then RecordSet.Close()
        RecordSet.Clear()
        '- RecordSet.CursorType = VBto.CursorTypeEnum.adOpenDynamic 'adOpenKeyset
        '- RecordSet.CursorLocation = adCursor 'adUseClient
        '- RecordSet.LockType = iLockType 'adLockOptimistic
        If strSelect = "*" And table1 <> "" And bAllowColumns Then
            If table2 = "" Then
                sTable2 = GetTableFromJoin(strWhere) ' Get table name from join statement.
            Else
                sTable2 = table2
            End If
            'If sTable2 = "" Then
            '    strSelect = AllColumns(ColumnsForTabName(table1), table1)
            'Else
            '    strSelect = AllColumns(ColumnsForTabName(table1), table1) & ", " & AllColumns(ColumnsForTabName(sTable2), sTable2)
            'End If
        End If
        strRead = "SELECT " & strSelect
        If (table1 <> "") Then strRead &= " FROM " & IIf(Not bFox, "(", "") & table1
        If (table2 <> "") Then
            strRead &= ", " & table2
        End If
        strRead &= IIf(InStr(strRead, "(") And Not bFox And (table1 <> ""), ") ", " ") & strWhere & ";"
        strSql = strRead
        Debug.WriteLine(Dos2Win(strRead))
        'If (bDisplayOnLOG Or bForceLog) Then CWLog("modActiveRecordSet", Dos2Win(strRead), bForceLog)
        'If (bDisplayOnStatus) Then
        '    ScreenCursorActive(Cursors.WaitCursor)
        'End If

        'RecordSet.Open(strRead, glblConnection)
        If (bFox) Then
            adaptor = New OleDbDataAdapter(strRead, glblConnection)
            adaptor.Fill(RecordSet)
        Else
            adaptorODBC = New OdbcDataAdapter(strRead, glblConnectionODBC)
            adaptorODBC.Fill(RecordSet)
        End If
        'dtLastMysqlAccess = Now
        strSql = "" 'needed only for error
        Try ' On Error GoTo ignore
            If (bDisplayOnLOG Or bForceLog) Then
                'If (RecordSet.rows.Count=0  Then
                '    CWLog("modActiveRecordSet", "recordcount 0", bForceLog)
                'Else
                '    CWLog("modActiveRecordSet", "recordcount" & Str(RecordSet.Rows.Count), bForceLog)
                'End If
            End If
            'If (bDisplayOnStatus) Then ScreenCursorActive(Cursors.Default)
            OpenReadRecordSet = True
            ' YI 20-JUN-2010: removed reference to fMain in order to exclude all forms from projects that use this module.
            'If (bMainSet And bDisplayOnStatus) Then SetMainFormPanelText(PanelsENUM.PANEL_REPORT_PRINTER, strREPORT_PRINTER)
            Exit Function
        Catch pOleDbException As OleDbException
            'ScreenCursorActive(Cursors.Default)
            Err.Clear()
            'ErrLogADO(pOleDbException, "OpenReadRecordSet", "modactiverecordset", strRead, KILLMODE.CONTINUE_MANDATORY)
        Catch ' ignore:
            'ErrLogLine(System.Reflection.MethodBase.GetCurrentMethod().Name, "", ex)
        End Try
    End Function
    Public Function Dos2Win(ByVal strIn As String, Optional ByVal bCompact As Boolean = False) As String
        Dos2Win = ""
        Dim i As Integer
        Dim strTemp As String
        Dim strDebug As String = ""

        Try
            strDebug = "Len(strIn)" & Str(Len(strIn))
            If bFox Then
                For i = 1 To Len(strIn)
                    strTemp = strIn(i - 1).ToString()
                    'strDebug = "i=" & CStr(i) & " " & strTemp
                    'Debug.Print Hex(Asc(strTemp)), strTemp
                    If ((Asc(strTemp) >= &H80) And (Asc(strTemp) <= &H9A)) Then
                        Dos2Win = Dos2Win & Convert.ToString(Chr(Asc(strTemp) + &H60))
                    Else
                        Dos2Win = Dos2Win & strTemp
                    End If
                Next i
                If (bCompact) Then
                    strTemp = Replace(Dos2Win, vbTab, " ")
                    Do
                        If (InStr(strTemp, Space(2)) = 0) Then Exit Do
                        Dos2Win = Replace(strTemp, Space(2), " ")
                        strTemp = Dos2Win
                        Application.DoEvents()
                    Loop
                End If
            Else
                Dos2Win = strIn
            End If
            Exit Function
        Catch ex As Exception
            ErrLogLine(System.Reflection.MethodBase.GetCurrentMethod().Name, "", ex)
        End Try
        'ErrLog("Dos2Win" & " " & strDebug & " " & Dos2Win, "module1", True)
    End Function

    Private Function GetTableFromJoin(ByVal strWhere As String) As String
        Dim i As Integer
        Dim nStage As Short ' 0 - start, 1 - on word inner|outer|left|right, 2 - on word JOIN, -1 - frustration
        Dim sTerms() As String
        Dim sTerm As String
        sTerms = Split(strWhere, " ")
        nStage = 0
        GetTableFromJoin = ""
        For i = LBound(sTerms) To UBound(sTerms)
            If (sTerms(i) <> "") Then
                If nStage = 0 Then
                    sTerm = UCase(sTerms(i))
                    If (sTerm = "OUTER" Or sTerm = "INNER" Or sTerm = "LEFT" Or sTerm = "RIGHT") Then
                        nStage = 1
                    ElseIf sTerm = "JOIN" Then
                        nStage = 2
                    Else
                        nStage = -1
                    End If
                ElseIf nStage = 1 Then
                    sTerm = UCase(sTerms(i))
                    If sTerm <> "JOIN" Then
                        nStage = -1
                    Else
                        nStage = 2
                    End If
                ElseIf nStage = 2 Then
                    GetTableFromJoin = sTerms(i)
                    Exit For
                End If
                If nStage = -1 Then Exit For
            End If
        Next i
    End Function


End Class