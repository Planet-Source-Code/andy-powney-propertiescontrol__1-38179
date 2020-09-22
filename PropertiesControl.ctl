VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.UserControl PropertiesControl 
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5370
   ScaleHeight     =   235.5
   ScaleMode       =   2  'Point
   ScaleWidth      =   268.5
   Begin VB.CommandButton cmdButton 
      Caption         =   "..."
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox lstList 
      Height          =   315
      Left            =   3240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstLists 
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstTypes 
      Height          =   450
      Left            =   4320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox lstTrueFalseEdit 
      Height          =   315
      ItemData        =   "PropertiesControl.ctx":0000
      Left            =   1440
      List            =   "PropertiesControl.ctx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid grdProp 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
   End
End
Attribute VB_Name = "PropertiesControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' PropertiesControl v1.1
' ======================
'
' By Andy Powney, The Perplexity Project
' 21 August 2002
'
'
'
'
' INTRODUCTION
' ============
'
' This CTL for VB5+ mimics the core functionality of a properties sheet, as used in DevStudio itself.
' This control can handle various types of property, and the specifics are outlined below. By default
' the properties that appear in your sheet are sorted alphabetically on the name.
'
'
'
'
' SUPPORTED PROPERTY TYPES
' ========================
'
' TEXT/NUMERIC PROPERTIES
' e.g. Name=Alfred
'
' Use the SetTextProperty to set or update a named property. The user will be presented with a standard
' textbox when editing the values for this property.
'
'
' BOOLEAN PROPERTIES
' e.g. HasScrollbars=True
'
' Use SetBooleanProperty to set or update the value of the named property. The user will be presented with
' a dropdown listbox containing the words "True" and "False".
'
'
' LIST PROPERTIES
' e.g. Today=Monday (where Monday is selected from the names of the weekdays)
'
' Use SetListProperty to set or update the value of the named property. You will also need to pass the set
' of strings to be used as possible values. This is done in a single string, where each term is separated
' using the vbCRLF characters. i.e. "Monday" & vbCRLF & "Tuesday" & vbCRLF & "Wednesday" & vbCRLF ...
'
'
' BUTTON PROPERTIES
' e.g. Filename=c:\test.txt
'
' Use SetButtonProperty to set or update the value of the named property. The user is presented with
' the text value you specified (which they cannot edit directly) and a button. You will receive the
' relevant event notification only when the user clicks this button.
'
'
'
'
'
' BASIC METHODS & EVENTS
' ======================
'
' SetTextProperty, SetBooleanProperty, SetListProperty, SetButtonProperty
'   Use these methods to add or edit a value in the property sheet.
'
' RemoveProperty
'   Use this method to remove the named property from the property sheet.
'
' Sort
'   When your property sheet is not automatically sorted, you can do it manually.
'
' GetPropertyValue
'   This will return a string containing the value for the named property.
'
' OnPropertyClick (Event)
'   This event will be fired when the user changes the value of a certain property. Every time the user
' chooses a new item from a list, or changes True to False, or clicks a button, or every keypress (during
' editing of a text property) this event will be fired.
'
' OnPropertyDblClick (Event)
'   This event is fired when the user double-clicks on the property name.
'
' Clear
'   This will remove all properties from the property sheet.
'
'

Public Event OnPropertyClick(strName As String)
Public Event OnPropertyDblClick(strName As String)
Public Event OnHighlightChange(strName As String)

Private blnIgnoreMessage As Boolean
Private blnSorted As Boolean
Private intEditIndex As Integer

Private Type Size
        cx As Long
        cy As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Const TYPE_TEXT = 1
Const TYPE_BOOLEAN = 2
Const TYPE_LIST = 3
Const TYPE_BUTTON = 4

Public Property Let Sorted(sort_flag As Boolean)
    blnSorted = sort_flag
    If blnSorted Then Sort
End Property

Public Property Get Sorted() As Boolean
    Sorted = blnSorted
End Property

Private Sub ResizeControls()
    With grdProp
        If .Visible = True Then
            .Left = 0
            .Top = 0
            .Width = ScaleWidth
            .Height = ScaleHeight
        End If
    End With
    ResizeColumns
    If grdProp.Rows > 0 Then
        If txtEdit.Visible = True Then
            ShowTextEdit
        ElseIf lstTrueFalseEdit.Visible = True Then
            ShowBooleanEdit
        ElseIf lstList.Visible = True Then
            ShowListEdit
        ElseIf cmdButton.Visible = True Then
            ShowButtonEdit
        Else
            blnIgnoreMessage = False
        End If
    Else
        txtEdit.Visible = False
        lstTrueFalseEdit.Visible = False
        lstList.Visible = False
        cmdButton.Visible = False
        blnIgnoreMessage = False
    End If
End Sub

Private Sub cmdButton_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then DoTabAction Shift
End Sub

Private Sub grdProp_Click()
    intEditIndex = -1
    blnIgnoreMessage = False
    grdProp_RowColChange
End Sub

Private Sub grdProp_DblClick()
    If blnIgnoreMessage = True Then Exit Sub
    If grdProp.Rows < 1 Then Exit Sub
    blnIgnoreMessage = True
    Dim strName As String
    grdProp.Col = 0
    strName = grdProp.Text
    grdProp.Col = 1
    blnIgnoreMessage = False
    RaiseEvent OnPropertyDblClick(strName)
End Sub

Private Sub grdProp_EnterCell()
    If blnIgnoreMessage = True Then Exit Sub
    blnIgnoreMessage = True
    Dim strName As String
    grdProp.Col = 0
    strName = grdProp.Text
    grdProp.Col = 1
    blnIgnoreMessage = False
    RaiseEvent OnHighlightChange(strName)
End Sub

Private Sub grdProp_GotFocus()
    HideEdit
End Sub

Private Sub grdProp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then DoTabAction Shift
End Sub

Private Sub grdProp_RowColChange()
    On Error GoTo error_handler
    If blnIgnoreMessage = True Then Exit Sub
    blnIgnoreMessage = True
    Dim strName As String
    grdProp.Col = 0
    strName = grdProp.Text
    grdProp.Col = 1
    blnIgnoreMessage = False
    DoDefaultEdit strName
    Exit Sub
error_handler:
    blnIgnoreMessage = False
End Sub

Private Sub grdProp_Scroll()
    HideEdit
End Sub

Private Sub grdProp_SelChange()
    If blnIgnoreMessage = True Then Exit Sub
    blnIgnoreMessage = True
    Dim strName As String
    grdProp.Col = 0
    strName = grdProp.Text
    grdProp.Col = 1
    DoDefaultEdit strName
    blnIgnoreMessage = False
End Sub

Private Sub DoTabAction(Shift As Integer)
    Dim intRow As Integer
    intRow = intEditIndex
    If Shift Then
        ' reverse
        intRow = intRow - 1
        If intRow < 0 Then intRow = grdProp.Rows - 1
    Else
        ' forward
        intRow = intRow + 1
        If intRow >= grdProp.Rows Then intRow = 0
    End If
    HideEdit
    grdProp.Row = intRow
    grdProp.SetFocus
    DoEvents
    DoDefaultEdit grdProp.TextMatrix(intRow, 0)
    DoEvents
    grdProp_Click
End Sub

Private Sub lstList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then DoTabAction Shift
End Sub

Private Sub lstTrueFalseEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then DoTabAction Shift
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Then
        DoTabAction Shift
    End If
End Sub

Private Sub UserControl_Initialize()
    blnSorted = True
    intEditIndex = -1
End Sub

Private Sub UserControl_Resize()
    ResizeControls
End Sub

Public Function GetPropertyValue(strName As String) As String
    Dim index As Integer
    index = IndexOf(strName)
    If index = -1 Then
        Let GetPropertyValue = Null
    Else
        GetPropertyValue = Mid(grdProp.TextMatrix(index, 1), 2)
    End If
End Function

Public Function RemoveProperty(strName As String)
    HideEdit
    Dim index As Integer
    index = IndexOf(strName)
    If index <> -1 Then
        If grdProp.Rows = 1 Then
            grdProp.Clear
            grdProp.Rows = 0
        Else
            grdProp.RemoveItem index
        End If
    End If
    If blnSorted Then Sort
    ResizeControls
    RemovePropertyType strName
End Function

Public Function Clear()
    HideEdit
    grdProp.Clear
    grdProp.Rows = 0
    lstList.Clear
    lstLists.Clear
    lstTypes.Clear
End Function

Public Function SetListProperty(strName As String, strValue As String, strItems As String)
    HideEdit
    SetTextDisplay strName, strValue
    SetListItems strName, strItems
    SetPropertyType strName, TYPE_LIST
    RaiseEvent OnPropertyClick(strName)
    DoDefaultEdit strName
End Function

Public Function SetBooleanProperty(strName As String, blnValue As Boolean)
    HideEdit
    If blnValue Then
        SetTextDisplay strName, "True"
    Else
        SetTextDisplay strName, "False"
    End If
    SetPropertyType strName, TYPE_BOOLEAN
    RaiseEvent OnPropertyClick(strName)
    DoDefaultEdit strName
End Function

Public Function SetButtonProperty(strName As String, strValue As String)
    HideEdit
    SetPropertyType strName, TYPE_BUTTON
    SetTextDisplay strName, strValue
    SetPropertyType strName, TYPE_BUTTON
    RaiseEvent OnPropertyClick(strName)
    DoDefaultEdit strName
End Function

Public Function SetTextProperty(strName As String, strValue As String)
    HideEdit
    SetTextDisplay strName, strValue
    SetPropertyType strName, TYPE_TEXT
    RaiseEvent OnPropertyClick(strName)
    DoDefaultEdit strName
End Function

Private Function SetTextDisplay(strName As String, strValue As String)
    blnIgnoreMessage = True
    Dim index As Integer
    index = IndexOf(strName)
    If index = -1 Then
        With grdProp
            .AddItem ""
            index = .Rows - 1
            .TextMatrix(index, 0) = strName
            .TextMatrix(index, 1) = " " & strValue
        End With
    Else
        With grdProp
            .TextMatrix(index, 1) = " " & strValue
        End With
    End If
    blnIgnoreMessage = False
    If blnSorted Then Sort
    grdProp.Row = IndexOf(strName)
    grdProp.Col = 1
    ResizeControls
End Function

Private Function IndexOf(strName As String) As Integer
    blnIgnoreMessage = True
    Dim i As Integer
    IndexOf = -1
    For i = 0 To grdProp.Rows - 1
        If grdProp.TextMatrix(i, 0) = strName Then IndexOf = i
    Next i
    blnIgnoreMessage = False
End Function

Public Function Sort()
    If grdProp.Rows < 2 Then Exit Function
    Dim i As Integer
    Dim j As Integer
    For i = 0 To grdProp.Rows - 1
        For j = 0 To grdProp.Rows - 1
            Dim strName As String
            Dim strValue As String
            If grdProp.TextMatrix(i, 0) < grdProp.TextMatrix(j, 0) Then
                strName = grdProp.TextMatrix(i, 0)
                strValue = grdProp.TextMatrix(i, 1)
                grdProp.TextMatrix(i, 0) = grdProp.TextMatrix(j, 0)
                grdProp.TextMatrix(i, 1) = grdProp.TextMatrix(j, 1)
                grdProp.TextMatrix(j, 0) = strName
                grdProp.TextMatrix(j, 1) = strValue
            End If
        Next j
    Next i
End Function

Private Sub ResizeColumns()
    blnIgnoreMessage = True
    Dim intMin As Integer
    Dim szTextSize As Size
    Dim i As Integer
    intMin = 0
    For i = 0 To grdProp.Rows - 1
        Dim strName As String
        strName = grdProp.TextMatrix(i, 0)
        GetTextExtentPoint UserControl.hdc, strName & " ", Len(strName & " "), szTextSize
        If szTextSize.cx > intMin Then intMin = szTextSize.cx
    Next i
    intMin = ScaleX(intMin, vbPoints, vbTwips)
    grdProp.ColWidth(0) = intMin
    grdProp.ColWidth(1) = Width - grdProp.ColWidth(0) - 64
    blnIgnoreMessage = False
End Sub

Public Function DoDefaultEdit(strName As String)
    ' Ensure the correct cell is selected
    Dim index As Integer
    index = IndexOf(strName)
    If index = -1 Then Exit Function
    grdProp.Row = index
    grdProp.Col = 1
    grdProp.RowSel = index
    grdProp.ColSel = 1
    
    ' Do the appropriate editing function
    If GetPropertyType(strName) = TYPE_TEXT Then
        intEditIndex = index
        ShowTextEdit
    ElseIf GetPropertyType(strName) = TYPE_BOOLEAN Then
        intEditIndex = index
        ShowBooleanEdit
    ElseIf GetPropertyType(strName) = TYPE_LIST Then
        intEditIndex = index
        PopulateList strName
        ShowListEdit
    ElseIf GetPropertyType(strName) = TYPE_BUTTON Then
        intEditIndex = index
        ShowButtonEdit
    End If
End Function

Private Sub HideEdit()
    txtEdit.Visible = False
    lstTrueFalseEdit.Visible = False
    lstList.Visible = False
    cmdButton.Visible = False
    intEditIndex = -1
    blnIgnoreMessage = False
End Sub

' -- Property type management --

Public Function SetPropertyType(strName As String, intType As Integer)
    Dim index As Integer
    index = GetPropertyTypeIndex(strName)
    If index = -1 Then
        lstTypes.AddItem strName
        lstTypes.ItemData(lstTypes.NewIndex) = intType
    Else
        lstTypes.ItemData(index) = intType
    End If
End Function

Public Function GetPropertyType(strName As String) As Integer
    GetPropertyType = 1
    Dim index As Integer
    index = GetPropertyTypeIndex(strName)
    If index <> -1 Then
        GetPropertyType = lstTypes.ItemData(index)
    End If
End Function

Private Function GetPropertyTypeIndex(strName As String) As Integer
    Dim i As Integer
    GetPropertyTypeIndex = -1
    For i = 0 To lstTypes.ListCount - 1
        If lstTypes.List(i) = strName Then GetPropertyTypeIndex = i
    Next i
End Function

Private Function RemovePropertyType(strName As String)
    Dim index As Integer
    index = GetPropertyTypeIndex(strName)
    If index <> -1 Then
        lstTypes.RemoveItem (index)
    End If
End Function

Private Function AreScrollBarsVisible() As Boolean
    Dim rectClient As RECT
    Dim rectWindow As RECT
    GetClientRect grdProp.hwnd, rectClient
    GetWindowRect grdProp.hwnd, rectWindow
    If rectClient.Right + 16 < rectWindow.Right - rectWindow.Left Then
        AreScrollBarsVisible = True
    Else
        AreScrollBarsVisible = False
    End If
End Function

' --- TEXT FIELD EDITING ---

Private Sub ShowTextEdit()
    On Error GoTo error_handler
    
    With txtEdit
        .Text = Mid(grdProp.TextMatrix(intEditIndex, 1), 2)
        .Left = ScaleX(grdProp.CellLeft, vbTwips, vbPoints)
        .Top = ScaleY(grdProp.CellTop, vbTwips, vbPoints)
        .Width = ScaleX(grdProp.CellWidth, vbTwips, vbPoints)
        .Height = ScaleY(grdProp.CellHeight, vbTwips, vbPoints)
        If AreScrollBarsVisible() Then .Width = .Width - 12
        .Visible = True
        .SetFocus
    End With
    Exit Sub
error_handler:
    txtEdit.Visible = False
    intEditIndex = -1
End Sub

Private Sub txtEdit_Change()
    If intEditIndex <> -1 Then
        grdProp.TextMatrix(intEditIndex, 1) = " " & txtEdit.Text
        RaiseEvent OnPropertyClick(grdProp.TextMatrix(intEditIndex, 0))
    End If
End Sub

Private Sub txtEdit_GotFocus()
    txtEdit.BackColor = vbHighlight
    txtEdit.ForeColor = vbHighlightText
    txtEdit.SelLength = 0
    txtEdit.SelStart = Len(txtEdit.Text)
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.BackColor = vbWindowBackground
    txtEdit.ForeColor = vbWindowText
    txtEdit.Visible = False
    blnIgnoreMessage = False
End Sub

' --- BOOLEAN FIELD EDITING ---

Private Sub ShowBooleanEdit()
    If blnIgnoreMessage = True Then Exit Sub
    blnIgnoreMessage = True
    On Error GoTo error_handler
    With lstTrueFalseEdit
        Dim strVal As String
        strVal = Mid(grdProp.TextMatrix(intEditIndex, 1), 2)
        If UCase(strVal) = "TRUE" Or UCase(strVal) = "YES" Then
            strVal = "True"
        Else
            strVal = "False"
        End If
        .Text = strVal
        .Left = ScaleX(grdProp.CellLeft, vbTwips, vbPoints)
        .Top = ScaleY(grdProp.CellTop, vbTwips, vbPoints)
        .Width = ScaleX(grdProp.CellWidth, vbTwips, vbPoints)
        If AreScrollBarsVisible() Then .Width = .Width - 12
        .Visible = True
        .SetFocus
    End With
    blnIgnoreMessage = False
    Exit Sub
error_handler:
    lstTrueFalseEdit.Visible = False
    intEditIndex = -1
    blnIgnoreMessage = False
End Sub

Private Sub lstTrueFalseEdit_Click()
    If blnIgnoreMessage = True Then Exit Sub
    blnIgnoreMessage = True
    If intEditIndex <> -1 Then
        grdProp.TextMatrix(intEditIndex, 1) = " " & lstTrueFalseEdit.Text
        RaiseEvent OnPropertyClick(grdProp.TextMatrix(intEditIndex, 0))
        DoDefaultEdit grdProp.TextMatrix(intEditIndex, 0)
    End If
    blnIgnoreMessage = False
End Sub

Private Sub lstTrueFalseEdit_LostFocus()
    lstTrueFalseEdit.Visible = False
    blnIgnoreMessage = False
End Sub

' --- LIST FIELD EDITING ---

Private Sub ShowListEdit()
    If blnIgnoreMessage = True Then Exit Sub
    blnIgnoreMessage = True
    On Error GoTo error_handler
    With lstList
        Dim strText As String
        strText = Mid(grdProp.TextMatrix(intEditIndex, 1), 2)
        Dim i As Integer
        For i = 0 To lstList.ListCount - 1
            If lstList.List(i) = strText Then lstList.ListIndex = i
        Next i
        If lstList.ListIndex = -1 Then lstList.ListIndex = 0
        .Left = ScaleX(grdProp.CellLeft, vbTwips, vbPoints)
        .Top = ScaleY(grdProp.CellTop, vbTwips, vbPoints)
        .Width = ScaleX(grdProp.CellWidth, vbTwips, vbPoints)
        If AreScrollBarsVisible() Then .Width = .Width - 12
        .Visible = True
        .SetFocus
    End With
    blnIgnoreMessage = False
    Exit Sub
error_handler:
    txtEdit.Visible = False
    intEditIndex = -1
    If lstList.ListCount > 0 Then lstList.ListIndex = 0
    blnIgnoreMessage = False
End Sub

Private Sub lstList_Click()
    If blnIgnoreMessage Then Exit Sub
    blnIgnoreMessage = True
    If intEditIndex <> -1 Then
        grdProp.TextMatrix(intEditIndex, 1) = " " & lstList.Text
        RaiseEvent OnPropertyClick(grdProp.TextMatrix(intEditIndex, 0))
        DoDefaultEdit grdProp.TextMatrix(intEditIndex, 0)
    End If
    blnIgnoreMessage = False
End Sub

Private Sub lstList_LostFocus()
    lstList.Visible = False
    blnIgnoreMessage = False
End Sub

' --- BUTTON FIELD EDITING ---

Private Sub ShowButtonEdit()
    On Error GoTo error_handler
    With cmdButton
        .Top = ScaleY(grdProp.CellTop, vbTwips, vbPoints)
        .Height = ScaleY(grdProp.CellHeight, vbTwips, vbPoints)
        .Left = ScaleX(grdProp.CellLeft + grdProp.CellWidth, vbTwips, vbPoints) - .Width
        If AreScrollBarsVisible() Then .Left = .Left - 12
        .Visible = True
        .SetFocus
    End With
    Exit Sub
error_handler:
    cmdButton.Visible = False
    intEditIndex = -1
End Sub

Private Sub cmdButton_Click()
    If blnIgnoreMessage Then Exit Sub
    blnIgnoreMessage = True
    If intEditIndex <> -1 Then
        RaiseEvent OnPropertyClick(grdProp.TextMatrix(intEditIndex, 0))
        DoDefaultEdit grdProp.TextMatrix(intEditIndex, 0)
    End If
    blnIgnoreMessage = False
End Sub

' Each item in a list must be terminated with vbCRLF

Public Function SetListItems(strName As String, strItems As String)
    Dim i As Integer
    Dim index As Integer
    index = -1
    For i = 0 To lstLists.ListCount - 1
        If Left(lstLists.List(i), Len(strName) + Len(vbCrLf)) = strName & vbCrLf Then
            index = i
        End If
    Next i
    If Right(strItems, Len(vbCrLf)) <> vbCrLf Then strItems = strItems & vbCrLf
    If index = -1 Then
        lstLists.AddItem strName & vbCrLf & strItems
    Else
        lstLists.List(index) = strName & vbCrLf & strItems
    End If
End Function

Private Sub PopulateList(strName As String)
    Dim i As Integer
    Dim index As Integer
    index = -1
    For i = 0 To lstLists.ListCount - 1
        If Left(lstLists.List(i), Len(strName) + Len(vbCrLf)) = strName & vbCrLf Then
            index = i
        End If
    Next i
    lstList.Clear
    If index <> -1 Then
        Dim strTemp As String
        strTemp = Mid(lstLists.List(index), Len(strName) + Len(vbCrLf) + 1)
        i = InStr(strTemp, vbCrLf)
        While i > 0
            lstList.AddItem Left(strTemp, i - Len(vbCrLf) + 1)
            strTemp = Mid(strTemp, i + Len(vbCrLf))
            i = InStr(strTemp, vbCrLf)
        Wend
    End If
End Sub
