VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB String Find Tool"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3600
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel Output"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton cmd_browse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   11
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CheckBox chkSchema 
      Caption         =   "Schema File"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   8
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "ocx"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4980
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "cls"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.CheckBox Check2 
      Caption         =   "bas"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2940
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "vbp"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   4815
   End
   Begin VB.TextBox txtSearchString 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label lblStatus 
      Caption         =   "."
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   7095
   End
   Begin VB.Label Label3 
      Caption         =   "Search in Folder"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Select File Types"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Search String"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intVBPCount As Integer
Dim intCLSCount As Integer
Dim intBASCount As Integer
Dim intOCXCount As Integer
Dim intSQLCount As Integer
Dim intFRMCount As Integer
Dim intRowIndex As Long
Dim lngTotalLineCnt As Long
Dim objExcel As Excel.Application
Dim objWorkBook As Excel.Workbook
Dim objSheet As Excel.Worksheet
Dim oShell As Shell
Dim oFolder As Object
Dim oFile As Object
Dim sNewLocal As String
Dim sNewRemote As String

Dim objTextOut As TextStream

Private Sub Check5_Click()

End Sub

Private Sub cmd_browse_Click()
    Dim strFolder As String

    strFolder = OpenDirectoryTV(Me, "Select Folder")
    If strFolder <> "" Then
        Me.txtPath = strFolder
    End If
End Sub

Private Sub cmdExcel_Click()
    If Trim(txtSearchString) <> "" And Trim(txtPath) <> "" Then
        
        intVBPCount = 0
        intCLSCount = 0
        intBASCount = 0
        intOCXCount = 0
        intSQLCount = 0
        intFRMCount = 0
        intRowIndex = 1
        
        Set objExcel = CreateObject("Excel.Application")
        Set objWorkBook = objExcel.Workbooks.Add
        objWorkBook.SaveAs App.Path & "\" & "Output.xls"
        
        Screen.MousePointer = vbHourglass
        SearchString txtPath, Trim(Me.txtSearchString), True
        Screen.MousePointer = vbNormal
        
        objWorkBook.Save
        objWorkBook.Close
        
        Set objExcel = Nothing
        Set objWorkBook = Nothing
        Set objSheet = Nothing

    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    If Trim(txtSearchString) <> "" And Trim(txtPath) <> "" Then
        
        intVBPCount = 0
        intCLSCount = 0
        intBASCount = 0
        intOCXCount = 0
        intSQLCount = 0
        intFRMCount = 0
        
        Set objTextOut = objFso.CreateTextFile("Output.txt", True, True)
        Screen.MousePointer = vbHourglass
        SearchString txtPath, Trim(Me.txtSearchString), False
        Screen.MousePointer = vbNormal
        
        objTextOut.WriteLine
        objTextOut.WriteLine "Summary : "
        objTextOut.WriteLine "  VBP File Count : " & CStr(intVBPCount)
        objTextOut.WriteLine "  FRM File Count : " & CStr(intFRMCount)
        objTextOut.WriteLine "  BAS File Count : " & CStr(intBASCount)
        objTextOut.WriteLine "  CLS File Count : " & CStr(intCLSCount)
        objTextOut.WriteLine "  OCX File Count : " & CStr(intOCXCount)
        objTextOut.WriteLine "  SCHEMA File Count : " & CStr(intSQLCount)
        objTextOut.WriteLine "  Total Line count: " & CStr(lngTotalLineCnt)
        objTextOut.Close
        Set objTextOut = Nothing
    End If
End Sub

Private Sub SearchString(strPath As String, strSearch As String, blnExcel As Boolean)
    Dim objFolder As Folder
    Dim objSubFolder As Folder
    Dim objStream As TextStream
    Dim strLine As String
    Dim objFile As File
    Dim intCnt As Integer
    Dim strExtention As String
    Dim blnWriteFileName As Boolean
    
    Set objFolder = objFso.GetFolder(strPath)
    
    For Each objFile In objFolder.Files
        If blnWriteFileName = True Then
            If blnExcel = False Then
               objTextOut.WriteLine ""
            End If
            blnWriteFileName = False
        End If
        strExtention = UCase(objFso.GetExtensionName(objFso.BuildPath(strPath, objFile.Name)))
        If strExtention = "VBP" Or strExtention = "FRM" Or strExtention = "CLS" Or strExtention = "CTL" Or strExtention = "BAS" Or strExtention = "SCHEMA" Then
        
            lblStatus = "Processing File : " & objFso.BuildPath(strPath, objFile.Name)
            
            DoEvents
            
            Set objStream = objFso.OpenTextFile(objFso.BuildPath(strPath, objFile.Name), ForReading)
            If Not objStream Is Nothing Then
                Do While Not objStream.AtEndOfStream
                    strLine = objStream.ReadLine
                    If InStr(1, UCase(strLine), UCase(strSearch), vbTextCompare) > 0 Then
                        If blnWriteFileName = False Then
                            blnWriteFileName = True
                            If blnExcel = False Then
                              objTextOut.WriteLine objFso.BuildPath(strPath, objFile.Name)
                            End If
                            If strExtention = "VBP" Then
                                intVBPCount = intVBPCount + 1
                            ElseIf strExtention = "FRM" Then
                                intFRMCount = intFRMCount + 1
                            ElseIf strExtention = "CLS" Then
                                intCLSCount = intCLSCount + 1
                            ElseIf strExtention = "BAS" Then
                                intBASCount = intBASCount + 1
                            ElseIf strExtention = "CTL" Then
                                intOCXCount = intOCXCount + 1
                            ElseIf strExtention = "SQL" Then
                                intSQLCount = intSQLCount + 1
                            End If
                        End If
                        If blnExcel = False Then
                           objTextOut.WriteLine strLine
                        Else
                           Call PrintinExcel(strPath, objFile.Name, strLine)
                        End If
                        'MsgBox strLine
                    End If
                     If strExtention = "FRM" Or strExtention = "CLS" Or strExtention = "CTL" Or strExtention = "BAS" Then
                        lngTotalLineCnt = lngTotalLineCnt + 1
                     End If
                Loop
                objStream.Close
            End If
        End If
    Next
    
    For Each objSubFolder In objFolder.SubFolders
         strPath = objSubFolder.Path
        Call SearchString(strPath, strSearch, blnExcel)
    Next
End Sub

Private Sub PrintinExcel(strPath As String, strFileName As String, strLine As String)
    Dim objRange As Excel.Range
 '   Set objRange = objWorkBook.Sheets(1).Range
    intRowIndex = intRowIndex + 1
    objWorkBook.Sheets(1).Cells(intRowIndex, 1) = CStr(intRowIndex)
    objWorkBook.Sheets(1).Cells(intRowIndex, 2) = strPath
    objWorkBook.Sheets(1).Cells.Cells(intRowIndex, 3) = strFileName
    objWorkBook.Sheets(1).Cells.Cells(intRowIndex, 4) = strLine
End Sub

Private Sub Form_Load()
 '  Dim sfile As String
  ' Dim objFolder As Folder
'   sfile = GetExtendedProperties("D:\ToMove\Tools\VBFindTool\", "VBStringSearchTool.exe")

'   MsgBox CStr(objFso.GetFileVersion(sfile))
'   Text1.Text = App.Major & "." & App.Minor & "." & App.Revision & "." & App.FileDescription
   
End Sub


Private Function GetExtendedProperties(ByRef strFolderName As String, ByRef strFileName As String) As String
    
    Dim arrHeaders(40) As String
    Dim msg As String
    Dim i As Integer
    
    Set oShell = CreateObject("Shell.Application")
    
    Set oFolder = oShell.Namespace(strFolderName)
    
    For i = 1 To 40
        arrHeaders(i) = oFolder.GetDetailsOf(oFolder.ParseName(Dir$(strFileName)), i)
'        arrHeaders(i) = oFolder.GetDetailsOf(oFolder.Items, i)
    Next
    
    Set oFile = oFolder.ParseName(strFileName)
    MsgBox oFolder.GetDetailsOf(oFile, 36)
    For Each oFile In oFolder.Items
        If oFile.Name = strFileName Then
            For i = 0 To 40
                msg = msg & arrHeaders(i) & ": " & oFolder.GetDetailsOf(oFile, i) & vbCrLf
            Next
            Exit For
        End If
    Next

    If msg = "" Then
        GetExtendedProperties = strFileName & " not found."
    Else
        GetExtendedProperties = msg
    End If

End Function
