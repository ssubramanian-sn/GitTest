VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{1AF1F43C-1DE4-44ED-B0FD-A49A4EAA03A6}#4.0#0"; "IGResizer40.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Fleet App Builder  "
   ClientHeight    =   7890
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13875
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSourceFolder 
      Height          =   375
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "E:\GIT\MWS_DEV"
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdCheckOut 
      Caption         =   "Start Build"
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   7200
      Width           =   1515
   End
   Begin VB.CommandButton cmdUndoCheckout 
      Caption         =   "Cancel Build"
      Height          =   615
      Left            =   12120
      TabIndex        =   12
      Top             =   7200
      Width           =   1515
   End
   Begin VB.CommandButton cmdCheckIn 
      Caption         =   "Finalize Build"
      Height          =   615
      Left            =   6180
      TabIndex        =   11
      Top             =   7200
      Width           =   1515
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   4755
      Left            =   60
      TabIndex        =   9
      Top             =   2280
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   8387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ActiveResizer.SSResizer SSResizer1 
      Left            =   14160
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   262144
      MinFontSize     =   8
      MaxFontSize     =   10
      DesignWidth     =   13875
      DesignHeight    =   7890
   End
   Begin VB.Frame Frame1 
      Caption         =   "Build Details"
      Height          =   1095
      Left            =   180
      TabIndex        =   2
      Top             =   660
      Width           =   13575
      Begin VB.TextBox txtBuildNo 
         Height          =   375
         Left            =   3720
         TabIndex        =   17
         Text            =   "Text4"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdSetVer 
         Caption         =   "Set Version"
         Height          =   495
         Left            =   4680
         TabIndex        =   10
         Top             =   480
         Width           =   1515
      End
      Begin VB.TextBox txtRevision 
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtMinor 
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtMajor 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Build"
         Height          =   375
         Left            =   3720
         TabIndex        =   18
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Revision"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Minor"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Major"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdGetIni 
      Caption         =   "Get Ini"
      Height          =   495
      Left            =   12840
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   13320
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "Build Ini File"
      Height          =   315
      Left            =   4920
      TabIndex        =   19
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   13455
   End
   Begin VB.Label Label13 
      Caption         =   "Source Folder"
      Height          =   315
      Left            =   180
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblRemakeIniFile 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   6315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function VerQueryValue Lib "version" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function GetFileVersionInfoSize Lib "version" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Const CONST_RELATIVE_PATH = "RelativePath"
Const CONST_HEADER_SECTION = "Header"
Const CONST_VB6PATH_KEYNAME = "VB6Path"
Const CONST_TFSPATH_KEYNAME = "TFSPath"
Const CONST_MAJORVER_KEYNAME = "MajorVersion"
Const CONST_MINORVER_KEYNAME = "MinorVersion"
Const CONST_REVISION_KEYNAME = "RevisionVersion"
Const CONST_BUILDNO_KEYNAME = "BuildNO"

Const Project_Pos = 0
Const ProjFile_Pos = 1
Const Version_Pos = 2
Const CMProjStat_Pos = 3
Const CommitProjStat_Pos = 4

Dim mProjList() As String
Dim msIniFile As String
Dim fso As FileSystemObject
Dim msIniPath As String
Dim mLogFile As String
Dim mIsFailure As Boolean
Dim msVB6Path As String
Dim msTFSPath As String

Private Sub cmdGetIni_Click()
   Dim sIniFile As String
   Dim sIniPath As String
   
   On Error GoTo SubError:
   
   'nEW LINE ADDED IN OLDER VERSION
   
   If Len(lblRemakeIniFile.Caption) > 0 Then
      sIniFile = lblRemakeIniFile.Caption
      cd1.FileName = sIniFile
      If fso.FileExists(sIniFile) Then
         msIniPath = fso.GetParentFolderName(sIniFile)
      End If
   Else
      cd1.FileName = ""
   End If

   cd1.DialogTitle = "Select Fleet App Builder Ini File"
   cd1.InitDir = msIniPath
   cd1.Flags = cdlOFNReadOnly
   cd1.CancelError = True
   cd1.Filter = "INI (*.ini)|snAppBuilder.ini"

   cd1.DefaultExt = "INI"
   cd1.ShowOpen
   
   lblRemakeIniFile.Caption = cd1.FileName
   
   msIniFile = cd1.FileName
   msIniPath = fso.GetParentFolderName(msIniFile)
   txtSourceFolder = fso.GetParentFolderName(msIniPath)

   SetRegVal "AppBuilderPath", msIniFile
   SetRegVal "SourceFolder", txtSourceFolder.Text

   msVB6Path = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_VB6PATH_KEYNAME)
   msTFSPath = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_TFSPATH_KEYNAME)
   txtMajor = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_MAJORVER_KEYNAME)
   txtMinor = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_MINORVER_KEYNAME)
   txtRevision = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_REVISION_KEYNAME)
   txtBuildNo = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_BUILDNO_KEYNAME)
   
   ReadIniFile
   
   Exit Sub
SubError:
   ' common dialog list canceled lookup...keep same value
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If

End Sub

Private Sub cmdCheckIn_Click()
   Dim FileHandle As Long
   Dim OutFName As String
   Dim itmX As ListItem
   Dim strProj As String
   Dim Parms As String
   Dim sPath As String
   Dim VerNum As String
   Dim sCommand As String
   Dim sBuildFile As String
   Dim sStatusFile As String
   
   On Error GoTo FunctionError
   
   mIsFailure = False
   
   mLogFile = App.Path & "\" & "FleetBuild.Log"
   
   SetAutoIncrement "1"
   
   OutFName = App.Path & "\" & "TFSBAT.BAT"
   VerNum = Trim(txtMajor.Text) & "." & Trim(txtMinor.Text) & "." & Trim(txtRevision.Text) & "." & Trim(txtBuildNo.Text)
   
   For Each itmX In LV1.ListItems
      itmX.EnsureVisible
      FinalizeFile itmX.Text, VerNum, itmX
     
   Next itmX
   
   sBuildFile = txtSourceFolder & "\Deploy\" & "mw_setup.wsi"
   FinalizeFile sBuildFile, VerNum
      
'   FileHandle = FreeFile()
'   Open OutFName For Output As FileHandle
'
'   Print #FileHandle, "git push ""git@github.com:ssubramanian-sn/GitTest.git"""
'   Print #FileHandle, "Pause"
'
'   Close FileHandle
'
' '  ShellAndWait OutFName, vbHide
'
'   RunBat App.Path & "\" & "TFSBAT.BAT"
   
   UpdateLogFile ("Completing Build|" & txtMajor.Text & "." & txtMinor.Text & "." & txtRevision.Text & "." & txtBuildNo.Text & "|")
   
   If mIsFailure Then
      lblStatus.Caption = "Build Faild. Please see the log for details"
      UpdateLogFile ("Build |" & txtMajor.Text & "." & txtMinor.Text & "." & txtRevision.Text & "." & txtBuildNo.Text & "|Success")
   Else
      UpdateLogFile ("Build |" & txtMajor.Text & "." & txtMinor.Text & "." & txtRevision.Text & "." & txtBuildNo.Text & "|Fail")
   End If
   
   mIsFailure = False
   
   'ReadIniFile
   Exit Sub
FunctionError:
   MsgBox Err.Description
   Resume Next

End Sub

Private Sub cmdCheckOut_Click()

   Dim itmX As ListItem
   Dim sBuildFile As String
   Dim OutFName As String
   Dim sStatusFile As String
   Dim Parms As String
   Dim FileHandle As Long

   
   On Error GoTo FunctionError
   
   
   mLogFile = App.Path & "\" & "FleetBuild.Log"
   
   If fso.FileExists(mLogFile) Then
      fso.DeleteFile (mLogFile)
   End If
   
   UpdateLogFile ("Starting Build|" & txtMajor.Text & "." & txtMinor.Text & "." & txtRevision.Text & "." & txtBuildNo.Text & "|")
   UpdateLogFile ("Source Folder|" & txtSourceFolder & "|")
   UpdateLogFile ("Build Bat File|" & lblRemakeIniFile & "|")
   
   mIsFailure = False
   
   DoEvents
   DoEvents
   DoEvents
   DoEvents
   
   For Each itmX In LV1.ListItems
      
      itmX.EnsureVisible
      
      ProcessFile itmX
   
   Next itmX
   
   sBuildFile = txtSourceFolder & "\Deploy\" & "mw_setup.wsi"
   
   If mIsFailure Then
      lblStatus.Caption = "Build Faild. Please see the log for details"
   End If
   mIsFailure = False
   Exit Sub
FunctionError:
   
End Sub

Private Sub cmdSetVer_Click()
   Dim itmX As ListItem
   Dim sVBPPath As String
   
   
   For Each itmX In LV1.ListItems
      sVBPPath = itmX.Text
   
      PutIniParameter sVBPPath, "MajorVer", txtMajor.Text
      PutIniParameter sVBPPath, "MinorVer", txtMinor.Text
      PutIniParameter sVBPPath, "RevisionVer", txtRevision.Text & Format(txtBuildNo.Text, "00")
      
      PutIniParameter sVBPPath, "AutoIncrementVer", "0"
      
   Next itmX
   
   ReadIniFile
   
End Sub
Private Sub SetAutoIncrement(OnOff As String)
   Dim itmX As ListItem
   Dim strVBPPath As String
   
   For Each itmX In LV1.ListItems
      strVBPPath = itmX.Text
   
      PutIniParameter strVBPPath, "AutoIncrementVer", OnOff
   Next itmX
   
End Sub



Private Sub Form_Load()
On Error GoTo Error_Trap:
   
   Set fso = New FileSystemObject
   
   LV1.ColumnHeaders.Add , , "Project", 0
   LV1.ColumnHeaders.Add , , "Project File", 4000
   LV1.ColumnHeaders.Add , , "Version", 1500
   LV1.ColumnHeaders.Add , , "Compile", 2000
   LV1.ColumnHeaders.Add , , "Commit", 2000


   txtSourceFolder = GetRegVal("SourceFolder", "E:\GIT\MWS_DEV")
   lblRemakeIniFile.Caption = GetRegVal("AppBuilderPath", txtSourceFolder & "\Build\snAppBuilder.ini")
   
   If fso.FileExists(lblRemakeIniFile.Caption) Then
      msIniFile = lblRemakeIniFile.Caption
      
      msVB6Path = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_VB6PATH_KEYNAME)
      msTFSPath = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_TFSPATH_KEYNAME)
      txtMajor = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_MAJORVER_KEYNAME)
      txtMinor = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_MINORVER_KEYNAME)
      txtRevision = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_REVISION_KEYNAME)
      txtBuildNo = GetIni(msIniFile, CONST_HEADER_SECTION, CONST_BUILDNO_KEYNAME)
   
      ReadIniFile
   End If
   Exit Sub
Error_Trap:
   MsgBox Err.Description
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   KillObject fso

End Sub


Private Sub ReadIniFile()
   Dim ts As TextStream
   Dim sBuffer As String
   Dim sVBPPath As String
   Dim sFolderName As String
   Dim sProjPath As String
   Dim nPos As Integer
   Dim sProjFileName As String
   Dim sRootFolder As String
   Dim sMajorVer As String
   Dim sMinorVer As String
   Dim sRevision As String
   Dim sBuildNo As String
   Dim sVersion As String
   Dim itmX As ListItem
   
   On Error GoTo SubError:
   
   LV1.ListItems.Clear
   
   If Not fso.FileExists(msIniFile) Then
'      RaiseErrorL "INI File Missing: " & msIniFile, Err
'      LoadINIConfig = False
      Exit Sub
   End If
   
   msIniPath = fso.GetParentFolderName(msIniFile)
   sRootFolder = txtSourceFolder
   
   ReDim mProjList(0) As String
   
   Set ts = fso.OpenTextFile(msIniFile)
   Do While Not ts.AtEndOfStream
      sBuffer = ts.ReadLine
      If UCase(sBuffer) = "[HEADER]" Then
         sBuffer = ""
      End If
      If Left(sBuffer, 1) = "[" Then
         sFolderName = Mid(sBuffer, 2, Len(sBuffer) - 2)
         ReDim Preserve mProjList(nPos + 1) As String
         mProjList(nPos) = sFolderName
         nPos = nPos + 1
      End If
   Loop
   ts.Close
   Set ts = Nothing
      
   For nPos = 0 To UBound(mProjList) - 1
      sProjPath = GetIni(msIniFile, mProjList(nPos), CONST_RELATIVE_PATH)
      sVBPPath = sRootFolder & Replace(sProjPath, "..", "")
'      sVBPPath = fso.GetAbsolutePathName(sVBPPath)
      sProjFileName = fso.GetFileName(sVBPPath)
            
      Set itmX = LV1.ListItems.Add
      itmX.Text = sVBPPath
      
      If fso.FileExists(sVBPPath) Then
         sMajorVer = GetIniParameter(sVBPPath, "", "MajorVer")
         sMinorVer = GetIniParameter(sVBPPath, "", "MinorVer")
         sRevision = GetIniParameter(sVBPPath, "", "RevisionVer")
         sBuildNo = Mid(sRevision, 2, 2)
         sVersion = sMajorVer & "." & sMinorVer & "." & Left(sRevision, 1) & "." & Val(sBuildNo)
         itmX.SubItems(Version_Pos) = sVersion
         itmX.SubItems(ProjFile_Pos) = sProjFileName
      End If
   
   Next nPos
            
   Exit Sub
   
SubError:
   MsgBox Err.Description
   If Not ts Is Nothing Then
      Set ts = Nothing
   End If

End Sub


Private Function CheckStatusLog(sStatusFile As String) As Boolean
   Dim ts As TextStream
   Dim sBuffer As String
   Dim sElements() As String
   Dim itmX As ListItem

   On Error GoTo FunctionError

   If Not fso.FileExists(sStatusFile) Then
      CheckStatusLog = False
      Exit Function
   End If

   Set ts = fso.OpenTextFile(sStatusFile)
   Do While Not ts.AtEndOfStream
      sBuffer = ts.ReadLine
      If Len(sBuffer) = 0 Then
         sBuffer = " "
      End If
      If InStr(1, sBuffer, "Status=", vbTextCompare) > 0 Then
         sElements = Split(sBuffer, "=", , vbBinaryCompare)
         If UBound(sElements) > 0 Then
            If Trim(sElements(1)) = "0" Then
               CheckStatusLog = True
            End If
         End If
      End If
   Loop
   ts.Close

   Set ts = Nothing

   Exit Function

FunctionError:

'   RaiseErrorL "Error processing Site INI File: " & strFile, Err
   If Not ts Is Nothing Then Set ts = Nothing

End Function


Private Function SetProjectVersion(sProjectFile As String) As Boolean
   Dim sAppTitle As String
   On Error GoTo FunctionError:
   
   sAppTitle = GetIniParameter(sProjectFile, "", "Title")
   sAppTitle = Replace(sAppTitle, """", "")
   
   PutIniParameter sProjectFile, "MajorVer", txtMajor.Text
   PutIniParameter sProjectFile, "MinorVer", txtMinor.Text
   PutIniParameter sProjectFile, "RevisionVer", txtRevision.Text & Format(txtBuildNo.Text, "00") & "0"
   PutIniParameter sProjectFile, "VersionComments", ""
   PutIniParameter sProjectFile, "VersionFileDescription", Trim(sAppTitle) & " Ver:" & Trim(txtMajor.Text) & "." & Trim(txtMinor.Text) & "." & Trim(txtRevision.Text) & " " & "Build " & Val(txtBuildNo.Text)
   
   PutIniParameter sProjectFile, "AutoIncrementVer", "0"
   
   SetProjectVersion = True
   Exit Function
FunctionError:
   MsgBox Err.Description
   SetProjectVersion = False
End Function

Private Function UpdateLogFile(strMsg As String) As Boolean
   On Error GoTo UpdateLogFile_error
   Dim sOut As String
   Dim ts As TextStream
   Set ts = fso.OpenTextFile(mLogFile, ForAppending, True)
   
   If ts Is Nothing Then
      MsgBox " Log file " & mLogFile & " could not be opened for writing: " & Err.Number & "-" & Err.Description
      UpdateLogFile = False
      Exit Function
   End If
   sOut = Format(Now, "yyyy-mm-dd hh:mm:ss") & "|" & strMsg
   ts.WriteLine sOut
   ts.Close
   UpdateLogFile = True
   Exit Function
UpdateLogFile_error:
   ts.Close
   UpdateLogFile = False
End Function



Private Function ProcessFile(lstItem As ListItem) As Boolean
On Error GoTo ErrorTrap:
   Dim FileHandle As Long
   Dim OutFName As String
   Dim sProj As String
   Dim sExec As String
   Dim Parms As String
'   Dim sPath As String
   Dim sCommand As String
   Dim sStatusFile As String
   Dim objFile As File
   Dim sVer() As String
   Dim sBuildInfo As String
   Dim sBuildNoText As String
   Dim sMajorVer As String
   Dim sMinorVer As String
   Dim sRevision As String
   Dim sBuildNo As String
   Dim sVersion As String
   
   OutFName = App.Path & "\" & "TFSBAT.BAT"
   sStatusFile = App.Path & "\" & "TFSStatus.log"
   
   sProj = lstItem.Text
   
   lblStatus = "Updating build Version : " & sProj
   
   If SetProjectVersion(sProj) = True Then
      sMajorVer = GetIniParameter(sProj, "", "MajorVer")
      sMinorVer = GetIniParameter(sProj, "", "MinorVer")
      sRevision = GetIniParameter(sProj, "", "RevisionVer")
      sBuildNo = Val(Mid(sRevision, 2, 2))
      sVersion = sMajorVer & "." & sMinorVer & "." & Left(sRevision, 1) & "." & sBuildNo
      lstItem.SubItems(Version_Pos) = sVersion
   
      UpdateLogFile ("Updating build Version|" & sProj & "|Success")
   Else
      UpdateLogFile ("Updating build Version|" & sProj & "|Fail")
      HighlightFailedItem lstItem
      mIsFailure = True
   End If
   
   DoEvents
   DoEvents
   DoEvents
   DoEvents
   
   lblStatus = "Compiling Project " & sProj
   
   FileHandle = FreeFile()
   Open OutFName For Output As FileHandle
   
   Parms = " /make " & sProj
   
   Print #FileHandle, msVB6Path & Parms
   '      Print #FileHandle, "ECHO CompileStatus=%ERRORLEVEL% >>" & sStatusFile
   
   Close FileHandle
   
   ShellAndWait OutFName, vbHide
   
   
   If fso.FileExists(sExec) Then   '11.206.0.2653)
      sVer = Split(fso.GetFileVersion(sExec), ".")
      
      If UBound(sVer) > 0 Then
        Set objFile = fso.GetFile(sExec)
'        sBuildNoText = Trim("Build " & txtBuildNo.Text)
'         sBuildInfo = GetExtendedProperties(sExec, "VersionComments")
       '  sBuildInfo = GetVersionProperty(sExec, "Comments")
         sBuildInfo = Mid(Format(sVer(3), "0000"), 2, 2)

         If sVer(0) = Trim(txtMajor.Text) And sVer(1) = Trim(txtMinor.Text) And Left(Format(sVer(3), "0000"), 1) = Trim(txtRevision.Text) And Val(sBuildInfo) = Val(Trim(txtBuildNo)) Then
            lstItem.SubItems(CMProjStat_Pos) = "P"
            UpdateLogFile ("Compiling Project|" & sProj & "|Success")
         Else
            lstItem.SubItems(CMProjStat_Pos) = "F"
            HighlightFailedItem lstItem
            UpdateLogFile ("Compiling Project|" & sProj & "|Fail")
            mIsFailure = True
         End If
      End If
   Else
      UpdateLogFile ("Compiling Project|" & sProj & "|Fail")
      mIsFailure = True
      HighlightFailedItem lstItem
   End If
   DoEvents
   DoEvents
   DoEvents
   DoEvents
   
   lblStatus = "Digitally Sign in File " & sExec
   
   FileHandle = FreeFile()
   Open OutFName For Output As FileHandle
   
   Parms = " signtool sign /a /du http://www.shipnet.no/ /t http://timestamp.verisign.com/scripts/timstamp.dll /q """ & sExec & """"
   
   Print #FileHandle, Parms
   '      Print #FileHandle, "ECHO CompileStatus=%ERRORLEVEL% >>" & sStatusFile
   
   Close FileHandle
   
   ShellAndWait OutFName, vbHide
   
   
   DoEvents
   DoEvents
   DoEvents
   DoEvents
   
   ProcessFile = True
   Exit Function
ErrorTrap:
   ProcessFile = False
End Function

Private Sub HighlightFailedItem(lstItem As ListItem)
   Dim i As Integer
   lstItem.Bold = True
   lstItem.ForeColor = vbRed
   For i = 1 To lstItem.ListSubItems.Count
      lstItem.ListSubItems(i).Bold = True
      lstItem.ListSubItems(i).ForeColor = vbRed
   Next i
   mIsFailure = True
End Sub


Private Function FinalizeFile(sFileName As String, VerNum As String, Optional lstItem As ListItem) As Boolean
On Error GoTo ErrorTrap:
   Dim FileHandle As Long
   Dim OutFName As String
   Dim sProj As String
   Dim sExec As String
   Dim Parms As String
'   Dim sPath As String
   Dim sCommand As String
   Dim sStatusFile As String
   Dim objFile As File
   Dim sVer() As String
   
   OutFName = App.Path & "\" & "TFSBAT.BAT"
   sStatusFile = App.Path & "\" & "TFSStatus.log"

   If fso.FileExists(sStatusFile) Then
      fso.DeleteFile (sStatusFile)
   End If

   FileHandle = FreeFile()
   Open OutFName For Output As FileHandle

   sProj = sFileName
   lblStatus = "Staging File" & sProj

   Parms = " git stage  " & sProj
   Parms = Parms & " > " & sStatusFile

   Print #FileHandle, Parms

   Print #FileHandle, "ECHO Status=%ERRORLEVEL% >>" & sStatusFile

   Close FileHandle

   ShellAndWait OutFName, vbHide

   If CheckStatusLog(sStatusFile) = True Then
      UpdateLogFile ("Stage File |" & sProj & "|Success")
      If Not lstItem Is Nothing Then
        lstItem.SubItems(CommitProjStat_Pos) = "P"
      End If
   Else
      UpdateLogFile ("Stage File|" & sProj & "|Fail")
      If Not lstItem Is Nothing Then
          lstItem.SubItems(CommitProjStat_Pos) = "F"
          HighlightFailedItem lstItem
      End If
      mIsFailure = True
   End If
   DoEvents
   DoEvents
   DoEvents
   DoEvents
   
   OutFName = App.Path & "\" & "TFSBAT.BAT"
   sStatusFile = App.Path & "\" & "TFSStatus.log"

   If fso.FileExists(sStatusFile) Then
      fso.DeleteFile (sStatusFile)
   End If

   FileHandle = FreeFile()
   Open OutFName For Output As FileHandle

   sProj = lstItem.Text
   lblStatus = "Commit File" & sProj

   Parms = " git commit -m " & VerNum
   Parms = Parms & " > " & sStatusFile

   Print #FileHandle, Parms
   Print #FileHandle, "ECHO Status=%ERRORLEVEL% >>" & sStatusFile

   Close FileHandle

   ShellAndWait OutFName, vbHide

   If CheckStatusLog(sStatusFile) = True Then
      UpdateLogFile ("Commit File |" & sProj & "|Success")
      If Not lstItem Is Nothing Then
          lstItem.SubItems(CommitProjStat_Pos) = "P"
      End If
   Else
      UpdateLogFile ("Commit File|" & sProj & "|Fail")
      If Not lstItem Is Nothing Then
          lstItem.SubItems(CommitProjStat_Pos) = "F"
          HighlightFailedItem lstItem
      End If
      mIsFailure = True
   End If
   DoEvents
   DoEvents
   DoEvents
   DoEvents
   
   FinalizeFile = True
   Exit Function
ErrorTrap:
   FinalizeFile = False
End Function

