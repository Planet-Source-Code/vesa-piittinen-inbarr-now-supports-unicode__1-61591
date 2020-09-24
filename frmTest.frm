VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InStr vs. InBArr (by Merri)"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdValidateInBArrayANSI 
      Caption         =   "Validate InBArr (ANSI)"
      Height          =   375
      Left            =   6960
      TabIndex        =   16
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdValidateInBArrayRevANSI 
      Caption         =   "Validate InBArrRev (ANSI)"
      Height          =   375
      Left            =   4680
      TabIndex        =   15
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdValidateInBArray 
      Caption         =   "Validate InBArr"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdValidateInBArrayRev 
      Caption         =   "Validate InBArrRev"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTest.frx":0000
      Left            =   120
      List            =   "frmTest.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   240
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "InBArr"
      Height          =   1935
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdTimeInTheString 
         Caption         =   "InTheString"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton cmdTimeInstr 
         Caption         =   "InStr"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdTimeInBArr 
         Caption         =   "InBArr"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "InBArrRev"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4455
      Begin VB.CommandButton cmdTimeInBArrRev 
         Caption         =   "InBArrRev"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdTimeInStrRev 
         Caption         =   "InStrRev"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TestFile As String, Iterations As Long, KeyWord As String
Private Sub cmdTimeInTheString_Click()
    Dim Buffer As String, FileNumber As Byte, FileLength As Long
    Dim A As Long, Result As Long, Compare As Byte
    Compare = Combo1.ListIndex
    'open file for input
    FileNumber = FreeFile
    FileLength = FileLen(TestFile)
    Open TestFile For Input As #FileNumber
        'read data
        Buffer = Input(FileLength, FileNumber)
    Close #FileNumber
    
    Me.MousePointer = vbHourglass
    Start
    For A = 1 To Iterations
        Result = InTheString(StrPtr(Buffer), LenB(Buffer), KeyWord, 1)
    Next A
    Label5 = Finish & " ms : " & Result
    Me.MousePointer = vbNormal
End Sub
Private Sub cmdValidateInBArray_Click()
    MsgBox IsGoodInBArr
End Sub
Private Sub cmdValidateInBArrayANSI_Click()
    MsgBox IsGoodInBArrANSI
End Sub
Private Sub cmdValidateInBArrayRev_Click()
    MsgBox IsGoodInBArrRev
End Sub
Private Sub cmdValidateInBArrayRevANSI_Click()
    MsgBox IsGoodInBArrRevANSI
End Sub
Private Sub Combo1_Click()
    Label1 = vbNullString
    Label2 = vbNullString
    Label3 = vbNullString
    Label4 = vbNullString
End Sub
Private Sub cmdTimeInBArrRev_Click()
    Dim Buffer() As Byte, FileNumber As Byte
    Dim A As Long, Result As Long, Compare As Byte
    Compare = Combo1.ListIndex
    'open file for input
    FileNumber = FreeFile
    Open TestFile For Binary Access Read As #FileNumber
        'resize buffer
        ReDim Preserve Buffer(LOF(FileNumber) - 1)
        'read data
        Get #FileNumber, , Buffer
    Close #FileNumber
    
    Buffer = StrConv(Buffer, vbUnicode)
    
    Me.MousePointer = vbHourglass
    Start
    For A = 1 To Iterations
        Result = InBArrRev(Buffer, KeyWord, , Compare)
    Next A
    Label1 = Finish & " ms : " & (Result \ 2)
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdTimeInstrRev_Click()
    Dim Buffer As String, FileNumber As Byte
    Dim A As Long, Result As Long, Compare As Byte
    Compare = Combo1.ListIndex
    'open file for input
    FileNumber = FreeFile
    Open TestFile For Input As #FileNumber
        'read data
        Buffer = Input(FileLen(TestFile), FileNumber)
    Close #FileNumber
    
    Me.MousePointer = vbHourglass
    Start
    For A = 1 To Iterations
        Result = InStrRev(Buffer, KeyWord, , Compare)
    Next A
    Label2 = Finish & " ms : " & Result
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdTimeInBArr_Click()
    Dim Buffer() As Byte, FileNumber As Byte, FileLength As Long
    Dim A As Long, Result As Long, Compare As Byte
    Compare = Combo1.ListIndex
    'open file for input
    FileNumber = FreeFile
    Open TestFile For Binary Access Read As #FileNumber
        FileLength = LOF(FileNumber)
        'resize buffer
        ReDim Preserve Buffer(FileLength - 1)
        'read data
        Get #FileNumber, , Buffer
    Close #FileNumber
    
    Buffer = StrConv(Buffer, vbUnicode)
    
    Me.MousePointer = vbHourglass
    Start
    For A = 1 To Iterations
        Result = InBArr(Buffer, KeyWord, , Compare)
    Next A
    Label3 = Finish & " ms : " & (Result \ 2)
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdTimeInstr_Click()
    Dim Buffer As String, FileNumber As Byte, FileLength As Long
    Dim A As Long, Result As Long, Compare As Byte
    Compare = Combo1.ListIndex
    'open file for input
    FileNumber = FreeFile
    FileLength = FileLen(TestFile)
    Open TestFile For Input As #FileNumber
        'read data
        Buffer = Input(FileLength, FileNumber)
    Close #FileNumber
    
    Me.MousePointer = vbHourglass
    Start
    For A = 1 To Iterations
        Result = InStr(1, Buffer, KeyWord, Compare)
    Next A
    Label4 = Finish & " ms : " & Result
    Me.MousePointer = vbNormal
End Sub
Private Sub Form_Load()
    Combo1.Clear
    Combo1.AddItem "vbBinaryCompare"
    Combo1.AddItem "vbTextCompare"
    Combo1.ListIndex = 1
    TestFile = "c:\hotfix.txt"
    KeyWord = "PARTICULAR"
    Iterations = 1000 '100000
End Sub
Public Function IsGoodInBArr(Optional fLigaturesToo As Boolean) As Boolean
    ' verify correct InBArr returns, 2005-07-08
    ' based on InStr test code available at VBspeed
    ' returns True if all tests are passed
    Dim fFailed As Boolean
    Dim Temp() As Byte
    Dim Test As Long
    
    ' replace "InBArr" with the name of your function
    Temp = "abc"
    If InBArr(Temp, "b") <> 2 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "ab") <> 0 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "aB") <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "aB", , vbTextCompare) <> 0 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "ab", 2) <> 4 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "ab", 4) <> 4 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "ab", 6) <> -1 Then Stop: fFailed = True
    Temp = "aaabcab"
    If InBArr(Temp, "abc", 6) <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "", 6) <> 6 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "", 8) <> 8 Then Stop: fFailed = True
    Erase Temp
    If InBArr(Temp, "", 6) <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "c") <> -1 Then Stop: fFailed = True
    
    Temp = "abcdabcd"
    If InBArr(Temp, "abcd") <> 0 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "Ab") <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArr(Temp, "Ab", , vbTextCompare) <> 0 Then Stop: fFailed = True
    
    Temp = "a" & String$(50000, "b")
    If InBArr(Temp, "a") <> 0 Then Stop: fFailed = True
    
    ' unicode
    Temp = "a€€c"
    If InBArr(Temp, "€") <> 2 Then Stop: fFailed = True
    
    ' the 4 stooges: š/Š, œ/Œ, ž/Ž, ÿ/Ÿ (154/138, 156/140, 158/142, 255/159)
    Temp = "Hašiš"
    If InBArr(Temp, "Š", , vbTextCompare) <> 4 Then Stop: fFailed = True
    ' ligatures  textcompare (VBspeed entries do NOT have to pass this test)
    If fLigaturesToo Then
        ' ligatures, a digraphemic fun house: ss/ß, ae/æ, oe/œ, th/þ
        Temp = "Straße"
        If InBArr(Temp, "ss", , vbTextCompare) <> 8 Then Stop: fFailed = True
    End If
    
    ' well done
    IsGoodInBArr = Not fFailed
End Function

Public Function IsGoodInBArrANSI() As Boolean
    ' verify correct InBArr returns (ANSI mode), 2005-07-08
    ' based on InStr test code available at VBspeed
    ' returns True if all tests are passed
    Dim fFailed As Boolean
    Dim Temp() As Byte
    Dim Test As Long
    
    ' replace ".InStr01" with the name of your function
    Temp = StrConv("abc", vbFromUnicode)
    If InBArr(Temp, "b", , , False) <> 1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "ab", , , False) <> 0 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "aB", , , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "aB", , vbTextCompare, False) <> 0 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "ab", 1, , False) <> 2 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "ab", 2, , False) <> 2 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "ab", 3, , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("aaabcab", vbFromUnicode)
    If InBArr(Temp, "abc", 3, , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "", 3, , False) <> 3 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "", 4, , False) <> 4 Then Stop: fFailed = True
    Erase Temp
    If InBArr(Temp, "", 3, , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "c", , , False) <> -1 Then Stop: fFailed = True
    
    Temp = StrConv("abcdabcd", vbFromUnicode)
    If InBArr(Temp, "abcd", , , False) <> 0 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "Ab", , , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArr(Temp, "Ab", , vbTextCompare, False) <> 0 Then Stop: fFailed = True
    
    Temp = StrConv("a" & String$(50000, "b"), vbFromUnicode)
    If InBArr(Temp, "a", , , False) <> 0 Then Stop: fFailed = True
    
    ' well done
    IsGoodInBArrANSI = Not fFailed
End Function
Public Function IsGoodInBArrRev(Optional fLigaturesToo As Boolean) As Boolean
    ' verify correct InBArrRev returns, 2005-07-08
    ' based on InStr test code available at VBspeed
    ' returns True if all tests are passed
    Dim fFailed As Boolean
    Dim Temp() As Byte
    Dim Test As Long
    
    ' replace "InBArrRev" with the name of your function
    Temp = "abc"
    If InBArrRev(Temp, "b") <> 2 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "ab") <> 4 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "aB") <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "aB", , vbTextCompare) <> 4 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "ab", 2) <> 0 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "ab", 4) <> 4 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "ab", 6) <> -1 Then Stop: fFailed = True
    Temp = "aaabcab"
    If InBArrRev(Temp, "abc", 6) <> 4 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "", 6) <> 6 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "", 8) <> 8 Then Stop: fFailed = True
    Erase Temp
    If InBArrRev(Temp, "", 6) <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "c") <> -1 Then Stop: fFailed = True
    
    Temp = "abcdabcd"
    If InBArrRev(Temp, "abcd") <> 8 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "Ab") <> -1 Then Stop: fFailed = True
    Temp = "abab"
    If InBArrRev(Temp, "Ab", , vbTextCompare) <> 4 Then Stop: fFailed = True
    
    Temp = "a" & String$(50000, "b")
    If InBArrRev(Temp, "a") <> 0 Then Stop: fFailed = True
    
    ' unicode
    Temp = "a€€c"
    If InBArrRev(Temp, "€") <> 4 Then Stop: fFailed = True
    
    ' the 4 stooges: š/Š, œ/Œ, ž/Ž, ÿ/Ÿ (154/138, 156/140, 158/142, 255/159)
    Temp = "Hašiš"
    If InBArrRev(Temp, "Š", , vbTextCompare) <> 8 Then Stop: fFailed = True
    ' ligatures  textcompare (VBspeed entries do NOT have to pass this test)
    If fLigaturesToo Then
        ' ligatures, a digraphemic fun house: ss/ß, ae/æ, oe/œ, th/þ
        Temp = "Straße"
        If InBArrRev(Temp, "ss", , vbTextCompare) <> 8 Then Stop: fFailed = True
    End If
    
    ' well done
    IsGoodInBArrRev = Not fFailed
End Function
Public Function IsGoodInBArrRevANSI() As Boolean
    ' verify correct InBArrRev returns, 2005-07-08
    ' based on InStr test code available at VBspeed
    ' returns True if all tests are passed
    Dim fFailed As Boolean
    Dim Temp() As Byte
    Dim Test As Long
    
    ' replace "InBArrRev" with the name of your function
    Temp = StrConv("abc", vbFromUnicode)
    If InBArrRev(Temp, "b", , , False) <> 1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "ab", , , False) <> 2 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "aB", , , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "aB", , vbTextCompare, False) <> 2 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "ab", 1, , False) <> 0 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "ab", 2, , False) <> 2 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "ab", 3, , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("aaabcab", vbFromUnicode)
    If InBArrRev(Temp, "abc", 3, , False) <> 2 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "", 3, , False) <> 3 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "", 4, , False) <> 4 Then Stop: fFailed = True
    Erase Temp
    If InBArrRev(Temp, "", 3, , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "c", , , False) <> -1 Then Stop: fFailed = True
    
    Temp = StrConv("abcdabcd", vbFromUnicode)
    If InBArrRev(Temp, "abcd", , , False) <> 4 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "Ab", , , False) <> -1 Then Stop: fFailed = True
    Temp = StrConv("abab", vbFromUnicode)
    If InBArrRev(Temp, "Ab", , vbTextCompare, False) <> 2 Then Stop: fFailed = True
    
    Temp = StrConv("a" & String$(50000, "b"), vbFromUnicode)
    If InBArrRev(Temp, "a", , , False) <> 0 Then Stop: fFailed = True
    
    ' well done
    IsGoodInBArrRevANSI = Not fFailed
End Function

