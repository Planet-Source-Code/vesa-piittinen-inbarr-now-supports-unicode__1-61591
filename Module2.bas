Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function lstrlenW Lib "kernel32" _
    (ByRef lpString As Long) _
    As Long

Private Declare Sub RtlMoveMemory Lib "ntdll.dll" _
    (ByRef lpvDest As Any, _
     ByRef lpvSrc As Any, _
     ByVal cbLen As Long)

Private Declare Function RtlCompareMemory Lib "ntdll.dll" _
    (ByRef lpvSource1 As Any, _
     ByRef lpvSource2 As Any, _
     ByVal cbLen As Long) _
    As Long
'

Public Function InTheString(ByRef lpszStringToSearch As Long, _
                            ByRef lngStringLen As Long, _
                            ByRef pstrSearchFor As String, _
                            Optional ByVal plngStartPos As Long = 1) _
                           As Long

Dim charBuffer()    As Byte
Dim lngString2Len   As Long
Dim i               As Long

    lngString2Len = LenB(pstrSearchFor)

    ReDim charBuffer(lngStringLen - 1)
    
    RtlMoveMemory charBuffer(0), _
                  ByVal lpszStringToSearch, _
                  lngStringLen

    If (lngString2Len < lngStringLen) Then
        For i = (plngStartPos + plngStartPos - 2) To (lngStringLen - lngString2Len) Step 2
            If (RtlCompareMemory(charBuffer(i), _
                                 ByVal StrPtr(pstrSearchFor), _
                                 lngString2Len _
                                ) = lngString2Len) Then
                InTheString = (i \ 2) + 1
                Exit Function
            End If
        Next i
    End If

    InTheString = -1
End Function

