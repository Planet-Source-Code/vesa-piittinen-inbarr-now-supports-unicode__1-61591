Attribute VB_Name = "MerriInBArr"
Option Explicit
Public Function InBArr(ByRef ByteArray() As Byte, ByRef KeyWord As String, Optional StartPos As Long = 0, Optional Compare As Byte = vbBinaryCompare, Optional IsUnicode As Boolean = True) As Long
    Static KeyBuffer() As Byte, KeyBufferU() As Byte
    Dim A As Long, B As Long, C As Long, KeyLen As Long, KeyUpper As Long
    Dim FirstKeyByte As Byte, LastKeyByte As Byte, TempByte As Byte
    Dim FirstKeyByte2 As Byte, LastKeyByte2 As Byte, TempByte2 As Byte
    Dim FirstKeyByteU As Byte, LastKeyByteU As Byte
    Dim FirstKeyByte2U As Byte, LastKeyByte2U As Byte
    If (Not ByteArray) = True Then InBArr = -1: Exit Function
    If LenB(KeyWord) = 0 Then InBArr = StartPos: Exit Function
    If Compare = vbBinaryCompare Then
        KeyBuffer = KeyWord
    Else
        KeyBufferU = UCase$(KeyWord)
        KeyBuffer = LCase$(KeyWord)
    End If
    KeyLen = UBound(KeyBuffer) - 1
    If StartPos < 0 Then StartPos = 0
    If IsUnicode Then
        If KeyLen > UBound(ByteArray) - 1 Then InBArr = -1: Exit Function
        If (StartPos Mod 2) = 1 Then StartPos = StartPos - (StartPos Mod 2) + 2
        If StartPos > UBound(ByteArray) - KeyLen Then InBArr = -1: Exit Function
        FirstKeyByte = KeyBuffer(0)
        LastKeyByte = KeyBuffer(KeyLen)
        FirstKeyByte2 = KeyBuffer(1)
        LastKeyByte2 = KeyBuffer(KeyLen + 1)
        If Compare = vbBinaryCompare Then
            'loop through the array
            For A = StartPos To UBound(ByteArray) - KeyLen - 1 Step 2
                If ByteArray(A) = FirstKeyByte And ByteArray(A + 1) = FirstKeyByte2 Then
                    If ByteArray(A + KeyLen) = LastKeyByte And ByteArray(A + KeyLen + 1) = LastKeyByte2 Then
                        If KeyLen > 4 Then
                            'check if keyword is found from the array
                            C = A + 2
                            For B = 2 To KeyLen - 1 Step 2
                                If Not (ByteArray(C) = KeyBuffer(B) And ByteArray(C + 1) = KeyBuffer(B + 1)) Then Exit For
                                C = C + 2
                            Next B
                            'keyword is found!
                            If B >= KeyLen Then
                                InBArr = A
                                Exit Function
                            End If
                        Else
                            InBArr = A
                            Exit Function
                        End If
                    End If
                End If
            Next A
        Else 'vbTextCompare
            FirstKeyByteU = KeyBufferU(0)
            LastKeyByteU = KeyBufferU(KeyLen)
            FirstKeyByte2U = KeyBufferU(1)
            LastKeyByte2U = KeyBufferU(KeyLen + 1)
            'loop through the array
            For A = StartPos To UBound(ByteArray) - KeyLen - 1 Step 2
                TempByte = ByteArray(A)
                TempByte2 = ByteArray(A + 1)
                If (TempByte = FirstKeyByte Or TempByte = FirstKeyByteU) And (TempByte2 = FirstKeyByte2 Or TempByte2 = FirstKeyByte2U) Then
                    TempByte = ByteArray(A + KeyLen)
                    TempByte2 = ByteArray(A + KeyLen + 1)
                    If (TempByte = LastKeyByte Or TempByte = LastKeyByteU) And (TempByte2 = LastKeyByte2 Or TempByte2 = LastKeyByte2U) Then
                        If KeyLen > 4 Then
                            'check if keyword is found from the array
                            C = A + 2
                            For B = 2 To KeyLen - 1 Step 2
                                TempByte = ByteArray(C)
                                TempByte2 = ByteArray(C + 1)
                                If Not ((TempByte = KeyBuffer(B) Or TempByte = KeyBufferU(B)) And (TempByte2 = KeyBuffer(B + 1) Or TempByte2 = KeyBufferU(B + 1))) Then Exit For
                                C = C + 2
                            Next B
                            'keyword is found!
                            If B >= KeyLen Then
                                InBArr = A
                                Exit Function
                            End If
                        Else
                            InBArr = A
                            Exit Function
                        End If
                    End If
                End If
            Next A
        End If
    Else
        KeyUpper = KeyLen \ 2
        If KeyUpper > UBound(ByteArray) Then InBArr = -1: Exit Function
        If StartPos > UBound(ByteArray) - KeyUpper Then InBArr = -1: Exit Function
        FirstKeyByte = KeyBuffer(0)
        LastKeyByte = KeyBuffer(KeyLen)
        If Compare = vbBinaryCompare Then
            'loop through the array
            Debug.Print StartPos
            For A = StartPos To UBound(ByteArray) - KeyUpper
                If ByteArray(A) = FirstKeyByte Then
                    If ByteArray(A + KeyUpper) = LastKeyByte Then
                        If KeyLen > 4 Then
                            'check if keyword is found from the array
                            C = A + 1
                            For B = 2 To KeyLen Step 2
                                If Not (ByteArray(C) = KeyBuffer(B)) Then Exit For
                                C = C + 1
                            Next B
                            'keyword is found!
                            If B > KeyLen Then
                                InBArr = A
                                Exit Function
                            End If
                        Else
                            InBArr = A
                            Exit Function
                        End If
                    End If
                End If
            Next A
        Else 'vbTextCompare
            FirstKeyByteU = KeyBufferU(0)
            LastKeyByteU = KeyBufferU(KeyLen)
            'loop through the array
            For A = StartPos To UBound(ByteArray) - KeyUpper
                TempByte = ByteArray(A)
                If TempByte = FirstKeyByte Or TempByte = FirstKeyByteU Then
                    TempByte = ByteArray(A + KeyUpper)
                    If TempByte = LastKeyByte Or TempByte = LastKeyByteU Then
                        If KeyLen > 4 Then
                            'check if keyword is found from the array
                            C = A + 1
                            For B = 2 To KeyLen Step 2
                                TempByte = ByteArray(C)
                                If Not (TempByte = KeyBuffer(B) Or TempByte = KeyBufferU(B)) Then Exit For
                                C = C + 1
                            Next B
                            'keyword is found!
                            If B > KeyLen Then
                                InBArr = A
                                Exit Function
                            End If
                        Else
                            InBArr = A
                            Exit Function
                        End If
                    End If
                End If
            Next A
        End If
    End If
    InBArr = -1
End Function
Public Function InBArrRev(ByRef ByteArray() As Byte, ByRef KeyWord As String, Optional StartPos As Long = -1, Optional Compare As Byte = vbBinaryCompare, Optional IsUnicode As Boolean = True) As Long
    Static KeyBuffer() As Byte, KeyBufferU() As Byte
    Dim A As Long, B As Long, C As Long, KeyLen As Long, KeyUpper As Long
    Dim FirstKeyByte As Byte, LastKeyByte As Byte, TempByte As Byte
    Dim FirstKeyByte2 As Byte, LastKeyByte2 As Byte, TempByte2 As Byte
    Dim FirstKeyByteU As Byte, LastKeyByteU As Byte
    Dim FirstKeyByte2U As Byte, LastKeyByte2U As Byte
    If (Not ByteArray) = True Then InBArrRev = -1: Exit Function
    If LenB(KeyWord) = 0 Then InBArrRev = StartPos: Exit Function
    If Compare = vbBinaryCompare Then
        KeyBuffer = KeyWord
    Else
        KeyBufferU = UCase$(KeyWord)
        KeyBuffer = LCase$(KeyWord)
    End If
    KeyLen = UBound(KeyBuffer) - 1
    If IsUnicode Then
        If KeyLen > UBound(ByteArray) - 1 Then InBArrRev = -1: Exit Function
        If StartPos < 0 Then StartPos = UBound(ByteArray) - KeyLen - 1
        If (StartPos Mod 2) = 1 Then StartPos = StartPos - (StartPos Mod 2) + 2
        If StartPos >= UBound(ByteArray) - KeyLen Then InBArrRev = -1: Exit Function
        FirstKeyByte = KeyBuffer(0)
        LastKeyByte = KeyBuffer(KeyLen)
        FirstKeyByte2 = KeyBuffer(1)
        LastKeyByte2 = KeyBuffer(KeyLen + 1)
        If Compare = vbBinaryCompare Then
            'loop through the array
            For A = StartPos To 0 Step -2
                If ByteArray(A) = FirstKeyByte And ByteArray(A + 1) = FirstKeyByte2 Then
                    If ByteArray(A + KeyLen) = LastKeyByte And ByteArray(A + KeyLen + 1) = LastKeyByte2 Then
                        If KeyLen > 4 Then
                            'check if keyword is found from the array
                            C = A + 2
                            For B = 2 To KeyLen - 1 Step 2
                                If Not (ByteArray(C) = KeyBuffer(B) And ByteArray(C + 1) = KeyBuffer(B + 1)) Then Exit For
                                C = C + 2
                            Next B
                            'keyword is found!
                            If B >= KeyLen Then
                                InBArrRev = A
                                Exit Function
                            End If
                        Else
                            InBArrRev = A
                            Exit Function
                        End If
                    End If
                End If
            Next A
        Else 'vbTextCompare
            FirstKeyByteU = KeyBufferU(0)
            LastKeyByteU = KeyBufferU(KeyLen)
            FirstKeyByte2U = KeyBufferU(1)
            LastKeyByte2U = KeyBufferU(KeyLen + 1)
            'loop through the array
            For A = StartPos To 0 Step -2
                TempByte = ByteArray(A)
                TempByte2 = ByteArray(A + 1)
                If (TempByte = FirstKeyByte Or TempByte = FirstKeyByteU) And (TempByte2 = FirstKeyByte2 Or TempByte2 = FirstKeyByte2U) Then
                    TempByte = ByteArray(A + KeyLen)
                    TempByte2 = ByteArray(A + KeyLen + 1)
                    If (TempByte = LastKeyByte Or TempByte = LastKeyByteU) And (TempByte2 = LastKeyByte2 Or TempByte2 = LastKeyByte2U) Then
                        If KeyLen > 4 Then
                            'check if keyword is found from the array
                            C = A + 2
                            For B = 2 To KeyLen - 1 Step 2
                                TempByte = ByteArray(C)
                                TempByte2 = ByteArray(C + 1)
                                If Not ((TempByte = KeyBuffer(B) Or TempByte = KeyBufferU(B)) And (TempByte2 = KeyBuffer(B + 1) Or TempByte2 = KeyBufferU(B + 1))) Then Exit For
                                C = C + 2
                            Next B
                            'keyword is found!
                            If B >= KeyLen Then
                                InBArrRev = A
                                Exit Function
                            End If
                        Else
                            InBArrRev = A
                            Exit Function
                        End If
                    End If
                End If
            Next A
        End If
    Else
        KeyUpper = KeyLen \ 2
        If KeyUpper > UBound(ByteArray) Then InBArrRev = -1: Exit Function
        If StartPos < 0 Then StartPos = UBound(ByteArray) - KeyUpper
        If StartPos > UBound(ByteArray) - KeyUpper Then InBArrRev = -1: Exit Function
        FirstKeyByte = KeyBuffer(0)
        LastKeyByte = KeyBuffer(KeyLen)
        If Compare = vbBinaryCompare Then
            'loop through the array
            Debug.Print StartPos
            For A = StartPos To 0 Step -1
                If ByteArray(A) = FirstKeyByte Then
                    If ByteArray(A + KeyUpper) = LastKeyByte Then
                        If KeyLen > 4 Then
                            'check if keyword is found from the array
                            C = A + 1
                            For B = 2 To KeyLen Step 2
                                If Not (ByteArray(C) = KeyBuffer(B)) Then Exit For
                                C = C + 1
                            Next B
                            'keyword is found!
                            If B > KeyLen Then
                                InBArrRev = A
                                Exit Function
                            End If
                        Else
                            InBArrRev = A
                            Exit Function
                        End If
                    End If
                End If
            Next A
        Else 'vbTextCompare
            FirstKeyByteU = KeyBufferU(0)
            LastKeyByteU = KeyBufferU(KeyLen)
            'loop through the array
            For A = StartPos To 0 Step -1
                TempByte = ByteArray(A)
                If TempByte = FirstKeyByte Or TempByte = FirstKeyByteU Then
                    TempByte = ByteArray(A + KeyUpper)
                    If TempByte = LastKeyByte Or TempByte = LastKeyByteU Then
                        If KeyLen > 4 Then
                            'check if keyword is found from the array
                            C = A + 1
                            For B = 2 To KeyLen Step 2
                                TempByte = ByteArray(C)
                                If Not (TempByte = KeyBuffer(B) Or TempByte = KeyBufferU(B)) Then Exit For
                                C = C + 1
                            Next B
                            'keyword is found!
                            If B > KeyLen Then
                                InBArrRev = A
                                Exit Function
                            End If
                        Else
                            InBArrRev = A
                            Exit Function
                        End If
                    End If
                End If
            Next A
        End If
    End If
    InBArrRev = -1
End Function
