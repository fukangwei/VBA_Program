Private Type rc4_key ' 定义一个结构体
    s(256) As Byte
    x As Byte
    y As Byte
End Type

' 初始化RC4
Private Sub prepare_key(ByRef key_data() As Byte, ByRef key As rc4_key)
    Dim i As Long, j As Byte, keylen As Long, c As Integer

    For c = 0 To 255
        key.s(c) = c
    Next

    key.x = 0
    key.y = 0

    i = 0
    j = 0
    keylen = UBound(key_data) - LBound(key_data) + 1

    For c = 0 To 255
        j = ((key_data(i) Mod 256) + key.s(c) + j) Mod 256
        key.s(c) = key.s(c) Xor key.s(j)
        key.s(j) = key.s(c) Xor key.s(j)
        key.s(c) = key.s(c) Xor key.s(j)
        i = (i + 1) Mod keylen
    Next
End Sub

Private Sub rc4(ByRef buff() As Byte, ByRef key As rc4_key) ' RC4加密和解密
    Dim x As Byte, y As Byte, z As Byte, c As Long, ub As Long, lb As Long
    x = key.x
    y = key.y
    ub = UBound(buff)
    lb = LBound(buff)

    For c = lb To ub
        x = (x + 1) Mod 256
        y = ((key.s(x) Mod 256) + y) Mod 256
        key.s(x) = key.s(x) Xor key.s(y)
        key.s(y) = key.s(x) Xor key.s(y)
        key.s(x) = key.s(x) Xor key.s(y)
        z = ((key.s(x) Mod 256) + key.s(y)) Mod 256
        buff(c) = buff(c) Xor key.s(z)
    Next

    key.x = x
    key.y = y
End Sub

Private Function Byte_To_Str(ByRef buff() As Byte)
    Dim Encode As String
    
    ub = UBound(buff)
    lb = LBound(buff)

    For c = lb To ub
        Encode = Encode + Right("0" & hex(buff(c)), 2)
    Next

    Byte_To_Str = Encode
End Function

Private Function Str_To_Byte(Num_String As String) ' 16进制字符转数字
    length_num = Len(Num_String)

    Dim s() As Byte
    ReDim Preserve s(length_num / 2 - 1)

    For i = 1 To (length_num / 2)
        temp = Mid(Num_String, 2 * (i - 1) + 1, 1)
        
        If (Asc(temp) - Asc("0")) < 10 Then
            temp_1 = (Asc(temp) - Asc("0")) * 16
        Else
            temp_1 = (Asc(temp) - Asc("A") + 10) * 16
        End If
        
        temp = Mid(Num_String, 2 * i, 1)
        
        If (Asc(temp) - Asc("0")) < 10 Then
            temp_2 = Asc(temp) - Asc("0")
        Else
            temp_2 = Asc(temp) - Asc("A") + 10
        End If

        s(i - 1) = temp_1 + temp_2
    Next i

    Str_To_Byte = s
End Function

Public Sub Tets()
    Dim s() As Byte, p() As Byte
    Dim enkey As rc4_key, denkey As rc4_key
    Dim Text_String As String

    Dim text() As Byte

    s = "AAAAAA" ' 这是要加密的数据
    p = "hhhh" ' 这是加密钥匙

    Call prepare_key(p, enkey)
    denkey = enkey ' 保留key
    Call rc4(s, enkey)

    Text_String = Byte_To_Str(s)
    MsgBox Text_String ' 显示解密后的数据
    text = Str_To_Byte(Text_String)
    Call rc4(text, denkey)
    MsgBox text ' 显示解密后的数据
End Sub