' 获取数组长度
Public Function ArrayLength(ByVal ary) As Integer
    ArrayLength = UBound(ary) - LBound(ary) + 1
End Function

' 判断Item_Name是否在集合Set_Name中，如果存在，则返回True
Function Is_In_Set(Set_Name, Item_Name) As Boolean
    Is_In_Set = False

    For Set_Index = 1 To Set_Name.Count
        If Set_Name(Set_Index) = Item_Name Then
            Is_In_Set = True
        End If
    Next
End Function

'将Set_Name中的数字转换为字符串
Function Set_To_String(Set_Name) As String
    Dim Temp_String As String

    For Set_Index = 1 To Set_Name.Count
        If Set_Index = 1 Then
            Temp_String = CStr(Set_Name(Set_Index))
        Else
            Temp_String = Temp_String + "," + CStr(Set_Name(Set_Index))
        End If
    Next

    Set_To_String = Temp_String
End Function

Sub test()
 
Source_Sheet = "Sheet2" ' 源Sheet
Target_Sheet = "Sheet1" ' 目标Sheet

Dim Type_Set As New Collection ' 存储Type值的集合
Dim Type_String As String ' 存储Type的字符串形式
Dim Split_Result() As String ' 存储Split后的字符串数组
'---------------------------------------------------------
For Row_Index_1 = 2 To Sheets(Source_Sheet).UsedRange.Rows.Count
    Set Type_Set = Nothing
    Type_String = ""
    Erase Split_Result ' 释放数组所用内存
    '-----------------------------------------------
    ID_List = Sheets(Source_Sheet).Cells(Row_Index_1, 1)
    Split_Result = VBA.Split(ID_List, "/")

    For Array_Index = 0 To (ArrayLength(Split_Result) - 1) ' 注意，数组的序号是从0开始的
        Dim Current_ID As Integer
        Dim Current_Type As String
        '------------------------------------------
        Current_ID = CInt(Split_Result(Array_Index))
        Current_Type = Sheets(Target_Sheet).Cells(Current_ID + 1, 3)

        If Is_In_Set(Type_Set, Current_Type) = False Then
            Type_Set.Add (Current_Type)
        End If
    Next

    Type_String = Set_To_String(Type_Set)
    '------------------------------------------------------
    Sheets(Source_Sheet).Cells(Row_Index_1, 2) = Type_String
Next

End Sub