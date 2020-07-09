' ��ȡ���鳤��
Public Function ArrayLength(ByVal ary) As Integer
    ArrayLength = UBound(ary) - LBound(ary) + 1
End Function

' �ж�Item_Name�Ƿ��ڼ���Set_Name�У�������ڣ��򷵻�True
Function Is_In_Set(Set_Name, Item_Name) As Boolean
    Is_In_Set = False

    For Set_Index = 1 To Set_Name.Count
        If Set_Name(Set_Index) = Item_Name Then
            Is_In_Set = True
        End If
    Next
End Function

'��Set_Name�е�����ת��Ϊ�ַ���
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
 
Source_Sheet = "Sheet2" ' ԴSheet
Target_Sheet = "Sheet1" ' Ŀ��Sheet

Dim Type_Set As New Collection ' �洢Typeֵ�ļ���
Dim Type_String As String ' �洢Type���ַ�����ʽ
Dim Split_Result() As String ' �洢Split����ַ�������
'---------------------------------------------------------
For Row_Index_1 = 2 To Sheets(Source_Sheet).UsedRange.Rows.Count
    Set Type_Set = Nothing
    Type_String = ""
    Erase Split_Result ' �ͷ����������ڴ�
    '-----------------------------------------------
    ID_List = Sheets(Source_Sheet).Cells(Row_Index_1, 1)
    Split_Result = VBA.Split(ID_List, "/")

    For Array_Index = 0 To (ArrayLength(Split_Result) - 1) ' ע�⣬���������Ǵ�0��ʼ��
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