Function Add_Header(shtName) As Boolean
    Sheets(shtName).Cells(1, 1) = "name"
    Sheets(shtName).Cells(1, 2) = "id"
    Sheets(shtName).Cells(1, 3) = "type"
End Function

Function Is_In_Set(Set_Name, Item_Name) As Boolean
    Is_In_Set = False

    For Set_Index = 1 To Set_Name.Count
        If Set_Name(Set_Index) = Item_Name Then
            Is_In_Set = True
        End If
    Next
End Function

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
    Source_Sheet = "Sheet1"
    Target_Sheet = "Sheet2"

    Add_Header(Target_Sheet)
    '---------------------------------------------------------------------
    Dim Name_Set As New Collection

    For Row_Index = 2 To Sheets(Source_Sheet).UsedRange.Rows.Count
        User_Name = Sheets(Source_Sheet).Cells(Row_Index, 2)

        If Is_In_Set(Name_Set, User_Name) = False Then
            Name_Set.Add(User_Name)
        End If
    Next
    '---------------------------------------------------------------------
    For Set_Index = 1 To Name_Set.Count
        Sheets(Target_Sheet).Cells(Set_Index + 1, 1) = Name_Set(Set_Index)
    Next
    '---------------------------------------------------------------------
    Dim ID_Set As New Collection
    Dim ID_String As String
    Dim Type_Set As New Collection
    Dim Type_String As String

    For Row_Index = 1 To Name_Set.Count
    Set ID_Set = Nothing
    Set Type_Set = Nothing
    ID_String = ""
        Type_String = ""
        '---------------------------------------------------------------------
        Current_Name = Name_Set(Row_Index)

        For Row_Index_1 = 2 To Sheets(Source_Sheet).UsedRange.Rows.Count
            If Sheets(Source_Sheet).Cells(Row_Index_1, 2) = Current_Name Then
                User_ID = Sheets(Source_Sheet).Cells(Row_Index_1, 1)
                ID_Set.Add(User_ID)
                '-------------------------------------------------------------
                User_Type = Sheets(Source_Sheet).Cells(Row_Index_1, 3)

                If Is_In_Set(Type_Set, User_Type) = False Then
                    Type_Set.Add(User_Type)
                End If
            End If
        Next
        '---------------------------------------------------------------------
        ID_String = Set_To_String(ID_Set)
        Type_String = Set_To_String(Type_Set)
        '---------------------------------------------------------------------
        For Row_Index_2 = 2 To Sheets(Target_Sheet).UsedRange.Rows.Count
            If Sheets(Target_Sheet).Cells(Row_Index_2, 1) = Current_Name Then
                Sheets(Target_Sheet).Cells(Row_Index_2, 2) = ID_String
                Sheets(Target_Sheet).Cells(Row_Index_2, 3) = Type_String
            End If
        Next
    Next
End Sub