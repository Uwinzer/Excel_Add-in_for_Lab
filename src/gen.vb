'Function version 1.1, and it can be used for results sheet and RTD sheets items search
'Combine the separete function from results sheest with RTD sheet(version 1.0)
'Add argument sheet_name, can be used for RTD search and results search
Function Find_Items_From_Sheets(Sheet_Name As String, Item_Name As String, Search_Order As String) As String()
Dim rng As Range, i As Integer, First As String, Last As String
Dim Addr_Str() As String

Dim Count As Integer
Count = 0

'Set up search order
If Search_Order = "ByColumns" Then Set rng = Worksheets(Sheet_Name).Cells.Find(Item_Name, LookAt:=xlPart, LookIn:=xlValues, SearchOrder:=xlByColumns)
If Search_Order = "ByRows" Then Set rng = Worksheets(Sheet_Name).Cells.Find(Item_Name, LookAt:=xlPart, LookIn:=xlValues, SearchOrder:=xlByRows)


'If the rng is empty
If Not rng Is Nothing Then

'Set up column offset or row offset
    If Search_Order = "ByColumns" Then Set Last_rng = Worksheets(Sheet_Name).Cells(rng.Row, rng.Column + 1)
    If Search_Order = "ByRows" Then Set Last_rng = Worksheets(Sheet_Name).Cells(rng.Row + 1, rng.Column)

'record fist postion and first repeative position
    First = rng.Address
    Last = Last_rng.Address
    Do
        Count = Count + 1
        ReDim Preserve Addr_Str(Count)
        Addr_Str(Count) = rng.Address
        'rng.Select
        Set rng = Worksheets(Sheet_Name).Cells.FindNext(rng)
    Loop Until rng.Address = Last
    Addr_Str(0) = Count
    Find_Items_From_Sheets = Addr_Str
    
Else
    MsgBox "No specific item found"
    
End If

End Function

Function Write_From_Results_to_ResultData(ParamArray Items() As Variant)
Dim Finish_Write As Integer
Dim Column_Count As Integer

Column_Index = 1
Finish_Write = 0

While Finish_Write = 0

If Worksheets("ResultData").Cells(1, Column_Index) = "" Then

For i = 1 To CInt(Items(0)(0))
Worksheets("ResultData").Cells(i, Column_Index) = Worksheets("Results").Range(Items(0)(i)).Offset(0, 3)
Next

Finish_Write = 1
Else
Finish_Write = 0
Column_Index = Column_Index + 1
End If

Wend
End Function
'Function version 1.0, and it can automatically generate the torque and speed table based on sepcifications data
Function Resize_Torques_Speeds_Table()

Dim Total_NewRows As Integer
Total_NewRows = 0
Dim Add_Str() As String
Dim Interval As Integer


Max_Speed = Worksheets("Specifications").Cells(200, "D").Value
Max_Torque = Worksheets("Specifications").Cells(201, "D").Value
Speed_Step = Worksheets("Specifications").Cells(202, "D").Value
Torque_Step = Worksheets("Specifications").Cells(203, "D").Value

Torque_num = Max_Torque / Torque_Step
speed_num = Max_Speed / Speed_Step

Total_NewRows = Torque_num * speed_num
ReDim Preserve Add_Str(Total_NewRows)
Interval = 0
For i = 1 To Total_NewRows - 1
        Worksheets("PassReport").Rows("16:16").Insert
    For j = 1 To 8
        Worksheets("PassReport").Range(Cells(Range(Worksheets("PassReport").Cells(16, "A"), Worksheets("PassReport").Cells(16, "A")).Row, Range(Worksheets("PassReport").Cells(16, "A"), Worksheets("PassReport").Cells(16, "A")).Column + Interval + j - 1), Cells(Range(Worksheets("PassReport").Cells(16, "A"), Worksheets("PassReport").Cells(16, "A")).Row, Range(Worksheets("PassReport").Cells(16, "A"), Worksheets("PassReport").Cells(16, "A")).Column + j + Interval)).Merge
        Interval = Interval + 1
    Next
        Interval = 0
Next

For Torque_Index = 1 To Torque_num
    For Speed_Index = 1 To speed_num
        Worksheets("PassReport").Range(Cells(16, "A"), Cells(16, "B")).Offset(Row_Index, 0).Value = Torque_Step * Torque_Index
        Worksheets("PassReport").Range(Cells(16, "A"), Cells(16, "B")).Offset(Row_Index, 3).Value = Speed_Step * Speed_Index
        Row_Index = Row_Index + 1
    Next
Next

End Function
' Function version 1.0, it can generate the map table automatically
Function Create_Maps(Map_Caption As String) As String

Dim Speeds() As Integer
Dim Torques() As Integer
Max_Speed = Worksheets("Specifications").Cells(200, "D").Value
Max_Torque = Worksheets("Specifications").Cells(201, "D").Value
Speed_Step = Worksheets("Specifications").Cells(202, "D").Value
Torque_Step = Worksheets("Specifications").Cells(203, "D").Value

Torque_num = Max_Torque / Torque_Step
speed_num = Max_Speed / Speed_Step
ReDim Preserve Speeds(speed_num)
ReDim Preserve Torques(Torque_num)

Speeds(0) = speed_num
Torques(0) = Torque_num
    
' check if the cell is taken up by other maps, if it is, then find next avaliable one
Dim Finish_Write As Integer
Dim Column_Count As Integer
Dim Avaliable_Cell As Range
Column_Index = 1
Row_Index = 1
Map_Count = 0
Finish_Write = 0

While Finish_Write = 0
    
    
    If Worksheets("Maps").Cells(Row_Index, Column_Index) = "" And Map_Count < 4 Then
        Set Avaliable_Cell = Worksheets("Maps").Cells(Row_Index, Column_Index)
        Finish_Write = 1
    Else
        Map_Count = Map_Count + 1
        Finish_Write = 0
            If Map_Count = 3 Then
                Row_Index = Row_Index + Torque_num + 2
                Column_Index = 1
                Map_Count = 0
            Else
                Column_Index = Column_Index + speed_num + 1
            End If
        
        
    End If

Wend

'Create the map table based on the torques and speeds
Data_Start = Avaliable_Cell.Offset(2, 1).Address
Range(Avaliable_Cell.Offset(0, 0), Avaliable_Cell.Offset(0, speed_num)).Merge
Avaliable_Cell.Value = Map_Caption
Avaliable_Cell.HorizontalAlignment = Excel.xlCenter
    For i = 1 To speed_num
        Worksheets("Maps").Cells(Avaliable_Cell.Row + 1, Avaliable_Cell.Column).Offset(0, i).Value = CStr(i * Speed_Step) & " rpm"
    Next
    For j = 1 To Torque_num
        Worksheets("Maps").Cells(Avaliable_Cell.Row + 1, Avaliable_Cell.Column).Offset(j, 0).Value = CStr(j * Torque_Step) & " Nm"
    Next
    
Create_Maps = Data_Start

End Function
'Function version 1.0, map the result data to sigle map in maps sheet
Function Data_Mapping(Column_Index As Integer, Start_Addr As String)
Max_Speed = Worksheets("Specifications").Cells(200, "D").Value
Max_Torque = Worksheets("Specifications").Cells(201, "D").Value
Speed_Step = Worksheets("Specifications").Cells(202, "D").Value
Torque_Step = Worksheets("Specifications").Cells(203, "D").Value

Torque_num = Max_Torque / Torque_Step
speed_num = Max_Speed / Speed_Step

Dim Source_Rng As Range
Dim Target_Rng As Range
Dim Next_Start As Integer

For i = 1 To Torque_num

    Set First_Cell = Worksheets("ResultData").Cells(1, Column_Index).Offset(Next_Start, 0)
    Set Source_Rng = Range(First_Cell, First_Cell.Offset(speed_num - 1, 0))
    Set Target_Rng = Worksheets("Maps").Range(Start_Addr).Offset(i - 1, 0)
    Source_Rng.Copy
    Target_Rng.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=True
    Next_Start = Next_Start + speed_num
   
Next

End Function
Function Maps_Coloring(Start_Addr As String)

Max_Speed = Worksheets("Specifications").Cells(200, "D").Value
Max_Torque = Worksheets("Specifications").Cells(201, "D").Value
Speed_Step = Worksheets("Specifications").Cells(202, "D").Value
Torque_Step = Worksheets("Specifications").Cells(203, "D").Value

Torque_num = Max_Torque / Torque_Step
speed_num = Max_Speed / Speed_Step
Set SP = Worksheets("Maps").Range(Start_Addr)
Set Target_Rng = Range(SP, SP.Offset(Torque_num - 1, speed_num - 1))

Max_Value = Application.WorksheetFunction.Max(Target_Rng)
Min_Value = Application.WorksheetFunction.Min(Target_Rng)
Interval_Value = Max_Value - Min_Value
Dim rng As Range

For Each rng In Target_Rng

Percentage_Value = (rng.Value2 - Min_Value) / Interval_Value * 100
'MsgBox Percentage_Value

If Percentage_Value >= 80 And Percentage_Value <= 100 Then
rng.Interior.Color = RGB(255, 0, 0)
ElseIf Percentage_Value >= 60 And Percentage_Value < 80 Then
rng.Interior.Color = RGB(255, 128, 0)
ElseIf Percentage_Value >= 40 And Percentage_Value < 60 Then
rng.Interior.Color = RGB(255, 255, 0)
ElseIf Percentage_Value >= 20 And Percentage_Value < 40 Then
rng.Interior.Color = RGB(128, 255, 0)
ElseIf Percentage_Value < 20 Then
rng.Interior.Color = RGB(0, 255, 0)
End If

Next

End Function
Function Surface_Plot_Filling(Start_Addr As String, Chart_Pos As String, Title_Name As String)

Dim cht As ChartObject

Max_Speed = Worksheets("Specifications").Cells(200, "D").Value
Max_Torque = Worksheets("Specifications").Cells(201, "D").Value
Speed_Step = Worksheets("Specifications").Cells(202, "D").Value
Torque_Step = Worksheets("Specifications").Cells(203, "D").Value

Torque_num = Max_Torque / Torque_Step
speed_num = Max_Speed / Speed_Step
Set SP = Worksheets("Maps").Range(Start_Addr)
Set Target_Rng = Range(SP.Offset(-1, -1), SP.Offset(Torque_num - 1, speed_num - 1))
Set Location_Start = Worksheets("PassReport").Range(Chart_Pos)
Set Chart_Location = Worksheets("PassReport").Range(Location_Start, Location_Start.Offset(17, 7))
Set Chart_obj = Worksheets("PassReport").ChartObjects.Add(0, 0, 0, 0)

With Chart_obj
    .Chart.HasTitle = True
    .Chart.ChartTitle.Text = Title_Name
    .Chart.SetSourceData Source:=Target_Rng
    .Chart.ChartType = xlSurface
    .Top = Chart_Location.Top
    .Left = Chart_Location.Left
    .Width = Chart_Location.Width
    .Height = Chart_Location.Height
End With

End Function
Function Perfomance_Filling(Paste_Addr As String, Copy_Addr As String)
Dim Source_Rng As Range
Dim Targe_Rng As Range
Dim i As Integer
Max_Speed = Worksheets("Specifications").Cells(200, "D").Value
Max_Torque = Worksheets("Specifications").Cells(201, "D").Value
Speed_Step = Worksheets("Specifications").Cells(202, "D").Value
Torque_Step = Worksheets("Specifications").Cells(203, "D").Value

Torque_num = Max_Torque / Torque_Step
speed_num = Max_Speed / Speed_Step

Set SP_RD = Worksheets("ResultData").Range(Copy_Addr)
Set Source_Rng = Range(SP_RD, SP_RD.Offset(Torque_num * speed_num - 1, 0))
Set SP_PR = Worksheets("PassReport").Range(Paste_Addr)
Set Target_Rng = Worksheets("PassReport").Range(SP_PR, SP_PR.Offset(Torque_num * speed_num - 1, 0))
i = 0
For Each rng In Source_Rng
SP_PR.Offset(i, 0).Value2 = rng.Value2
i = i + 1
Next

'Source_Rng.Copy
'Target_Rng.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False


End Function
'This function is used for calculating average value for RTD data version1.0
Function Read_Data_From_RTD(ParamArray Items() As Variant) As String()

Dim Average_Value() As String
ReDim Preserve Average_Value(CInt(Items(0)(0)))
For i = 1 To CInt(Items(0)(0))

ColumnIndex = Range(Items(0)(i)).Column
LastRow = Sheets("RTD").Columns(ColumnIndex).Find("*", , xlValues, , , 2).Row
Average_Value(i) = "=SUM(" & "RTD!" & Range(Cells(7, ColumnIndex), Cells(LastRow, ColumnIndex)).Address(True, True) & ")" _
                                                & "/RTD!" & Range(Cells(5, ColumnIndex), Cells(5, ColumnIndex)).Address(True, True)
Next
Average_Value(0) = Items(0)(0)
Read_Data_From_RTD = Average_Value
                                                
End Function
'This function is used for wrting the average value to Result Data sheet version1.0
Function Write_Avg_to_ResultData(ParamArray Items() As Variant)
Dim Finish_Write As Integer
Dim Column_Count As Integer

Column_Index = 1
Finish_Write = 0

While Finish_Write = 0

If Worksheets("ResultData").Cells(1, Column_Index) = "" Then

For i = 1 To CInt(Items(0)(0))
Worksheets("ResultData").Cells(i, Column_Index) = Items(0)(i)
Next
MsgBox "Data has been written"
Finish_Write = 1
Else
Finish_Write = 0
Column_Index = Column_Index + 1
End If

Wend
End Function

Public Sub Report_Gen()

Dim Copper_Loss_S_Addr As String
Dim Rotational_Loss_S_Addr As String
Dim Total_Loss_S_Addr As String
Dim Inverter_Loss_S_Addr As String
Dim Motor_Loss_S_Addr As String
Dim Sys_Loss_S_Addr As String
Dim Cal_Sys_Eff_S_Addr As String
Dim Cal_Motor_Eff_S_Addr As String
Dim Cal_Inv_Eff_S_Addr As String


'losses Data
Copper_Loss = Find_Items_From_Sheets("Results", "Copper Loss", "ByColumns")
Rotational_Loss = Find_Items_From_Sheets("Results", "Rotational Loss", "ByColumns")
Inverter_Loss = Find_Items_From_Sheets("Results", "Inverter Loss", "ByColumns")
Motor_Loss = Find_Items_From_Sheets("Results", "Motor Loss", "ByColumns")
System_Loss = Find_Items_From_Sheets("Results", "System Loss", "ByColumns")
Total_Loss = Find_Items_From_Sheets("Results", "Total Loss", "ByColumns")

'Efficiency Data
Calculated_System_Eff = Find_Items_From_Sheets("Results", "Calculated System Efficiency", "ByColumns")
Calculated_Motor_Eff = Find_Items_From_Sheets("Results", "Calculated Motor Efficiency", "ByColumns")
Calculated_Inverter_Eff = Find_Items_From_Sheets("Results", "Calculated Inverter Efficiency", "ByColumns")

'write to result in sequence
Write_From_Results_to_ResultData (Copper_Loss)
Write_From_Results_to_ResultData (Rotational_Loss)
Write_From_Results_to_ResultData (Total_Loss)
Write_From_Results_to_ResultData (Inverter_Loss)
Write_From_Results_to_ResultData (Motor_Loss)
Write_From_Results_to_ResultData (System_Loss)
Write_From_Results_to_ResultData (Calculated_System_Eff)
Write_From_Results_to_ResultData (Calculated_Motor_Eff)
Write_From_Results_to_ResultData (Calculated_Inverter_Eff)

MsgBox "Data has been written"

Copper_Loss_S_Addr = Create_Maps("Copper Loss")
Rotational_Loss_S_Addr = Create_Maps("Rotational Loss")
Total_Loss_S_Addr = Create_Maps("Total Loss")
Inverter_Loss_S_Addr = Create_Maps("Inverter Loss")
Motor_Loss_S_Addr = Create_Maps("Motor Loss")
Sys_Loss_S_Addr = Create_Maps("System Loss")
Cal_Sys_Eff_S_Addr = Create_Maps("Calculated System Effciency")
Cal_Motor_Eff_S_Addr = Create_Maps("Calculated Motor Efficiency")
Cal_Inv_Eff_S_Addr = Create_Maps("Calculated Inverter Efficiency")

Call Data_Mapping(1, Copper_Loss_S_Addr)
Call Data_Mapping(2, Rotational_Loss_S_Addr)
Call Data_Mapping(3, Total_Loss_S_Addr)
Call Data_Mapping(4, Inverter_Loss_S_Addr)
Call Data_Mapping(5, Motor_Loss_S_Addr)
Call Data_Mapping(6, Sys_Loss_S_Addr)
Call Data_Mapping(7, Cal_Sys_Eff_S_Addr)
Call Data_Mapping(8, Cal_Motor_Eff_S_Addr)
Call Data_Mapping(9, Cal_Inv_Eff_S_Addr)

Call Maps_Coloring(Copper_Loss_S_Addr)
Call Maps_Coloring(Rotational_Loss_S_Addr)
Call Maps_Coloring(Total_Loss_S_Addr)
Call Maps_Coloring(Inverter_Loss_S_Addr)
Call Maps_Coloring(Motor_Loss_S_Addr)
Call Maps_Coloring(Sys_Loss_S_Addr)
Call Maps_Coloring(Cal_Sys_Eff_S_Addr)
Call Maps_Coloring(Cal_Motor_Eff_S_Addr)
Call Maps_Coloring(Cal_Inv_Eff_S_Addr)

Call Surface_Plot_Filling(Copper_Loss_S_Addr, "a18", "Copper Loss")
Call Surface_Plot_Filling(Rotational_Loss_S_Addr, "i18", "Rotational Loss")
Call Surface_Plot_Filling(Total_Loss_S_Addr, "a35", "Total Loss")
Call Surface_Plot_Filling(Inverter_Loss_S_Addr, "i35", "Inverter Loss")
Call Surface_Plot_Filling(Motor_Loss_S_Addr, "a52", "Motor Loss")
Call Surface_Plot_Filling(Sys_Loss_S_Addr, "i52", "System Loss")
Call Surface_Plot_Filling(Cal_Sys_Eff_S_Addr, "a69", "Calculated System Efficiency")
Call Surface_Plot_Filling(Cal_Motor_Eff_S_Addr, "i69", "Calculated Motor Efficiency")
'Call Surface_Plot_Filling(Cal_Inv_Eff_S_Addr, "a18")
Resize_Torques_Speeds_Table

Torque = Find_Items_From_Sheets("RTD", "Torque", "ByRows")
Average_Torque = Read_Data_From_RTD(Torque)
Write_Avg_to_ResultData (Average_Torque)

DynoSpeed = Find_Items_From_Sheets("RTD", "DynoSpeed", "ByRows")
Average_Speed = Read_Data_From_RTD(DynoSpeed)
Write_Avg_to_ResultData (Average_Speed)


Input_Power = Find_Items_From_Sheets("Results", "Input Power", "ByColumns")
Output_Power = Find_Items_From_Sheets("Results", "Output Power", "ByColumns")
Write_From_Results_to_ResultData (Input_Power)
Write_From_Results_to_ResultData (Output_Power)

Call Perfomance_Filling("C16", "J1")
Call Perfomance_Filling("G16", "K1")
Call Perfomance_Filling("I16", "L1")
Call Perfomance_Filling("K16", "M1")
Call Perfomance_Filling("O16", "C1")
Call Perfomance_Filling("M16", "H1")


MsgBox "Report Generated"




End Sub
