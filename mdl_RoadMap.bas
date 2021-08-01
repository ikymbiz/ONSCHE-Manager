Attribute VB_Name = "mdl_RoadMap"
Option Explicit

'__Setting__
    Dim ProjectName        As String
    Dim ProjectStartDate  As Date
    Dim ProjectEndDate    As Date
    
    '�V�[�g��
    Const wsSetting         As String = "Setting"
    Const wsRoadMap     As String = "RoadMap"
    Const wsMember       As String = "Member"
    Const wsCalender      As String = "Calender"
    
    '�s
    Const rowHead                 As Integer = 1
    Const rowTItleMonth        As Integer = 2
    Const rowTitleDay            As Integer = 3
    Const rowTitleWeekDay   As Integer = 4
    Const rowActivityStart      As Integer = 5

    '��
    Const colActID                  As Integer = 1
    Const colActA                   As Integer = 2
    Const colActB                   As Integer = 3
    Const colActC                   As Integer = 4
    
    Const colPlanStart             As Integer = 5
    Const colPlanEnd               As Integer = 6
    Const colPlanDays             As Integer = 7
    Const colResultStart           As Integer = 8
    Const colResultEnd            As Integer = 9
    Const colResultDays          As Integer = 10
    
    Const colDepartment         As Integer = 11
    Const colMember               As Integer = 12
    Const colStatus                  As Integer = 13
    Const colSubSequence       As Integer = 14
    
    Const colCalenderStart      As Integer = 15

'���t��
    Const DateWidth = 3.5

'��
    Const BasicWidth = 12

'�s��
    Const Height = 25

'�ǉ��s��
    Const CntAddRows As Integer = 2

Sub RESET()
'�C�x���g���������f�����ꍇ�A��ʕ\���ƃC�x���g�������ď�������

Application.ScreenUpdating = False
Application.EnableEvents = False

    Call GetInitSetting
    Call SetBorderLine
    
    Range("ProjectName") = ProjectName
    Range("ProjectStartDate") = ProjectStartDate
    Range("ProjectEndDate") = ProjectEndDate
    
Application.ScreenUpdating = True
Application.EnableEvents = True
    
End Sub

Sub SetInitFormat()
'�����ݒ�

Dim i As Integer

Application.ScreenUpdating = False
Application.EnableEvents = False

On Error Resume Next

'__init__
    Sheets(wsRoadMap).Select
    
    '�J�����_�[�̕��ݒ�
    Cells.ColumnWidth = DateWidth
    Cells.RowHeight = Height
    Call GetInitSetting
    Call DelNameDef
    Sheets(wsRoadMap).Cells.Delete
    Call DelAllShape
    
'__main__
'# RoadMap
    Sheets(wsRoadMap).Range(Cells(rowHead, colActID), Cells(rowHead, colActC)).Name = "Title"
    With Range("Title")
        .Merge
        .Font.Bold = True
        .Font.Size = 16
    End With
    Range("Title") = "RoadMap"
    Call SetHeight(20, rowHead, rowHead)
    
'# ProjectName
    Sheets(wsRoadMap).Range(Cells(rowHead, colPlanStart), Cells(rowHead, colResultEnd)).Name = "ProjectName"
    With Range("ProjectName")
        .Merge
        .Font.Bold = False
        .Font.Size = 12
        .ShrinkToFit = True
        .HorizontalAlignment = xlLeft
    End With
    Range("ProjectName") = ProjectName
    
'# ����
    Sheets(wsRoadMap).Range(Cells(rowHead, colResultDays), Cells(rowHead, colResultDays)).Name = "ProjectStartDate"
    Sheets(wsRoadMap).Range(Cells(rowHead, colMember), Cells(rowHead, colMember)).Name = "ProjectEndDate"
    
    Range("ProjectStartDate") = ProjectStartDate
    Range("ProjectStartDate").NumberFormatLocal = "yyyy/m/d"
    Range("ProjectEndDate") = ProjectEndDate
    Range("ProjectEndDate").NumberFormatLocal = "yyyy/m/d"
    
    Columns(colPlanStart).NumberFormatLocal = "yyyy/m/d"
    Columns(colPlanEnd).NumberFormatLocal = "yyyy/m/d"
    Columns(colResultStart).NumberFormatLocal = "yyyy/m/d"
    Columns(colResultEnd).NumberFormatLocal = "yyyy/m/d"
    
'# ActID
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colActID), Cells(rowTitleWeekDay, colActID)).Name = "ActID"
    With Range("ActID")
        .Merge
        .ShrinkToFit = True
    End With
    Range("ActID") = "ActID"
    Call SetWidth(6, colActID, colActID)

'# ActIDCounter
    Sheets(wsRoadMap).Cells(rowHead, colStatus).Name = "ActIDCounter"
    Range("ActIDCounter") = 1

'# Activity
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colActA), Cells(rowTitleWeekDay, colActC)).Name = "Activity"
    With Range("Activity")
        .Merge
    End With
    Range("Activity") = "�A�N�e�B�r�e�B"
    Call SetWidth(2, colActA, colActB)
    Call SetWidth(25, colActC, colActC)

'# �\��
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colPlanStart), Cells(rowTItleMonth, colPlanDays))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colPlanStart), Cells(rowTItleMonth, colPlanDays)) = "�\��"
    
    Call SetHeight(15, rowTItleMonth, rowTitleWeekDay)
    
'# �\��J�n��
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanStart), Cells(rowTitleWeekDay, colPlanStart))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanStart), Cells(rowTitleWeekDay, colPlanStart)) = "�J�n"
    Call SetWidth(BasicWidth, colPlanStart, colPlanStart)
'# �\��I����
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanEnd), Cells(rowTitleWeekDay, colPlanEnd))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanEnd), Cells(rowTitleWeekDay, colPlanEnd)) = "�I��"
    Call SetWidth(BasicWidth, colPlanEnd, colPlanEnd)
'# �\��H��
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanDays), Cells(rowTitleWeekDay, colPlanDays))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanDays), Cells(rowTitleWeekDay, colPlanDays)) = "�\��H��"
    Call SetWidth(BasicWidth, colPlanDays, colPlanDays)
'# ����
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colResultStart), Cells(rowTItleMonth, colResultDays))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colResultStart), Cells(rowTItleMonth, colResultDays)) = "����"

'# ���ъJ�n��
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultStart), Cells(rowTitleWeekDay, colResultStart))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultStart), Cells(rowTitleWeekDay, colResultStart)) = "�J�n"
    Call SetWidth(BasicWidth, colResultStart, colResultStart)
'# ���яI����
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultEnd), Cells(rowTitleWeekDay, colResultEnd))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultEnd), Cells(rowTitleWeekDay, colResultEnd)) = "�I��"
    Call SetWidth(BasicWidth, colResultEnd, colResultEnd)
'# ���эH��
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultDays), Cells(rowTitleWeekDay, colResultDays))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultDays), Cells(rowTitleWeekDay, colResultDays)) = "���эH��"
    Call SetWidth(BasicWidth, colResultDays, colResultDays)
'# �S������
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colDepartment), Cells(rowTitleWeekDay, colDepartment))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colDepartment), Cells(rowTitleWeekDay, colDepartment)) = "�S������"
    Call SetWidth(BasicWidth, colDepartment, colDepartment)
'# �S����
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colMember), Cells(rowTitleWeekDay, colMember))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colMember), Cells(rowTitleWeekDay, colMember)) = "�S����"
    Call SetWidth(BasicWidth, colMember, colMember)
'# �X�e�[�^�X
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colStatus), Cells(rowTitleWeekDay, colStatus))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colStatus), Cells(rowTitleWeekDay, colStatus)) = "�X�e�[�^�X"
    Call SetWidth(BasicWidth, colStatus, colStatus)
'# �㑱���
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colSubSequence), Cells(rowTitleWeekDay, colSubSequence))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colSubSequence), Cells(rowTitleWeekDay, colSubSequence)) = "�㑱"
    Call SetWidth(BasicWidth, colSubSequence, colSubSequence)
    
'�J�����_�[�ݒ�
    Call SetCalender(ProjectStartDate, ProjectEndDate)
    
'�E�B���h�E�g�̌Œ�
    ActiveWindow.FreezePanes = False
    Sheets(wsRoadMap).Cells(rowActivityStart, colCalenderStart).Select
    
'�^�C�g���Z���F�t��
    Call CellColor(Range(Sheets(wsRoadMap).Cells(rowTItleMonth, colActID), _
                            Sheets(wsRoadMap).Cells(rowTitleWeekDay, _
                                                                        colLast(wsRoadMap, rowTitleWeekDay))), 259, 230, 160, 0)     'SoftYellow

'ActID�̍̔�
    Sheets(wsRoadMap).Cells(rowActivityStart, colActID) = Range("ActIDCounter")
    Range("ActIDCounter").VerticalAlignment = xlBottom

'�g��
    Call SetBorderLine

Application.ScreenUpdating = True
Application.EnableEvents = True

MsgBox ("���������܂����B")

End Sub

Sub Auto_Open()
'�t�@�C�����J�����ۂ̃C�x���g

Dim i As Long

'SetUp
    Call SetUp
            
On Error Resume Next
Application.ScreenUpdating = False
Application.EnableEvents = False
    
        Call GetInitSetting
        
        If Sheets(wsRoadMap).Range("ProjectEndDate") > ProjectEndDate Then
            Sheets(wsRoadMap).Range(Columns(GetColDate(ProjectEndDate) + 1), Columns(colLast(wsRoadMap, rowTItleMonth))).Clear
        ElseIf Range("ProjectEndDate") < ProjectEndDate Then
            Call SetCalender(ProjectStartDate, ProjectEndDate)
        Else
        End If
        
        Sheets(wsRoadMap).Select
        Application.ScreenUpdating = False
        
            For i = rowActivityStart To rowLast(wsRoadMap, colActID)
                ChangeEvent (Cells(i, colPlanStart))
            Next
            
            Call RESET
            Application.ScreenUpdating = True
            Application.EnableEvents = True
End Sub

Sub ChangeEvent(Target As Range)
'�Z���̒l��ύX�����Ƃ��̃C�x���g

Dim i As Long
Dim SubSequenceFlg As Boolean

On Error Resume Next
    Application.EnableEvents = False
    Application.ScreenUpdating = False

'__init__
    Call GetInitSetting
    
'__main__
    '�\����
    If colPlanStart <= Target.column And _
        Target.column <= colResultDays And _
        Target.row >= rowActivityStart Then
        
        If Cells(Target.row, colActID) <> Empty Then
        
            '�o�[�̃Z�b�g
            Call DelBar(Target.row, True)
            Call DelBar(Target.row, False)
            Call SetPlanBar(Target.row)
            Call SetResultBar(Target.row)
            Call SetStatus(Target)
            Call GetDevelopDays(Target, True)
            
        Else
            'ActID���Ȃ��ꍇ�͉������Ȃ�
            GoTo exitProc
        End If
        
        '�v���W�F�N�g���ԓ��̓��t���m�F
        If Target.column <> colPlanDays And _
            Target.column <> colResultDays Then
            
            If ProjectStartDate <= Target And Target.Value <= ProjectEndDate Then
                If Cells(Target.row, colStatus) = "����" Then
                    Call CellColor(Target, 160, 160, 160)   'Gray
                Else
                    Call ClearColor(Target)
                End If
            Else
                If Target = Empty Then
                    ClearColor (Target)
                Else
                    '�v���W�F�N�g���ԊO
                    Call CellColor(Target, 255, 180, 10)    'Orange
                End If
            End If
        Else
        End If
        
        '�A�N�e�B�r�e�B��
        ElseIf colActA <= Target.column And _
                Target.column <= colActC And _
                Target.row >= rowActivityStart Then
                
                'ActID�̍폜
                If Cells(Target.row, colActA) = Empty And _
                    Cells(Target.row, colActB) = Empty And _
                    Cells(Target.row, colActC) = Empty Then
                        
                        Sheets(wsRoadMap).Cells(Target.row, colActID).ClearContents

                        '�X�e�[�^�X�̃Z�b�g
                        Call SetStatus(Target)
                Else
                    '���̃v���Z�X�̎��s
                End If
        '����ȊO
        Else: GoTo exitProc
        End If
        
        'ActID�̐ݒ�
        If Cells(Target.row, colActID) = Empty Then
            Call SetActID
        Else
        End If
        
        '�g���̐ݒ�
        Application.ScreenUpdating = False
        Call SetBorderLine
        
        For i = rowActivityStart To rowLast(wsRoadMap, colActID)
            
            '�㑱�����̓��t�_��
            If CheckSubsequence(i) = True And Target.row = i Then
                SubSequenceFlg = False
            Else
                SubSequenceFlg = True
            End If
            
            '�v���W�F�N�g�̒x���m�F
            '�\��J�n���A�I�����̂����ꂩ�����͂���Ă��Ȃ��ꍇ
            If Cells(Target.row, colPlanStart) = Empty Or Cells(Target.row, colPlanEnd) = Empty Then
                Call ClearColor(Range(Cells(Target.row, colActID), Cells(Target.row, colSubSequence)))
            
            '�x�����Ă���X�e�[�^�X�������ł͂Ȃ��ꍇ
            ElseIf Cells(Target.row, colPlanEnd) < Date And Cells(Target.row, colStatus) <> "����" Then
                Call CellColor(Range(Cells(Target.row, colActID), Cells(Target.row, colStatus)), 250, 100, 100)     'Pink
            
            '�X�e�[�^�X�������ƂȂ��Ă���ꍇ
            ElseIf Cells(Target.row, colStatus) = "����" Then
                Call CellColor(Range(Cells(Target.row, colActID), Cells(Target.row, colSubSequence)), 160, 160, 160)    'Gray
            
            Else    '�㑱�����Ƃ̐������m�F
                '�㑱�����Ƃ̐���������
                If SubSequenceFlg = False Then
                    Call ClearColor(Range(Cells(Target.row, colActID), Cells(Target.row, colSubSequence)))
                    
                '�㑱�����Ƃ̐������Ȃ�
                Else
                    '�x���F�̂܂܂ɂ���
                End If
            End If
        Next
exitProc:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub BeforeSaveRoadMap()
'�ۑ����̃C�x���g

Dim DayGap As Integer
Dim i As Integer

Application.ScreenUpdating = False

'__init__
    Call GetInitSetting
    Sheets(wsRoadMap).Cells.FormatConditions.Delete
    DayGap = ProjectEndDate - ProjectStartDate
    
    '�ǉ��̘g�ݒ�
    Call SetBorderLine
    
    For i = 0 To DayGap
        '�����t������
        Call SetFormatConditions(Cells(rowHead, colCalenderStart + i), rowLast(wsRoadMap, colActID) + CntAddRows)
    Next
    
    Application.ScreenUpdating = True
End Sub

Sub SetUp()
'Excel�V�[�g�Ƀv���O������W�J���Z�b�g�A�b�v����

Dim ws As Worksheet
Dim arrSheetName As Variant
Dim i As Integer

'�V�[�g�̒ǉ�
'    On Error GoTo errEnd
    If Sheets.Count > 2 Then GoTo errEnd
    Sheets(1).Name = "Setting"
    arrSheetName = Array("RoadMap", "Member")
    
    For i = LBound(arrSheetName) To UBound(arrSheetName)
        Set ws = Worksheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = arrSheetName(i)
        Set ws = Nothing
    Next

'Setting�V�[�g�̐ݒ�
    Worksheets("Setting").Select
    Range("A1") = "Project��"
    Range("B1") = "Initial Project"
    Range("A2") = "����"
    Range("B2") = Date
    Range("C2") = "~"
    Range("D2") = DateAdd("m", 3, Date)
    
'���W���[���̃C���|�[�g
MsgBox ("F5�ŏ������p�����Ă��������B")
Stop
 Call ImportModule

'RoadMap�̐ݒ�
    Call SetInitFormat
errEnd:

End Sub

Function ImportModule()
    With ThisWorkbook.VBProject.VBComponents(Sheets("RoadMap").CodeName).CodeModule
        .AddFromFile ThisWorkbook.Path & "\SheetRoadMap"
    End With
    With ThisWorkbook.VBProject.VBComponents("ThisWorkbook").CodeModule
        .AddFromFile ThisWorkbook.Path & "\ThisWorkbook"
    End With
End Function

Function GetInitSetting()
'�����ݒ��Ǎ���

    ProjectName = Sheets(wsSetting).Range("B1").Value
    ProjectStartDate = Sheets(wsSetting).Range("B2").Value
    ProjectEndDate = Sheets(wsSetting).Range("D2").Value
    
End Function

Function DelNameDef()
'���O�̒�`���폜����

Dim nm As Name

    '���O�̒�`������폜����
    On Error Resume Next
        For Each nm In ActiveWorkbook.Names
            nm.Delete
        Next
End Function

Function SetCalender(ProjectStartDate As Date, ProjectEndDate As Date)
'�J�����_�[���Z�b�g����

Dim DayGap As Integer
Dim i As Integer, init_i As Integer
Dim M As Integer, D As Integer
Dim W As String

'__init__
    Sheets(wsRoadMap).Cells.FormatConditions.Delete
    
    '�v���W�F�N�g�̍��v�����Z�o
    DayGap = ProjectEndDate - ProjectStartDate
    
    '�J�n�s���w�肷��
    If Month(ProjectStartDate) = Sheets(wsRoadMap).Cells(rowTItleMonth, colCalenderStart) And _
        Day(ProjectStartDate) = Sheets(wsRoadMap).Cells(rowTitleDay, colCalenderStart) Then
            init_i = colLast(wsRoadMap, rowTItleMonth) - colCalenderStart
    Else
        init_i = 0
    End If
    
'__main__
    '�J�����_�[�̓��t���͂��J��Ԃ�
    For i = init_i To DayGap
        M = Month(ProjectStartDate + i)
        D = Day(ProjectStartDate + i)
        W = WeekdayName(Weekday(ProjectStartDate + i, vbSunday), True)
        
        With Sheets(wsRoadMap)
            'Month
            .Cells(rowTItleMonth, colCalenderStart + i) = M
            .Cells(rowTItleMonth, colCalenderStart + i).HorizontalAlignment = xlCenter
            
            'Date
            .Cells(rowTitleDay, colCalenderStart + i) = D
            .Cells(rowTitleDay, colCalenderStart + i).HorizontalAlignment = xlCenter
            
            'Weekday
            .Cells(rowTitleWeekDay, colCalenderStart + i) = W
            .Cells(rowTitleWeekDay, colCalenderStart + i).HorizontalAlignment = xlCenter
            
            '�x���ݒ�
            If W = "�y" Or W = "��" Then
                .Cells(rowHead, colCalenderStart + i) = "�x"
                .Cells(rowHead, colCalenderStart + i).HorizontalAlignment = xlCenter
                .Cells(rowHead, colCalenderStart + i).VerticalAlignment = xlBottom
            Else
                Call GetHolidays(ProjectStartDate, ProjectEndDate)
            End If
        End With
        
        '�����t�������ݒ�
        Call SetFormatConditions(Sheets(wsRoadMap).Cells(rowHead, colCalenderStart + i), rowTitleWeekDay + CntAddRows + 1)
    Next
End Function

Function GetHolidays(ProjectStartDate As Date, ProojectEndDate As Date)
'�����̋x�������擾��RoadMap�ɔ��f����

Const rowStart As Long = 1
Const colStart  As Long = 1
Dim DayGap     As Integer
Dim i As Long
Dim j As Integer
Dim TargetDay As Variant
Dim M As Integer, D As Integer

'__main__
    '�v���W�F�N�g���Ԓ��̋x�����擾����
        DayGap = ProjectEndDate - ProjectStartDate
        
        '�v���W�F�N�g���ԓ��̓��t�̂݌����ΏۂƂ���
        For i = rowStart To rowLast(wsCalender, rowStart)
            TargetDay = Sheets(wsCalender).Cells(i, colStart)
            
            If ProjectStartDate <= TargetDay And _
                TargetDay <= ProjectEndDate <= ProjectEndDate Then
                    
                    M = Month(TargetDay)
                    D = Day(TargetDay)
                    
                    '�Y���̓��t���x���Ƃ��ăJ�����_�[�V�[�g�ɓo�ڂ���Ă��邩�m�F����
                    For j = 0 To DayGap
                        If Sheets(wsRoadMap).Cells(rowTItleMonth, colCalenderStart + j) = M And _
                            Sheets(wsRoadMap).Cells(rowTitleDay, colCalenderStart + j) = D Then
                                
                                '�Y���̓��t���x���������ꍇ�A"�x"�ƕ\������
                                With Sheets(wsRoadMap).Cells(rowHead, colCalenderStart + j)
                                    .Value = "�x"
                                    .HorizontalAlignment = xlCenter
                                    .VerticalAlignment = xlBottom
                                End With
                        Else
                        End If
                    Next
            Else
            End If
        Next
End Function

Function SetFormatConditions(TargetRng As Range, LastRow As Long)
'�����t��������ݒ肷��

'__init__
    '�Q�ƃZ���̐ݒ�i�����Z���j
    Dim strTarget As String
        strTarget = Mid(TargetRng.Address, 2, Len(TargetRng.Address) - 1)
        
    '�������f�͈�
    Dim SettingRng As Range
        Set SettingRng = Range(Sheets(wsRoadMap).Cells(rowTItleMonth, TargetRng.column), _
                                            Sheets(wsRoadMap).Cells(LastRow, TargetRng.column))
                                            
'__main__
    With SettingRng
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, Formula1:="=" & strTarget & "=""�x"""
        .FormatConditions(1).SetFirstPriority
        
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = RGB(200, 200, 220)     'LightGray
            .TintAndShade = 0.2
        End With
    End With
End Function

Function SetBorderLine()
'�g���̐ݒ�

Dim i As Integer, DayGap As Integer
Dim LastRow As Long, LastCol As Integer, colToday As Integer

'__init__
    LastRow = rowLast(wsRoadMap, colActID) + CntAddRows
    LastCol = colLast(wsRoadMap, rowTitleWeekDay)
    colToday = GetColDate(Date)
    

    Range(Sheets(wsRoadMap).Cells(LastRow + 1, colActID), _
                Sheets(wsRoadMap).Cells(LastRow + 1, LastCol)) _
                .Borders.LineStyle = xlNone
                
    Sheets(wsRoadMap).Cells.FormatConditions.Delete

'__main__
    '�S�̘̂g����ݒ�
    Range(Sheets(wsRoadMap).Cells(rowTItleMonth, colActID), _
                Sheets(wsRoadMap).Cells(LastRow, LastCol)) _
            .Borders.LineStyle = xlContinuous
                
    '�A�N�e�B�r�e�B�̘g��������
    With Range(Sheets(wsRoadMap).Cells(rowActivityStart, colActB), _
                        Sheets(wsRoadMap).Cells(LastRow, colActB))
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
    End With
                        
    '���t���̐ݒ�
    With Range(Sheets(wsRoadMap).Cells(rowTItleMonth, colToday), Sheets(wsRoadMap).Cells(LastRow, colToday))
        With .Borders(xlEdgeRight)
                .LineStyle = xlDash
                .Color = RGB(255, 0, 0) 'Red
                .Weight = xlMedium
        End With
    End With
    
    DayGap = ProjectEndDate - ProjectStartDate
    
    For i = 0 To DayGap
        '�����t������
        Call SetFormatConditions(Cells(rowHead, colCalenderStart + i), LastRow)
    Next
End Function

Function SetPlanBar(row As Long)
'�\��̃o�[��`�o����

Dim StartDate As Date, EndDate As Date

'__init__
    StartDate = Sheets(wsRoadMap).Cells(row, colPlanStart)
    EndDate = Sheets(wsRoadMap).Cells(row, colPlanEnd)
    
'__main__
    '���t�̓��͂�����ꍇ
    If StartDate <> Empty And EndDate <> Empty Then
        Call SetBar(row, GetColDate(StartDate), GetColDate(EndDate), True)
    
    '���t�̓��͂��Ȃ��ꍇ
    Else
        Call DelBar(row, True)      '�o�[������
    End If
End Function

Function SetResultBar(row As Long)
'���т̃o�[��`�o����

Dim StartDate As Date, EndDate As Date

'__init__
    StartDate = Sheets(wsRoadMap).Cells(row, colResultStart)
    EndDate = Sheets(wsRoadMap).Cells(row, colResultEnd)
    
'__main__
    '�J�n���t�A�I�����t�̓��͂�����ꍇ
    If StartDate <> Empty And EndDate <> Empty Then
        Call SetBar(row, GetColDate(StartDate), GetColDate(EndDate), False)
    
    '�J�n���t�̂ݓ��͂���Ă���ꍇ
    ElseIf StartDate <> Empty Then
        Call SetBar(row, GetColDate(StartDate), GetColDate(Date), False)
    
    '���t�̓��͂��Ȃ��ꍇ
    Else
        Call DelBar(row, False)      '�o�[������
    End If
End Function

Function DelBar(row As Long, PlanFlg As Boolean)
'�o�[���폜����

Dim Shp As Shape
    For Each Shp In Sheets(wsRoadMap).Shapes
        If Shp.Name = "PLAN" & Sheets(wsRoadMap).Cells(row, colActID) Then
            If PlanFlg = True Then Shp.Delete
        ElseIf Shp.Name = "RESULT" & Sheets(wsRoadMap).Cells(row, colActID) Then
            If PlanFlg = False Then Shp.Delete
        End If
    Next
End Function

Function GetColDate(tgDate As Date) As Integer
'�Ώۓ��t�̗�ԍ���Ԃ�

Dim M As Integer, D As Integer
Dim i As Integer

    Call GetInitSetting
    
    M = Month(tgDate)
    D = Day(tgDate)
    
    If tgDate = Empty Then Exit Function
    
    GetColDate = colCalenderStart + tgDate - ProjectStartDate
    
    If tgDate < ProjectStartDate Then
        GetColDate = colCalenderStart
    ElseIf tgDate > ProjectEndDate Then
        GetColDate = colCalenderStart + ProjectEndDate - ProjectStartDate
    Else
    End If
End Function

Function SetBar(row As Long, colStart As Integer, colEnd As Integer, PlanFlg As Boolean)
'�o�[��`�o����

Dim bar As Shape

'On Error GoTo errExit
'__main__
    If PlanFlg = True Then  '�\��
        Set bar = ActiveSheet.Shapes.AddShape( _
                        Type:=1, _
                        Left:=Cells(row, colStart).Left, _
                        Top:=Cells(row, colStart).Top + 3, _
                        Width:=Cells(row, colStart).Width * (colEnd - colStart + 1), _
                        Height:=Cells(row, colStart).Height / 3)
                        
        With bar.Fill
            .ForeColor.RGB = RGB(80, 50, 220)   'Blue
            .Transparency = 0
        End With
        
        bar.Name = "PLAN" & Cells(row, colActID)
    Else    '����
        Set bar = ActiveSheet.Shapes.AddShape( _
                        Type:=1, _
                        Left:=Cells(row, colStart).Left, _
                        Top:=Cells(row + 1, colStart).Top - Cells(row, colStart).Height / 3 - 3, _
                        Width:=Cells(row, colStart).Width * (colEnd - colStart + 1), _
                        Height:=Cells(row, colStart).Height / 3)
                        
        With bar.Fill
            .ForeColor.RGB = RGB(220, 5, 5)   'Red
            .Transparency = 0
        End With
        
        bar.Name = "RESULT" & Cells(row, colActID)
    End If
    Exit Function
errExit:
Stop
End Function

Function DelAllShape()
'���ׂĂ̐}�`���폜����

Dim Shp As Shape

    For Each Shp In Sheets(wsRoadMap).Shapes
        Shp.Delete
    Next
End Function

Function CntActID() As Integer
'ActID���擾����

    CntActID = WorksheetFunction.Max(Sheets(wsRoadMap).Columns(colActID)) + 1
    Range("ActIDCounter") = CntActID
End Function

Function SetActID()
'ActID���Z�b�g����

Dim i As Long
Dim LastRow As Long

    LastRow = rowLast(wsRoadMap, colActA)
    If LastRow < rowLast(wsRoadMap, colActB) Then LastRow = rowLast(wsRoadMap, colActB)
    If LastRow < rowLast(wsRoadMap, colActC) Then LastRow = rowLast(wsRoadMap, colActC)
        
        For i = rowActivityStart To LastRow
            If Cells(i, colActID) = Empty Then
                Cells(i, colActID) = CntActID
            Else
                Call CntActID
            End If
        Next
End Function

Function SetStatus(Target As Range)
'�X�e�[�^�X���Z�b�g����

Dim rowTG As Long
Dim PlanStartDate As Date, PlanEndDate As Date
Dim ResultStartDate As Date, ResultEndDate As Date

'__init__
    rowTG = Target.row
    
    PlanStartDate = Sheets(wsRoadMap).Cells(rowTG, colPlanStart)
    PlanEndDate = Sheets(wsRoadMap).Cells(rowTG, colPlanEnd)
    ResultStartDate = Sheets(wsRoadMap).Cells(rowTG, colResultStart)
    ResultEndDate = Sheets(wsRoadMap).Cells(rowTG, colResultEnd)

'__check__
    '�v���W�F�N�g���ԊO
    If PlanStartDate <> Empty Then
        If PlanStartDate < ProjectStartDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultStart).Select
            MsgBox ("�\��J�n�����v���W�F�N�g�J�n�O�ł��B")
        ElseIf PlanStartDate > ProjectEndDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultStart).Select
            MsgBox ("�\��J�n�����v���W�F�N�g�I����ł��B")
        Else
        End If
    Else
    End If
    
    If PlanEndDate <> Empty Then
        If PlanEndDate < ProjectStartDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("�\��I�������v���W�F�N�g�J�n�O�ł��B")
        ElseIf PlanEndDate > ProjectEndDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("�\��I�������v���W�F�N�g�I����ł��B")
        Else
        End If
    Else
    End If
    
    If ResultStartDate <> Empty Then
        If ResultStartDate < ProjectStartDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("���ъJ�n�����v���W�F�N�g�J�n�O�ł��B")
        ElseIf ResultStartDate > ProjectEndDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("���ъJ�n�����v���W�F�N�g�I����ł��B")
        Else
        End If
    Else
    End If
    
    If ResultEndDate <> Empty Then
        If ResultEndDate < ProjectStartDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("���яI�������v���W�F�N�g�J�n�O�ł��B")
        ElseIf ResultEndDate > ProjectEndDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("���яI�������v���W�F�N�g�I����ł��B")
        Else
        End If
    Else
    End If
    
    '����t
    If ResultStartDate > Date Then
        Sheets(wsRoadMap).Cells(rowTG, colResultStart).Select
        MsgBox ("���ъJ�n��������t�ł��B")
    ElseIf ResultEndDate > Date Then
        Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
        MsgBox ("���яI����������t�ł��B")
    Else
    End If
    
    '�O��֌W
    '�I���\����t
    If PlanEndDate = Empty Then
    ElseIf PlanStartDate > PlanEndDate Then
        Sheets(wsRoadMap).Cells(rowTG, colPlanEnd).Select
        MsgBox ("�I���\������������Ă��������B")
    Else
    End If
    
    '�I�����ѓ��t
    If ResultEndDate = Empty Then
    ElseIf ResultStartDate > ResultEndDate Then
        Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
        MsgBox ("�I�����ѓ��t���������Ă��������B")
    Else
    End If
    
'__main__
    '�\��̓��͊m�F
    If PlanStartDate <> Empty And PlanEndDate <> Empty Then
        
        '���т̓��͊m�F
        If ResultStartDate <> Empty And ResultEndDate <> Empty Then
                Sheets(wsRoadMap).Cells(rowTG, colStatus) = "����"
        
        ElseIf ResultStartDate <> Empty Then
            Sheets(wsRoadMap).Cells(rowTG, colStatus) = "�d�|"
        Else
            Sheets(wsRoadMap).Cells(rowTG, colStatus) = "������"
        End If
    Else
        Sheets(wsRoadMap).Cells(rowTG, colStatus).ClearContents
    End If
    
    If Sheets(wsRoadMap).Cells(rowTG, colStatus) = "����" Then
        Call CellColor(Range(Sheets(wsRoadMap).Cells(rowTG, colActID), _
                                        Sheets(wsRoadMap).Cells(rowTG, colLast(wsRoadMap, rowTitleWeekDay))), _
                                        160, 160, 160)      'Gray
    Else
        Call ClearColor(Sheets(wsRoadMap).Rows(rowTG))
    End If
End Function


Function GetDevelopDays(Target As Range, PlanFlg As Boolean)
'�c�Ɠ����Z�o

    '�\��H���̍X�V   ->  �\��I�����̍X�V
    If Target.column = colPlanDays Then
'        Application.EnableEvents = True
'        Cells(Target.row, colPlanEnd) = GetDevelopDays(Cells(Target.row, colPlanStart), Target.Value)
'        Application.EnableEvents = False
    
    ElseIf Target.column = colPlanStart Or Target.column = colPlanEnd Then
        '�\����t���̍X�V   ->  �\��H���̍X�V
        If Cells(Target.row, colPlanStart) <> Empty And Cells(Target.row, colPlanEnd) <> Empty Then
            Cells(Target.row, colPlanDays) = CntWorkingDays(Cells(Target.row, colPlanStart), Cells(Target.row, colPlanEnd))
        Else
            Cells(Target.row, colPlanDays).ClearContents
        End If
    Else
    End If
    
    '���эH�����̍X�V   ->  ���яI�����̍X�V
    If Target.column = colResultDays Then
'        Application.EnableEvents = True
'        Cells(Target.row, colResultEnd) = CntWorkingDays(Cells(Target.row, colResultStart), Target.Value)
'        Application.EnableEvents = False
    ElseIf Target.column = colResultStart Or Target.column = colResultEnd Then
        
        '���ѓ��t���̍X�V   ->  ���эH���X�V
        If Cells(Target.row, colResultStart) <> Empty And Cells(Target.row, colResultEnd) <> Empty Then
            Cells(Target.row, colResultDays) = CntWorkingDays(Cells(Target.row, colResultStart), Cells(Target.row, colResultEnd))
        Else
            Cells(Target.row, colResultDays).ClearContents
        End If
    Else
    End If
    
    Application.EnableEvents = True
End Function

Function CntWorkingDays(StartDate As Date, EndDate As Date)
'�c�Ɠ����Z�o

Dim CntHolidays As Integer

    '�x���̓������Z�o
    CntHolidays = WorksheetFunction.CountA(Range(Cells(rowHead, GetColDate(StartDate)), Cells(rowHead, GetColDate(EndDate))))
    
    '�x���̓���������������������Ԃ�
    CntWorkingDays = EndDate - StartDate + 1 - CntHolidays

End Function

Function GetWorkingEndDate(StartDate As Date, DevDays As Integer)
'�H������I�������Z�o����

Dim CntHolidays As Integer
Dim EndDate As Date

    EndDate = StartDate
    If DevDays = 0 Then GoTo exitProc
    
    '�I�����܂łɋx��������ꍇ�͋x���̓��������Z����
    Do Until CntWorkingDays(StartDate, EndDate) = DevDays
        EndDate = EndDate + 1
    Loop
    
exitProc:
    GetWorkingEndDate = EndDate
End Function

Function CheckSubsequence(row As Long) As Boolean
'�㑱�����Ƃ̓����d���m�F

Dim CurrentPlanEndDate As Date
Dim CurrentStatus As String
Dim CurrentResultEndDate As Date
Dim CurrentEndDate As Date
Dim NextSubSequenceNo As Integer
Dim NextPlanStartDate As Date

    CurrentPlanEndDate = Sheets(wsRoadMap).Cells(row, colPlanEnd)
    CurrentResultEndDate = Sheets(wsRoadMap).Cells(row, colResultEnd)
    CurrentStatus = Sheets(wsRoadMap).Cells(row, colStatus)
    
    '�X�e�[�^�X�ɉ����Ĕ�r������t��ς���
    Select Case CurrentStatus
    Case "����"
        CurrentResultEndDate = CurrentResultEndDate
    Case "�d�|"
        CurrentResultEndDate = Date
    Case Else
        CurrentResultEndDate = CurrentPlanEndDate
    End Select
    
    NextSubSequenceNo = Sheets(wsRoadMap).Cells(row, colSubSequence)
    NextPlanStartDate = Sheets(wsRoadMap).Cells(RowActID(NextSubSequenceNo), colPlanStart)
    
    '�s�̓��t�擾
    If CurrentResultEndDate <> Empty And CurrentResultEndDate > CurrentPlanEndDate Then
        CurrentEndDate = CurrentResultEndDate
    Else
        CurrentEndDate = CurrentPlanEndDate
    End If
    
    '�㑱��Ƃ̓��t�Əƍ�
    If NextSubSequenceNo = Empty Then
        '�㑱�ԍ��Ȃ�
        Exit Function
    ElseIf CurrentEndDate > NextPlanStartDate Then
        '�O��֌W����
        Call CellColor(Cells(row, colSubSequence), 255, 255, 0) 'Yellow
        Call CellColor(Cells(RowActID(NextSubSequenceNo), colPlanStart), 255, 255, 0) 'Yellow
        CheckSubsequence = True
    Else
        '�O��֌W�ɖ��Ȃ�
        If Cells(row, colStatus) = "����" Then
            Call CellColor(Range(Cells(row, colActID), Cells(row, colSubSequence)), 160, 160, 160)  'Gray
            Call ClearColor(Cells(RowActID(NextSubSequenceNo), colPlanStart))
        ElseIf Cells(row, colPlanEnd) < Date Then
            If CurrentEndDate > NextPlanStartDate Then
                Call CellColor(Cells(row, colSubSequence), 255, 255, 0) 'Yellow
            Else
                Call CellColor(Range(Cells(row, colActID), Cells(row, colSubSequence)), 250, 100, 100)  'Pinkl
            End If
        Else
            Call ClearColor(Cells(row, colSubSequence))
            Call ClearColor(Cells(RowActID(NextSubSequenceNo), colPlanStart))
        End If
    End If
End Function

Function RowActID(ActID As Integer)
'ActID����s�ԍ����擾����

    RowActID = Sheets(wsRoadMap).Columns(colActID).Find(What:=ActID, _
                                                                                                LookIn:=xlFormulas, _
                                                                                                SearchOrder:=xlByRows, _
                                                                                                Serchdirection:=xlNext).row
End Function

'--- Import From BasicLib ---
Function rowLast(sheetName As String, column As Long) As Long
'�ŏI�s�����߂�

'Arg
'sheetName     ��������V�[�g��
'column            ���������ԍ�

On Error GoTo errExit
    rowLast = Sheets(sheetName).Columns(column).Find(What:="*", _
                                                                                            LookIn:=xlFormulas, _
                                                                                            SearchOrder:=xlByRows, _
                                                                                            SearchDirection:=xlPrevious).row
                                                                                
    Exit Function
errExit:
    rowLast = 0
End Function

Function colLast(sheetName As String, row As Long) As Integer
'�ŏI������߂�

'Arg
'sheetName     ��������V�[�g��
'column            ���������ԍ�

On Error GoTo errExit
    colLast = Sheets(sheetName).Rows(row).Find(What:="*", _
                                                                                LookIn:=xlFormulas, _
                                                                                SearchOrder:=xlByColumns, _
                                                                                SearchDirection:=xlPrevious).column
                                                                                
    Exit Function
errExit:
    colLast = 0

End Function

Function CellColor(rngR As Range, _
                                intColorR As Long, intColorG As Long, intColorB As Long, _
                                Optional dblTintAndShade As Double)
'RGB�X�P�[���ŃZ���̐F��ς���

'RGB�p�����[�^
'   https://ironodata.info/

    With rngR.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = RGB(intColorR, intColorG, intColorB)
                    .TintAndShade = dblTintAndShade
                    .PatternTintAndShade = 0
    End With
End Function
                                
Function ClearColor(rngR As Range)
'�Z���̐F�ݒ���N���A����
    With rngR.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
    End With
End Function

Function SetWidth(Width As Integer, StartCol As Integer, LastCol As Integer)
'�Z���̕���ݒ肷��

    Range(Columns(StartCol), Columns(LastCol)).ColumnWidth = Width
End Function

Function SetHeight(Height As Integer, StartRow As Integer, LastRow As Integer)
'�Z���̍�����ݒ肷��

    Range(Rows(StartRow), Rows(LastRow)).RowHeight = Height
End Function
