Attribute VB_Name = "mdl_RoadMap"
Option Explicit

'__Setting__
    Dim ProjectName        As String
    Dim ProjectStartDate  As Date
    Dim ProjectEndDate    As Date
    
    'シート名
    Const wsSetting         As String = "Setting"
    Const wsRoadMap     As String = "RoadMap"
    Const wsMember       As String = "Member"
    Const wsCalender      As String = "Calender"
    
    '行
    Const rowHead                 As Integer = 1
    Const rowTItleMonth        As Integer = 2
    Const rowTitleDay            As Integer = 3
    Const rowTitleWeekDay   As Integer = 4
    Const rowActivityStart      As Integer = 5

    '列
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

'日付列幅
    Const DateWidth = 3.5

'列幅
    Const BasicWidth = 12

'行幅
    Const Height = 25

'追加行数
    Const CntAddRows As Integer = 2

Sub RESET()
'イベント処理が中断した場合、画面表示とイベント処理を再処理する

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
'初期設定

Dim i As Integer

Application.ScreenUpdating = False
Application.EnableEvents = False

On Error Resume Next

'__init__
    Sheets(wsRoadMap).Select
    
    'カレンダーの幅設定
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
    
'# 期間
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
    Range("Activity") = "アクティビティ"
    Call SetWidth(2, colActA, colActB)
    Call SetWidth(25, colActC, colActC)

'# 予定
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colPlanStart), Cells(rowTItleMonth, colPlanDays))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colPlanStart), Cells(rowTItleMonth, colPlanDays)) = "予定"
    
    Call SetHeight(15, rowTItleMonth, rowTitleWeekDay)
    
'# 予定開始日
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanStart), Cells(rowTitleWeekDay, colPlanStart))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanStart), Cells(rowTitleWeekDay, colPlanStart)) = "開始"
    Call SetWidth(BasicWidth, colPlanStart, colPlanStart)
'# 予定終了日
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanEnd), Cells(rowTitleWeekDay, colPlanEnd))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanEnd), Cells(rowTitleWeekDay, colPlanEnd)) = "終了"
    Call SetWidth(BasicWidth, colPlanEnd, colPlanEnd)
'# 予定工数
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanDays), Cells(rowTitleWeekDay, colPlanDays))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colPlanDays), Cells(rowTitleWeekDay, colPlanDays)) = "予定工数"
    Call SetWidth(BasicWidth, colPlanDays, colPlanDays)
'# 実績
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colResultStart), Cells(rowTItleMonth, colResultDays))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colResultStart), Cells(rowTItleMonth, colResultDays)) = "実績"

'# 実績開始日
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultStart), Cells(rowTitleWeekDay, colResultStart))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultStart), Cells(rowTitleWeekDay, colResultStart)) = "開始"
    Call SetWidth(BasicWidth, colResultStart, colResultStart)
'# 実績終了日
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultEnd), Cells(rowTitleWeekDay, colResultEnd))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultEnd), Cells(rowTitleWeekDay, colResultEnd)) = "終了"
    Call SetWidth(BasicWidth, colResultEnd, colResultEnd)
'# 実績工数
    With Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultDays), Cells(rowTitleWeekDay, colResultDays))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTitleDay, colResultDays), Cells(rowTitleWeekDay, colResultDays)) = "実績工数"
    Call SetWidth(BasicWidth, colResultDays, colResultDays)
'# 担当部署
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colDepartment), Cells(rowTitleWeekDay, colDepartment))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colDepartment), Cells(rowTitleWeekDay, colDepartment)) = "担当部署"
    Call SetWidth(BasicWidth, colDepartment, colDepartment)
'# 担当者
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colMember), Cells(rowTitleWeekDay, colMember))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colMember), Cells(rowTitleWeekDay, colMember)) = "担当者"
    Call SetWidth(BasicWidth, colMember, colMember)
'# ステータス
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colStatus), Cells(rowTitleWeekDay, colStatus))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colStatus), Cells(rowTitleWeekDay, colStatus)) = "ステータス"
    Call SetWidth(BasicWidth, colStatus, colStatus)
'# 後続作業
    With Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colSubSequence), Cells(rowTitleWeekDay, colSubSequence))
        .Merge
    End With
    Sheets(wsRoadMap).Range(Cells(rowTItleMonth, colSubSequence), Cells(rowTitleWeekDay, colSubSequence)) = "後続"
    Call SetWidth(BasicWidth, colSubSequence, colSubSequence)
    
'カレンダー設定
    Call SetCalender(ProjectStartDate, ProjectEndDate)
    
'ウィンドウ枠の固定
    ActiveWindow.FreezePanes = False
    Sheets(wsRoadMap).Cells(rowActivityStart, colCalenderStart).Select
    
'タイトルセル色付け
    Call CellColor(Range(Sheets(wsRoadMap).Cells(rowTItleMonth, colActID), _
                            Sheets(wsRoadMap).Cells(rowTitleWeekDay, _
                                                                        colLast(wsRoadMap, rowTitleWeekDay))), 259, 230, 160, 0)     'SoftYellow

'ActIDの採番
    Sheets(wsRoadMap).Cells(rowActivityStart, colActID) = Range("ActIDCounter")
    Range("ActIDCounter").VerticalAlignment = xlBottom

'枠線
    Call SetBorderLine

Application.ScreenUpdating = True
Application.EnableEvents = True

MsgBox ("初期化しました。")

End Sub

Sub Auto_Open()
'ファイルを開いた際のイベント

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
'セルの値を変更したときのイベント

Dim i As Long
Dim SubSequenceFlg As Boolean

On Error Resume Next
    Application.EnableEvents = False
    Application.ScreenUpdating = False

'__init__
    Call GetInitSetting
    
'__main__
    '予実欄
    If colPlanStart <= Target.column And _
        Target.column <= colResultDays And _
        Target.row >= rowActivityStart Then
        
        If Cells(Target.row, colActID) <> Empty Then
        
            'バーのセット
            Call DelBar(Target.row, True)
            Call DelBar(Target.row, False)
            Call SetPlanBar(Target.row)
            Call SetResultBar(Target.row)
            Call SetStatus(Target)
            Call GetDevelopDays(Target, True)
            
        Else
            'ActIDがない場合は何もしない
            GoTo exitProc
        End If
        
        'プロジェクト期間内の日付か確認
        If Target.column <> colPlanDays And _
            Target.column <> colResultDays Then
            
            If ProjectStartDate <= Target And Target.Value <= ProjectEndDate Then
                If Cells(Target.row, colStatus) = "完了" Then
                    Call CellColor(Target, 160, 160, 160)   'Gray
                Else
                    Call ClearColor(Target)
                End If
            Else
                If Target = Empty Then
                    ClearColor (Target)
                Else
                    'プロジェクト期間外
                    Call CellColor(Target, 255, 180, 10)    'Orange
                End If
            End If
        Else
        End If
        
        'アクティビティ欄
        ElseIf colActA <= Target.column And _
                Target.column <= colActC And _
                Target.row >= rowActivityStart Then
                
                'ActIDの削除
                If Cells(Target.row, colActA) = Empty And _
                    Cells(Target.row, colActB) = Empty And _
                    Cells(Target.row, colActC) = Empty Then
                        
                        Sheets(wsRoadMap).Cells(Target.row, colActID).ClearContents

                        'ステータスのセット
                        Call SetStatus(Target)
                Else
                    '次のプロセスの実行
                End If
        'それ以外
        Else: GoTo exitProc
        End If
        
        'ActIDの設定
        If Cells(Target.row, colActID) = Empty Then
            Call SetActID
        Else
        End If
        
        '枠線の設定
        Application.ScreenUpdating = False
        Call SetBorderLine
        
        For i = rowActivityStart To rowLast(wsRoadMap, colActID)
            
            '後続処理の日付点検
            If CheckSubsequence(i) = True And Target.row = i Then
                SubSequenceFlg = False
            Else
                SubSequenceFlg = True
            End If
            
            'プロジェクトの遅延確認
            '予定開始日、終了日のいずれかが入力されていない場合
            If Cells(Target.row, colPlanStart) = Empty Or Cells(Target.row, colPlanEnd) = Empty Then
                Call ClearColor(Range(Cells(Target.row, colActID), Cells(Target.row, colSubSequence)))
            
            '遅延しておりステータスが完了ではない場合
            ElseIf Cells(Target.row, colPlanEnd) < Date And Cells(Target.row, colStatus) <> "完了" Then
                Call CellColor(Range(Cells(Target.row, colActID), Cells(Target.row, colStatus)), 250, 100, 100)     'Pink
            
            'ステータスが完了となっている場合
            ElseIf Cells(Target.row, colStatus) = "完了" Then
                Call CellColor(Range(Cells(Target.row, colActID), Cells(Target.row, colSubSequence)), 160, 160, 160)    'Gray
            
            Else    '後続処理との整合性確認
                '後続処理との整合性あり
                If SubSequenceFlg = False Then
                    Call ClearColor(Range(Cells(Target.row, colActID), Cells(Target.row, colSubSequence)))
                    
                '後続処理との整合性なし
                Else
                    '警告色のままにする
                End If
            End If
        Next
exitProc:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub BeforeSaveRoadMap()
'保存時のイベント

Dim DayGap As Integer
Dim i As Integer

Application.ScreenUpdating = False

'__init__
    Call GetInitSetting
    Sheets(wsRoadMap).Cells.FormatConditions.Delete
    DayGap = ProjectEndDate - ProjectStartDate
    
    '追加の枠設定
    Call SetBorderLine
    
    For i = 0 To DayGap
        '条件付き書式
        Call SetFormatConditions(Cells(rowHead, colCalenderStart + i), rowLast(wsRoadMap, colActID) + CntAddRows)
    Next
    
    Application.ScreenUpdating = True
End Sub

Sub SetUp()
'Excelシートにプログラムを展開しセットアップする

Dim ws As Worksheet
Dim arrSheetName As Variant
Dim i As Integer

'シートの追加
'    On Error GoTo errEnd
    If Sheets.Count > 2 Then GoTo errEnd
    Sheets(1).Name = "Setting"
    arrSheetName = Array("RoadMap", "Member")
    
    For i = LBound(arrSheetName) To UBound(arrSheetName)
        Set ws = Worksheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = arrSheetName(i)
        Set ws = Nothing
    Next

'Settingシートの設定
    Worksheets("Setting").Select
    Range("A1") = "Project名"
    Range("B1") = "Initial Project"
    Range("A2") = "期間"
    Range("B2") = Date
    Range("C2") = "~"
    Range("D2") = DateAdd("m", 3, Date)
    
'モジュールのインポート
MsgBox ("F5で処理を継続してください。")
Stop
 Call ImportModule

'RoadMapの設定
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
'初期設定を読込む

    ProjectName = Sheets(wsSetting).Range("B1").Value
    ProjectStartDate = Sheets(wsSetting).Range("B2").Value
    ProjectEndDate = Sheets(wsSetting).Range("D2").Value
    
End Function

Function DelNameDef()
'名前の定義を削除する

Dim nm As Name

    '名前の定義を一つずつ削除する
    On Error Resume Next
        For Each nm In ActiveWorkbook.Names
            nm.Delete
        Next
End Function

Function SetCalender(ProjectStartDate As Date, ProjectEndDate As Date)
'カレンダーをセットする

Dim DayGap As Integer
Dim i As Integer, init_i As Integer
Dim M As Integer, D As Integer
Dim W As String

'__init__
    Sheets(wsRoadMap).Cells.FormatConditions.Delete
    
    'プロジェクトの合計日数算出
    DayGap = ProjectEndDate - ProjectStartDate
    
    '開始行を指定する
    If Month(ProjectStartDate) = Sheets(wsRoadMap).Cells(rowTItleMonth, colCalenderStart) And _
        Day(ProjectStartDate) = Sheets(wsRoadMap).Cells(rowTitleDay, colCalenderStart) Then
            init_i = colLast(wsRoadMap, rowTItleMonth) - colCalenderStart
    Else
        init_i = 0
    End If
    
'__main__
    'カレンダーの日付入力を繰り返す
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
            
            '休日設定
            If W = "土" Or W = "日" Then
                .Cells(rowHead, colCalenderStart + i) = "休"
                .Cells(rowHead, colCalenderStart + i).HorizontalAlignment = xlCenter
                .Cells(rowHead, colCalenderStart + i).VerticalAlignment = xlBottom
            Else
                Call GetHolidays(ProjectStartDate, ProjectEndDate)
            End If
        End With
        
        '条件付き書式設定
        Call SetFormatConditions(Sheets(wsRoadMap).Cells(rowHead, colCalenderStart + i), rowTitleWeekDay + CntAddRows + 1)
    Next
End Function

Function GetHolidays(ProjectStartDate As Date, ProojectEndDate As Date)
'国民の休日等を取得しRoadMapに反映する

Const rowStart As Long = 1
Const colStart  As Long = 1
Dim DayGap     As Integer
Dim i As Long
Dim j As Integer
Dim TargetDay As Variant
Dim M As Integer, D As Integer

'__main__
    'プロジェクト期間中の休日を取得する
        DayGap = ProjectEndDate - ProjectStartDate
        
        'プロジェクト期間内の日付のみ検索対象とする
        For i = rowStart To rowLast(wsCalender, rowStart)
            TargetDay = Sheets(wsCalender).Cells(i, colStart)
            
            If ProjectStartDate <= TargetDay And _
                TargetDay <= ProjectEndDate <= ProjectEndDate Then
                    
                    M = Month(TargetDay)
                    D = Day(TargetDay)
                    
                    '該当の日付が休日としてカレンダーシートに登載されているか確認する
                    For j = 0 To DayGap
                        If Sheets(wsRoadMap).Cells(rowTItleMonth, colCalenderStart + j) = M And _
                            Sheets(wsRoadMap).Cells(rowTitleDay, colCalenderStart + j) = D Then
                                
                                '該当の日付が休日だった場合、"休"と表示する
                                With Sheets(wsRoadMap).Cells(rowHead, colCalenderStart + j)
                                    .Value = "休"
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
'条件付き書式を設定する

'__init__
    '参照セルの設定（条件セル）
    Dim strTarget As String
        strTarget = Mid(TargetRng.Address, 2, Len(TargetRng.Address) - 1)
        
    '書式反映範囲
    Dim SettingRng As Range
        Set SettingRng = Range(Sheets(wsRoadMap).Cells(rowTItleMonth, TargetRng.column), _
                                            Sheets(wsRoadMap).Cells(LastRow, TargetRng.column))
                                            
'__main__
    With SettingRng
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, Formula1:="=" & strTarget & "=""休"""
        .FormatConditions(1).SetFirstPriority
        
        With .FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = RGB(200, 200, 220)     'LightGray
            .TintAndShade = 0.2
        End With
    End With
End Function

Function SetBorderLine()
'枠線の設定

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
    '全体の枠線を設定
    Range(Sheets(wsRoadMap).Cells(rowTItleMonth, colActID), _
                Sheets(wsRoadMap).Cells(LastRow, LastCol)) _
            .Borders.LineStyle = xlContinuous
                
    'アクティビティの枠線を消す
    With Range(Sheets(wsRoadMap).Cells(rowActivityStart, colActB), _
                        Sheets(wsRoadMap).Cells(LastRow, colActB))
            .Borders(xlEdgeLeft).LineStyle = xlNone
            .Borders(xlEdgeRight).LineStyle = xlNone
    End With
                        
    '日付線の設定
    With Range(Sheets(wsRoadMap).Cells(rowTItleMonth, colToday), Sheets(wsRoadMap).Cells(LastRow, colToday))
        With .Borders(xlEdgeRight)
                .LineStyle = xlDash
                .Color = RGB(255, 0, 0) 'Red
                .Weight = xlMedium
        End With
    End With
    
    DayGap = ProjectEndDate - ProjectStartDate
    
    For i = 0 To DayGap
        '条件付き書式
        Call SetFormatConditions(Cells(rowHead, colCalenderStart + i), LastRow)
    Next
End Function

Function SetPlanBar(row As Long)
'予定のバーを描出する

Dim StartDate As Date, EndDate As Date

'__init__
    StartDate = Sheets(wsRoadMap).Cells(row, colPlanStart)
    EndDate = Sheets(wsRoadMap).Cells(row, colPlanEnd)
    
'__main__
    '日付の入力がある場合
    If StartDate <> Empty And EndDate <> Empty Then
        Call SetBar(row, GetColDate(StartDate), GetColDate(EndDate), True)
    
    '日付の入力がない場合
    Else
        Call DelBar(row, True)      'バーを消す
    End If
End Function

Function SetResultBar(row As Long)
'実績のバーを描出する

Dim StartDate As Date, EndDate As Date

'__init__
    StartDate = Sheets(wsRoadMap).Cells(row, colResultStart)
    EndDate = Sheets(wsRoadMap).Cells(row, colResultEnd)
    
'__main__
    '開始日付、終了日付の入力がある場合
    If StartDate <> Empty And EndDate <> Empty Then
        Call SetBar(row, GetColDate(StartDate), GetColDate(EndDate), False)
    
    '開始日付のみ入力されている場合
    ElseIf StartDate <> Empty Then
        Call SetBar(row, GetColDate(StartDate), GetColDate(Date), False)
    
    '日付の入力がない場合
    Else
        Call DelBar(row, False)      'バーを消す
    End If
End Function

Function DelBar(row As Long, PlanFlg As Boolean)
'バーを削除する

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
'対象日付の列番号を返す

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
'バーを描出する

Dim bar As Shape

'On Error GoTo errExit
'__main__
    If PlanFlg = True Then  '予定
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
    Else    '実績
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
'すべての図形を削除する

Dim Shp As Shape

    For Each Shp In Sheets(wsRoadMap).Shapes
        Shp.Delete
    Next
End Function

Function CntActID() As Integer
'ActIDを取得する

    CntActID = WorksheetFunction.Max(Sheets(wsRoadMap).Columns(colActID)) + 1
    Range("ActIDCounter") = CntActID
End Function

Function SetActID()
'ActIDをセットする

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
'ステータスをセットする

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
    'プロジェクト期間外
    If PlanStartDate <> Empty Then
        If PlanStartDate < ProjectStartDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultStart).Select
            MsgBox ("予定開始日がプロジェクト開始前です。")
        ElseIf PlanStartDate > ProjectEndDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultStart).Select
            MsgBox ("予定開始日がプロジェクト終了後です。")
        Else
        End If
    Else
    End If
    
    If PlanEndDate <> Empty Then
        If PlanEndDate < ProjectStartDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("予定終了日がプロジェクト開始前です。")
        ElseIf PlanEndDate > ProjectEndDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("予定終了日がプロジェクト終了後です。")
        Else
        End If
    Else
    End If
    
    If ResultStartDate <> Empty Then
        If ResultStartDate < ProjectStartDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("実績開始日がプロジェクト開始前です。")
        ElseIf ResultStartDate > ProjectEndDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("実績開始日がプロジェクト終了後です。")
        Else
        End If
    Else
    End If
    
    If ResultEndDate <> Empty Then
        If ResultEndDate < ProjectStartDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("実績終了日がプロジェクト開始前です。")
        ElseIf ResultEndDate > ProjectEndDate Then
            Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
            MsgBox ("実績終了日がプロジェクト終了後です。")
        Else
        End If
    Else
    End If
    
    '先日付
    If ResultStartDate > Date Then
        Sheets(wsRoadMap).Cells(rowTG, colResultStart).Select
        MsgBox ("実績開始日が先日付です。")
    ElseIf ResultEndDate > Date Then
        Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
        MsgBox ("実績終了日が先日付です。")
    Else
    End If
    
    '前後関係
    '終了予定日付
    If PlanEndDate = Empty Then
    ElseIf PlanStartDate > PlanEndDate Then
        Sheets(wsRoadMap).Cells(rowTG, colPlanEnd).Select
        MsgBox ("終了予定日を見直してください。")
    Else
    End If
    
    '終了実績日付
    If ResultEndDate = Empty Then
    ElseIf ResultStartDate > ResultEndDate Then
        Sheets(wsRoadMap).Cells(rowTG, colResultEnd).Select
        MsgBox ("終了実績日付を見直してください。")
    Else
    End If
    
'__main__
    '予定の入力確認
    If PlanStartDate <> Empty And PlanEndDate <> Empty Then
        
        '実績の入力確認
        If ResultStartDate <> Empty And ResultEndDate <> Empty Then
                Sheets(wsRoadMap).Cells(rowTG, colStatus) = "完了"
        
        ElseIf ResultStartDate <> Empty Then
            Sheets(wsRoadMap).Cells(rowTG, colStatus) = "仕掛"
        Else
            Sheets(wsRoadMap).Cells(rowTG, colStatus) = "未着手"
        End If
    Else
        Sheets(wsRoadMap).Cells(rowTG, colStatus).ClearContents
    End If
    
    If Sheets(wsRoadMap).Cells(rowTG, colStatus) = "完了" Then
        Call CellColor(Range(Sheets(wsRoadMap).Cells(rowTG, colActID), _
                                        Sheets(wsRoadMap).Cells(rowTG, colLast(wsRoadMap, rowTitleWeekDay))), _
                                        160, 160, 160)      'Gray
    Else
        Call ClearColor(Sheets(wsRoadMap).Rows(rowTG))
    End If
End Function


Function GetDevelopDays(Target As Range, PlanFlg As Boolean)
'営業日数算出

    '予定工数の更新   ->  予定終了日の更新
    If Target.column = colPlanDays Then
'        Application.EnableEvents = True
'        Cells(Target.row, colPlanEnd) = GetDevelopDays(Cells(Target.row, colPlanStart), Target.Value)
'        Application.EnableEvents = False
    
    ElseIf Target.column = colPlanStart Or Target.column = colPlanEnd Then
        '予定日付欄の更新   ->  予定工数の更新
        If Cells(Target.row, colPlanStart) <> Empty And Cells(Target.row, colPlanEnd) <> Empty Then
            Cells(Target.row, colPlanDays) = CntWorkingDays(Cells(Target.row, colPlanStart), Cells(Target.row, colPlanEnd))
        Else
            Cells(Target.row, colPlanDays).ClearContents
        End If
    Else
    End If
    
    '実績工数欄の更新   ->  実績終了日の更新
    If Target.column = colResultDays Then
'        Application.EnableEvents = True
'        Cells(Target.row, colResultEnd) = CntWorkingDays(Cells(Target.row, colResultStart), Target.Value)
'        Application.EnableEvents = False
    ElseIf Target.column = colResultStart Or Target.column = colResultEnd Then
        
        '実績日付欄の更新   ->  実績工数更新
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
'営業日数算出

Dim CntHolidays As Integer

    '休日の日数を算出
    CntHolidays = WorksheetFunction.CountA(Range(Cells(rowHead, GetColDate(StartDate)), Cells(rowHead, GetColDate(EndDate))))
    
    '休日の日数を差し引いた日数を返す
    CntWorkingDays = EndDate - StartDate + 1 - CntHolidays

End Function

Function GetWorkingEndDate(StartDate As Date, DevDays As Integer)
'工数から終了日を算出する

Dim CntHolidays As Integer
Dim EndDate As Date

    EndDate = StartDate
    If DevDays = 0 Then GoTo exitProc
    
    '終了日までに休日がある場合は休日の日数を加算する
    Do Until CntWorkingDays(StartDate, EndDate) = DevDays
        EndDate = EndDate + 1
    Loop
    
exitProc:
    GetWorkingEndDate = EndDate
End Function

Function CheckSubsequence(row As Long) As Boolean
'後続処理との日程重複確認

Dim CurrentPlanEndDate As Date
Dim CurrentStatus As String
Dim CurrentResultEndDate As Date
Dim CurrentEndDate As Date
Dim NextSubSequenceNo As Integer
Dim NextPlanStartDate As Date

    CurrentPlanEndDate = Sheets(wsRoadMap).Cells(row, colPlanEnd)
    CurrentResultEndDate = Sheets(wsRoadMap).Cells(row, colResultEnd)
    CurrentStatus = Sheets(wsRoadMap).Cells(row, colStatus)
    
    'ステータスに応じて比較する日付を変える
    Select Case CurrentStatus
    Case "完了"
        CurrentResultEndDate = CurrentResultEndDate
    Case "仕掛"
        CurrentResultEndDate = Date
    Case Else
        CurrentResultEndDate = CurrentPlanEndDate
    End Select
    
    NextSubSequenceNo = Sheets(wsRoadMap).Cells(row, colSubSequence)
    NextPlanStartDate = Sheets(wsRoadMap).Cells(RowActID(NextSubSequenceNo), colPlanStart)
    
    '行の日付取得
    If CurrentResultEndDate <> Empty And CurrentResultEndDate > CurrentPlanEndDate Then
        CurrentEndDate = CurrentResultEndDate
    Else
        CurrentEndDate = CurrentPlanEndDate
    End If
    
    '後続作業の日付と照合
    If NextSubSequenceNo = Empty Then
        '後続番号なし
        Exit Function
    ElseIf CurrentEndDate > NextPlanStartDate Then
        '前後関係相違
        Call CellColor(Cells(row, colSubSequence), 255, 255, 0) 'Yellow
        Call CellColor(Cells(RowActID(NextSubSequenceNo), colPlanStart), 255, 255, 0) 'Yellow
        CheckSubsequence = True
    Else
        '前後関係に問題なし
        If Cells(row, colStatus) = "完了" Then
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
'ActIDから行番号を取得する

    RowActID = Sheets(wsRoadMap).Columns(colActID).Find(What:=ActID, _
                                                                                                LookIn:=xlFormulas, _
                                                                                                SearchOrder:=xlByRows, _
                                                                                                Serchdirection:=xlNext).row
End Function

'--- Import From BasicLib ---
Function rowLast(sheetName As String, column As Long) As Long
'最終行を求める

'Arg
'sheetName     検索するシート名
'column            検索する列番号

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
'最終列を求める

'Arg
'sheetName     検索するシート名
'column            検索する列番号

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
'RGBスケールでセルの色を変える

'RGBパラメータ
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
'セルの色設定をクリアする
    With rngR.Interior
                    .Pattern = xlNone
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
    End With
End Function

Function SetWidth(Width As Integer, StartCol As Integer, LastCol As Integer)
'セルの幅を設定する

    Range(Columns(StartCol), Columns(LastCol)).ColumnWidth = Width
End Function

Function SetHeight(Height As Integer, StartRow As Integer, LastRow As Integer)
'セルの高さを設定する

    Range(Rows(StartRow), Rows(LastRow)).RowHeight = Height
End Function
