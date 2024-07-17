VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMakeFlow 
   Caption         =   "UserForm1"
   ClientHeight    =   2670
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4845
   OleObjectBlob   =   "frmMakeFlow.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMakeFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' 定数定義========================================================================================
Private Const RATE_ROW = 2
Private Const FLOW_START_ADDRESS_COL = "H"
Private Const FLOW_START_ADDRESS_ROW = "5"
Private Const FLOW_START_ADDRESS = FLOW_START_ADDRESS_COL + FLOW_START_ADDRESS_ROW
Private Const FLOW_GUID_ADDRESS = FLOW_START_ADDRESS_COL + ":" + FLOW_START_ADDRESS_COL

'■変数定義========================================================================================
'連想配列
Private mFlowTpeList As New Scripting.Dictionary
Private mRowsSize As New Scripting.Dictionary
'ファイル読込用
Private mFlowList As Variant
Private mFlowListIndex As Integer





'再設定
Private Sub CommandButton10_Click()
    
    Dim shpObj As Shape
    For Each shpObj In ActiveSheet.Shapes
        Call SetShapeFormat(shpObj, True)
    Next
    
End Sub

Private Sub CommandButton6_Click()
    mFlowListIndex = TextBox3.Text - 1
    NextFlowList
End Sub




Private Sub NextFlowList()
    
    If CheckBox2.Value = False Then Exit Sub
    If UBound(mFlowList) > mFlowListIndex Then
        ComboBox1.Text = Split(mFlowList(mFlowListIndex), ",")(0)
        TextBox1.Text = Split(mFlowList(mFlowListIndex), ",")(1)
        TextBox3.Text = mFlowListIndex
        Label9.Caption = mFlowList(mFlowListIndex)
        mFlowListIndex = mFlowListIndex + 1
    Else
        'ComboBox1.Text = ""
        'TextBox1.Text = ""
    End If
End Sub









Private Sub MultiPage1_Change()

End Sub

'■■■■■■■■■■■■■■■■■■■■■■■■■■
'■ ページ：フロー作成
'■■■■■■■■■■■■■■■■■■■■■■■■■■

'====================================================
'= 起動時
'====================================================
Private Sub UserForm_Initialize()
    
    ' フォームの初期化
    OptionButton1.Value = True
    MultiPage1.Value = 0
    
    ' ページ１の初期化
    mFlowTpeList.Add "開始", msoShapeFlowchartTerminator
    mFlowTpeList.Add "処理", msoShapeFlowchartProcess
    mFlowTpeList.Add "関数", msoShapeFlowchartPredefinedProcess
    mFlowTpeList.Add "入出力", msoShapeFlowchartData
    mFlowTpeList.Add "判断", msoShapeFlowchartDecision
    mFlowTpeList.Add "判断（合流）", msoShapeFlowchartConnector
    mFlowTpeList.Add "準備", msoShapeFlowchartPreparation
    mFlowTpeList.Add "ループ", msoShapeSnip2SameRectangle
    mFlowTpeList.Add "内部記憶", msoShapeFlowchartInternalStorage
    mFlowTpeList.Add "結合端子", msoShapeFlowchartConnector
    mFlowTpeList.Add "結合端子(他ページ)", msoShapeFlowchartOffpageConnector
    mFlowTpeList.Add "データベース", msoShapeFlowchartMagneticDisk
    mFlowTpeList.Add "表示", msoShapeFlowchartDisplay
    mFlowTpeList.Add "手入力", msoShapeFlowchartManualInput
    mFlowTpeList.Add "終了", msoShapeFlowchartTerminator
    For Each ft In mFlowTpeList
        ComboBox6.AddItem (ft)
        ComboBox1.AddItem (ft)
    Next
    ComboBox1.ListIndex = 1
    
    
    '-------------------
    ' ページ２の初期化
    '-------------------
    'Dictionaryオブジェクトの初期化、要素の追加
    mRowsSize.Add "全角１行", 11
    mRowsSize.Add "半角２行（標準）", 18
    mRowsSize.Add "半角・全角２行", 20
    mRowsSize.Add "全角２行", 22
    
    'Dictionaryオブジェクトの要素の参照
    Dim i As Integer
    Dim Keys() As Variant
    Keys = mRowsSize.Keys
    For i = 0 To mRowsSize.Count - 1
        ComboBox2.AddItem Keys(i)
    Next i
    ComboBox2.ListIndex = 2
    
    '図の幅（セル数）
    ComboBox3.AddItem 6
    ComboBox3.AddItem 8
    ComboBox3.AddItem 10
    ComboBox3.ListIndex = 0
    
    'フォント
    ComboBox4.AddItem ("游ゴシック 本文")
    ComboBox4.AddItem ("ＭＳ Ｐゴシック")
    ComboBox4.AddItem ("ＭＳ ゴシック")
    ComboBox4.AddItem ("ＭＳ Ｐ明朝")
    ComboBox4.AddItem ("ＭＳ 明朝")
    ComboBox4.ListIndex = 0
    
    'フォント
    ComboBox5.AddItem (6)
    ComboBox5.AddItem (7)
    ComboBox5.AddItem (8)
    ComboBox5.AddItem (9)
    ComboBox5.AddItem (10)
    ComboBox5.AddItem (11)
    ComboBox5.AddItem (12)
    ComboBox5.ListIndex = 4
    
    
    
    '-------------------
    ' ページ：読込
    '-------------------
    ListView1.Visible = False
    
End Sub


'===================================================================
'= フロー図生成
'===================================================================
Private Sub CommandButton1_Click()
    MakeDraw
    NextFlowList
End Sub

Private Sub MakeDraw()
    
    Call MakeDrawExec
    Call UpdateActiveCellPos
    
End Sub


'図のサイズ１
Private Sub GetMakeFlowSize(ByRef top As Double, ByRef left As Double, ByRef width As Double, ByRef height As Double)
    top = ActiveCell.top
    left = ActiveCell.left
    width = ActiveCell.width * ComboBox3.Text
    height = ActiveCell.height * RATE_ROW
End Sub

'図のサイズ２
Private Sub GetMakeFlowMinSize(nRate As Integer, ByRef left As Double, ByRef width As Double)
    left = ActiveCell.left + (ActiveCell.width * ((ComboBox3.Text / nRate) - (nRate / 2)))
    width = ActiveCell.width * nRate
End Sub

'図のサイズ３
Private Sub GetMakeFlowMinSize2(nRate As Integer, ByRef left As Double, ByRef width As Double)
    left = ActiveCell.left + (ActiveCell.width * nRate)
    width = ActiveCell.width * (ComboBox3.Text - (nRate * 2))
End Sub


Private Sub MakeDrawExec()
    
    
    Dim top As Double
    Dim left As Double
    Dim width As Double
    Dim height As Double
    Dim shapeType As MsoAutoShapeType
    Call GetMakeFlowSize(top, left, width, height)
    
    shapeType = mFlowTpeList.Item(ComboBox1.Text)
    Select Case ComboBox1.Text
    Case "開始"
        Call DrawFlowNormal(shapeType, left, top, width, height)
    Case "表示"
        'Call GetMakeFlowMinSize2(1, left, width)
        Call DrawFlowNormal(shapeType, left, top, width, height)
    Case "手入力"
        Call DrawFlowNormal(shapeType, left, top, width, height)
    Case "準備"
        Call DrawFlowNormal(shapeType, left, top, width, height)
        
    Case "処理"
        Call DrawFlowNormal(shapeType, left, top, width, height)
    
    Case "入出力"
        Call DrawFlowNormal(shapeType, left, top, width, height)
    
    Case "内部記憶"
        Call DrawFlowNormal(shapeType, left, top, width, height)
    
    Case "関数"
        Call DrawFlowNormal(shapeType, left, top, width, height)
    
    Case "判断"
        Call GetMakeFlowMinSize(2, left, width)
        'left = ActiveCell.left + (ActiveCell.width * (combobox3.text / 2 - 1))
        'width = ActiveCell.width * 2
        Call DrawFlowCondition(shapeType, left, top, width, height)
        
    Case "ループ"
        Call DrawFlowLoop(shapeType, left, top, width, height, True)
        UpdateActiveCellPos
        Call GetMakeFlowSize(top, left, width, height)
        Call DrawFlowLoop(shapeType, left, top, width, height, False)

    Case "判断（合流）"
        Call GetMakeFlowMinSize(2, left, width)
        width = 2
        height = 2
        left = left + ActiveCell.width - (width / 2)
        Call DrawFlowJoinTerminal(shapeType, left, top, width, height)
        
    Case "結合端子"
        Call GetMakeFlowMinSize(2, left, width)
        Call DrawFlowNormal(shapeType, left, top, width, height)
    
    Case "結合端子(他ページ)"
        Call GetMakeFlowMinSize(2, left, width)
        Call DrawFlowNormal(shapeType, left, top, width, height)
        
    Case "データベース"
        Call GetMakeFlowMinSize(2, left, width)
        Call DrawFlowNormal(shapeType, left, top, width, height)
        
    Case "終了"
        Call DrawFlowEnd(shapeType, left, top, width, height)
        
    Case Else
    
    End Select

End Sub

' 共通
Private Function DrawFlowCommon(shapeType As MsoAutoShapeType, X As Double, Y As Double, width As Double, height As Double) As Shape

    Dim shpObj As Shape
    Set shpObj = ActiveSheet.Shapes.AddShape(shapeType, X, Y, width, height)
    Call SetShapeFormat(shpObj, False)
    Set DrawFlowCommon = shpObj

End Function

' 処理など
Private Sub DrawFlowNormal(shapeType As MsoAutoShapeType, X As Double, Y As Double, width As Double, height As Double)
    
    Dim shpObj As Shape
    Set shpObj = DrawFlowCommon(shapeType, X, Y, width, height)
    shpObj.TextFrame.Characters.Text = TextBox1.Text
    
End Sub

' 終了
Private Sub DrawFlowEnd(shapeType As MsoAutoShapeType, X As Double, Y As Double, width As Double, height As Double)
    
    Call DrawFlowCommon(shapeType, X, Y, width, height)
End Sub

' 判断
Private Sub DrawFlowCondition(shapeType As MsoAutoShapeType, X As Double, Y As Double, width As Double, height As Double)
    
    Dim shpObj As Shape
    Dim shpObjText As Shape
    Dim shpObjYes As Shape
    Dim shpObjNo As Shape
    
    Set shpObj = ActiveSheet.Shapes.AddShape(shapeType, X, Y, width, height)
    Call SetShapeFormat(shpObj, False)

    '条件はテキストボックスで表示させる                                                       '↓小さくするとと右方向へ
    Set shpObjText = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, X + (width / 1.4), Y - height / 4, width, height)
    shpObjText.Fill.Visible = msoFalse                                                                          '↑大きくすると上方向へ
    shpObjText.Line.Visible = msoFalse
    shpObjText.TextFrame2.TextRange.Text = TextBox1.Text
    shpObjText.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
    shpObjText.TextFrame2.WordWrap = msoFalse
    
    
    'Yes
    If MsgBox("Yes・Noをつけますか?", vbYesNo) = vbYes Then
    
        Set shpObjYes = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, X + (width / 1.2), Y + height / 1.8, 30, 20)
        shpObjYes.Fill.Visible = msoFalse
        shpObjYes.Line.Visible = msoFalse
        shpObjYes.TextFrame2.TextRange.Text = "No"
        shpObjYes.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        shpObjYes.TextFrame2.WordWrap = msoFalse
        
        'No
        Set shpObjNo = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, X, Y + height / 1, 30, 20)
        shpObjNo.Fill.Visible = msoFalse
        shpObjNo.Line.Visible = msoFalse
        shpObjNo.TextFrame2.TextRange.Text = "Yes"
        shpObjNo.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        shpObjNo.TextFrame2.WordWrap = msoFalse
    
    End If
    
    
End Sub

'ループ
Private Sub DrawFlowLoop(shapeType As MsoAutoShapeType, X As Double, Y As Double, width As Double, height As Double, startFlg As Boolean)

    Dim shpObj As Shape
    Dim shpsObj As Shapes
    Set shpObj = ActiveSheet.Shapes.AddShape(shapeType, X, Y, width, height)
    shpObj.TextFrame.Characters.Text = TextBox1.Text
    
    If startFlg = True Then
        shpObj.Adjustments.Item(1) = 0.3
        shpObj.Adjustments.Item(2) = 0
    Else
        shpObj.Adjustments.Item(1) = 0
        shpObj.Adjustments.Item(2) = 0.3
    End If
    Call SetShapeFormat(shpObj, False)
    

End Sub

'判断（合流）
Private Sub DrawFlowJoinTerminal(shapeType As MsoAutoShapeType, X As Double, Y As Double, width As Double, height As Double)
    
    Dim shpObj As Shape
    Set shpObj = ActiveSheet.Shapes.AddShape(shapeType, X, Y, width, height)
    Call SetShapeFormat(shpObj, False)
End Sub




Private Sub SetShapeFormat(shpObj As Shape, bReset)
    
    If bReset = True And shpObj.TextFrame2.HasText = msoFalse Then Exit Sub
    
    If shpObj.AutoShapeType <> msoTextOrientationHorizontal Then
        shpObj.ShapeStyle = msoShapeStylePreset1
    End If
    
    
    If InStr(TextBox1.Text, vbLf) > 1 Or shpObj.AutoShapeType = msoShapeFlowchartMagneticDisk Then
        shpObj.TextFrame2.WordWrap = msoFalse
    End If
    
    With shpObj.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .HorizontalAnchor = msoAnchorCenter
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    shpObj.TextFrame.HorizontalOverflow = xlOartHorizontalOverflowOverflow
    shpObj.TextFrame.VerticalOverflow = xlOartVerticalOverflowOverflow
    
    With shpObj.TextFrame2.TextRange.Font
        .NameComplexScript = ComboBox4.Text
        .NameFarEast = ComboBox4.Text
        .Name = ComboBox4.Text
    End With
    
    Dim bkLeft As Double
    bkLeft = shpObj.left
    shpObj.TextFrame2.TextRange.Font.Size = ComboBox5.Text
    shpObj.left = bkLeft
    
End Sub


Private Sub UpdateActiveCellPos()
    If OptionButton1.Value = True Then
        ActiveCell.offset(RATE_ROW + 1, 0).Activate
    Else
        ActiveCell.offset(0, ComboBox3.Text + 1).Activate
    End If

End Sub


'===================================================================
'= フロー線生成
'===================================================================
Private Sub CommandButton2_Click()
    
    Dim msg As String
    If OptionButton1.Value = True Then msg = "↓"
    If OptionButton2.Value = True Then msg = "→"
    If OptionButton3.Value = True Then msg = "←"
    If MsgBox("方向は「" + msg + "」です。正しいですか？", vbYesNo) = vbNo Then Exit Sub
    
    If VarType(Selection) = vbObject Then
        DrawLine
    Else
    
        If Selection.Count > 1 Then
            DrawLineFree
        End If
    End If
    
End Sub

Private Sub DrawLine()
    Dim shpObj As Shape
    Dim shpType As MsoAutoShapeType
    Dim shpBeginPos As Integer
    Dim shpEndPos As Integer
    Dim ArrowStyle As MsoArrowheadStyle
    ArrowStyle = msoArrowheadTriangle
    
    
    If OptionButton1.Value = True Then          '↓方向
        shpType = msoConnectorStraight
        shpBeginPos = 3
        shpEndPos = 1
    ElseIf OptionButton2.Value = True Then      '→方向
        shpType = msoConnectorElbow
        shpBeginPos = 4
        shpEndPos = 1
    Else
        shpType = msoConnectorElbow             '←方向
        shpBeginPos = 3
        shpEndPos = 7
        'ArrowStyle = msoArrowheadNone
    End If
    
    
    For i = 1 To Selection.Count - 1
        
        '線を結合
        If OptionButton1.Value = True Then
            '開始線の設定
            Select Case Selection.ShapeRange(i).AutoShapeType
            Case msoShapeFlowchartConnector
                shpBeginPos = 5
            Case msoShapeSnip2SameRectangle
                shpBeginPos = 2
            Case msoShapeFlowchartData
                shpBeginPos = 5
            Case msoShapeFlowchartMagneticDisk
                shpBeginPos = 4
            Case Else
                shpBeginPos = 3
            End Select
        
            '終了線の設定
            Select Case Selection.ShapeRange(i + 1).AutoShapeType
            Case msoShapeFlowchartConnector
                shpEndPos = 1
                If Selection.ShapeRange(i + 1).width < 10 Then
                    ArrowStyle = msoArrowheadNone
                Else
                    ArrowStyle = msoArrowheadTriangle
                End If
            Case msoShapeSnip2SameRectangle
                shpEndPos = 4
                ArrowStyle = msoArrowheadTriangle
            Case msoShapeFlowchartData
                shpEndPos = 2
            Case msoShapeFlowchartMagneticDisk
                shpEndPos = 2
            Case Else
                shpEndPos = 1
                ArrowStyle = msoArrowheadTriangle
            End Select
        
        ElseIf OptionButton3.Value = True Then
'            Select Case Selection.ShapeRange(i + 1).AutoShapeType
'            Case msoShapeFlowchartConnector
'                shpEndPos = 7
'            Case Else
'                shpEndPos = 1
'            End Select
        End If
        
        Dim shpObjBegin As Shape
        If OptionButton2.Value = True Then
            Set shpObjBegin = Selection.ShapeRange(1)
        Else
            Set shpObjBegin = Selection.ShapeRange(i)
        End If
        
        
        Dim shpObjEnd As Shape
        If OptionButton3.Value = True Then
            Set shpObjEnd = Selection.ShapeRange(Selection.Count)
        Else
            Set shpObjEnd = Selection.ShapeRange(i + 1)
        End If
        
        
        '線を描画
        Set shpObj = ActiveSheet.Shapes.AddConnector(shpType, ActiveCell.left, ActiveCell.top, ActiveCell.left, ActiveCell.top + ActiveCell.height)
        shpObj.Line.EndArrowheadStyle = ArrowStyle
        
        
        shpObj.ConnectorFormat.BeginConnect GetShape(shpObjBegin), shpBeginPos
        shpObj.ConnectorFormat.EndConnect GetShape(shpObjEnd), shpEndPos
    Next

End Sub

Private Sub DrawLineFree()
    Dim rngObj1 As Range
    Dim rngObj2 As Range
    Dim shpObj As Shape
    
    Dim vRange As Variant
    vRange = Split(Selection.Address, ",")
    For i = 0 To UBound(vRange) - 1
        Set rngObj1 = Range(vRange(i))
        Set rngObj2 = Range(vRange(i + 1))
        Set shpObj = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, rngObj1.left, rngObj1.top, rngObj2.left, rngObj2.top)
    Next
    
    shpObj.Line.EndArrowheadStyle = msoArrowheadTriangle
    
End Sub


'===================================================================
'= 関数生成
'===================================================================
Private Sub CommandButton4_Click()
    CreateFunc
End Sub

Private Sub CreateFunc()
    
    '事前処理
    Dim funcNameJ As String  ' 和名
    Dim funcNameS As String  ' ソース
    Dim shpObj As Shape
    
    '図形を選択していなかったら作業中のシートに作成する
    If ActiveSheet.Shapes.Count = 0 Then
        TextBox1.Text = InputBox("関数名を入力してください。", , "")
        ActiveSheet.Name = TextBox1.Text
        GoTo CreateFlow
    End If
    
    On Error GoTo AbortEnd
    Set shpObj = Selection.ShapeRange(1)
    If shpObj.AutoShapeType <> msoShapeFlowchartPredefinedProcess Then
        If MsgBox("図形が関数ではありませんが、処理を継続しますか？", vbYesNo) <> vbYes Then Exit Sub
    End If
    If shpObj.TextFrame.Characters.Text = "" Then Exit Sub
    Dim wsObjSrc As Worksheet
    Set wsObjSrc = ActiveSheet
    Dim rngLinkSrc As Range
    Set rngLinkSrc = Selection.ShapeRange.Item(1).TopLeftCell.offset(0, ComboBox3.Text)
    
    '関数名取得
    funcNameJ = Split(shpObj.TextFrame.Characters.Text, vbLf)(0)
    If InStr(shpObj.TextFrame.Characters.Text, vbLf) > 0 Then
        funcNameS = Split(shpObj.TextFrame.Characters.Text, vbLf)(1)
    End If
    TextBox1.Text = shpObj.TextFrame.Characters.Text
    
    '新たにシートを作成
    Dim wsObj As Worksheet
    For Each wsObj In Worksheets
        If wsObj.Name = funcNameJ Then
            If MsgBox("既に関数「" & funcNameJ & "」が存在します。" + vbCrLf + "リンクを設定しますか？", vbYesNo) = vbYes Then
                wsObjSrc.Hyperlinks.Add Anchor:=rngLinkSrc, Address:="", SubAddress:=funcNameJ & "!A1", TextToDisplay:="参照"
            End If
            Exit Sub
        End If
    Next
    Dim wsObjDst As Worksheet
    Set wsObjDst = Worksheets.Add(after:=Worksheets(Worksheets.Count))
    wsObjDst.Name = funcNameJ
    
    '元のシートにリンクを設定
    wsObjSrc.Hyperlinks.Add Anchor:=rngLinkSrc, Address:="", SubAddress:=funcNameJ & "!A1", TextToDisplay:="参照"
    '新しいシートにリンク設定
    ActiveSheet.Hyperlinks.Add Anchor:=ActiveCell, Address:="", SubAddress:=wsObjSrc.Name & "!" & rngLinkSrc.Address(False, False), TextToDisplay:=wsObjSrc.Name & "へ"
    
CreateFlow:
    
    '新しいシートの初期化
    SheetInit
    
    '新しいシートに開始のフローを書く
    Range(FLOW_START_ADDRESS).Activate
    OptionButton1.Value = True
    ComboBox1.Text = "開始"
    'TextBox1.Text = funcNameJ
    MakeDraw
    ComboBox1.Text = "処理"
    MultiPage1.Value = 0
    
    Range("B2").Value = "概要"
    Range("B3").Value = "入力"
    Range("B4").Value = "出力"
    
    Exit Sub
AbortEnd:
    If Err <> 438 Then
        Call MsgBox(Error$(Err))
    End If
    On Error GoTo 0
    
    Exit Sub
End Sub

' ガイド線
Private Sub CheckBox1_Click()
        
    Dim j As Integer
    j = 0
    Dim rngObj As Range
    Set rngObj = Range(FLOW_GUID_ADDRESS)
    For i = 0 To 9
        If CheckBox1.Value Then
            rngObj.Borders(xlEdgeLeft).LineStyle = xlContinuous
            rngObj.Borders(xlEdgeLeft).Weight = xlHairline
        Else
            rngObj.Borders(xlEdgeLeft).LineStyle = xlNone
        End If
        j = 1
        Set rngObj = rngObj.offset(0, ComboBox3.Text + 1)
    Next
    
End Sub



'■■■■■■■■■■■■■■■■■■■■■■■■■■
'■ ページ：読込
'■■■■■■■■■■■■■■■■■■■■■■■■■■

'読込
Private Sub CommandButton5_Click()
    Dim buf As String
        
    If TextBox2.Text = "" Then
        If MsgBox("サンプルファイルを作成しますか？", vbYesNo) = vbYes Then
            Dim fn As String
            fn = ActiveWorkbook.Path + "\FlowSample.txt"
            
            
            With CreateObject("Scripting.FileSystemObject")
                If .FileExists(fn) Then
                    If MsgBox("既にファイルが存在します。上書きしますか？", vbYesNo) = vbNo Then Exit Sub
                End If
            End With
            
            
            Open fn For Append As #1
            Print #1, "処理,1111"
            Print #1, "関数,2222"
            Print #1, "入出力,3333"
            Print #1, "判断,4444"
            Print #1, "判断（合流）,aaaaaaa"
            Print #1, "準備,5555"
            Print #1, "ループ,66666"
            Print #1, "内部記憶,77777"
            Print #1, "結合端子,8888"
            Print #1, "結合端子(他ページ),9999"
            Print #1, "データベース,10"
            Print #1, "表示,11"
            Print #1, "終了,12"
            Close #1
            
            
            Dim WSH
            Set WSH = CreateObject("Wscript.Shell")
            WSH.Run fn, 3
            Set WSH = Nothing
            
            'Call MsgBox("FlowSample.txtを作成しました。")
        End If
    
        Exit Sub
        
    End If
    
    
    '読込インデックスを初期化
    mFlowListIndex = 0
    
    'パスのダブルコーテーションを削除
    If Mid(TextBox2.Text, 1, 1) = """" Then
        TextBox2.Text = Mid(TextBox2.Text, 2, Len(TextBox2.Text) - 2)
    End If
    
    
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(TextBox2.Text).OpenAsTextStream
            buf = .ReadAll
            mFlowList = Split(buf, vbCrLf)
            CheckBox2.Value = True
            .Close
        End With
    End With
    
    Call NextFlowList
    
End Sub

'ドラッグ
Private Sub TextBox2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    ListView1.Visible = True
End Sub
'ドロップ
Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LineText As String
    Dim AllText As String
    Dim i As Integer
    
    If Data.Files.Count > 0 Then
        TextBox2.Text = Data.Files(1)
        ListView1.Visible = False
    End If
End Sub




'■■■■■■■■■■■■■■■■■■■■■■■■■■
'■ ページ：その他
'■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Sub CommandButton7_Click()
    Call MakeTableOfContents
End Sub

'選択範囲削除
Private Sub CommandButton9_Click()
    
    If MsgBox("選択範囲の図形を削除します。よりしいですか?", vbYesNo) <> vbYes Then Exit Sub
    
    '選択範囲ループ
    Dim rngObj As Range
    For Each rngObj In Selection
        
        'シート上のshapeをループ
        Dim shpObj As Shape
        For Each shpObj In ActiveSheet.Shapes
            
            'セル範囲が重なれば削除
            If Not Intersect(rngObj, shpObj.TopLeftCell) Is Nothing Then
                shpObj.Delete
            End If
        Next
    Next
End Sub





Private Function GetShape(shpObj As Shape) As Shape
    Dim c As Shape
    Dim group As Shapes
    
    If shpObj.Type = msoGroup Then
        For Each c In shpObj.GroupItems
            If InStr(c.Name, "TextBox") = 0 Then
                Set GetShape = c
                Exit Function
            End If
        Next
    Else
        Set GetShape = shpObj
    End If
End Function




'■■■■■■■■■■■■■■■■■■■■■■■■■■
'■ シート初期化
'■■■■■■■■■■■■■■■■■■■■■■■■■■

Private Sub CommandButton3_Click()
    SheetInit
End Sub


'====================================================
'= シート初期化
'====================================================
Private Sub SheetInit()
    Columns.ColumnWidth = 3
'    Rows.RowHeight = 10.8  '18ピクセル
'    Rows.RowHeight = 18    '30ピクセル
'    Rows.RowHeight = 20.4  '34ピクセル
'    Rows.RowHeight = 22.2  '37ピクセル
'    Rows.RowHeight = ComboBox2.Text
    
    Dim i As Integer
    Dim Keys() As Variant
    Keys = mRowsSize.Keys
    Rows.RowHeight = mRowsSize.Item(Keys(ComboBox2.ListIndex))
    ActiveWindow.DisplayGridlines = False
End Sub


'■■■■■■■■■■■■■■■■■■■■■■■■■■
'■ その他
'■■■■■■■■■■■■■■■■■■■■■■■■■■
Private Sub MakeTableOfContents()
    
    Dim wsObj As Worksheet
    
    '目次シートがあれば消す
    For Each wsObj In Worksheets
        If wsObj.Name = "目次" Then
            If MsgBox("目次がすでに存在します。作り直しますか？", vbYesNo) <> vbYes Then Exit Sub
            '作り直すため目次を削除
            Worksheets("目次").Delete
            Exit For
        End If
    Next
    
    
    Set wsObj = Worksheets.Add(before:=Worksheets(1))
    wsObj.Name = "目次"
    
    '開始セル設定
    Dim rngStart As Range
    Set rngStart = Range("B3")
    
    rngStart.offset(0, 0) = "No"
    rngStart.offset(0, 1) = "関数"
    rngStart.offset(0, 2) = "概要"
    
    Dim i As Integer
    
    For i = 2 To Worksheets.Count
        rngStart.offset(i - 1, 0).Formula = "=row()- " & rngStart.Row
        wsObj.Hyperlinks.Add Anchor:=rngStart.offset(i - 1, 1), _
                             Address:="", _
                             SubAddress:=Worksheets(i).Name & "!A1", TextToDisplay:=Worksheets(i).Name
        
        rngStart.offset(i - 1, 2) = Worksheets(i).Range("D2").Value
    Next
    
    Range("A:A").ColumnWidth = 3.2
    Range("B:B").ColumnWidth = 3.2
    Range("C:C").ColumnWidth = 30
    Range("D:D").ColumnWidth = 50
    
    Range(rngStart.Address & ":" & rngStart.offset(0, 2).Address).Interior.ColorIndex = 35
    
    Range(rngStart.Address & ":" & rngStart.offset(Worksheets.Count - 1, 2).Address).Borders.LineStyle = xlContinuous
    
    
End Sub

'図形の変更
Private Sub CommandButton8_Click()
    If ComboBox6.Text = "" Then Exit Sub
    'If VarType(Selection) <> vbObject Then Exit Sub
    Selection.ShapeRange.AutoShapeType = mFlowTpeList.Item(ComboBox6.Text)

End Sub
Private Sub ComboBox6_Change()
    If ComboBox6.Text = "" Then Exit Sub
    'If VarType(Selection) <> vbObject Then Exit Sub
    Selection.ShapeRange.AutoShapeType = mFlowTpeList.Item(ComboBox6.Text)
End Sub
