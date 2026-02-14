VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "セル範囲"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Excel")
' セル範囲オブジェクトのラッパークラス
' Excel VBAの Range オブジェクトの機能を日本語でラップ

Private pRangeObject As Range

' ========================================
' 初期化
' ========================================

Public Sub 初期化(rangeObj As Range)
    Set pRangeObject = rangeObj
End Sub

' ========================================
' プロパティ: アドレス関連
' ========================================

''' <summary>
''' セルのアドレスを取得
''' </summary>
Public Property Get セルアドレス() As String
    セルアドレス = pRangeObject.Address
End Property

''' <summary>
''' 列番号を取得
''' </summary>
Public Property Get 列番号() As Long
    列番号 = pRangeObject.Column
End Property

''' <summary>
''' 行番号を取得
''' </summary>
Public Property Get 行番号() As Long
    行番号 = pRangeObject.Row
End Property

''' <summary>
''' 最後の列番号を取得（複数列の場合）
''' </summary>
Public Property Get 最後の列番号() As Long
    最後の列番号 = pRangeObject.Column + pRangeObject.Columns.Count - 1
End Property

''' <summary>
''' 最後の行番号を取得（複数行の場合）
''' </summary>
Public Property Get 最後の行番号() As Long
    最後の行番号 = pRangeObject.Row + pRangeObject.Rows.Count - 1
End Property

' ========================================
' プロパティ: サイズ関連
' ========================================

''' <summary>
''' 幅をポイント単位で取得/設定
''' </summary>
Public Property Get 幅() As Double
    幅 = pRangeObject.Width
End Property

Public Property Let 幅(value As Double)
    pRangeObject.Width = value
End Property

''' <summary>
''' 高さをポイント単位で取得/設定
''' </summary>
Public Property Get 高さ() As Double
    高さ = pRangeObject.Height
End Property

Public Property Let 高さ(value As Double)
    pRangeObject.Height = value
End Property

''' <summary>
''' 列数を取得
''' </summary>
Public Property Get 列数() As Long
    列数 = pRangeObject.Columns.Count
End Property

''' <summary>
''' 行数を取得
''' </summary>
Public Property Get 行数() As Long
    行数 = pRangeObject.Rows.Count
End Property

''' <summary>
''' セル数を取得
''' </summary>
Public Property Get セル数() As Long
    セル数 = pRangeObject.Count
End Property

' ========================================
' プロパティ: 値と数式
' ========================================

''' <summary>
''' セルの値を取得/設定
''' </summary>
Public Property Get 値() As Variant
    値 = pRangeObject.value
End Property

Public Property Let 値(value As Variant)
    pRangeObject.value = value
End Property

''' <summary>
''' セルの数式を取得/設定
''' </summary>
Public Property Get 数式() As String
    数式 = pRangeObject.Formula
End Property

Public Property Let 数式(value As String)
    pRangeObject.Formula = value
End Property

''' <summary>
''' セルのR1C1参照スタイル数式を取得/設定
''' </summary>
Public Property Get R1C1数式() As String
    R1C1数式 = pRangeObject.FormulaR1C1
End Property

Public Property Let R1C1数式(value As String)
    pRangeObject.FormulaR1C1 = value
End Property

''' <summary>
''' セルの左上隅のテキストを取得
''' </summary>
Public Property Get テキスト() As String
    テキスト = pRangeObject.Text
End Property

''' <summary>
''' セルの表示値（計算結果）を取得
''' </summary>
Public Property Get 表示値() As Variant
    表示値 = pRangeObject.value
End Property

' ========================================
' プロパティ: 書式関連
' ========================================

''' <summary>
''' 数値形式を取得/設定
''' </summary>
Public Property Get 数値形式() As String
    数値形式 = pRangeObject.NumberFormat
End Property

Public Property Let 数値形式(value As String)
    pRangeObject.NumberFormat = value
End Property

''' <summary>
''' 背景色を取得/設定（RGB値）
''' </summary>
Public Property Get 背景色() As Long
    背景色 = pRangeObject.Interior.Color
End Property

Public Property Let 背景色(value As Long)
    pRangeObject.Interior.Color = value
End Property

''' <summary>
''' 文字色を取得/設定（RGB値）
''' </summary>
Public Property Get 文字色() As Long
    文字色 = pRangeObject.Font.Color
End Property

Public Property Let 文字色(value As Long)
    pRangeObject.Font.Color = value
End Property

''' <summary>
''' フォントサイズを取得/設定
''' </summary>
Public Property Get フォントサイズ() As Long
    フォントサイズ = pRangeObject.Font.Size
End Property

Public Property Let フォントサイズ(value As Long)
    pRangeObject.Font.Size = value
End Property

''' <summary>
''' フォント名を取得/設定
''' </summary>
Public Property Get フォント名() As String
    フォント名 = pRangeObject.Font.Name
End Property

Public Property Let フォント名(value As String)
    pRangeObject.Font.Name = value
End Property

''' <summary>
''' 太字かどうかを取得/設定
''' </summary>
Public Property Get 太字() As Boolean
    太字 = pRangeObject.Font.Bold
End Property

Public Property Let 太字(value As Boolean)
    pRangeObject.Font.Bold = value
End Property

''' <summary>
''' 斜体かどうかを取得/設定
''' </summary>
Public Property Get 斜体() As Boolean
    斜体 = pRangeObject.Font.Italic
End Property

Public Property Let 斜体(value As Boolean)
    pRangeObject.Font.Italic = value
End Property

''' <summary>
''' 下線かどうかを取得/設定
''' </summary>
Public Property Get 下線() As Boolean
    下線 = pRangeObject.Font.Underline <> xlNone
End Property

Public Property Let 下線(value As Boolean)
    pRangeObject.Font.Underline = IIf(value, xlUnderlineStyleSingle, xlNone)
End Property

''' <summary>
''' 水平配置を取得/設定（-4152: 左, -4108: 中央, -4160: 右）
''' </summary>
Public Property Get 水平配置() As Long
    水平配置 = pRangeObject.HorizontalAlignment
End Property

Public Property Let 水平配置(value As Long)
    pRangeObject.HorizontalAlignment = value
End Property

''' <summary>
''' 垂直配置を取得/設定（-4107: 上, -4108: 中央, -4161: 下）
''' </summary>
Public Property Get 垂直配置() As Long
    垂直配置 = pRangeObject.VerticalAlignment
End Property

Public Property Let 垂直配置(value As Long)
    pRangeObject.VerticalAlignment = value
End Property

''' <summary>
''' テキスト折返しを取得/設定
''' </summary>
Public Property Get テキスト折返し() As Boolean
    テキスト折返し = pRangeObject.WrapText
End Property

Public Property Let テキスト折返し(value As Boolean)
    pRangeObject.WrapText = value
End Property

''' <summary>
''' インデント有無を取得/設定
''' </summary>
Public Property Get インデント() As Boolean
    インデント = pRangeObject.AddIndent
End Property

Public Property Let インデント(ByVal インデント有り As Boolean)
    pRangeObject.AddIndent = インデント有り
End Property

' ========================================
' プロパティ: 保護/ロック関連
' ========================================

''' <summary>
''' セルがロックされているかを取得/設定
''' </summary>
Public Property Get ロック() As Boolean
    ロック = pRangeObject.Locked
End Property

Public Property Let ロック(value As Boolean)
    pRangeObject.Locked = value
End Property

''' <summary>
''' セルが非表示かどうかを取得/設定
''' </summary>
Public Property Get 非表示() As Boolean
    非表示 = pRangeObject.Hidden
End Property

Public Property Let 非表示(value As Boolean)
    pRangeObject.Hidden = value
End Property

' ========================================
' プロパティ: 特殊
' ========================================

''' <summary>
''' Range オブジェクト本体への直接アクセス
''' </summary>
Public Property Get 元のオブジェクト() As Object
    Set 元のオブジェクト = pRangeObject
End Property

' ========================================
' メソッド: 基本操作
' ========================================

''' <summary>
''' セルを選択
''' </summary>
Public Sub 選択()
    pRangeObject.Select
End Sub

''' <summary>
''' セルをアクティブにする
''' </summary>
Public Sub アクティブ化()
    pRangeObject.Activate
End Sub

''' <summary>
''' セルをコピー
''' </summary>
Public Sub コピー()
    pRangeObject.Copy
End Sub

''' <summary>
''' セルを切り取り
''' </summary>
Public Sub 切り取り()
    pRangeObject.Cut
End Sub

''' <summary>
''' すべてのコンテンツと書式をクリア
''' </summary>
Public Sub クリア()
    pRangeObject.Clear
End Sub

''' <summary>
''' コンテンツのみをクリア（書式は保持）
''' </summary>
Public Sub 内容をクリア()
    pRangeObject.ClearContents
End Sub

''' <summary>
''' 書式のみをクリア（コンテンツは保持）
''' </summary>
Public Sub 書式をクリア()
    pRangeObject.ClearFormats
End Sub

''' <summary>
''' セルを削除
''' </summary>
Public Sub 削除()
    pRangeObject.Delete
End Sub

' ========================================
' メソッド: セルの配置
' ========================================

''' <summary>
''' セルをオフセット位置から取得
''' オフセット行数、オフセット列数を指定
''' </summary>
Public Function オフセット(Optional 行offset As Long = 0, Optional 列offset As Long = 0) As Object
    Set オフセット = pRangeObject.Offset(行offset, 列offset)
End Function

''' <summary>
''' セルのサイズを変更した新しい Range を取得
''' </summary>
Public Function サイズ変更(サイズ行数 As Long, サイズ列数 As Long) As Object
    Set サイズ変更 = pRangeObject.Resize(サイズ行数, サイズ列数)
End Function

''' <summary>
''' 複数のセル範囲の共通部分を取得
''' </summary>
Public Function 共通部分(他の範囲 As Object) As Object
    On Error Resume Next
    Set 共通部分 = Application.Intersect(pRangeObject, 他の範囲)
    On Error GoTo 0
End Function

''' <summary>
''' 複数のセル範囲を統合した新しい Range を取得
''' </summary>
Public Function 統合(他の範囲 As Object) As Object
    On Error Resume Next
    Set 統合 = Application.Union(pRangeObject, 他の範囲)
    On Error GoTo 0
End Function

' ========================================
' メソッド: 検索・置換
' ========================================

''' <summary>
''' セル範囲内にある値を検索
''' 見つかった場合は Range オブジェクトを返す
''' </summary>
Public Function 検索(検索値 As String, Optional 検索タイプ As Long = xlValues) As Object
    On Error Resume Next
    Set 検索 = pRangeObject.Find(検索値)
    On Error GoTo 0
End Function

''' <summary>
''' 前回の検索から次の一致を検索
''' </summary>
Public Function 次を検索(検索値 As String) As Object
    On Error Resume Next
    Set 次を検索 = pRangeObject.FindNext(pRangeObject.Find(検索値))
    On Error GoTo 0
End Function

''' <summary>
''' セル範囲内の値を置換
''' </summary>
Public Sub 置換(検索値 As String, 置換値 As String)
    pRangeObject.Replace 検索値, 置換値
End Sub

' ========================================
' メソッド: オートフィル
' ========================================

''' <summary>
''' オートフィルで値を埋める
''' 方向: xlRight(右), xlDown(下), xlLeft(左), xlUp(上)
''' </summary>
Public Sub オートフィル(宛先 As Object, Optional 方向 As Long = xlDown)
    pRangeObject.AutoFill 宛先, 方向
End Sub

' ========================================
' メソッド: フィルター
' ========================================

''' <summary>
''' オートフィルターを適用
''' </summary>
Public Sub オートフィルター適用()
    pRangeObject.AutoFilter
End Sub

''' <summary>
''' オートフィルターを削除
''' </summary>
Public Sub オートフィルター解除()
    pRangeObject.AutoFilter
End Sub

' ========================================
' メソッド: 結合・分割
' ========================================

''' <summary>
''' セルを結合
''' </summary>
Public Sub セルを結合(Optional 中央配置 As Boolean = True)
    With pRangeObject
        .Merge
        If 中央配置 Then
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End If
    End With
End Sub

''' <summary>
''' 結合されたセルを分割
''' </summary>
Public Sub 結合を解除()
    pRangeObject.UnMerge
End Sub

''' <summary>
''' セルが結合されているかチェック
''' </summary>
Public Function 結合済み() As Boolean
    結合済み = pRangeObject.MergeCells
End Function

' ========================================
' メソッド: グループ化
' ========================================

''' <summary>
''' 行または列をグループ化
''' </summary>
Public Sub グループ化()
    pRangeObject.Group
End Sub

''' <summary>
''' グループ化を解除
''' </summary>
Public Sub グループ解除()
    pRangeObject.Ungroup
End Sub

' ========================================
' メソッド: ソート・フィルター
' ========================================

''' <summary>
''' セル範囲をソート
''' キー列: ソートのキーとなる列番号
''' 昇順: True の場合は昇順、False の場合は降順
''' </summary>
Public Sub ソート(Optional キー列 As Long = 1, Optional 昇順 As Boolean = True)
    pRangeObject.Sort pRangeObject.Columns(キー列), IIf(昇順, xlAscending, xlDescending)
End Sub

' ========================================
' メソッド: 特殊な貼り付け
''' </summary>
Public Sub 値だけ貼り付け()
    pRangeObject.PasteSpecial xlValues
End Sub

''' <summary>
''' 書式だけを貼り付け
''' </summary>
Public Sub 書式だけ貼り付け()
    pRangeObject.PasteSpecial xlFormats
End Sub

''' <summary>
''' 数式だけを貼り付け
''' </summary>
Public Sub 数式だけ貼り付け()
    pRangeObject.PasteSpecial xlFormulas
End Sub

' ========================================
' メソッド: 特殊な範囲取得
' ========================================

''' <summary>
''' 空白セルのみを取得（SpecialCells）
''' </summary>
Public Function 空白セル取得() As Object
    On Error Resume Next
    Set 空白セル取得 = pRangeObject.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
End Function

''' <summary>
''' 定数のみを取得
''' </summary>
Public Function 定数セル取得() As Object
    On Error Resume Next
    Set 定数セル取得 = pRangeObject.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
End Function

''' <summary>
''' 数式を含むセルのみを取得
''' </summary>
Public Function 数式セル取得() As Object
    On Error Resume Next
    Set 数式セル取得 = pRangeObject.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
End Function

''' <summary>
''' エラー値を含むセルのみを取得
''' </summary>
'Public Function エラーセル取得() As Object
'    On Error Resume Next
'
'    Set エラーセル取得 = pRangeObject.SpecialCells(xlCellTypeErrors)
'    On Error GoTo 0
'End Function

' ========================================
' メソッド: データ検証
' ========================================

''' <summary>
''' データ検証ルールを追加
''' 型: xlWhole（整数）, xlDecimal（小数）, xlList（リスト）など
''' 最小値, 最大値, または リスト項目を指定
''' </summary>
Public Sub データ検証追加(データ型 As Long, 条件 As Long, Optional 数値1, Optional 数値2)
    With pRangeObject.Validation
        .Delete
        .Add データ型, , 条件, 数値1, 数値2
        .ShowInput = True
        .ShowError = True
    End With
End Sub

''' <summary>
''' データ検証ルールをクリア
''' </summary>
Public Sub データ検証クリア()
    pRangeObject.Validation.Delete
End Sub

' ========================================
' メソッド: 条件付き書式
' ========================================

''' <summary>
''' 条件付き書式を追加
''' </summary>
Public Sub 条件付き書式追加(条件タイプ As Long, 演算子 As Long, 数式1 As String, Optional 数式2)
    pRangeObject.FormatConditions.Add 条件タイプ, 演算子, 数式1, 数式2
End Sub

''' <summary>
''' 条件付き書式をすべてクリア
''' </summary>
Public Sub 条件付き書式クリア()
    pRangeObject.FormatConditions.Delete
End Sub

' ========================================
' ユーティリティメソッド
' ========================================

''' <summary>
''' 隣のセルを取得（右）
''' </summary>
Public Function 右のセル(Optional オフセット As Long = 1) As Object
    Set 右のセル = pRangeObject.Offset(0, オフセット)
End Function

''' <summary>
''' 隣のセルを取得（左）
''' </summary>
Public Function 左のセル(Optional オフセット As Long = 1) As Object
    Set 左のセル = pRangeObject.Offset(0, -オフセット)
End Function

''' <summary>
''' 隣のセルを取得（下）
''' </summary>
Public Function 下のセル(Optional オフセット As Long = 1) As Object
    Set 下のセル = pRangeObject.Offset(オフセット, 0)
End Function

''' <summary>
''' 隣のセルを取得（上）
''' </summary>
Public Function 上のセル(Optional オフセット As Long = 1) As Object
    Set 上のセル = pRangeObject.Offset(-オフセット, 0)
End Function
