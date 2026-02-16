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


Private 元のオブジェクト As Range

' ========================================
' 初期化
' ========================================

Public Sub 初期化する(オブジェクト As Range)
    Set 元のオブジェクト = オブジェクト
End Sub

Private Function 新しいインスタンスを生成する(ByRef オブジェクト As Range) As セル範囲
    Set 新しいインスタンスを生成する = New セル範囲
    新しいインスタンスを生成する.初期化する オブジェクト
End Function

' ========================================
' デストラクタ
' ========================================

Private Sub Class_Terminate()
    Set 元のオブジェクト = Nothing
End Sub

' ========================================
' プロパティ: ラップ元のオブジェクト
' ========================================

''' <summary>
''' ラップ元のオブジェクトを取得
''' </summary>
Public Property Get ラップ元のオブジェクト() As Range
    ラップ元のオブジェクト = 元のオブジェクト
End Property

' ========================================
' プロパティ: アドレス関連
' ========================================

''' <summary>
''' セルのアドレスを取得
''' </summary>
Public Property Get セルアドレス() As String
    セルアドレス = 元のオブジェクト.Address
End Property

''' <summary>
''' 列番号を取得
''' </summary>
Public Property Get 列番号() As Long
    列番号 = 元のオブジェクト.Column
End Property

''' <summary>
''' 行番号を取得
''' </summary>
Public Property Get 行番号() As Long
    行番号 = 元のオブジェクト.Row
End Property

''' <summary>
''' 最後の列番号を取得（複数列の場合）
''' </summary>
Public Property Get 最後の列番号() As Long
    最後の列番号 = 元のオブジェクト.Column + 元のオブジェクト.Columns.Count - 1
End Property

''' <summary>
''' 最後の行番号を取得（複数行の場合）
''' </summary>
Public Property Get 最後の行番号() As Long
    最後の行番号 = 元のオブジェクト.Row + 元のオブジェクト.Rows.Count - 1
End Property

' ========================================
' プロパティ: サイズ関連
' ========================================

''' <summary>
''' 幅をポイント単位で取得/設定
''' </summary>
Public Property Get 幅() As Double
    幅 = 元のオブジェクト.Width
End Property

Public Property Let 幅(ポイント数 As Double)
    元のオブジェクト.Width = ポイント数
End Property

''' <summary>
''' 高さをポイント単位で取得/設定
''' </summary>
Public Property Get 高さ() As Double
    高さ = 元のオブジェクト.Height
End Property

Public Property Let 高さ(ポイント数 As Double)
    元のオブジェクト.Height = ポイント数
End Property

''' <summary>
''' 列数を取得
''' </summary>
Public Property Get 列数() As Long
    列数 = 元のオブジェクト.Columns.Count
End Property

''' <summary>
''' 行数を取得
''' </summary>
Public Property Get 行数() As Long
    行数 = 元のオブジェクト.Rows.Count
End Property

''' <summary>
''' セル数を取得
''' </summary>
Public Property Get セル数() As Long
    セル数 = 元のオブジェクト.Count
End Property

' ========================================
' プロパティ: 値と数式
' ========================================

''' <summary>
''' セルの値を取得/設定
''' </summary>
Public Property Get 値() As Variant
    値 = 元のオブジェクト.value
End Property

Public Property Let 値(設定値 As Variant)
    元のオブジェクト.value = 設定値
End Property

''' <summary>
''' セルの数式を取得/設定
''' </summary>
Public Property Get 数式() As String
    数式 = 元のオブジェクト.Formula
End Property

Public Property Let 数式(設定値 As String)
    元のオブジェクト.Formula = 設定値
End Property

''' <summary>
''' セルのR1C1参照スタイル数式を取得/設定
''' </summary>
Public Property Get R1C1数式() As String
    R1C1数式 = 元のオブジェクト.FormulaR1C1
End Property

Public Property Let R1C1数式(設定値 As String)
    元のオブジェクト.FormulaR1C1 = 設定値
End Property

''' <summary>
''' セルの左上隅のテキストを取得
''' </summary>
Public Property Get テキスト() As String
    テキスト = 元のオブジェクト.Text
End Property

''' <summary>
''' セルの表示値（計算結果）を取得
''' </summary>
Public Property Get 表示値() As Variant
    表示値 = 元のオブジェクト.value
End Property

' ========================================
' プロパティ: 書式関連
' ========================================

''' <summary>
''' 数値形式を取得/設定
''' </summary>
Public Property Get 数値形式() As String
    数値形式 = 元のオブジェクト.NumberFormat
End Property

Public Property Let 数値形式(設定値 As String)
    元のオブジェクト.NumberFormat = 設定値
End Property

''' <summary>
''' 背景色を取得/設定（RGB値）
''' </summary>
Public Property Get 背景色() As Long
    背景色 = 元のオブジェクト.Interior.Color
End Property

Public Property Let 背景色(色 As Long)
    元のオブジェクト.Interior.Color = 色
End Property

''' <summary>
''' 文字色を取得/設定（RGB値）
''' </summary>
Public Property Get 文字色() As Long
    文字色 = 元のオブジェクト.Font.Color
End Property

Public Property Let 文字色(色 As Long)
    元のオブジェクト.Font.Color = 色
End Property

''' <summary>
''' フォントサイズを取得/設定
''' </summary>
Public Property Get フォントサイズ() As Long
    フォントサイズ = 元のオブジェクト.Font.Size
End Property

Public Property Let フォントサイズ(サイズ As Long)
    元のオブジェクト.Font.Size = サイズ
End Property

''' <summary>
''' フォント名を取得/設定
''' </summary>
Public Property Get フォント名() As String
    フォント名 = 元のオブジェクト.Font.Name
End Property

Public Property Let フォント名(名前 As String)
    元のオブジェクト.Font.Name = 名前
End Property

''' <summary>
''' 太字かどうかを取得/設定
''' </summary>
Public Property Get 太字() As Boolean
    太字 = 元のオブジェクト.Font.Bold
End Property

Public Property Let 太字(太字にする As Boolean)
    元のオブジェクト.Font.Bold = 太字にする
End Property

''' <summary>
''' 斜体かどうかを取得/設定
''' </summary>
Public Property Get 斜体() As Boolean
    斜体 = 元のオブジェクト.Font.Italic
End Property

Public Property Let 斜体(斜体にする As Boolean)
    元のオブジェクト.Font.Italic = 斜体にする
End Property

''' <summary>
''' 下線かどうかを取得/設定
''' </summary>
Public Property Get 下線() As Boolean
    下線 = 元のオブジェクト.Font.Underline <> xlNone
End Property

Public Property Let 下線(下線を付ける As Boolean)
    元のオブジェクト.Font.Underline = IIf(下線を付ける, xlUnderlineStyleSingle, xlNone)
End Property

''' <summary>
''' 水平配置を取得/設定（-4152: 左, -4108: 中央, -4160: 右）
''' </summary>
Public Property Get 水平配置() As Long
    水平配置 = 元のオブジェクト.HorizontalAlignment
End Property

Public Property Let 水平配置(設定値 As Long)
    元のオブジェクト.HorizontalAlignment = 設定値
End Property

''' <summary>
''' 垂直配置を取得/設定（-4107: 上, -4108: 中央, -4161: 下）
''' </summary>
Public Property Get 垂直配置() As Long
    垂直配置 = 元のオブジェクト.VerticalAlignment
End Property

Public Property Let 垂直配置(設定値 As Long)
    元のオブジェクト.VerticalAlignment = 設定値
End Property

''' <summary>
''' テキスト折返しを取得/設定
''' </summary>
Public Property Get テキスト折返し() As Boolean
    テキスト折返し = 元のオブジェクト.WrapText
End Property

Public Property Let テキスト折返し(テキストを折り返す As Boolean)
    元のオブジェクト.WrapText = テキストを折り返す
End Property

''' <summary>
''' インデント有無を取得/設定
''' </summary>
Public Property Get インデント() As Boolean
    インデント = 元のオブジェクト.AddIndent
End Property

Public Property Let インデント(ByVal インデント有り As Boolean)
    元のオブジェクト.AddIndent = インデント有り
End Property

' ========================================
' プロパティ: 保護/ロック関連
' ========================================

''' <summary>
''' セルがロックされているかを取得/設定
''' </summary>
Public Property Get ロック() As Boolean
    ロック = 元のオブジェクト.Locked
End Property

Public Property Let ロック(ロックする As Boolean)
    元のオブジェクト.Locked = ロックする
End Property

''' <summary>
''' セルが非表示かどうかを取得/設定
''' </summary>
Public Property Get 非表示() As Boolean
    非表示 = 元のオブジェクト.Hidden
End Property

Public Property Let 非表示(非表示にする As Boolean)
    元のオブジェクト.Hidden = 非表示にする
End Property

' ========================================
' メソッド: 基本操作
' ========================================

''' <summary>
''' セルを選択
''' </summary>
Public Sub 選択する()
    元のオブジェクト.Select
End Sub

''' <summary>
''' セルをアクティブにする
''' </summary>
Public Sub アクティブにする()
    元のオブジェクト.Activate
End Sub

''' <summary>
''' セルをコピー
''' </summary>
Public Sub コピーする()
    元のオブジェクト.Copy
End Sub

''' <summary>
''' セルを切り取り
''' </summary>
Public Sub 切り取る()
    元のオブジェクト.Cut
End Sub

''' <summary>
''' すべてのコンテンツと書式をクリア
''' </summary>
Public Sub クリアする()
    元のオブジェクト.Clear
End Sub

''' <summary>
''' コンテンツのみをクリア（書式は保持）
''' </summary>
Public Sub 内容をクリアする()
    元のオブジェクト.ClearContents
End Sub

''' <summary>
''' 書式のみをクリア（コンテンツは保持）
''' </summary>
Public Sub 書式をクリアする()
    元のオブジェクト.ClearFormats
End Sub

''' <summary>
''' セルを削除
''' </summary>
Public Sub 削除する()
    元のオブジェクト.Delete
End Sub

' ========================================
' メソッド: セルの配置
' ========================================

''' <summary>
''' セルをオフセット位置から取得
''' オフセット行数、オフセット列数を指定
''' </summary>
Public Function オフセットした範囲を取得する(Optional 行オフセット As Long = 0, Optional 列オフセット As Long = 0) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = 元のオブジェクト.Offset(行オフセット, 列オフセット)
    Set オフセットした範囲を取得する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' セルのサイズを変更した新しい Range を取得
''' </summary>
Public Function サイズを変更した範囲を取得する(新しい範囲の行数 As Long, 新しい範囲の列数 As Long) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = 元のオブジェクト.Resize(新しい範囲の行数, 新しい範囲の列数)
    Set サイズを変更した範囲を取得する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' 複数のセル範囲の共通部分を取得
''' </summary>
Public Function 共通セル範囲を取得する(他の範囲 As セル範囲) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = Application.Intersect(元のオブジェクト, 他の範囲)
    Set 共通セル範囲を取得する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' 複数のセル範囲を統合した新しい Range を取得
''' </summary>
Public Function 結合したセル範囲を取得する(他の範囲 As Object) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = Application.Union(元のオブジェクト, 他の範囲)
    Set 結合したセル範囲を取得する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

' ========================================
' メソッド: 検索・置換
' ========================================

''' <summary>
''' セル範囲内にある値を検索
''' 見つかった場合は Range オブジェクトを返す
''' </summary>
Public Function 検索する(検索値 As String, Optional 検索タイプ As Long = xlValues) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = 元のオブジェクト.Find(検索値)
    Set 検索する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' 前回の検索から次の一致を検索
''' </summary>
Public Function 次を検索する(検索値 As String) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = 元のオブジェクト.FindNext(元のオブジェクト.Find(検索値))
    Set 次を検索する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' セル範囲内の値を置換
''' </summary>
Public Sub 置換する(検索値 As String, 置換値 As String)
    元のオブジェクト.Replace 検索値, 置換値
End Sub

' ========================================
' メソッド: オートフィル
' ========================================

''' <summary>
''' オートフィルで値を埋める
''' 方向: xlRight(右), xlDown(下), xlLeft(左), xlUp(上)
''' </summary>
Public Sub オートフィルする(宛先 As Object, Optional 方向 As Long = xlDown)
    元のオブジェクト.AutoFill 宛先, 方向
End Sub

' ========================================
' メソッド: フィルター
' ========================================

''' <summary>
''' オートフィルターを適用
''' </summary>
Public Sub オートフィルターを適用する()
    元のオブジェクト.AutoFilter
End Sub

''' <summary>
''' オートフィルターを削除
''' </summary>
Public Sub オートフィルターを解除する()
    元のオブジェクト.AutoFilter
End Sub

' ========================================
' メソッド: 結合・分割
' ========================================

''' <summary>
''' セルを結合
''' </summary>
Public Sub セルを結合する(Optional 中央配置 As Boolean = True)
    With 元のオブジェクト
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
Public Sub 結合を解除する()
    元のオブジェクト.UnMerge
End Sub

''' <summary>
''' セルが結合されているかチェック
''' </summary>
Public Function 結合済み() As Boolean
    結合済み = 元のオブジェクト.MergeCells
End Function

' ========================================
' メソッド: グループ化
' ========================================

''' <summary>
''' 行または列をグループ化
''' </summary>
Public Sub グループ化する()
    元のオブジェクト.Group
End Sub

''' <summary>
''' グループ化を解除
''' </summary>
Public Sub グループ解除する()
    元のオブジェクト.Ungroup
End Sub

' ========================================
' メソッド: ソート・フィルター
' ========================================

''' <summary>
''' セル範囲をソート
''' キー列: ソートのキーとなる列番号
''' 昇順: True の場合は昇順、False の場合は降順
''' </summary>
Public Sub ソートする(Optional キー列 As Long = 1, Optional 昇順 As Boolean = True)
    元のオブジェクト.Sort 元のオブジェクト.Columns(キー列), IIf(昇順, xlAscending, xlDescending)
End Sub

' ========================================
' メソッド: 特殊な貼り付け
''' </summary>
Public Sub 値だけ貼り付ける()
    元のオブジェクト.PasteSpecial xlValues
End Sub

''' <summary>
''' 書式だけを貼り付け
''' </summary>
Public Sub 書式だけ貼り付ける()
    元のオブジェクト.PasteSpecial xlFormats
End Sub

''' <summary>
''' 数式だけを貼り付け
''' </summary>
Public Sub 数式だけ貼り付ける()
    元のオブジェクト.PasteSpecial xlFormulas
End Sub

' ========================================
' メソッド: 特殊な範囲取得
' ========================================

''' <summary>
''' 空白セルのみを取得（SpecialCells）
''' </summary>
Public Function 空白セルを取得する() As Object
    On Error Resume Next
    Set 空白セルを取得する = 元のオブジェクト.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
End Function

''' <summary>
''' 定数のみを取得
''' </summary>
Public Function 定数セルを取得する() As Object
    On Error Resume Next
    Set 定数セルを取得する = 元のオブジェクト.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
End Function

''' <summary>
''' 数式を含むセルのみを取得
''' </summary>
Public Function 数式セルを取得する() As Object
    On Error Resume Next
    Set 数式セルを取得する = 元のオブジェクト.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
End Function

''' <summary>
''' エラー値を含むセルのみを取得
''' </summary>
'Public Function エラーセル取得() As Object
'    On Error Resume Next
'
'    Set エラーセル取得 = 元のオブジェクト.SpecialCells(xlCellTypeErrors)
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
Public Sub データ検証ルールを追加する(データ型 As Long, 条件 As Long, Optional 数値1, Optional 数値2)
    With 元のオブジェクト.Validation
        .Delete
        .Add データ型, , 条件, 数値1, 数値2
        .ShowInput = True
        .ShowError = True
    End With
End Sub

''' <summary>
''' データ検証ルールをクリア
''' </summary>
Public Sub データ検証ルールをクリアする()
    元のオブジェクト.Validation.Delete
End Sub

' ========================================
' メソッド: 条件付き書式
' ========================================

''' <summary>
''' 条件付き書式を追加
''' </summary>
Public Sub 条件付き書式を追加する(条件タイプ As Long, 演算子 As Long, 数式1 As String, Optional 数式2)
    元のオブジェクト.FormatConditions.Add 条件タイプ, 演算子, 数式1, 数式2
End Sub

''' <summary>
''' 条件付き書式をすべてクリア
''' </summary>
Public Sub 条件付き書式をクリアする()
    元のオブジェクト.FormatConditions.Delete
End Sub

' ========================================
' ユーティリティメソッド
' ========================================

''' <summary>
''' 隣のセルを取得（右）
''' </summary>
Public Function 右のセル(Optional オフセット As Long = 1) As Object
    Set 右のセル = 元のオブジェクト.Offset(0, オフセット)
End Function

''' <summary>
''' 隣のセルを取得（左）
''' </summary>
Public Function 左のセル(Optional オフセット As Long = 1) As Object
    Set 左のセル = 元のオブジェクト.Offset(0, -オフセット)
End Function

''' <summary>
''' 隣のセルを取得（下）
''' </summary>
Public Function 下のセル(Optional オフセット As Long = 1) As Object
    Set 下のセル = 元のオブジェクト.Offset(オフセット, 0)
End Function

''' <summary>
''' 隣のセルを取得（上）
''' </summary>
Public Function 上のセル(Optional オフセット As Long = 1) As Object
    Set 上のセル = 元のオブジェクト.Offset(-オフセット, 0)
End Function
