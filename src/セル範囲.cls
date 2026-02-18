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
' セル範囲クラス
' Microsoft Excel Object Library の Range クラスの日本語ラッパークラス

' ========================================
' メンバ変数
' ========================================

''' <summary>
''' ラップ元のインスタンス
''' </summary>
Private ラップ元のインスタンス As Range


' ========================================
' 初期化
' ========================================

''' <summary>
''' 初期化する
''' </summary>
Public Sub 初期化する(ByRef オブジェクト As Range)
    Set ラップ元のインスタンス = オブジェクト
End Sub

''' <summary>
''' 新しいインスタンスを生成する
''' </summary>
Private Function 新しいインスタンスを生成する(ByRef オブジェクト As Range) As セル範囲
    Set 新しいインスタンスを生成する = New セル範囲
    新しいインスタンスを生成する.初期化する オブジェクト
End Function


' ========================================
' デストラクタ
' ========================================

Private Sub Class_Terminate()
    Set ラップ元のインスタンス = Nothing
End Sub


' ========================================
' プロパティ: ラップ元のオブジェクト
' ========================================

''' <summary>
''' ラップ元のオブジェクトを取得
''' </summary>
Public Property Get ラップ元のオブジェクト() As Range
    ラップ元のオブジェクト = ラップ元のインスタンス
End Property


' ========================================
' メソッド
' ========================================

''' <summary>
''' Addressメソッドのラッパー
''' </summary>
Public Sub アクティブにする()

End Sub

''' <summary>
''' AddCommentメソッドのラッパー
''' </summary>
Public Sub コメントを追加する()

End Sub

''' <summary>
''' AddCommentThreadedメソッドのラッパー
''' </summary>
Public Sub スレッドコメントを追加する()

End Sub

''' <summary>
''' AdvancedFilterメソッドのラッパー
''' </summary>
Public Sub フィルター処理する()

End Sub

''' <summary>
''' AllocateChangesメソッドのラッパー
''' </summary>
Public Sub 書き戻し操作を実行する()

End Sub

''' <summary>
''' ApplyNamesメソッドのラッパー
''' </summary>
Public Sub 名前を付ける()

End Sub


''' <summary>
''' ApplyOutlineStylesメソッドのラッパー
''' </summary>
Public Sub アウトラインの既定のスタイルを適用する()

End Sub

''' <summary>
''' AutoCompleteメソッドのラッパー
''' </summary>
Public Sub オートコンプリートする()

End Sub

''' <summary>
''' AutoFillメソッドのラッパー
''' 方向: xlRight(右), xlDown(下), xlLeft(左), xlUp(上)
''' </summary>
Public Sub オートフィルする(宛先 As Object, Optional 方向 As Long = xlDown)
    ラップ元のインスタンス.AutoFill 宛先, 方向
End Sub

''' <summary>
''' AutoFilterメソッドのラッパー
''' </summary>
Public Sub オートフィルターする()

End Sub

''' <summary>
''' AutoFitメソッドのラッパー
''' </summary>
Public Sub 幅と高さを自動調整する()

End Sub

''' <summary>
''' AutoOutlineメソッドのラッパー
''' </summary>
Public Sub アウトラインを自動的に作成する()

End Sub

''' <summary>
''' BorderAroundメソッドのラッパー
''' </summary>
Public Sub 罫線で囲む()

End Sub

''' <summary>
''' Calculateメソッドのラッパー
''' </summary>
Public Sub 計算する()

End Sub

''' <summary>
''' CalculateRowMajorOrderメソッドのラッパー
''' </summary>
Public Sub 行メジャー順に計算する()

End Sub

''' <summary>
''' CheckSpellingメソッドのラッパー
''' </summary>
Public Sub スペルチェックする()

End Sub

''' <summary>
''' Clearメソッドのラッパー
''' </summary>
Public Sub クリアする()

End Sub

''' <summary>
''' ClearCommentsメソッドのラッパー
''' </summary>
Public Sub コメントを削除する()

End Sub

''' <summary>
''' ClearContentsメソッドのラッパー
''' </summary>
Public Sub 数式と値をクリアする()

End Sub

''' <summary>
''' ClearFormatsメソッドのラッパー
''' </summary>
Public Sub 書式設定を削除する()

End Sub

''' <summary>
''' ClearHyperlinksメソッドのラッパー
''' </summary>
Public Sub ハイパーリンクを削除する()

End Sub

''' <summary>
''' ClearNotesメソッドのラッパー
''' </summary>
Public Sub ノートを削除する()

End Sub

''' <summary>
''' ClearOutlineメソッドのラッパー
''' </summary>
Public Sub アウトラインを消去する()

End Sub

''' <summary>
''' ColumnDifferencesメソッドのラッパー
''' </summary>
Public Sub 対象セル範囲内のセルと各列の比較セルで､内容の異なるすべてのセルを取得する()

End Sub

''' <summary>
''' Consolidateメソッドのラッパー
''' </summary>
Public Sub 統合する()

End Sub

''' <summary>
''' ConvertToLinkedDataTypeメソッドのラッパー
''' </summary>
Public Sub リンクされたデータ型に変換する()

End Sub

''' <summary>
''' Copyメソッドのラッパー
''' </summary>
Public Sub コピーする()

End Sub

''' <summary>
''' CopyFromRecordsetメソッドのラッパー
''' </summary>
Public Sub Recordsetをコピーする()

End Sub

''' <summary>
''' CopyPictureメソッドのラッパー
''' </summary>
Public Sub ピクチャをコピーする()

End Sub

''' <summary>
''' CreateNamesメソッドのラッパー
''' </summary>
Public Sub 名前を作成する()

End Sub

''' <summary>
''' Cutメソッドのラッパー
''' </summary>
Public Sub 切り取る()

End Sub

''' <summary>
''' DataTypeToTextメソッドのラッパー
''' </summary>
Public Sub リンクされたデータ型をテキストに変換する()

End Sub

''' <summary>
''' DataSeriesメソッドのラッパー
''' </summary>
Public Sub 連続データを作成する()

End Sub

''' <summary>
''' Deleteメソッドのラッパー
''' </summary>
Public Sub 削除する()

End Sub

''' <summary>
''' DialogBoxメソッドのラッパー
''' </summary>
Public Sub ダイアログボックスを表示する()

End Sub

''' <summary>
''' Dirtyメソッドのラッパー
''' </summary>
Public Sub 再計算する範囲を指定する()

End Sub

''' <summary>
''' DiscardChangesメソッドのラッパー
''' </summary>
Public Sub 変更を破棄する()

End Sub

''' <summary>
''' EditionOptionsメソッドのラッパー
''' </summary>
Public Sub キーワードのヘルプを表示する()

End Sub

''' <summary>
''' ExportAsFixedFormatメソッドのラッパー
''' </summary>
Public Sub 指定された形式のファイルにエクスポートする()

End Sub

''' <summary>
''' FillDownメソッドのラッパー
''' </summary>
Public Sub 下にフィルする()

End Sub

''' <summary>
''' FillLeftメソッドのラッパー
''' </summary>
Public Sub 左にフィルする()

End Sub

''' <summary>
''' FillRightメソッドのラッパー
''' </summary>
Public Sub 右にフィルする()

End Sub

''' <summary>
''' FillUpメソッドのラッパー
''' </summary>
Public Sub 上にフィルする()

End Sub

''' <summary>
''' Findメソッドのラッパー
''' </summary>
Public Function 検索する(検索値 As String, Optional 検索タイプ As Long = xlValues) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = ラップ元のインスタンス.Find(検索値)
    Set 検索する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' FindNextメソッドのラッパー
''' </summary>
Public Function 次を検索する(検索値 As String) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = ラップ元のインスタンス.FindNext(ラップ元のインスタンス.Find(検索値))
    Set 次を検索する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' FindPreviousメソッドのラッパー
''' </summary>
Public Sub 前を検索する()

End Sub

''' <summary>
''' FlashFillメソッドのラッパー
''' </summary>
Public Sub フラッシュフィルする()

End Sub

''' <summary>
''' FunctionWizardメソッドのラッパー
''' </summary>
Public Sub 関数ウィザードを起動する()

End Sub

''' <summary>
''' Groupメソッドのラッパー
''' </summary>
Public Sub グループ化する()
    ラップ元のインスタンス.Group
End Sub

''' <summary>
''' Insertメソッドのラッパー
''' </summary>
Public Sub 挿入する()

End Sub

''' <summary>
''' InsertIndentメソッドのラッパー
''' </summary>
Public Sub インデントを変更する()

End Sub

''' <summary>
''' Justifyメソッドのラッパー
''' </summary>
Public Sub 割り付けする()

End Sub

''' <summary>
''' ListNamesメソッドのラッパー
''' </summary>
Public Sub 名前の一覧を貼り付ける()

End Sub

''' <summary>
''' Mergeメソッドのラッパー
''' </summary>
Public Sub マージする()

End Sub

''' <summary>
''' NavigateArrowメソッドのラッパー
''' </summary>
Public Sub トレース矢印をたどって移動する()

End Sub

''' <summary>
''' NoteTextメソッドのラッパー
''' </summary>
Public Sub ノートを取得または設定する()

End Sub

''' <summary>
''' Parseメソッドのラッパー
''' </summary>
Public Sub データを分割する()

End Sub

''' <summary>
''' PasteSpecialメソッドのラッパー
''' </summary>
Public Sub 特殊な貼り付けを実行する()

End Sub

''' <summary>
''' PrintOutメソッドのラッパー
''' </summary>
Public Sub 印刷する()

End Sub

''' <summary>
''' PrintPreviewメソッドのラッパー
''' </summary>
Public Sub 印刷プレビューを表示する()

End Sub

''' <summary>
''' RemoveDuplicatesメソッドのラッパー
''' </summary>
Public Sub 重複した値を削除する()

End Sub

''' <summary>
''' RemoveSubtotalメソッドのラッパー
''' </summary>
Public Sub 小計を削除する()

End Sub

''' <summary>
''' Replaceメソッドのラッパー
''' </summary>
Public Sub 置換する(検索値 As String, 置換値 As String)
    ラップ元のインスタンス.Replace 検索値, 置換値
End Sub

''' <summary>
''' RowDifferencesメソッドのラッパー
''' </summary>
Public Sub 各行の比較セルから内容が異なるすべてのセルを返する()

End Sub

''' <summary>
''' Runメソッドのラッパー
''' </summary>
Public Sub マクロを実行する()

End Sub

''' <summary>
''' Selectメソッドのラッパー
''' </summary>
Public Sub 選択する()

End Sub

''' <summary>
''' SetCellDataTypeFromCellメソッドのラッパー
''' </summary>
Public Sub リンクされたデータをコピーする()

End Sub

''' <summary>
''' SetPhoneticメソッドのラッパー
''' </summary>
Public Sub ふりがなを表示する()

End Sub

''' <summary>
''' Showメソッドのラッパー
''' </summary>
Public Sub 可視位置に来るようスクロールする()

End Sub

''' <summary>
''' ShowCardメソッドのラッパー
''' </summary>
Public Sub カードを表示する()

End Sub

''' <summary>
''' ShowDependentsメソッドのラッパー
''' </summary>
Public Sub 参照先へトレース矢印を引く()

End Sub

''' <summary>
''' ShowErrorsメソッドのラッパー
''' </summary>
Public Sub エラー矢印を表示する()

End Sub

''' <summary>
''' ShowPrecedentsメソッドのラッパー
''' </summary>
Public Sub 参照元から矢印を引く()

End Sub

''' <summary>
''' Sortメソッドのラッパー
''' キー列: ソートのキーとなる列番号
''' 昇順: True の場合は昇順、False の場合は降順
''' </summary>
Public Sub ソートする(Optional キー列 As Long = 1, Optional 昇順 As Boolean = True)
    ラップ元のインスタンス.Sort ラップ元のインスタンス.Columns(キー列), IIf(昇順, xlAscending, xlDescending)
End Sub

''' <summary>
''' SortSpecialメソッドのラッパー
''' </summary>
Public Sub 言語でソートする()

End Sub

''' <summary>
''' Speakメソッドのラッパー
''' </summary>
Public Sub 読み上げる()

End Sub

''' <summary>
''' SpecialCellsメソッドのラッパー
''' </summary>
Public Sub 特定の種類のセルを取得する()

End Sub

''' <summary>
''' SubscribeToメソッドのラッパー
''' </summary>
Public Sub サブスクライブする()

End Sub

''' <summary>
''' Subtotalメソッドのラッパー
''' </summary>
Public Sub 集計する()

End Sub

''' <summary>
''' Tableメソッドのラッパー
''' </summary>
Public Sub テーブルを作成する()

End Sub

''' <summary>
''' TextToColumnsメソッドのラッパー
''' </summary>
Public Sub データ区切る()

End Sub

''' <summary>
''' Ungroupメソッドのラッパー
''' </summary>
Public Sub グループ化を解除する()

End Sub

''' <summary>
''' UnMergeメソッドのラッパー
''' </summary>
Public Sub 分割する()

End Sub


' ========================================
' プロパティ
' ========================================

''' <summary>
''' インデント有無を取得/設定
''' </summary>
Public Property Get インデント() As Boolean
    インデント = ラップ元のインスタンス.AddIndent
End Property

Public Property Let インデント(ByVal インデント有り As Boolean)
    ラップ元のインスタンス.AddIndent = インデント有り
End Property

''' <summary>
''' セルの値を取得/設定
''' </summary>
Public Property Get 値() As Variant
    値 = ラップ元のインスタンス.value
End Property

Public Property Let 値(設定値 As Variant)
    ラップ元のインスタンス.value = 設定値
End Property

''' <summary>
''' セルをオフセット位置から取得
''' オフセット行数、オフセット列数を指定
''' </summary>
Public Function オフセットした範囲を取得する(Optional 行オフセット As Long = 0, Optional 列オフセット As Long = 0) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = ラップ元のインスタンス.Offset(行オフセット, 列オフセット)
    Set オフセットした範囲を取得する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

' ========================================
' プロパティ: アドレス関連
' ========================================

''' <summary>
''' セルのアドレスを取得
''' </summary>
Public Property Get セルアドレス() As String
    セルアドレス = ラップ元のインスタンス.Address
End Property

''' <summary>
''' 列番号を取得
''' </summary>
Public Property Get 列番号() As Long
    列番号 = ラップ元のインスタンス.Column
End Property

''' <summary>
''' 行番号を取得
''' </summary>
Public Property Get 行番号() As Long
    行番号 = ラップ元のインスタンス.Row
End Property

''' <summary>
''' 最後の列番号を取得（複数列の場合）
''' </summary>
Public Property Get 最後の列番号() As Long
    最後の列番号 = ラップ元のインスタンス.Column + ラップ元のインスタンス.Columns.Count - 1
End Property

''' <summary>
''' 最後の行番号を取得（複数行の場合）
''' </summary>
Public Property Get 最後の行番号() As Long
    最後の行番号 = ラップ元のインスタンス.Row + ラップ元のインスタンス.Rows.Count - 1
End Property

' ========================================
' プロパティ: サイズ関連
' ========================================

''' <summary>
''' 幅をポイント単位で取得/設定
''' </summary>
Public Property Get 幅() As Double
    幅 = ラップ元のインスタンス.Width
End Property

Public Property Let 幅(ポイント数 As Double)
    ラップ元のインスタンス.Width = ポイント数
End Property

''' <summary>
''' 高さをポイント単位で取得/設定
''' </summary>
Public Property Get 高さ() As Double
    高さ = ラップ元のインスタンス.Height
End Property

Public Property Let 高さ(ポイント数 As Double)
    ラップ元のインスタンス.Height = ポイント数
End Property

''' <summary>
''' 列数を取得
''' </summary>
Public Property Get 列数() As Long
    列数 = ラップ元のインスタンス.Columns.Count
End Property

''' <summary>
''' 行数を取得
''' </summary>
Public Property Get 行数() As Long
    行数 = ラップ元のインスタンス.Rows.Count
End Property

''' <summary>
''' セル数を取得
''' </summary>
Public Property Get セル数() As Long
    セル数 = ラップ元のインスタンス.Count
End Property

' ========================================
' プロパティ: 値と数式
' ========================================

''' <summary>
''' セルの数式を取得/設定
''' </summary>
Public Property Get 数式() As String
    数式 = ラップ元のインスタンス.Formula
End Property

Public Property Let 数式(設定値 As String)
    ラップ元のインスタンス.Formula = 設定値
End Property

''' <summary>
''' セルのR1C1参照スタイル数式を取得/設定
''' </summary>
Public Property Get R1C1数式() As String
    R1C1数式 = ラップ元のインスタンス.FormulaR1C1
End Property

Public Property Let R1C1数式(設定値 As String)
    ラップ元のインスタンス.FormulaR1C1 = 設定値
End Property

''' <summary>
''' セルの左上隅のテキストを取得
''' </summary>
Public Property Get テキスト() As String
    テキスト = ラップ元のインスタンス.Text
End Property

''' <summary>
''' セルの表示値（計算結果）を取得
''' </summary>
Public Property Get 表示値() As Variant
    表示値 = ラップ元のインスタンス.value
End Property

' ========================================
' プロパティ: 書式関連
' ========================================

''' <summary>
''' 数値形式を取得/設定
''' </summary>
Public Property Get 数値形式() As String
    数値形式 = ラップ元のインスタンス.NumberFormat
End Property

Public Property Let 数値形式(設定値 As String)
    ラップ元のインスタンス.NumberFormat = 設定値
End Property

''' <summary>
''' 背景色を取得/設定（RGB値）
''' </summary>
Public Property Get 背景色() As Long
    背景色 = ラップ元のインスタンス.Interior.Color
End Property

Public Property Let 背景色(色 As Long)
    ラップ元のインスタンス.Interior.Color = 色
End Property

''' <summary>
''' 文字色を取得/設定（RGB値）
''' </summary>
Public Property Get 文字色() As Long
    文字色 = ラップ元のインスタンス.Font.Color
End Property

Public Property Let 文字色(色 As Long)
    ラップ元のインスタンス.Font.Color = 色
End Property

''' <summary>
''' フォントサイズを取得/設定
''' </summary>
Public Property Get フォントサイズ() As Long
    フォントサイズ = ラップ元のインスタンス.Font.Size
End Property

Public Property Let フォントサイズ(サイズ As Long)
    ラップ元のインスタンス.Font.Size = サイズ
End Property

''' <summary>
''' フォント名を取得/設定
''' </summary>
Public Property Get フォント名() As String
    フォント名 = ラップ元のインスタンス.Font.Name
End Property

Public Property Let フォント名(名前 As String)
    ラップ元のインスタンス.Font.Name = 名前
End Property

''' <summary>
''' 太字かどうかを取得/設定
''' </summary>
Public Property Get 太字() As Boolean
    太字 = ラップ元のインスタンス.Font.Bold
End Property

Public Property Let 太字(太字にする As Boolean)
    ラップ元のインスタンス.Font.Bold = 太字にする
End Property

''' <summary>
''' 斜体かどうかを取得/設定
''' </summary>
Public Property Get 斜体() As Boolean
    斜体 = ラップ元のインスタンス.Font.Italic
End Property

Public Property Let 斜体(斜体にする As Boolean)
    ラップ元のインスタンス.Font.Italic = 斜体にする
End Property

''' <summary>
''' 下線かどうかを取得/設定
''' </summary>
Public Property Get 下線() As Boolean
    下線 = ラップ元のインスタンス.Font.Underline <> xlNone
End Property

Public Property Let 下線(下線を付ける As Boolean)
    ラップ元のインスタンス.Font.Underline = IIf(下線を付ける, xlUnderlineStyleSingle, xlNone)
End Property

''' <summary>
''' 水平配置を取得/設定（-4152: 左, -4108: 中央, -4160: 右）
''' </summary>
Public Property Get 水平配置() As Long
    水平配置 = ラップ元のインスタンス.HorizontalAlignment
End Property

Public Property Let 水平配置(設定値 As Long)
    ラップ元のインスタンス.HorizontalAlignment = 設定値
End Property

''' <summary>
''' 垂直配置を取得/設定（-4107: 上, -4108: 中央, -4161: 下）
''' </summary>
Public Property Get 垂直配置() As Long
    垂直配置 = ラップ元のインスタンス.VerticalAlignment
End Property

Public Property Let 垂直配置(設定値 As Long)
    ラップ元のインスタンス.VerticalAlignment = 設定値
End Property

''' <summary>
''' テキスト折返しを取得/設定
''' </summary>
Public Property Get テキスト折返し() As Boolean
    テキスト折返し = ラップ元のインスタンス.WrapText
End Property

Public Property Let テキスト折返し(テキストを折り返す As Boolean)
    ラップ元のインスタンス.WrapText = テキストを折り返す
End Property


' ========================================
' プロパティ: 保護/ロック関連
' ========================================

''' <summary>
''' セルがロックされているかを取得/設定
''' </summary>
Public Property Get ロック() As Boolean
    ロック = ラップ元のインスタンス.Locked
End Property

Public Property Let ロック(ロックする As Boolean)
    ラップ元のインスタンス.Locked = ロックする
End Property

''' <summary>
''' セルが非表示かどうかを取得/設定
''' </summary>
Public Property Get 非表示() As Boolean
    非表示 = ラップ元のインスタンス.Hidden
End Property

Public Property Let 非表示(非表示にする As Boolean)
    ラップ元のインスタンス.Hidden = 非表示にする
End Property

' ========================================
' メソッド: セルの配置
' ========================================

''' <summary>
''' セルのサイズを変更した新しい Range を取得
''' </summary>
Public Function サイズを変更した範囲を取得する(新しい範囲の行数 As Long, 新しい範囲の列数 As Long) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = ラップ元のインスタンス.Resize(新しい範囲の行数, 新しい範囲の列数)
    Set サイズを変更した範囲を取得する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' 複数のセル範囲の共通部分を取得
''' </summary>
Public Function 共通セル範囲を取得する(他の範囲 As セル範囲) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = Application.Intersect(ラップ元のインスタンス, 他の範囲)
    Set 共通セル範囲を取得する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' 複数のセル範囲を統合した新しい Range を取得
''' </summary>
Public Function 結合したセル範囲を取得する(他の範囲 As Object) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = Application.Union(ラップ元のインスタンス, 他の範囲)
    Set 結合したセル範囲を取得する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

' ========================================
' メソッド: フィルター
' ========================================

''' <summary>
''' オートフィルターを適用
''' </summary>
Public Sub オートフィルターを適用する()
    ラップ元のインスタンス.AutoFilter
End Sub

''' <summary>
''' オートフィルターを削除
''' </summary>
Public Sub オートフィルターを解除する()
    ラップ元のインスタンス.AutoFilter
End Sub

' ========================================
' メソッド: 結合・分割
' ========================================

''' <summary>
''' セルを結合
''' </summary>
Public Sub セルを結合する(Optional 中央配置 As Boolean = True)
    With ラップ元のインスタンス
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
    ラップ元のインスタンス.UnMerge
End Sub

''' <summary>
''' セルが結合されているかチェック
''' </summary>
Public Function 結合済み() As Boolean
    結合済み = ラップ元のインスタンス.MergeCells
End Function


''' <summary>
''' グループ化を解除
''' </summary>
Public Sub グループ解除する()
    ラップ元のインスタンス.Ungroup
End Sub

' ========================================
' メソッド: 特殊な貼り付け
''' </summary>
Public Sub 値だけ貼り付ける()
    ラップ元のインスタンス.PasteSpecial xlValues
End Sub

''' <summary>
''' 書式だけを貼り付け
''' </summary>
Public Sub 書式だけ貼り付ける()
    ラップ元のインスタンス.PasteSpecial xlFormats
End Sub

''' <summary>
''' 数式だけを貼り付け
''' </summary>
Public Sub 数式だけ貼り付ける()
    ラップ元のインスタンス.PasteSpecial xlFormulas
End Sub

' ========================================
' メソッド: 特殊な範囲取得
' ========================================

''' <summary>
''' 空白セルのみを取得（SpecialCells）
''' </summary>
Public Function 空白セルを取得する() As Object
    On Error Resume Next
    Set 空白セルを取得する = ラップ元のインスタンス.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
End Function

''' <summary>
''' 定数のみを取得
''' </summary>
Public Function 定数セルを取得する() As Object
    On Error Resume Next
    Set 定数セルを取得する = ラップ元のインスタンス.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
End Function

''' <summary>
''' 数式を含むセルのみを取得
''' </summary>
Public Function 数式セルを取得する() As Object
    On Error Resume Next
    Set 数式セルを取得する = ラップ元のインスタンス.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
End Function

''' <summary>
''' エラー値を含むセルのみを取得
''' </summary>
'Public Function エラーセル取得() As Object
'    On Error Resume Next
'
'    Set エラーセル取得 = ラップ元のインスタンス.SpecialCells(xlCellTypeErrors)
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
    With ラップ元のインスタンス.Validation
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
    ラップ元のインスタンス.Validation.Delete
End Sub

' ========================================
' メソッド: 条件付き書式
' ========================================

''' <summary>
''' 条件付き書式を追加
''' </summary>
Public Sub 条件付き書式を追加する(条件タイプ As Long, 演算子 As Long, 数式1 As String, Optional 数式2)
    ラップ元のインスタンス.FormatConditions.Add 条件タイプ, 演算子, 数式1, 数式2
End Sub

''' <summary>
''' 条件付き書式をすべてクリア
''' </summary>
Public Sub 条件付き書式をクリアする()
    ラップ元のインスタンス.FormatConditions.Delete
End Sub

' ========================================
' ユーティリティメソッド
' ========================================

''' <summary>
''' 隣のセルを取得（右）
''' </summary>
Public Function 右のセル(Optional オフセット As Long = 1) As Object
    Set 右のセル = ラップ元のインスタンス.Offset(0, オフセット)
End Function

''' <summary>
''' 隣のセルを取得（左）
''' </summary>
Public Function 左のセル(Optional オフセット As Long = 1) As Object
    Set 左のセル = ラップ元のインスタンス.Offset(0, -オフセット)
End Function

''' <summary>
''' 隣のセルを取得（下）
''' </summary>
Public Function 下のセル(Optional オフセット As Long = 1) As Object
    Set 下のセル = ラップ元のインスタンス.Offset(オフセット, 0)
End Function

''' <summary>
''' 隣のセルを取得（上）
''' </summary>
Public Function 上のセル(Optional オフセット As Long = 1) As Object
    Set 上のセル = ラップ元のインスタンス.Offset(-オフセット, 0)
End Function

