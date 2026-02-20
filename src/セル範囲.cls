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
''' Activate メソッドのラッパー
''' </summary>
Public Sub アクティブにする()
    ラップ元のオブジェクト.Activate
End Sub

''' <summary>
''' AddComment メソッドのラッパー
''' </summary>
Public Function コメントを追加する(ByVal コメント As String) As Comment
    Set コメントを追加する = ラップ元のオブジェクト.AddComment(コメント)
End Function

''' <summary>
''' AddCommentThreaded メソッドのラッパー
''' </summary>
Public Function スレッドコメントを追加する(ByVal スレッドコメント As String) As CommentThreaded
    Set スレッドコメントを追加する = ラップ元のオブジェクト.AddCommentThreaded(スレッドコメント)
End Function

''' <summary>
''' AdvancedFilter メソッドのラッパー
''' </summary>
Public Function フィルター処理する(ByVal 操作 As XlFilterAction, Optional ByVal 検索条件範囲 As セル範囲, Optional ByVal 宛先範囲 As セル範囲, Optional ByVal 重複するレコードは無視する As Boolean) As Variant
    Set フィルター処理する = ラップ元のオブジェクト.AdvancedFilter(操作, 検索条件範囲, 宛先範囲, 重複するレコードは無視する)
End Function

''' <summary>
''' AllocateChanges メソッドのラッパー
''' </summary>
Public Sub 書き戻し操作を実行する()
    ラップ元のオブジェクト.AllocateChanges
End Sub

''' <summary>
''' ApplyNames メソッドのラッパー
''' </summary>
Public Function 名前を付ける(Optional ByVal 名前の配列 As Variant, Optional ByVal 参照を名前に置き換える As Boolean, Optional ByVal 見つからない時に行と列の名前を使う As Boolean, Optional ByVal 行方向の名前に置き換える As Boolean, Optional ByVal 列方向の名前に置き換える As Boolean, Optional ByVal 行と列のどちらを優先するか As XlApplyNamesOrder, Optional ByVal 最後に定義された名前の定義を置き換える As Boolean) As Variant
    Set 名前を付ける = ラップ元のオブジェクト.ApplyNames(名前の配列, 参照を名前に置き換える, 見つからない時に行と列の名前を使う, 行方向の名前に置き換える, 列方向の名前に置き換える, 行と列のどちらを優先するか, 最後に定義された名前の定義を置き換える)
End Function


''' <summary>
''' ApplyOutlineStyles メソッドのラッパー
''' </summary>
Public Sub アウトラインの既定のスタイルを適用する()
    ラップ元のオブジェクト.ApplyOutlineStyles
End Sub

''' <summary>
''' AutoComplete メソッドのラッパー
''' </summary>
Public Function オートコンプリートする(ByVal 文字列 As String) As String
    Let オートコンプリートする = ラップ元のオブジェクト.AutoComplete(文字列)
End Function

''' <summary>
''' AutoFill メソッドのラッパー
''' 方向: xlRight(右), xlDown(下), xlLeft(左), xlUp(上)
''' </summary>
Public Function 連続データを作成する(ByVal 範囲 As セル範囲, Optional 方向 As XlAutoFillType) As Variant
    Set 連続データを作成する = ラップ元のインスタンス.AutoFill(範囲, 方向)
End Function

''' <summary>
''' AutoFilter メソッドのラッパー
''' </summary>
Public Function オートフィルターする(Optional ByVal フィールド番号 As Integer, Optional ByVal 抽出条件 As String, Optional ByVal 演算子 As XlAutoFilterOperator, Optional ByVal 抽出条件2 As String, Optional ByVal データ型のフィールド As Variant, Optional ByVal ドロップダウン矢印を表示する As Boolean) As Variant
    Set オートフィルターする = ラップ元のオブジェクト.AutoFilter(フィールド番号, 抽出条件, 演算子, 抽出条件2, データ型のフィールド, ドロップダウン矢印を表示する)
End Function

''' <summary>
''' AutoFit メソッドのラッパー
''' </summary>
Public Sub 幅と高さを自動調整する()
    ラップ元のオブジェクト.AutoFit
End Sub

''' <summary>
''' AutoOutline メソッドのラッパー
''' </summary>
Public Function 自動アウトラインを実行する() As Variant
    Set 自動アウトラインを実行する = ラップ元のオブジェクト.AutoOutline
End Function

''' <summary>
''' BorderAround メソッドのラッパー
''' </summary>
Public Function 罫線で囲む(Optional ByVal LineStyle As Variant, Optional ByVal Weight As XlBorderWeight, Optional ByVal ColorIndex As XlColorIndex, Optional ByVal Color As Variant, Optional ByVal ThemeColor As Variant) As Variant
    Set 罫線で囲む = ラップ元のオブジェクト.BorderAround(LineStyle, Weight, ColorIndex, Color, ThemeColor)
End Function

''' <summary>
''' Calculate メソッドのラッパー
''' </summary>
Public Function 計算する() As Variant
    Set 計算する = ラップ元のオブジェクト.Calculate()
End Function

''' <summary>
''' CalculateRowMajorOrder メソッドのラッパー
''' </summary>
Public Function 行メジャー順に計算する() As Variant
    Set 行メジャー順に計算する = ラップ元のオブジェクト.CalculateRowMajorOrder()
End Function

''' <summary>
''' CheckSpelling  メソッドのラッパー
''' </summary>
Public Function スペルチェックする(CustomDictionary, IgnoreUppercase, AlwaysSuggest, SpellLang) As Variant
    Set スペルチェックする = ラップ元のオブジェクト.CheckSpelling(CustomDictionary, IgnoreUppercase, AlwaysSuggest, SpellLang)
End Function

''' <summary>
''' Clear メソッドのラッパー
''' </summary>
Public Sub クリアする()
    ラップ元のオブジェクト.Clear
End Sub

''' <summary>
''' ClearComments メソッドのラッパー
''' </summary>
Public Sub コメントを削除する()
    ラップ元のオブジェクト.ClearComments
End Sub

''' <summary>
''' ClearContents メソッドのラッパー
''' </summary>
Public Function 数式と値をクリアする() As Variant
    Set 数式と値をクリアする = ラップ元のオブジェクト.ClearContents()
End Function

''' <summary>
''' ClearFormats メソッドのラッパー
''' </summary>
Public Function 書式設定を削除する() As Variant
    Set 書式設定を削除する = ラップ元のオブジェクト.ClearFormats()
End Function

''' <summary>
''' ClearHyperlinks メソッドのラッパー
''' </summary>
Public Sub ハイパーリンクを削除する()
    ラップ元のオブジェクト.ClearHyperlinks
End Sub

''' <summary>
''' ClearNotes メソッドのラッパー
''' </summary>
Public Function ノートを削除する() As Variant
    Set ノートを削除する = ラップ元のオブジェクト.ClearNotes()
End Function

''' <summary>
''' ClearOutline メソッドのラッパー
''' </summary>
Public Function アウトラインを消去する() As Variant
    Set アウトラインを消去する = ラップ元のオブジェクト.ClearOutline()
End Function

''' <summary>
''' ColumnDifferences メソッドのラッパー
''' </summary>
Public Function アクティブ列との相違を取得する(Comparison) As セル範囲
    Set アクティブ列との相違を取得する = ラップ元のオブジェクト.ColumnDifferences(Comparison)
End Function

''' <summary>
''' Consolidate メソッドのラッパー
''' </summary>
Public Function 統合する(Sources, Function_, TopRow, LeftColumn, CreateLinks) As Variant
    Set 統合する = ラップ元のオブジェクト.Consolidate(Sources, Function_, TopRow, LeftColumn, CreateLinks)
End Function

''' <summary>
''' ConvertToLinkedDataType メソッドのラッパー
''' </summary>
Public Sub リンクされたデータ型に変換する(ServiceID, LanguageCulture)
    ラップ元のオブジェクト.ConvertToLinkedDataType ServiceID, LanguageCulture
End Sub

''' <summary>
''' Copy メソッドのラッパー
''' </summary>
Public Function コピーする(Destination) As Variant
    Set コピーする = ラップ元のオブジェクト.Copy(Destination)
End Function

''' <summary>
''' CopyFromRecordset メソッドのラッパー
''' </summary>
Public Function データを取得して変換する(Data, MaxRows, MaxColumns) As Long
    Let データを取得して変換する = ラップ元のオブジェクト.CopyFromRecordset(Data, MaxRows, MaxColumns)
End Function

''' <summary>
''' CopyPicture メソッドのラッパー
''' </summary>
Public Function ピクチャをコピーする(外観, 形式) As Variant
    Set ピクチャをコピーする = ラップ元のオブジェクト.CopyPicture(外観, 形式)
End Function

''' <summary>
''' CreateNames メソッドのラッパー
''' </summary>
Public Function 名前を作成する(上, 左, 下, 右) As Variant
    Set 名前を作成する = ラップ元のオブジェクト.CreateNames(上, 左, 下, 右)
End Function

''' <summary>
''' Cut メソッドのラッパー
''' </summary>
Public Function 切り取る(宛先) As Variant
    Set 切り取る = ラップ元のオブジェクト.Cut(宛先)
End Function

''' <summary>
''' DataSeries メソッドのラッパー
''' </summary>
Public Function 連続デ_タを作成する(Optional ByVal Rowcol As Variant, Optional ByVal Type_ As XlDataSeriesType, Optional ByVal Date_ As XlDataSeriesDate, Optional ByVal Step As Variant, Optional ByVal Stop_ As Variant, Optional ByVal Trend As Variant) As Variant
    Set 連続デ_タを作成する = ラップ元のオブジェクト.DataSeries(Rowcol, Type_, Date_, Step, Stop_, Trend)
End Function

''' <summary>
''' DataTypeToText メソッドのラッパー
''' </summary>
Public Sub リンクされたデータ型をテキストに変換する()
    ラップ元のオブジェクト.DataTypeToText
End Sub

''' <summary>
''' Delete メソッドのラッパー
''' </summary>
Public Function 削除する(Shift) As Variant
    Set 削除する = ラップ元のオブジェクト.Delete(Shift)
End Function

''' <summary>
''' DialogBox メソッドのラッパー
''' </summary>
Public Function ダイアログボックスを表示する() As Variant
    Set ダイアログボックスを表示する = ラップ元のオブジェクト.DialogBox()
End Function

''' <summary>
''' Dirty メソッドのラッパー
''' </summary>
Public Sub 再計算対象にする()
    ラップ元のオブジェクト.Dirty
End Sub

''' <summary>
''' DiscardChanges メソッドのラッパー
''' </summary>
Public Sub 変更を破棄する()
    ラップ元のオブジェクト.DiscardChanges
End Sub

''' <summary>
''' EditionOptions メソッドのラッパー
''' </summary>
Public Function キーワードのヘルプを表示する(Type_, Option_, Name, Reference, Appearance, ChartSize, Format) As Variant
    Set キーワードのヘルプを表示する = ラップ元のオブジェクト.EditionOptions(Type_, Option_, Name, Reference, Appearance, ChartSize, Format)
End Function

''' <summary>
''' ExportAsFixedFormat メソッドのラッパー
''' </summary>
Public Sub 形式を指定してエクスポートする(Type_, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To_, OpenAfterPublish, FixedFormatExtClassPtr)
    ラップ元のオブジェクト.ExportAsFixedFormat Type_, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To_, OpenAfterPublish, FixedFormatExtClassPtr
End Sub

''' <summary>
''' FillDown メソッドのラッパー
''' </summary>
Public Function 下にフィルする() As Variant
    Set 下にフィルする = ラップ元のオブジェクト.FillDown()
End Function

''' <summary>
''' FillLeft メソッドのラッパー
''' </summary>
Public Function 左にフィルする() As Variant
    Set 左にフィルする = ラップ元のオブジェクト.FillLeft()
End Function

''' <summary>
''' FillRight メソッドのラッパー
''' </summary>
Public Function 右にフィルする() As Variant
    Set 右にフィルする = ラップ元のオブジェクト.FillRight()
End Function

''' <summary>
''' FillUp メソッドのラッパー
''' </summary>
Public Function 上にフィルする() As Variant
    Set 上にフィルする = ラップ元のオブジェクト.FillUp()
End Function

''' <summary>
''' Find メソッドのラッパー
''' </summary>
Public Function 検索する(検索値 As String, Optional 検索タイプ As Long = xlValues) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = ラップ元のインスタンス.Find(検索値, 検索タイプ)
    Set 検索する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' FindNext メソッドのラッパー
''' </summary>
Public Function 次を検索する(検索値 As String) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = ラップ元のインスタンス.FindNext(検索値)
    Set 次を検索する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' FindPrevious メソッドのラッパー
''' </summary>
Public Function 前を検索する(検索値 As String) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = ラップ元のインスタンス.FindPrevious(検索値)
    Set 前を検索する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' FlashFill メソッドのラッパー
''' </summary>
Public Sub フラッシュフィルする()
    ラップ元のオブジェクト.FlashFill
End Sub

''' <summary>
''' FunctionWizard メソッドのラッパー
''' </summary>
Public Function 関数ウィザードを起動する() As Variant
    Set 関数ウィザードを起動する = ラップ元のオブジェクト.FunctionWizard()
End Function

''' <summary>
''' Group メソッドのラッパー
''' </summary>
Public Sub グループ化する(Start, End_, By, Periods)
    ラップ元のインスタンス.Group Start, End_, By, Periods
End Sub

''' <summary>
''' Insert メソッドのラッパー
''' </summary>
Public Function 挿入する(Shift, CopyOrigin) As Variant
    Set 挿入する = ラップ元のオブジェクト.Insert(Shift, CopyOrigin)
End Function

''' <summary>
''' InsertIndent メソッドのラッパー
''' </summary>
Public Sub インデントを変更する(InsertAmount)
    ラップ元のオブジェクト.InsertIndent InsertAmount
End Sub

''' <summary>
''' Justify メソッドのラッパー
''' </summary>
Public Function 割り付けする() As Variant
    Set 割り付けする = ラップ元のオブジェクト.Justify()
End Function

''' <summary>
''' ListNames メソッドのラッパー
''' </summary>
Public Function 名前の一覧を貼り付ける() As Variant
    Set 名前の一覧を貼り付ける = ラップ元のオブジェクト.ListNames()
End Function

''' <summary>
''' Merge メソッドのラッパー
''' セルを結合する
''' </summary>
Public Sub 結合する(全体)
    ラップ元のインスタンス.Merge 全体
End Sub

''' <summary>
''' NavigateArrow メソッドのラッパー
''' </summary>
Public Function トレース矢印をたどる(TowardPrecedent, ArrowNumber, LinkNumber) As Variant
    Set トレース矢印をたどる = ラップ元のオブジェクト.NavigateArrow(TowardPrecedent, ArrowNumber, LinkNumber)
End Function

''' <summary>
''' NoteText メソッドのラッパー
''' </summary>
Public Function ノートを取得または設定する(テキスト, 開始, 長さ) As String
    Let ノートを取得または設定する = ラップ元のオブジェクト.NoteText(テキスト, 開始, 長さ)
End Function

''' <summary>
''' Parse メソッドのラッパー
''' </summary>
Public Function データを分割する(ParseLine, Destination) As Variant
    Set データを分割する = ラップ元のオブジェクト.Parse(ParseLine, Destination)
End Function

''' <summary>
''' PasteSpecial メソッドのラッパー
''' </summary>
Public Function 特殊な貼り付けを実行する(貼り付け, 操作, SkipBlanks, Transpose) As Variant
    Set 特殊な貼り付けを実行する = ラップ元のオブジェクト.PasteSpecial(貼り付け, 操作, SkipBlanks, Transpose)
End Function

''' <summary>
''' PrintOut メソッドのラッパー
''' </summary>
Public Function 印刷する(From, To_, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName) As Variant
    Set 印刷する = ラップ元のオブジェクト.PrintOut(From, To_, Copies, Preview, ActivePrinter, PrintToFile, Collate, PrToFileName)
End Function

''' <summary>
''' PrintPreview メソッドのラッパー
''' </summary>
Public Function 印刷プレビューを表示する(EnableChanges) As Variant
    Set 印刷プレビューを表示する = ラップ元のオブジェクト.PrintPreview(EnableChanges)
End Function

''' <summary>
''' RemoveDuplicates メソッドのラッパー
''' </summary>
Public Sub 重複した値を削除する(列, ヘッダー)
    ラップ元のオブジェクト.RemoveDuplicates 列, ヘッダー
End Sub

''' <summary>
''' RemoveSubtotal メソッドのラッパー
''' </summary>
Public Function 小計を削除する() As Variant
    Set 小計を削除する = ラップ元のオブジェクト.RemoveSubtotal()
End Function

''' <summary>
''' Replace メソッドのラッパー
''' </summary>
Public Function 置換する(ByVal 検索する値 As Variant, ByVal 置換後の値 As Variant, Optional ByVal セル内容が完全に同一であるものを検索する As Variant, Optional ByVal 検索方向 As Variant, Optional ByVal 大文字と小文字を区別する As Variant, Optional ByVal 半角と全角を区別する As Variant, Optional ByVal 検索書式 As Variant, Optional ByVal 置換書式 As Variant, Optional ByVal 数式バージョン As Variant) As Boolean
    Let 置換する = ラップ元のインスタンス.Replace(検索する値, 置換後の値, セル内容が完全に同一であるものを検索する, 検索方向, 大文字と小文字を区別する, 半角と全角を区別する, 検索書式, 置換書式, 数式バージョン)
End Function

''' <summary>
''' ResetContents メソッドのラッパー
''' </summary>
Public Sub 内容をクリアする()
    ラップ元のオブジェクト.ResetContents
End Sub

''' <summary>
''' RowDifferences メソッドのラッパー
''' </summary>
Public Function アクティブ行との相違を取得する(比較) As Range
    Set アクティブ行との相違を取得する = ラップ元のオブジェクト.RowDifferences(比較)
End Function

''' <summary>
''' Run メソッドのラッパー
''' </summary>
Public Function マクロを実行する( _
Optional ByVal Arg1 As Variant, Optional ByVal Arg2 As Variant, Optional ByVal Arg3 As Variant, Optional ByVal Arg4 As Variant, Optional ByVal Arg5 As Variant, Optional ByVal Arg6 As Variant, Optional ByVal Arg7 As Variant, Optional ByVal Arg8 As Variant, Optional ByVal Arg9 As Variant, Optional ByVal Arg10 As Variant, _
Optional ByVal Arg11 As Variant, Optional ByVal Arg12 As Variant, Optional ByVal Arg13 As Variant, Optional ByVal Arg14 As Variant, Optional ByVal Arg15 As Variant, Optional ByVal Arg16 As Variant, Optional ByVal Arg17 As Variant, Optional ByVal Arg18 As Variant, Optional ByVal Arg19 As Variant, Optional ByVal Arg20 As Variant, _
Optional ByVal Arg21 As Variant, Optional ByVal Arg22 As Variant, Optional ByVal Arg23 As Variant, Optional ByVal Arg24 As Variant, Optional ByVal Arg25 As Variant, Optional ByVal Arg26 As Variant, Optional ByVal Arg27 As Variant, Optional ByVal Arg28 As Variant, Optional ByVal Arg29 As Variant, Optional ByVal Arg30) As Variant
    Set マクロを実行する = ラップ元のオブジェクト.Run(Arg1, Arg2, Arg3, Arg4, Arg5, Arg6, Arg7, Arg8, Arg9, Arg10, Arg11, Arg12, Arg13, Arg14, Arg15, Arg16, Arg17, Arg18, Arg19, Arg20, Arg21, Arg22, Arg23, Arg24, Arg25, Arg26, Arg27, Arg28, Arg29, Arg30)
End Function

''' <summary>
''' Select メソッドのラッパー
''' </summary>
Public Function 選択する() As Variant
    Set 選択する = ラップ元のオブジェクト.Select()
End Function

''' <summary>
''' SetCellDataTypeFromCell メソッドのラッパー
''' </summary>
Public Sub リンクされたデータ型をコピーする(ByVal コピー先 As セル範囲)
    ラップ元のオブジェクト.SetCellDataTypeFromCell コピー先.ラップ元のオブジェクト
End Sub

''' <summary>
''' SetPhonetic メソッドのラッパー
''' </summary>
Public Sub ふりがなを表示する()
    ラップ元のオブジェクト.SetPhonetic
End Sub

''' <summary>
''' Show メソッドのラッパー
''' </summary>
Public Function 見える位置に来るようスクロールする() As Variant
    Set 見える位置に来るようスクロールする = ラップ元のオブジェクト.Show()
End Function

''' <summary>
''' ShowCard メソッドのラッパー
''' </summary>
Public Sub カードを表示する()
    ラップ元のオブジェクト.ShowCard
End Sub

''' <summary>
''' ShowDependents メソッドのラッパー
''' </summary>
Public Function 依存先をトレースする(Remove) As Variant
    Set 依存先をトレースする = ラップ元のオブジェクト.ShowDependents(Remove)
End Function

''' <summary>
''' ShowErrors メソッドのラッパー
''' </summary>
Public Function エラーをトレースする() As Variant
    Set エラーをトレースする = ラップ元のオブジェクト.ShowErrors()
End Function

''' <summary>
''' ShowPrecedents メソッドのラッパー
''' </summary>
Public Function 参照元をトレースする(Remove) As Variant
    Set 参照元をトレースする = ラップ元のオブジェクト.ShowPrecedents(Remove)
End Function

''' <summary>
''' Sort メソッドのラッパー
''' キー列: ソートのキーとなる列番号
''' 昇順: True の場合は昇順、False の場合は降順
''' </summary>
Public Function ソートする(Key1, Order1, Key2, Type_, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, SortMethod, DataOption1, DataOption2, DataOption3) As Variant
    Set ソートする = ラップ元のインスタンス.Sort(Key1, Order1, Key2, Type_, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, SortMethod, DataOption1, DataOption2, DataOption3)
End Function

''' <summary>
''' SortSpecial メソッドのラッパー
''' </summary>
Public Function 並べ替える(SortMethod, Key1, Order1, Type_, Key2, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, DataOption1, DataOption2, DataOption3) As Variant
    Set 並べ替える = ラップ元のオブジェクト.SortSpecial(SortMethod, Key1, Order1, Type_, Key2, Order2, Key3, Order3, Header, OrderCustom, MatchCase, Orientation, DataOption1, DataOption2, DataOption3)
End Function

''' <summary>
''' Speak メソッドのラッパー
''' </summary>
Public Sub 読み上げる(SpeakDirection, SpeakFormulas)
    ラップ元のオブジェクト.Speak SpeakDirection, SpeakFormulas
End Sub

''' <summary>
''' SpecialCells メソッドのラッパー
''' </summary>
Public Function 条件を選択してジャンプする(型, 値) As セル範囲
    Set 条件を選択してジャンプする = ラップ元のオブジェクト.SpecialCells(型, 値)
End Function

''' <summary>
''' SubscribeTo メソッドのラッパー
''' </summary>
Public Function 外部データの変更を反映する(ByVal ファイル名 As String, Optional ByVal フォーマット As XlEditionFormat) As Variant
    Set 外部データの変更を反映する = ラップ元のインスタンス.SubscribeTo(ファイル名, フォーマット)
End Function

''' <summary>
''' Subtotal メソッドのラッパー
''' </summary>
Public Function 集計する(GroupBy, 関数, TotalList, Replace, PageBreaks, SummaryBelowData) As Variant
    Set 集計する = ラップ元のオブジェクト.Subtotal(GroupBy, 関数, TotalList, Replace, PageBreaks, SummaryBelowData)
End Function

''' <summary>
''' Table メソッドのラッパー
''' </summary>
Public Function データテーブルを作成する(RowInput, ColumnInput) As Variant
    Set データテーブルを作成する = ラップ元のオブジェクト.Table(RowInput, ColumnInput)
End Function

''' <summary>
''' TextToColumns メソッドのラッパー
''' </summary>
Public Function データを区切る(Destination, DataType, TextQualifier, consecutiveDelimiter, Tab_, セミコロン, コンマ, スペース, Other, OtherChar, FieldInfo, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers) As Variant
    Set データを区切る = ラップ元のオブジェクト.TextToColumns(Destination, DataType, TextQualifier, consecutiveDelimiter, Tab_, セミコロン, コンマ, スペース, Other, OtherChar, FieldInfo, DecimalSeparator, ThousandsSeparator, TrailingMinusNumbers)
End Function

''' <summary>
''' Ungroup メソッドのラッパー
''' </summary>
Public Function グループ解除する() As Variant
    Set グループ解除する = ラップ元のインスタンス.Ungroup
End Function

''' <summary>
''' UnMerge メソッドのラッパー
''' 結合されたセルを分割
''' </summary>
Public Sub 結合を解除する()
    ラップ元のインスタンス.UnMerge
End Sub


' ========================================
' プロパティ
' ========================================

''' <summary>
''' AddIndent プロパティのラッパー
''' インデント有無を取得/設定
''' </summary>
Public Property Get インデント() As Boolean
    インデント = ラップ元のインスタンス.AddIndent
End Property

Public Property Let インデント(ByVal インデント有り As Boolean)
    ラップ元のインスタンス.AddIndent = インデント有り
End Property

''' <summary>
''' セルのアドレスを取得する
''' Address プロパティのラッパー
''' </summary>
Public Property Get アドレス() As String
    アドレス = ラップ元のインスタンス.Address
End Property

''' <summary>
''' AddressLocal プロパティのラッパー
''' </summary>

''' <summary>
''' AllowEdit プロパティのラッパー
''' </summary>

''' <summary>
''' Application プロパティのラッパー
''' </summary>

''' <summary>
''' Areas プロパティのラッパー
''' </summary>

''' <summary>
''' Borders プロパティのラッパー
''' </summary>

''' <summary>
''' CellControl プロパティのラッパー
''' </summary>

''' <summary>
''' Cells プロパティのラッパー
''' </summary>

''' <summary>
''' Characters プロパティのラッパー
''' </summary>

''' <summary>
''' 列番号を取得する
''' Column プロパティのラッパー
''' </summary>
Public Property Get 列番号() As Long
    列番号 = ラップ元のインスタンス.Column
End Property

''' <summary>
''' Columns プロパティのラッパー
''' </summary>

''' <summary>
''' ColumnWidth プロパティのラッパー
''' </summary>

''' <summary>
''' Comment プロパティのラッパー
''' </summary>

''' <summary>
''' CommentThreaded プロパティのラッパー
''' </summary>

''' <summary>
''' セル数を取得する
''' Count プロパティのラッパー
''' </summary>
Public Property Get セル数() As Long
    セル数 = ラップ元のインスタンス.Count
End Property

''' <summary>
''' CountLarge プロパティのラッパー
''' </summary>

''' <summary>
''' Creator プロパティのラッパー
''' </summary>

''' <summary>
''' CurrentArray プロパティのラッパー
''' </summary>

''' <summary>
''' CurrentRegion プロパティのラッパー
''' </summary>

''' <summary>
''' Dependents プロパティのラッパー
''' </summary>

''' <summary>
''' DirectDependents プロパティのラッパー
''' </summary>

''' <summary>
''' DirectPrecedents プロパティのラッパー
''' </summary>

''' <summary>
''' DisplayFormat プロパティのラッパー
''' </summary>

''' <summary>
''' End プロパティのラッパー
''' </summary>

''' <summary>
''' EntireColumn プロパティのラッパー
''' </summary>

''' <summary>
''' EntireRow プロパティのラッパー
''' </summary>

''' <summary>
''' Errors プロパティのラッパー
''' </summary>

''' <summary>
''' Font プロパティのラッパー
''' </summary>

''' <summary>
''' FormatConditions プロパティのラッパー
''' </summary>

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


''' <summary>
''' 数式を取得/設定する
''' Formula プロパティのラッパー
''' </summary>

Public Property Get 数式() As String
    数式 = ラップ元のインスタンス.Formula
End Property

Public Property Let 数式(設定値 As String)
    ラップ元のインスタンス.Formula = 設定値
End Property


''' <summary>
''' FormulaArray プロパティのラッパー
''' </summary>

''' <summary>
''' FormulaHidden プロパティのラッパー
''' </summary>

''' <summary>
''' FormulaLocal プロパティのラッパー
''' </summary>

''' <summary>
''' R1C1参照スタイルの数式を取得/設定する
''' FormulaR1C1 プロパティのラッパー
''' </summary>
''' <summary>
''' </summary>

Public Property Get R1C1参照スタイルの数式() As String
    R1C1参照スタイルの数式 = ラップ元のインスタンス.FormulaR1C1
End Property

Public Property Let R1C1参照スタイルの数式(設定値 As String)
    ラップ元のインスタンス.FormulaR1C1 = 設定値
End Property

''' <summary>
''' FormulaR1C1Local プロパティのラッパー
''' </summary>

''' <summary>
''' HasArray プロパティのラッパー
''' </summary>

''' <summary>
''' HasFormula プロパティのラッパー
''' </summary>

''' <summary>
''' HasRichDataType プロパティのラッパー
''' </summary>

''' <summary>
''' 高さを取得/設定する（ポイント単位）
''' Height プロパティのラッパー
''' </summary>
Public Property Get 高さ() As Double
    高さ = ラップ元のインスタンス.Height
End Property

Public Property Let 高さ(ポイント数 As Double)
    ラップ元のインスタンス.Height = ポイント数
End Property

''' <summary>
''' セルが非表示かどうかを取得/設定する
''' Hidden プロパティのラッパー
''' </summary>

Public Property Get 非表示である() As Boolean
    非表示である = ラップ元のインスタンス.Hidden
End Property

Public Property Let 非表示である(非表示にする As Boolean)
    ラップ元のインスタンス.Hidden = 非表示にする
End Property

''' <summary>
''' 横位置を取得/設定する（-4152: 左, -4108: 中央, -4160: 右）
''' HorizontalAlignment プロパティのラッパー
''' </summary>
Public Property Get 横位置() As Long
    横位置 = ラップ元のインスタンス.HorizontalAlignment
End Property

Public Property Let 横位置(設定値 As Long)
    ラップ元のインスタンス.HorizontalAlignment = 設定値
End Property

''' <summary>
''' Hyperlinks プロパティのラッパー
''' </summary>

''' <summary>
''' ID プロパティのラッパー
''' </summary>

''' <summary>
''' IndentLevel プロパティのラッパー
''' </summary>

''' <summary>
''' Interior プロパティのラッパー
''' </summary>

''' <summary>
''' Item プロパティのラッパー
''' </summary>

''' <summary>
''' Left プロパティのラッパー
''' </summary>

''' <summary>
''' LinkedDataTypeState プロパティのラッパー
''' </summary>

''' <summary>
''' ListHeaderRows プロパティのラッパー
''' </summary>

''' <summary>
''' ListObject プロパティのラッパー
''' </summary>

''' <summary>
''' LocationInTable プロパティのラッパー
''' </summary>

''' <summary>
''' セルがロックされているかを取得/設定する
''' Locked プロパティのラッパー
''' </summary>
Public Property Get ロックされている() As Boolean
    Let ロックされている = ラップ元のインスタンス.Locked
End Property

Public Property Let ロックされている(ロックする As Boolean)
    ラップ元のインスタンス.Locked = ロックする
End Property


''' <summary>
''' MDX プロパティのラッパー
''' </summary>

''' <summary>
''' MergeArea プロパティのラッパー
''' </summary>

''' <summary>
''' MergeCells プロパティのラッパー
''' セルが結合されているかチェック
''' </summary>
Public Function 結合されている() As Boolean
    Let 結合されている = ラップ元のインスタンス.MergeCells
End Function

''' <summary>
''' Name プロパティのラッパー
''' </summary>

''' <summary>
''' Next プロパティのラッパー
''' </summary>

''' <summary>
''' 数値の表示形式を取得/設定する
''' NumberFormat プロパティのラッパー
''' </summary>
Public Property Get 数値の表示形式() As String
    数値の表示形式 = ラップ元のインスタンス.NumberFormat
End Property

Public Property Let 数値の表示形式(設定値 As String)
    ラップ元のインスタンス.NumberFormat = 設定値
End Property

''' <summary>
''' NumberFormatLocal プロパティのラッパー
''' </summary>

''' <summary>
''' Offset プロパティのラッパー
''' </summary>
Public Function オフセットした範囲を取得する(Optional 行オフセット As Long = 0, Optional 列オフセット As Long = 0) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = ラップ元のインスタンス.Offset(行オフセット, 列オフセット)
    Set オフセットした範囲を取得する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' Orientation プロパティのラッパー
''' </summary>

''' <summary>
''' OutlineLevel プロパティのラッパー
''' </summary>

''' <summary>
''' PageBreak プロパティのラッパー
''' </summary>

''' <summary>
''' Parent プロパティのラッパー
''' </summary>

''' <summary>
''' Phonetic プロパティのラッパー
''' </summary>

''' <summary>
''' Phonetics プロパティのラッパー
''' </summary>

''' <summary>
''' PivotCell プロパティのラッパー
''' </summary>

''' <summary>
''' PivotField プロパティのラッパー
''' </summary>

''' <summary>
''' PivotItem プロパティのラッパー
''' </summary>

''' <summary>
''' PivotTable プロパティのラッパー
''' </summary>

''' <summary>
''' Precedents プロパティのラッパー
''' </summary>

''' <summary>
''' PrefixCharacter プロパティのラッパー
''' </summary>

''' <summary>
''' Previous プロパティのラッパー
''' </summary>

''' <summary>
''' QueryTable プロパティのラッパー
''' </summary>

''' <summary>
''' Range プロパティのラッパー
''' </summary>

''' <summary>
''' ReadingOrder プロパティのラッパー
''' </summary>

''' <summary>
''' Resize プロパティのラッパー
''' セルのサイズを変更した新しい Range を取得
''' </summary>
Public Function サイズを変更した範囲を取得する(新しい範囲の行数 As Long, 新しい範囲の列数 As Long) As セル範囲
    Dim 戻り値のオブジェクト As Range
    Set 戻り値のオブジェクト = ラップ元のインスタンス.Resize(新しい範囲の行数, 新しい範囲の列数)
    Set サイズを変更した範囲を取得する = 新しいインスタンスを生成する(戻り値のオブジェクト)
End Function

''' <summary>
''' 行番号を取得する
''' Row プロパティのラッパー
''' </summary>
Public Property Get 行番号() As Long
    行番号 = ラップ元のインスタンス.Row
End Property

''' <summary>
''' RowHeight プロパティのラッパー
''' </summary>

''' <summary>
''' Rows プロパティのラッパー
''' </summary>

''' <summary>
''' ServerActions プロパティのラッパー
''' </summary>

''' <summary>
''' ShowDetail プロパティのラッパー
''' </summary>

''' <summary>
''' ShrinkToFit プロパティのラッパー
''' </summary>

''' <summary>
''' SoundNote プロパティのラッパー
''' </summary>

''' <summary>
''' SparklineGroups プロパティのラッパー
''' </summary>

''' <summary>
''' Style プロパティのラッパー
''' </summary>

''' <summary>
''' Summary プロパティのラッパー
''' </summary>

''' <summary>
''' 左上隅セルのテキストを取得する
''' Text プロパティのラッパー
''' </summary>
Public Property Get テキスト() As String
    テキスト = ラップ元のインスタンス.Text
End Property

''' <summary>
''' Top プロパティのラッパー
''' </summary>

''' <summary>
''' UseStandardHeight プロパティのラッパー
''' </summary>

''' <summary>
''' UseStandardWidth プロパティのラッパー
''' </summary>

''' <summary>
''' Validation プロパティのラッパー
''' </summary>

''' <summary>
''' value プロパティのラッパー
''' </summary>

Public Property Get 値() As Variant
    値 = ラップ元のインスタンス.value
End Property

Public Property Let 値(設定値 As Variant)
    ラップ元のインスタンス.value = 設定値
End Property


''' <summary>
''' Value2 プロパティのラッパー
''' </summary>

''' <summary>
''' 縦位置を取得/設定（-4107: 上, -4108: 中央, -4161: 下）
''' VerticalAlignment プロパティのラッパー
''' </summary>

Public Property Get 縦位置() As Long
    縦位置 = ラップ元のインスタンス.VerticalAlignment
End Property

Public Property Let 縦位置(設定値 As Long)
    ラップ元のインスタンス.VerticalAlignment = 設定値
End Property


''' <summary>
''' 幅を取得/設定する（ポイント単位）
''' Width プロパティのラッパー
''' </summary>
Public Property Get 幅() As Double
    幅 = ラップ元のインスタンス.Width
End Property

Public Property Let 幅(ポイント数 As Double)
    ラップ元のインスタンス.Width = ポイント数
End Property

''' <summary>
''' Worksheet プロパティのラッパー
''' </summary>

''' <summary>
''' 折り返して全体を表示するを取得/設定
''' WrapText プロパティのラッパー
''' </summary>

Public Property Get 折り返して全体を表示している() As Boolean
    折り返して全体を表示している = ラップ元のインスタンス.WrapText
End Property

Public Property Let 折り返して全体を表示している(折り返して全体を表示する As Boolean)
    ラップ元のインスタンス.WrapText = 折り返して全体を表示する
End Property

''' <summary>
''' XPath プロパティのラッパー
''' </summary>

