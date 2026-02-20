VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ワークシート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("Excel")

' Excelオブジェクトモデルのワークシート(Worksheet)をラップするクラス

Private m_ハンドル As Worksheet

' ========================================
' 初期化
' ========================================

Public Sub 新規(ByVal ワークシート As Worksheet)
    Set m_ハンドル = ワークシート
End Sub

Private Sub Class_Terminate()
    Set m_ハンドル = Nothing
End Sub

' ========================================
' 基本プロパティ
' ========================================

' 名前を取得/設定
Public Property Get シート名() As String
    シート名 = m_ハンドル.Name
End Property

Public Property Let シート名(ByVal シート名 As String)
    m_ハンドル.Name = シート名
End Property

' インデックスを取得
Public Property Get インデックス() As Long
    インデックス = m_ハンドル.Index
End Property

' コード名を取得
Public Property Get コード名() As String
    コード名 = m_ハンドル.CodeName
End Property

' タイプを取得
Public Property Get タイプ() As XlSheetType
    タイプ = m_ハンドル.Type
End Property

' 可視性を取得/設定
Public Property Get 可視化() As XlSheetVisibility
    可視化 = m_ハンドル.Visible
End Property

Public Property Let 可視化(ByVal 値 As XlSheetVisibility)
    m_ハンドル.Visible = 値
End Property

' ========================================
' セル・範囲プロパティ
' ========================================

' すべてのセルを取得
Public Property Get セル() As Range
    Set セル = m_ハンドル.Cells
End Property

' 指定したセルを取得
Public Function セルを取得(ByVal 行番号 As Long, Optional ByVal 列番号 As Long = 1) As Range
    Set セルを取得 = m_ハンドル.Cells(行番号, 列番号)
End Function

' 範囲を取得
Public Function 範囲を取得(ByVal アドレス As String) As Range
    Set 範囲を取得 = m_ハンドル.Range(アドレス)
End Function

' 使用済み範囲を取得
Public Property Get 使用済み範囲() As Range
    Set 使用済み範囲 = m_ハンドル.UsedRange
End Property

' 行を取得
Public Property Get 行() As Range
    Set 行 = m_ハンドル.Rows
End Property

' 列を取得
Public Property Get 列() As Range
    Set 列 = m_ハンドル.Columns
End Property

' ========================================
' ページ設定・印刷プロパティ
' ========================================

' ページ設定を取得
Public Property Get ページ設定() As PageSetup
    Set ページ設定 = m_ハンドル.PageSetup
End Property

' 標準行高を取得/設定
Public Property Get 標準行高() As Double
    標準行高 = m_ハンドル.StandardHeight
End Property

' 標準列幅を取得/設定
Public Property Get 標準列幅() As Double
    標準列幅 = m_ハンドル.StandardWidth
End Property

Public Property Let 標準列幅(ByVal 値 As Double)
    m_ハンドル.StandardWidth = 値
End Property

' ========================================
' 保護・フィルタプロパティ
' ========================================

' 内容を保護しているかを取得
Public Property Get 内容を保護中() As Boolean
    内容を保護中 = m_ハンドル.ProtectContents
End Property

' シナリオを保護しているかを取得
Public Property Get シナリオを保護中() As Boolean
    シナリオを保護中 = m_ハンドル.ProtectScenarios
End Property

' 描画オブジェクトを保護しているかを取得
Public Property Get 描画オブジェクトを保護中() As Boolean
    描画オブジェクトを保護中 = m_ハンドル.ProtectDrawingObjects
End Property

' オートフィルターモードを取得/設定
Public Property Get オートフィルターモード() As Boolean
    オートフィルターモード = m_ハンドル.AutoFilterMode
End Property

Public Property Let オートフィルターモード(ByVal 値 As Boolean)
    m_ハンドル.AutoFilterMode = 値
End Property

' フィルターモードを取得
Public Property Get フィルターモード() As Boolean
    フィルターモード = m_ハンドル.FilterMode
End Property

' ========================================
' 計算・表示プロパティ
' ========================================

' 計算を有効にするかを取得/設定
Public Property Get 計算を有効化() As Boolean
    計算を有効化 = m_ハンドル.EnableCalculation
End Property

Public Property Let 計算を有効化(ByVal 値 As Boolean)
    m_ハンドル.EnableCalculation = 値
End Property

' 表示ページ区切りを取得/設定
Public Property Get 表示ページ区切り() As Boolean
    表示ページ区切り = m_ハンドル.DisplayPageBreaks
End Property

Public Property Let 表示ページ区切り(ByVal 値 As Boolean)
    m_ハンドル.DisplayPageBreaks = 値
End Property

' 右から左に表示するかを取得/設定
Public Property Get 右から左に表示() As Boolean
    右から左に表示 = m_ハンドル.DisplayRightToLeft
End Property

Public Property Let 右から左に表示(ByVal 値 As Boolean)
    m_ハンドル.DisplayRightToLeft = 値
End Property

' スクロール領域を取得/設定
Public Property Get スクロール領域() As String
    スクロール領域 = m_ハンドル.ScrollArea
End Property

Public Property Let スクロール領域(ByVal 値 As String)
    m_ハンドル.ScrollArea = 値
End Property

' ========================================
' メソッド：基本操作
' ========================================

' ワークシートをアクティブにする
Public Sub アクティブ化()
    m_ハンドル.Activate
End Sub

' ワークシートを選択する
Public Sub 選択()
    m_ハンドル.Select
End Sub

' ワークシートを計算する
Public Sub 計算()
    m_ハンドル.Calculate
End Sub

' ========================================
' メソッド：コピー・削除
' ========================================

' ワークシートをコピーする
Public Sub コピー(Optional ByVal 前に_以下のシートの前 As Variant, Optional ByVal 後に_このシートの後 As Variant)
    If IsMissing(前に_以下のシートの前) And IsMissing(後に_このシートの後) Then
        m_ハンドル.Copy
    ElseIf Not IsMissing(前に_以下のシートの前) Then
        m_ハンドル.Copy before:=前に_以下のシートの前
    ElseIf Not IsMissing(後に_このシートの後) Then
        m_ハンドル.Copy after:=後に_このシートの後
    End If
End Sub

' ワークシートを削除する
Public Sub 削除()
    m_ハンドル.Delete
End Sub

' ワークシートを移動する
Public Sub 移動(Optional ByVal 前に_以下のシートの前 As Variant, Optional ByVal 後に_このシートの後 As Variant)
    If IsMissing(前に_以下のシートの前) And IsMissing(後に_このシートの後) Then
        m_ハンドル.Move
    ElseIf Not IsMissing(前に_以下のシートの前) Then
        m_ハンドル.Move before:=前に_以下のシートの前
    ElseIf Not IsMissing(後に_このシートの後) Then
        m_ハンドル.Move after:=後に_このシートの後
    End If
End Sub

' ========================================
' メソッド：印刷
' ========================================

' ワークシートを印刷する
Public Sub 印刷(Optional ByVal 開始ページ As Variant, Optional ByVal 終了ページ As Variant, Optional ByVal コピー数 As Variant = 1, Optional ByVal このようにプレビュー As Variant = False)
    Dim 引数 As String
    
    If Not IsMissing(開始ページ) Then
        m_ハンドル.PrintOut From:=開始ページ, To:=終了ページ, Copies:=コピー数, Preview:=このようにプレビュー
    Else
        m_ハンドル.PrintOut
    End If
End Sub

' 印刷プレビューを表示する
Public Sub 印刷プレビューを表示()
    m_ハンドル.PrintPreview
End Sub

' ========================================
' メソッド：保護
' ========================================

' ワークシートを保護する
Public Sub 保護(Optional ByVal パスワード As String = "", Optional ByVal 内容 As Boolean = True, Optional ByVal シナリオ As Boolean = False, Optional ByVal オブジェクト As Boolean = True)
    m_ハンドル.Protect password:=パスワード, contents:=内容, scenarios:=シナリオ, drawingObjects:=オブジェクト
End Sub

' ワークシートの保護を解除する
Public Sub 保護を解除(Optional ByVal パスワード As String = "")
    m_ハンドル.Unprotect password:=パスワード
End Sub

' ========================================
' メソッド：検索・置換
' ========================================

' スペルチェックを実行する
Public Sub スペルチェックする(Optional ByVal 言語 As Variant)
    If IsMissing(言語) Then
        m_ハンドル.CheckSpelling
    Else
        m_ハンドル.CheckSpelling CustomDictionary:=言語
    End If
End Sub

' ========================================
' メソッド：矢印
' ========================================

' 無効な計算セルを円で囲む
Public Sub 無効なセルを円で囲む()
    m_ハンドル.CircleInvalid
End Sub

' すべての矢印をクリアする
Public Sub すべての矢印をクリア()
    m_ハンドル.ClearArrows
End Sub

' すべての円をクリアする
Public Sub すべての円をクリア()
    m_ハンドル.ClearCircles
End Sub

' ========================================
' メソッド：データ操作
' ========================================

' すべてのデータを表示する
Public Sub すべてのデータを表示()
    m_ハンドル.ShowAllData
End Sub

' データフォームを表示する
Public Sub データフォームを表示()
    m_ハンドル.ShowDataForm
End Sub

' ========================================
' メソッド：ページ区切り
' ========================================

' すべてのページ区切りをリセットする
Public Sub すべてのページ区切りをリセット()
    m_ハンドル.ResetAllPageBreaks
End Sub

' ========================================
' コレクション
' ========================================

' ハイパーリンクコレクションを取得
Public Property Get ハイパーリンク() As Hyperlinks
    Set ハイパーリンク = m_ハンドル.Hyperlinks
End Property

' ========================================
' その他プロパティ
' ========================================

' 親オブジェクト(Workbook)を取得
Public Property Get 親() As Workbook
    Set 親 = m_ハンドル.Parent
End Property

' 前のワークシートを取得
Public Property Get 前のシート() As Worksheet
    If m_ハンドル.Previous Is Nothing Then
        Set 前のシート = Nothing
    Else
        Set 前のシート = m_ハンドル.Previous
    End If
End Property

' 次のワークシートを取得
Public Property Get 次のシート() As Worksheet
    If m_ハンドル.Next Is Nothing Then
        Set 次のシート = Nothing
    Else
        Set 次のシート = m_ハンドル.Next
    End If
End Property

' 基になるExcelオブジェクトを取得
Public Property Get ハンドル() As Worksheet
    Set ハンドル = m_ハンドル
End Property
