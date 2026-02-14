VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ファイル"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("ファイルシステム")
Option Explicit

Private インスタンス As file

Public Function 生成する(ByVal ハンドル As file) As ファイル
    Set 生成する = New ファイル
    Call 生成する.初期化する(ハンドル)
End Function
Public Function 初期化する(ByVal ハンドル As file) As ファイル
    Set 初期化する = Me
    Set インスタンス = ハンドル
End Function

Public Property Get 属性() As Long
    属性 = インスタンス.Attributes
End Property

Public Property Let 属性(ByVal 設定値 As Long)
    インスタンス.Attributes = 設定値
End Property

Public Property Get 作成日時() As Date
    作成日時 = インスタンス.DateCreated
End Property

Public Property Get 最終アクセス日時() As Date
    最終アクセス日時 = インスタンス.DateLastAccessed
End Property

Public Property Get 最終更新日時() As Date
    最終更新日時 = インスタンス.DateLastModified
End Property

Public Property Get ドライブ() As String
    ドライブ = インスタンス.Drive
End Property

Public Property Get 名前() As String
    名前 = インスタンス.Name
End Property

Public Property Let 名前(ByVal 設定値 As String)
    インスタンス.Name = 設定値
End Property

Public Property Get 親フォルダ() As String
    親フォルダ = インスタンス.ParentFolder
End Property

Public Property Get パス() As String
    パス = インスタンス.Path
End Property

Public Property Get 短縮名() As String
    短縮名 = インスタンス.ShortName
End Property

Public Property Get 短縮パス() As String
    短縮パス = インスタンス.ShortName
End Property

Public Property Get サイズ() As Long
    サイズ = インスタンス.Size
End Property

Public Property Get 種類() As String
    種類 = インスタンス.Type
End Property

Public Sub コピーする(ByVal コピー先 As String, Optional ByVal 上書きする As Boolean)
    インスタンス.Copy コピー先, 上書きする
End Sub

Public Sub 削除する(Optional ByVal 読み取り専用でも削除 As Boolean = False)
    インスタンス.Delete 読み取り専用でも削除
End Sub

Public Sub 移動する(ByVal 移動先 As String)
    インスタンス.Move 移動先
End Sub

Public Function 文字列ストリームで開く(Optional ByVal IOモード As IOMode, Optional ByVal フォーマット As Tristate) As 文字列ストリーム
    Dim TextStreamのハンドル As TextStream
    Set TextStreamのハンドル = インスタンス.OpenAsTextStream(IOモード, フォーマット)
    Dim 文字列ストリームのインスタンス As New 文字列ストリーム
    Set 文字列ストリームで開く = 文字列ストリームのインスタンス.初期化する(TextStreamのハンドル)
End Function
