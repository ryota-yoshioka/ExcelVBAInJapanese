VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ファイル"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public Sub コピーする(ByVal コピー先 As String, Optional ByVal 上書きする As Boolean)
    インスタンス.Copy コピー先, 上書きする
End Sub

Public Sub 削除する(Optional ByVal 読み取り専用でも削除する As Boolean = False)
    インスタンス.Delete 読み取り専用でも削除する
End Sub

Public Sub 移動する(ByVal 移動先 As String)
    インスタンス.Move 移動先
End Sub

Public Function 文字列ストリームで開く(Optional ByVal モード As IOMode, Optional ByVal フォーマット As Tristate) As TextStream
    Dim ハンドル As TextStream
    Set ハンドル = インスタンス.OpenAsTextStream(モード, フォーマット)
    Dim ストリーム As 文字列ストリーム
    Set 文字列ストリームで開く = ストリーム.生成する(ハンドル)
End Function
