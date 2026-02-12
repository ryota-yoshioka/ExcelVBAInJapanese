VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "文字列ストリーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'@Folder("ファイルシステム")

Public インスタンス As TextStream

Public Function 生成する(ByVal ハンドル As TextStream) As 文字列ストリーム
    Set 生成する = New 文字列ストリーム
    Set 生成する.インスタンス = ハンドル
End Function
Public Function 初期化する(ByVal ハンドル As TextStream) As 文字列ストリーム
    Set 初期化する = Me
    初期化する.インスタンス = ハンドル
End Function

Public Property Get 行末である() As Boolean
    行末である = インスタンス.AtEndOfLine
End Property

Public Property Get ストリームの最後である() As Boolean
    ストリームの最後である = インスタンス.AtEndOfStream
End Property

Public Property Get 桁() As Long
    桁 = インスタンス.Column
End Property

Public Property Get 行() As Long
    行 = インスタンス.Line
End Property

Public Sub 閉じる()
    インスタンス.Close
End Sub

Public Function 読む(ByVal 文字数 As Long) As String
    読む = インスタンス.Read(文字数)
End Function

Public Function 全て読む() As String
    全て読む = インスタンス.ReadAll()
End Function

Public Function 一行読む() As String
    一行読む = インスタンス.ReadLine()
End Function

Public Sub スキップする(ByVal 文字数 As Long)
    インスタンス.Skip (文字数)
End Sub

Public Sub 一行スキップする()
    インスタンス.SkipLine
End Sub

Public Sub 書く(ByVal 文字列 As String)
    インスタンス.Write (文字列)
End Sub

Public Sub 空行を書く(ByVal 行数 As Long)
    インスタンス.WriteBlankLines (行数)
End Sub

Public Sub 一行書く(Optional ByVal 文字列 As String = vbCrLf)
    インスタンス.WriteLine (文字列)
End Sub
