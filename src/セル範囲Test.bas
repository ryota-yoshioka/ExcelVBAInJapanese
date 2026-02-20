Attribute VB_Name = "セル範囲Test"
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Uncategorized")
Private Sub TestインデントGet()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ハンドル As Range
    Set ハンドル = Worksheets("Sheet1").Range("A1:A1")
    Dim ラッパー As New セル範囲
    ラッパー.初期化する ハンドル
    
    'Act:
    'Assert:
    ラッパー.インデント = True
    Assert.AreEqual ハンドル.AddIndent, True
    ラッパー.インデント = False
    Assert.AreEqual ハンドル.AddIndent, False
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Test値Let()
    On Error GoTo TestFail
    
    'Arrange:
    Dim ハンドル As Range
    Set ハンドル = Worksheets("Sheet1").Range("A1:A1")
    Dim ラッパー As New セル範囲
    ラッパー.初期化する ハンドル
    
    'Act:
    'Assert:
    ラッパー.値 = "値"
    Assert.AreEqual ハンドル.value, "値"
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Uncategorized")
Private Sub Testオフセットした範囲を取得する()
    On Error GoTo TestFail
    
    'Arrange:
    Dim 期待値 As Range
    Set 期待値 = Worksheets("Sheet1").Range("A1:A1")
    Dim ラッパー As New セル範囲
    ラッパー.初期化する 期待値
    
    'Act:
    Dim 実際値 As セル範囲
    Set 実際値 = ラッパー.オフセットした範囲を取得する(0, 0)
    
    'Assert:
    Assert.AreEqual 期待値.Address, 実際値.アドレス
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub









