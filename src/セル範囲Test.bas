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
    Dim handle As Range
    Set handle = Worksheets("Sheet1").Range("A1:A1")
    Dim target As New セル範囲
    target.初期化 handle
    
    'Act:
    'Assert:
    target.インデント = True
    Assert.AreEqual handle.AddIndent, True
    target.インデント = False
    Assert.AreEqual handle.AddIndent, False
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
