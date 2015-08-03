VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExternalInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private i As Long

Sub output(path As String)

'    Cells(i + 1, 1).Value = path
    
    ActiveSheet.Hyperlinks.Add Anchor:=Cells(i + 1, 1), Address:=path, _
        TextToDisplay:=path
    
'    MsgBox path & " " & i
    
    i = i + 1
    
End Sub
