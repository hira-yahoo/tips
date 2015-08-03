Attribute VB_Name = "Module1"
Option Explicit

Sub tree()

    Dim fso As New FileSystemObject
    
    Dim d As Folder
    
'    Set d = fso.GetFolder("C:\work\memo\ì‹Æƒƒ‚\0803\macro\folder1")
    Set d = fso.GetFolder("C:\Users\hirayama\Google ƒhƒ‰ƒCƒu")
    
    Call read_folder(d, New ExternalInterface)
    

End Sub

Sub read_folder(d As Folder, i As ExternalInterface)

    Dim f As File
    Dim sub_d As Folder
    
    i.output (d.path)
    
    For Each sub_d In d.SubFolders
        Call read_folder(sub_d, i)
    Next
    
    For Each f In d.Files
        i.output (f.path)
    Next

End Sub

