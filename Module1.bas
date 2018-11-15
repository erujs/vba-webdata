Attribute VB_Name = "Module1"
'Modules to add in Tools/References:
'   Microsoft Internet Controls
'   Microsoft HTML Object Library
Sub way()
    Dim i As Integer
    Dim o As Object
    Dim ie As Object
    
    Dim url As String
    
    url = "https://www.w3schools.com/html/html_tables.asp"
    'Ctrl+G to open Immediate window
    Debug.Print url
    
    'On Error GoTo handling
    On Error Resume Next
    Set ie = New internetexplorermedium
    
    'ie running in background, set to True to make ie visible
    ie.Visible = False
    
    ie.navigate (url)
    
    Application.StatusBar = "Loading ..."
    
    Do While ie.busy = True Or ie.readystate <> 4: Loop

    For Each o In ie.document.getElementsByTagName("table")(0).Children(0). _
        getElementsByTagName("tr")
        i = i + 1
        Sheets("Sheet1").Range("A" & i).Value = o.Children(0).textContent
        Sheets("Sheet1").Range("B" & i).Value = o.Children(1).textContent
        Sheets("Sheet1").Range("C" & i).Value = o.Children(2).textContent
    Next
    
    'Clean-up
    Set o = Nothing
    Set ie = Nothing
    'Call clear_all
    Application.StatusBar = ""

'handling:
    
End Sub

Sub clear_all()
    Shell "RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 255"
End Sub
