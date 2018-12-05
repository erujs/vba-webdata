Attribute VB_Name = "Module1"
'Modules to add in Tools/References:
'   Microsoft Internet Controls
'   Microsoft HTML Object Library
Sub way()
    Dim i As Integer
    Dim o As Object
    
    'On Error GoTo handling
    On Error Resume Next
    Dim ie As InternetExplorer: Set ie = New InternetExplorer
    
    Dim url As String
    
    url = "https://www.w3schools.com/html/html_tables.asp"
    'Ctrl+G to open Immediate window
    Debug.Print url
    
    'ie running in background, set to True to make ie visible
    ie.Visible = False
    
    ie.navigate (url)
    
    Application.StatusBar = "Loading ..."
    
    Do While ie.Busy = True Or ie.readyState <> 4: Loop
    
    'Scraping data using getElement function
    For Each o In ie.document.getElementById("customers").getElementsByTagName("tr")
        i = i + 1
        Sheets("Sheet1").Range("A" & i).Value = o.Children(0).textContent
        Sheets("Sheet1").Range("B" & i).Value = o.Children(1).textContent
        Sheets("Sheet1").Range("C" & i).Value = o.Children(2).textContent
        'Sheets("Sheet1").Range("D" & i).Value = o.Children(3).textContent
    Next
    
    'Clean-up dimensions
    Set o = Nothing
    Set ie = Nothing
    Call clear_all
    Application.StatusBar = ""

'handling:
    MsgBox "Error on InternetExplorer!"
End Sub
'Subroutine function that cleans browsing history
Sub clear_all()
    Shell "RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 255"
End Sub
