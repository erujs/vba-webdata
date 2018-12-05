Attribute VB_Name = "Module1"
' To learn more about getElements visit:
'   http://automatetheweb.net/vba-getelementsbytagname-method/
'
' Modules to add in Tools/References:
'   Microsoft Internet Controls
'   Microsoft HTML Object Library
Sub way()
    Dim i As Integer
    Dim o As Object
    
    'On Error GoTo handling
    'Continue to next line on error
    On Error Resume Next
    'Declare a new browser instance
    Dim ie As InternetExplorer: Set ie = New InternetExplorer
    
    Dim url As String: url = "https://www.w3schools.com/html/html_tables.asp"
    'Ctrl+G to open Immediate window
    Debug.Print url
    
    'IE running in background, set to True to make ie visible
    ie.Visible = False
    'Navigating to assigned webpage
    ie.navigate (url)
    
    Application.StatusBar = "Loading ..."
    'Waiting for webpage to load.
    Do While ie.Busy = True Or ie.readyState <> 4: Loop
    
    'Evaluating 'tr' elements data in the 'table' with id 'customers'
    For Each o In ie.document.getElementById("customers").getElementsByTagName("tr")
        i = i + 1
        'In the table each (table row) element has 3 data/children
        'Getting data content and assign value to specified cells.
        Sheets("Sheet1").Range("A" & i).Value = o.Children(0).textContent
        Sheets("Sheet1").Range("B" & i).Value = o.Children(1).textContent
        Sheets("Sheet1").Range("C" & i).Value = o.Children(2).textContent
    Next
    
    'Clean-up dimensions
    Set o = Nothing
    Set ie = Nothing
    Call clear_all
    Application.StatusBar = ""

'Optional handling method
'handling:
    'MsgBox "Error on InternetExplorer!"
End Sub
'Subroutine function that cleans browsing history
Sub clear_all()
    Shell "RunDll32.exe InetCpl.cpl, ClearMyTracksByProcess 255"
End Sub
