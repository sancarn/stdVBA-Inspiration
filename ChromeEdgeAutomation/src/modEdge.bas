Attribute VB_Name = "modEdge"
Option Explicit

'This is anexample of how to use the classes
Sub runedge()

    'Start Browser
    Dim objBrowser As clsEdge
    Set objBrowser = New clsEdge
    Call objBrowser.start
    
    'Attach to any ("") or a specific page
    Call objBrowser.attach("")
    
    'navigate
    Call objBrowser.navigate("https://google.de")
    
    Call objBrowser.waitCompletion
    
    'evaluate javascript
    Call objBrowser.jsEval("alert(""hi"")")
    
    'fill search form (textbox is named q)
    Call objBrowser.jsEval("document.getElementsByName(""q"")[0].value=""automate edge vba""")
    
    'run search
    Call objBrowser.jsEval("document.getElementsByName(""q"")[0].form.submit()")
    
    'wait till search has finished
    Call objBrowser.waitCompletion
    

    'click on codeproject link
    Call objBrowser.jsEval("document.evaluate("".//h3[text()='Automate Chrome / Edge using VBA - CodeProject']"", document).iterateNext().click()")
    
    Call objBrowser.waitCompletion
    
    Dim strVotes As String
    strVotes = objBrowser.jsEval("ctl00_RateArticle_VountCountHist.innerText")
    
    MsgBox ("finish! Vote count is " & strVotes)
    
    objBrowser.closeBrowser
    
    
End Sub


'the following two snippets show the serialization of the object
Sub runedge2()

    'Start Browser
    Dim objBrowser As clsEdge
    Set objBrowser = New clsEdge
    Call objBrowser.start(True)
    
    'Attach to any ("") or a specific page
    Call objBrowser.attach("")
    
    'navigate
    Call objBrowser.navigate("https://google.de")
    
    'evaluate javascript
    Call objBrowser.jsEval("alert(""hi"")")
    
    MsgBox ("finish1!")
    
    Dim strSerialized As String
    strSerialized = objBrowser.serialize()
    Tabelle1.Cells(1, 1) = strSerialized
End Sub

Sub runedge3()
    
    Dim objBrowser2 As clsEdge
    Set objBrowser2 = New clsEdge
    
    
    Call objBrowser2.deserialize(Tabelle1.Cells(1, 1))
    
    If Not objBrowser2.connectionAlive Then Stop
    
    Call objBrowser2.jsEval("alert(""hi again"")")
   
    MsgBox ("finish2!")
    
    
End Sub


