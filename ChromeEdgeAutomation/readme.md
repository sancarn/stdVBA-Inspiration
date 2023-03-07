# Source Information

Author: [ChrisK23](https://www.codeproject.com/script/Membership/View.aspx?mid=7634041)
Source: [On Codeproject](https://www.codeproject.com/Tips/5307593/Automate-Chrome-Edge-using-VBA)
TAGS: CDP, Chrome, Dev, Protocol, Microsoft, Edge, IO, Pipes, CreateProcessA, 

# Automate Chrome / Edge using VBA

A method to automate Chrome (based) browsers using VBA

Microsoft Internet Explorer was fully scriptable using OLE Automation. This functionality is no longer available with the new Microsoft Edge browser. This tip presents a way to automate Edge and other Chrome based browsers using only VBA.

## Introduction

Internet Explorer classic (IE in the following) was based on ActiveX technology. It was very easy to automate IE for tasks like Webscraping or testing from OLE-aware programming languages like VBA. But Microsoft will end support for IE in the near future and wants users to move to newer browsers like Microsoft Edge.

Microsoft Edge is no longer based on ActiveX technology. Microsoft seems uninterested in creating a drop-in replacement for the IE OLE Object. There are libraries that try to fill this gap using Selenium, see Seleniumbasic as an example. But this requires the installation of a Webdriver, which might not be feasible in some environments. The following solution needs no additional software, apart from a Chrome-based browser.

Keep in mind, that all running Edge procceses must be terminated before running the code. Otherwise the tabs are opened in the currently running process, not the one that has been started and subsequent communication between VBA and Edge fails.


## CDP Protocol

The code uses the Chrome Devtools Protocol (CDP) to communicate with the browser. A full documentation of the protocol can be found here. The code implements only a very narrow set of functions:

1. Basic functions to set up the communication channel
2. Navigation to a url
3. Evaluate arbitrary JavaScript expressions in the context of a page and return the result

But these functions should suffice to do basic Webscraping. The main code is as follows:

```vb
'This is an example of how to use the classes
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
'if a javascript expression evaluates to a plain type it is passed back to VBA
    strVotes = objBrowser.jsEval("ctl00_RateArticle_VountCountHist.innerText")
    
    MsgBox ("finish! Vote count is " & strVotes)
    
    objBrowser.closeBrowser

    
End Sub
```

The class clsEdge implements the CDP protocol. The CDP protocol is a message-based protocol. Messages are encoded as JSON. To generate and parse JSON, the code uses the VBA-JSON library from here.

## Low-Level Communication with Pipes

The low-level access to the CDP protocol is avaible by two means: Either Edge starts a small Webserver on a specific port or via pipes. The Webserver lacks any security features. Any user on the computer has access to the webserver. This may pose no risks on single user computers or dedicated virtual containers. But if the process is run on a terminal server with more than one user, this is not acceptable. That's why the code uses pipes to communicate with Edge.

Edge uses the third file descriptor (fd) for reading messages and the fourth fd for writing messages. Passing fds from a parent process to child process is common under Unix, but not under Windows. The WinApi call to create a child process (`CreateProcess`) allows to setup pipes for the three common fds (stdin, stdout, stderr) using the STARTUPINFO structure, see [`CreateProcessA` function (processthreadsapi.h)](https://docs.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-createprocessa) and [`STARTUPINFOA` structure (processthreadsapi.h)](https://docs.microsoft.com/en-us/windows/win32/api/processthreadsapi/ns-processthreadsapi-startupinfoa). Other fds cannot be passed to the child process.

In order to set up the fourth and fifth fds, one must use an undocumented feature of the Microsoft Visual C Runtime (`MSVCRT`): If an application is compiled with Microsoft C, than one can pass the pipes using the `lpReserved2` parameter of the STARTUPINFO structure. See ["Undocumented CreateProcess"](http://www.catch22.net/tuts/undocumented-createprocess#) for more details (scroll down the page).

The structure that can be passed in `lpReserved2` is defined in the module modExec.

```vb
Public Type STDIO_BUFFER
    number_of_fds As Long
    crt_flags(0 To 4) As Byte
    os_handle(0 To 4) As LongPtr
End Type
```

The structure is defined to pass five fds in the os_handle array. The values for the crt_flags array can be obtained from [`libuv`](https://github.com/libuv/libuv/blob/v1.x/src/win/process-stdio.c). The fields of the struct must lie contiguously in memory (packed). VBA aligns struct fields to 4 byte boundaries (on 32-bit systems). That's why a second struct with raw types is defined.

```vb
Public Type STDIO_BUFFER2
    number_of_fds As Long
    raw_bytes(0 To 24) As Byte
End Type
```

After populating the `STDIO_BUFFER` struct, the content is copied using MoveMemory to the `STDIO_BUFFER2` struct. The size of 25 bytes is enought to hold `crt_flags` (5 bytes) and the pointers (20 bytes). 

## History

* 8th July, 2021: Initial version
* 18th August 2021, added support for 64bit Office
* 3rd November 2021, some minor improvements






