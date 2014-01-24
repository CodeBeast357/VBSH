VBSH
==
Overview
--
This script provides <abbr title="Read-Eval-Print-Loop">REPL</abbr> features to <abbr title="Visual Basic Script">VBS</abbr> commands in a Windows shell prompt. It uses the [`Execute`](http://msdn.microsoft.com/en-us/library/03t418d2.aspx) statement in order to evaluate the commands passed to the prompt.

Features
--
* Scope separation: simply use `VBSH` within
* Loop detection: loop statements trigger multiline support
* [`Stop`](http://msdn.microsoft.com/en-us/library/zw86czy2.aspx) statement: to exit the script

Requirements
--
This script has solely been tested on Windows 7. It should work fine for pretty much all popular Windows releases.

Usage
--
Download and call it from a prompt. For a greater deployment, move the script to a folder contained in the [PATH](https://en.wikipedia.org/wiki/Path_%28variable%29) environment variable.

Examples
--
Printing "Hello world"

    VBSH> WScript.Echo "Hello world"
    Hello world

Summing numbers from 0 to 3

    VBSH> sum = 0
    
    VBSH> For i = 0 To 3
    ...       sum = sum + i
    ...   Next
    
    VBSH> WScript.Echo sum
    6

Starting an Excel application and making it shown

    VBSH> Set oExcel = CreateObject("Excel.Application")
    
    VBSH> oExcel.Visible = True
