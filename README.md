VBSH
==
Overview
--
This script provides [REPL](https://en.wikipedia.org/wiki/REPL "Read-Eval-Print-Loop") features to [VBScript](https://en.wikipedia.org/wiki/VBScript "Visual Basic Script") commands in a Windows shell prompt. It uses the [`Execute`](http://msdn.microsoft.com/en-us/library/03t418d2.aspx) statement in order to evaluate the commands passed to the prompt.

Features
--
* Loop detection: loop statements trigger multiline support
* [`Stop`](http://msdn.microsoft.com/en-us/library/zw86czy2.aspx) statement: to exit the script
* Automatic detection of output value

Requirements
--
For the time being, it only works with the console based [WSH](https://en.wikipedia.org/wiki/Windows_Script_Host), CScript that is. This script has solely been tested on Windows 7. It should work fine for pretty much all popular Windows releases.

Usage
--
Download and run it. For a greater deployment, move the script to a folder contained in the [PATH](https://en.wikipedia.org/wiki/Path_%28variable%29) environment variable.

Examples
--
Printing "Hello world"

    VBSH> "Hello world"
    Hello world

Summing numbers from 0 to 3

    VBSH> sum = 0
    VBSH> For i = 0 To 3
    ...       sum = sum + i
    ...   Next
    VBSH> sum
    6

Starting an Excel application and making it shown

    VBSH> Set oExcel = CreateObject("Excel.Application")
    VBSH> oExcel.Visible = True
