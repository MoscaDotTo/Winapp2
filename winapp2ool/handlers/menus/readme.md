
# MenuMaker

**MenuMaker** is a singleton driver module for making dynamic finite state console applications with numbered menus

# Properties

*italicized* properties are *private*

|Name|Type|Default Value|Description|
|:-|:-|:-|:-|
|anyKeyStr|`String`|`"Press any key to return to the menu."`|An instruction to press any key to return to the previous menu|
invInpStr|`String`|`"Invalid input. Please try again."`|An error message informing the user their input was invalid|
promptStr|`String`|`"Enter a number, or leave blank to run the default: "`|An instruction for the user to provide input|
menuItemLength|`Integer`|`None`, but every winapp2ool module initalizes using `35`|The maximum length of the `Name` half of a `#. Name - Description` style menu option|
ColorHeader|`Boolean`|`Nothing`, but modules initialize it with `False` by default|Indicates that the menu header should be printed with color|
HeaderColor|`ConsoleColor`|`Nothing`|The color with which the next header should be printed if `ColorHeader` is `True`|
SuppressOutput|`Boolean`|`False`|Indicates that the application should not output or ask input from the user except when encountering exceptions
ExitCode|`Boolean`|`False`|Indicates that an exit from the current menu is pending
MenuHeaderText|`String`|`Nothing`|The text that appears in the top block of the menu|
*OptNum*|`Integer`|`0`|The number associated with the next `Menu Option` that will be printed (if any)|
*Openers*|`String()`|`{"║", "╔", "╚", "╠"}`|Frame characters used to open a menu line|
*Closers*|`String()`|`{"║", "╗", "╝", "╣"}`|Frame characters used to close a menu line

# Creating a menu

The general structure of a menu looks something like this

```
 ╔══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╗
 ║                                                  Menu Header Text Goes Here                                              ║
 ╠══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╣
 ║                                                Menu description items go here                                            ║
 ║                                                   No fixed number of lines                                               ║
 ║                                                                                                                          ║
 ║                                                Menu: Enter a number to select                                            ║
 ║                                                                                                                          ║
 ║ 0. Exit                            - Return to the menu                                                                  ║
 ║ 1. Some First Option               - Execute some function or toggle some setting                                        ║
 ╚══════════════════════════════════════════════════════════════════════════════════════════════════════════════════════════╝

Enter a number, or leave blank to run the default:
 ```

Creating this menu is very simple:

```vb
Public Sub printMenu()
    setHeaderText("Menu Header Text Goes Here")
    printMenuTop({"Menu description items go here", "No fixed number of lines"})
    print(1, "Some First Option", "Execute some function or toggle some setting", closeMenu:=True)
```

## Notes

* The `Console` class in Windows is *extremely* limited in terms of its capacity to show colors. The new Windows terminal has much better support for colors but isn't currently supported by **MenuMaker**
* Generated frames will be fit to the current console width as best they can be.
* Lines that are longer than the current console width will not have a closing frame character
* The above snippet calls `setHeaderText` but this is not necessary when instantiating menus through `initModule` (see below)

# Displaying menus

Menus are displayed and handled by the `initModule` subroutine

## initModule

### Displays a menu to and passes the user's input over to be handled until the exit command is given

Exiting a menu returns exactly one level up in the stack to the menu that called it
Effectively the main event loop for anything built with **MenuMaker**

```vb
Public Sub initModule(name As String, showMenu As Action, handleInput As Action(Of String), Optional itmLen As Integer = 35)
    ExitPending = False
    setHeaderText(name)
    menuItemLength = itmLen
    Do Until ExitPending
        clrConsole()
        showMenu()
        Console.Write(Environment.NewLine & promptStr)
        handleInput(Console.ReadLine)
    Loop
    ExitPending = False
    setHeaderText($"{name} closed")
End Sub
```

|Parameter|Type|Description|Optional|
|:-|:-|:-|:-|
|Name|`String`|The name of the module as it will be displayed to the user|No
|showMenu|`Action`|The subroutine that prints the module's menu|No
|handleInput|`Action(Of String)`|The subroutine that handles the module's input|No
|itmLen|`Integer`|Indicates the maximum length of menu option names|Yes, default: `35`

## Of Note

* The `Name` parameter is used to populate the menu header when `initMenu` is called and when its loop terminates. Beyond this, the module being initalized has control of the header at all times through the `setHeaderText` subroutine
* Prompts the user for input using the value provided by the `promptStr` module property
* Exiting a menu returns the user exactly one level up in the stack to the menu that called it
* Applications (generally) terminate when the last level of the stack is exited, but any post-run code you may wish to execute can be placed after the very first `initModule` call in your `main`

## Handling input

`initMenu` requires that you provide it with a subroutine that accepts a `String` value to handle the user's input. A snippet for this typically begins like this

```vb
Public Sub handleInput(input As String)
    Select Case True
        Case input = "0"
            exitModule()
        Case input = "1"
            ' HANDLE THE FIRST OPTION
        Case else
            setHeaderText(invInpStr,True)
    End Select
End Sub
```

Selecting for case `input` is also a good option if the menu doesn't need to consider any further conditions, such as online status or the state of module parameters (see below)

## (Actually) Displaying Menus

Printing menus is handled entirely through the `print` subroutine. It takes many parameters, but most are optional flags whose effects are described below, and whose default values cause them to be disabled

```vb
Public Sub print(printType As Integer, menuText As String, Optional optString As String = "", Optional cond As Boolean = True,
                    Optional leadingBlank As Boolean = False, Optional trailingBlank As Boolean = False, Optional isCentered As Boolean = False,
                    Optional closeMenu As Boolean = False, Optional openMenu As Boolean = False, Optional enStrCond As Boolean = False,
                    Optional colorLine As Boolean = False, Optional useArbitraryColor As Boolean = False, Optional arbitraryColor As ConsoleColor = Nothing,
                    Optional buffr As Boolean = False, Optional trailr As Boolean = False, Optional conjoin As Boolean = False)
    If Not cond Then Return
    cwl(cond:=buffr)
    If colorLine Then Console.ForegroundColor = If(useArbitraryColor, arbitraryColor, If(enStrCond, ConsoleColor.Green, ConsoleColor.Red))
    print(0, Nothing, cond:=leadingBlank)
    print(0, getFrame(1), cond:=openMenu)
    Select Case printType
        ' Prints lines
        Case 0
            printMenuLine(menuText, isCentered)
        ' Prints options
        Case 1
            printMenuOpt(menuText, optString)
        ' Prints the Reset Settings option
        Case 2
            print(1, "Reset Settings", $"Restore {menuText}'s settings to their default state", leadingBlank:=True)
        ' Prints a box with centered text
        Case 3
            print(4, menuText, closeMenu:=True)
        ' The top of a menu with a header
        Case 4
            print(0, menuText, isCentered:=True, openMenu:=True)
        ' Colored line printing for enable/disable menu options
        Case 5
            print(1, menuText, $"{enStr(enStrCond)} {optString}", colorLine:=True, enStrCond:=enStrCond)
    End Select
    print(0, getFrame(3), cond:=conjoin)
    print(0, Nothing, cond:=trailingBlank)
    print(0, getFrame(2), cond:=closeMenu)
    If colorLine Then Console.ResetColor()
    cwl(cond:=trailr)
End Sub
```

 Who needs ENUMS?

|Parameter|Type|Description|Optional|
|:-|:-|:-|:-|
|printType|`Integer`|The type of menu information to print \*|No
|menuText|`String`|The text to be printed \*\*|No|
|optString|`String`|The description of the menu option|Yes, default: `""`
|cond|`Boolean`|Indicates that the line should be printed|Yes, default: `True`
|leadingBlank|`Boolean`|Indicates that a blank menu line should be printed immediately before the printed lines|Yes, default: `False`
|trailingBlank|`Boolean`|Indicates that a blank menu line should be printed immediately after the printed lines|Yes, default: `False`
|isCentered|`Boolean`|Indicates that the printed text should be centered|Yes, default: `False`
|closeMenu|`Boolean`|Indicates that the bottom menu frame should be printed|Yes, default: `False`
|openMenu|`Boolean`|Indicates that the top menu frame should be printed|Yes, default: `False`
|enStrCond|`Boolean`|A module setting whose menu text will include an Enable/Disable toggle \*\*\*|Yes, default: `False`
|colorLine|`Boolean`|Indicates that lines should be printed using color|Yes, default: `False`
|useArbitraryColor|`Boolean`|Indicates that the line should be colored using the value provided by `arbitraryColor`|Yes, default: `False`
|arbitraryColor|`ConsoleColor`|Foreground `ConsoleColor` to be used when printing with color but using a color other than `Red` or `Green`|Yes, default: `False`
|buffr|`Boolean`|Indicates that a leading newline should be printed before the menu  lines|Yes, default: `False`
|trailr|`Boolean`|Indicates that a trailing newline should be printed after the menu lines|Yes, default: `False`
|conjoin|`Boolean`|Indicates that a conjoining menu frame should be printed after the printed lines|Yes, default: `False`

\* **PrintTypes**

|Integer Value|Result|
|:-|:-|
|0|Menu Line|
|1|Menu Option|
|2|Menu Option containing a "Reset Settings" prompt|
|3|Box with centered text|
|4|Menu Header|
|5|Menu Option containing an Enable/Disable toggle|

\** When `printType` is `1` or `5`, `menuText` contains the name of the menu option, when `printType` is `3`, `menuText` contains the name of the module whose settings are being reset

\*** If `colorLine` is `True`and `useArbitraryColor` is `False`, the line will be printed based on the value of `enStrCond`. This is `ConsoleColor.Green` when it is `True` and `ConsoleColor.Red` when it is False. The default value is `False` and as a result, the default color in this case will be red.

# Example Program

 The following code can be used alongside MenuMaker.vb to create a simple demo menu

 ```vb
 Module Module1
    Property toggleBool As Boolean = False
    Property ModuleSettingsChanged = False

    Sub Main()
        initModule("Menu Maker Demo", AddressOf printMenu, AddressOf handleInput)
    End Sub

    Sub printMenu()
        printMenuTop({"This menu uses all the different types of menu frames", "that MenuMaker can provide"})
        print(1, "Run (Default)", "Execute the function gated by ToggleBool")
        print(5, "ToggleBool", "the gating parameter from the run option", closeMenu:=Not ModuleSettingsChanged, enStrCond:=toggleBool)
        print(2, "Menu Maker Demo", closeMenu:=True, cond:=ModuleSettingsChanged)
    End Sub

    Sub handleInput(input As String)
        Select Case True
            Case input = "0"
                exitModule()
            Case input = "1" Or input = ""
                If Not denyActionWithHeader(Not toggleBool, "ToggleBool must be enabled to run") Then setHeaderText("Run Complete", True, printColor:=ConsoleColor.Green)
            Case input = "2"
                toggleBool = Not toggleBool
                ModuleSettingsChanged = True
            Case input = "3" And ModuleSettingsChanged
                toggleBool = False
                ModuleSettingsChanged = False
                setHeaderText("Menu Maker Demo settings restored to their defaults", True, printColor:=ConsoleColor.Green)
            Case Else
                setHeaderText(invInpStr, True)
        End Select
    End Sub
End Module
```

# Helper Functions

MenuMaker has some helper functions that should help format menu content

## denyActionWithHeader

### Informs a user when an action is unable to proceed due to a condition

```vb
Public Function denyActionWithHeader(cond As Boolean, errText As String) As Boolean
    setHeaderText(errText, True, cond)
    Return cond
End Function
```

|Parameter|Type|Description|
|:-|:-|:-|
|cond|`Boolean`|Indicates that an action should be denied|
|errText|`String`|The error text to be printed in the menu header

As demonstrated in the example program above, this function can be used to gate execution of a function behind a `Boolean` state. It accepts the condition that, when `True` should prevent further execution (eg. `toggleBool = False`). It returns the `Boolean` it is fed, and if `True`, sets the menu header's text to a given error string.

## enStr

### Returns the inverse state of a given boolean as a String, ie. `"Disable"` if `setting` is `True`, `"Enable"` otherwise. Useful for creating toggles

```vb
Public Function enStr(setting As Boolean) As String
    Return If(setting, "Disable", "Enable")
End Function
 ```

|Parameter|Type|Description|
|:-|:-|:-|
|setting|`Boolean`|A module setting whose state will be observed

## printMenuTop

### Prints the top of the menu, the header, a conjoiner, any description text provided, the menu prompt, and the exit option (optional)

```vb
Public Sub printMenuTop(descriptionItems As String(), Optional printExit As Boolean = True)
    print(4, MenuHeaderText, colorLine:=ColorHeader, useArbitraryColor:=True, arbitraryColor:=HeaderColor, conjoin:=True)
    For Each line In descriptionItems
        print(0, line, isCentered:=True)
    Next
    print(0, "Menu: Enter a number to select", leadingBlank:=True, trailingBlank:=True, isCentered:=True)
    OptNum = 0
    print(1, "Exit", "Return to the menu", printExit)
End Sub
```

|Parameter|Type|Description|Optional|
|:-|:-|:-|:-
|descriptionItems|`String()`|Text describing the current menu or module functions being presented to the user, each array will be displayed on a separate line|No
|printExit|`Boolean`|Indicates that an option to exit to the previous menu should be printed|Yes, default: `True`

## clrConsole

### Clears the console if `cond` is `True`, `SuppressOutput` is `False`, and the caller is not an x86 unit testing console window

```vb
Public Sub clrConsole(Optional cond As Boolean = True)
    If cond And Not SuppressOutput And Not Console.Title.Contains("testhost.x86") Then Console.Clear()
End Sub
```

|Parameter|Type|Description|Optional|
|:-|:-|:-|:-
cond|`Boolean`|Indicates that the console should be cleared|Yes, default: `True`

## cwl

### Prints a line to the console window if output is not currently being suppressed and the given `cond` is met

```vb
Public Sub cwl(Optional msg As String = Nothing, Optional cond As Boolean = True)
    If cond And Not SuppressOutput Then Console.WriteLine(msg)
End Sub
```

|Parameter|Type|Description|Optional|
|:-|:-|:-|:-
msg|`String`|The line to be printed|Yes, default: `Nothing`
cond|`Boolean`|Indicates the line should be printed|Yes, default: `True`

## crk

### Waits for the user to press a key if output is not currently being suppressed

```vb
Public Sub crk()
    If Not SuppressOutput Then Console.ReadKey()
End Sub
```

## setHeaderText

### Saves a menu header to be printed atop the next menu, optionally with color

```vb
Public Sub setHeaderText(txt As String, Optional cHeader As Boolean = False, Optional cond As Boolean = True, Optional printColor As ConsoleColor = ConsoleColor.Red)
    If Not cond Then Return
    MenuHeaderText = txt
    ColorHeader = cHeader
    HeaderColor = printColor
End Sub
```

|Parameter|Type|Description|Optional|
|:-|:-|:-|:-
txt|`String`|The text to appear in the header|No
cHeader|`Boolean`|Indicates that the header should be colored using the color given by `printColor`|Yes, default: `False`|
cond|`Boolean`|Indicates that the header text should be set|Yes, default: `True`
printColor|`ConsoleColor`|`ConsoleColor` with which the header should be colored when `cHeader` is `True`|Yes, default: `ConsoleColor.Red`

## replDir

### Replaces instances of the current directory in a path string with `".."`

```vb
Public Function replDir(dirStr As String) As String
    Return dirStr.Replace(Environment.CurrentDirectory, "..")
End Function
```

|Parameter|Type|Description
|:-|:-|:-
dirStr|`String`|A windows filesystem path

# Private members

The following are private member subroutines and functions of MenuMaker

## getFrame

### Returns an empty menu line, or a variety of filled menu lines

```vb
Private Function getFrame(Optional frameNum As Integer = 0) As String
    Return mkMenuLine("", 2, frameNum)
End Function
```

|Parameter|Type|Description|Optional|
|:-|:-|:-|:-
frameNum|`Integer`|Indicates which frame should be returned \*|Yes, default: `0`

\* frameNums:

|frameNum|Frame Type|
|:-|:-|
|`0`|Empty menu line with vertical frames `║     ║`
|`1`|Filled menu line with downward opening 90° angle frames `╔═════╗`
|`2`|Filled menu line with upward opening 90° angle frames `╚═════╝`
|`3`|Filled menu line with inward facing T-frames `╠═════╣`

## printMenuLine

### Prints a line bounded by vertical menu frames, or an empty menu line if `lineString` is `Nothing`

```vb
Private Sub printMenuLine(Optional lineString As String = Nothing, Optional isCentered As Boolean = False)
    If lineString = Nothing Then lineString = getFrame()
    cwl(mkMenuLine(lineString, If(isCentered, 0, 1)))
End Sub
```

|Parameter|Type|Description|Optional|
|:-|:-|:-|:-
lineString|`String`|The text to be printed|Yes, default: `Nothing`
isCentered|`Boolean`|Indicates that the printed text should be centered|Yes, default: `False`

## printMenuOpt

### Prints a numbered menu option after padding it to a set length

```vb
Private Sub printMenuOpt(lineString1 As String, lineString2 As String)
    lineString1 = $"{OptNum}. {lineString1}"
    padToEnd(lineString1, menuItemLength, "")
    cwl(mkMenuLine($"{lineString1}- {lineString2}", 1))
    OptNum += 1
End Sub
```

|Parameter|Type|Description
|:-|:-|:-
|lineString1|`String`|The name of the menu option
|lineString2|`String`|The description of the menu option

## mkMenuLine

### Constructs a menu line fit to the width of the console

```vb
Private Function mkMenuLine(line As String, align As Integer, Optional borderInd As Integer = 0) As String
    If line.Length >= Console.WindowWidth - 1 Then Return line
    Dim out = $" {Openers(borderInd)}"
    Select Case align
        Case 0
            padToEnd(out, CInt((((Console.WindowWidth - line.Length) / 2) + 2)), Closers(borderInd))
            out += line
            padToEnd(out, Console.WindowWidth - 2, Closers(borderInd))
        Case 1
            out += " " & line
            padToEnd(out, Console.WindowWidth - 2, Closers(borderInd))
        Case 2
            padToEnd(out, Console.WindowWidth - 2, Closers(borderInd), If(borderInd = 0, " ", "═"))
    End Select
    Return out
End Function
```

|Parameter|Type|Description|Optional|
|:-|:-|:-|:-
line|`String`|The text to be printed|No
align|`Integer`|The alignment of the line to be printed *|No
borderInd|`Integer`|Determines which characters should create the border for the menuline **|Yes, default: `0`

\* aligns

|align|Effect|
|:-|:-|
|`0`|Centers the string
|`1`|Aligns the string to the left side of the screen
|`2`|Creates a menu frame

** `borderInd` refers to the index within the `openers` and `closers` property for the desired frame character

## padToEnd

### Pads a given string with spaces until it is a target length

```vb
Private Sub padToEnd(ByRef out As String, targetLen As Integer, endline As String, Optional padChar As String = " ")
    While out.Length < targetLen
        out += padChar
    End While
    If targetLen = Console.WindowWidth - 2 Then out += endline
End Sub
```

|Parameter|Type|Description|Optional|
|:-|:-|:-|:-|
out|`String`|The text to be padded|No|
targetLen|`Integer`|The length to which the text should be padded|No|
endLine|`String`|The closer character for the type of frame being built|No|
padChar|`String`|The single length String with which to pad the text|Yes, default: `" "`|
