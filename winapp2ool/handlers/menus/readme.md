# MenuMaker
**MenuMaker** is a singleton driver module for making dynamic finite state console applications with numbered menus

# Properties

| Name|Type|Default Value|Description|
|:-|:-|:-|:-|
|anyKeyStr|`String`|`"Press any key to return to the menu."` | An instruction to press any key to return to the previous menu|
invInpStr|`String`|`"Invalid input. Please try again."`|An error message informing the user their input was invalid|
promptStr|`String`|`"Enter a number, or leave blank to run the default: "`|An instruction for the user to provide input|
menuItemLength|`Integer`|`None`, but every winapp2ool module initalizes using `35` | The maximum length of the portion of the first half of a '#. Option - Description' style menu line|
ColorHeader|`Boolean`|`Nothing`, but modules initialize it with `False` by default | Indicates that the menu header should be printed with color|
HeaderColor|`ConsoleColor`|`Nothing`|The color with which the next header should be printed if `ColorHeader` is `True`|
OptNum|`Integer`|`0`|Holds the current option number for the menu instance|
SuppressOutput|`Boolean`|`False`|Indicates that the application should not output or ask input from the user except when encountering exceptions 
ExitCode|`Boolean`|`False`|Indicates that an exit from the current menu is pending 
MenuHeaderText|`String`|`Nothing`|Holds the text that appears in the top block of the menu|
Openers|`String()`|`{"║", "╔", "╚", "╠"}`|Frame characters used to open a menu line|
Closers|`String()`|`{"║", "╗", "╝", "╣"}`|Frame characters used to close a menu line


# Creating a menu

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
###### The general structure of a menu looks something like this
Creating this menu is very simple:
```vb
    Public Sub printMenu()
        setHeaderText("Menu Header Text Goes Here")
        printMenuTop({"Menu description items go here", "No fixed number of lines"})
        print(1, "Some First Option", "Execute some function or toggle some setting", closeMenu:=True)
```
### Some notes
* The `Console` class in Windows is *extremely* limited in terms of its capacity to show colors. The new Windows terminal has much better support for colors but isn't currently supported by **MenuMaker**
* Generated frames will be fit to the current console width as best they can be. 
* Lines that are longer than the current console width will not have a closing frame character 
* The above snippet calls `setHeaderText` but this is not necessary when instantiating menus through `initModule` (see below)

# Displaying menus
Menus are displayed and handled by the `initModule` subroutine 

Code
```vb
Public Sub initModule(name As String, showMenu As Action, handleInput As Action(Of String), Optional itmLen As Integer = 35)
        initMenu(name, itmLen)
        Try
            Do Until ExitCode
                clrConsole()
                showMenu()
                Console.Write(Environment.NewLine & promptStr)
                handleInput(Console.ReadLine)
            Loop
            ExitCode = False
            setHeaderText($"{name} closed")
        Catch ex As Exception
            exc(ex)
        End Try
    End Sub
```
###### 

| Parameter|Type|Description |Optional|
|:-|:-|:-|:-|
|Name|`String`|The name of the module as it will be displayed to the user|No
|showMenu|`Action`|The subroutine prints the module's menu|No
|handleInput|`Action(Of String)`|The subroutine that handles the module's input|No
|itmLen|`Integer`|Indicates the maximum length of menu option names|Yes

#### Notes
* The `Name` parameter is used to populate the menu header when `initMenu` is called and when its loop terminates. Beyond this, the module being initalized has control of the header at all times through the `setHeaderText` subroutine 
* Prompts the user for input using the value provided by the `promptStr` module property
* Exiting a menu returns the user exactly one level up in the stack

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
###### Selecting for case `input` is also a good option if the menu doesn't need to consider any further conditions, such as online status or the state of module parameters (see below)

## (Actually) Displaying Menus

Printing menus is handled entirely through the `print` subroutine. It takes many parameters, but most are optional flags whose effects are described below, and whose default values cause them to be disabled 

Code
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
###### Who needs ENUMS? 
| Parameter|Type|Description |Optional|
|:-|:-|:-|:-|
|printType|`Integer`|The type of menu information to print*|No
|menuText|`String`|The text to be printed **|No|
|optString|`String`|The description of the menu option|Yes, Default: ""
|cond|`Boolean`|Indicates that the line should be printed|Yes, Default: True
|leadingBlank|`Boolean`|Indicates that a blank menu line should be printed immediately before the printed line|Yes, Default: False
|trailingBlank|`Boolean`|Indicates that a blank menu line should be printed immediately after the printed line|Yes, Default: False
|isCentered|`Boolean`|Indicates that the printed text should be centered |Yes, Default: False
|closeMenu|`Boolean`|Indicates that the bottom menu frame should be printed|Yes, Default: False
|openMenu|`Boolean`|Indicates that the top menu frame should be printed|Yes, Default: False
|enStrCond|`Boolean`|A module setting whose menu text will include an Enable/Disable toggle ***|Yes, Default: False
|colorLine|`Boolean`|Indicates we want to color any lines we print|Yes, Default: False
|useArbitraryColor|`Boolean`|Indicates that the line should be colored using the value provided by `arbitraryColor`|Yes, Default: False
|buffr|`Boolean`|Indicates that a leading newline should be printed before the menu lines|Yes, Default: False
|trailr|`Boolean`|Indicates that a trailing newline should be printed after the menu lines|Yes, Default: False
|conjoin|`Boolean`|Indicates that a conjoining menu frame should be printed|Yes, Default: False

\* **PrintTypes**

|Integer Value|Result|
|:-|:-|
|0|Menu Line|
|1|Menu Option|
|2|Menu Option containing a "Reset Settings" prompt|
|3|Box with centered text|
|4|Menu Header|
|5|Menu Option containing an Enable/Disable toggle|

\** When `printType` is `1` or `5`, `menuText` contains the name of the menu option, when `printType` is `3`, `menuText` contains the name of the module whose settings will be reset  

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
