
# Individual Scan/Repair state management

This page covers the code aspect of how WinappDebug's scan/repair settings are managed under the hood. If you are looking for documentation on the scan settings themselves, please see the  **WinappDebug** readme

# Menu & Input

## printMenu

### Prints the scan/repair management menu to the user

```vb
Public Sub printMenu()
    Console.WindowHeight = 49
    printMenuTop({"Enable or disable specific scans or repairs"})
    print(0, "Scan Options", leadingBlank:=True, trailingBlank:=True)
    Rules.ForEach(Sub(rule) print(5, rule.LintName, rule.ScanText, enStrCond:=rule.ShouldScan))
    ' Print all repairs except the last one
    print(0, "Repair Options", leadingBlank:=True, trailingBlank:=True)
    For i = 0 To Rules.Count - 2
        Dim rule = Rules(i)
        print(5, rule.LintName, rule.RepairText, enStrCond:=rule.ShouldRepair)
    Next
    ' Special case for the last repair option (closemenu flag)
    Dim lastRule = Rules.Last
    print(5, lastRule.LintName, lastRule.RepairText, closeMenu:=Not ScanSettingsChanged, enStrCond:=lastRule.ShouldRepair)
    print(2, "Scan And Repair", cond:=ScanSettingsChanged, closeMenu:=True)
End Sub
```

## handleUserInput

### Handles the user input for the scan/repair management menu

```vb
Public Sub handleUserInput(input As String)
    ' Determine the current state of the lint rules
    determineScanSettings()
    ' Get the input as an integer so we can index it against our rules
    Dim intInput = -1
    Integer.TryParse(input, intInput)
    ' The index of the rule assoicated with the user's input
    Dim ind = intInput - 1
    Select Case True
        Case input = "0"
            If ScanSettingsChanged Then WinappDebug.ModuleSettingsChanged = True
            exitModule()
        ' Enable/Disable individual scans
        Case intInput > 0 And intInput <= Rules.Count
            toggleSettingParam(Rules(ind).ShouldScan, "Scan", ScanSettingsChanged)
            ' Force repair off if the scan is off
            If Not Rules(ind).ShouldScan Then Rules(ind).turnOff()
        ' Enable/Disable individual repairs
        Case intInput > Rules.Count And intInput <= 2 * Rules.Count
            ind -= (Rules.Count)
            toggleSettingParam(Rules(ind).ShouldRepair, "Repair", ScanSettingsChanged)
            ' Force scan on if the repair is on
            If Rules(ind).ShouldRepair Then Rules(ind).turnOn()
        Case intInput = 2 * Rules.Count + 1 And ScanSettingsChanged
            resetScanSettings()
            setHeaderText("Settings Reset")
        ' This isn't documented anywhere and is mostly intended as a debugging shortcut
        Case input = "alloff"
            Rules.ForEach(Sub(rule) rule.turnOff())
            ScanSettingsChanged = True
        Case Else
            setHeaderText(invInpStr, True)
    End Select
End Sub
```

|Parameter|Type|Description|
|:-|:-|:-
input|`String`|The string containing the user's input

# Determining whether or not scans settings have changed

## determineScanSettings

### Determines which if any lint rules have been modified and whether or not only some repairs are scheduled to run

```vb
Private Sub determineScanSettings()
    Dim repairAll = True
    Dim repairAny = False
    For Each rule In Rules
        If rule.hasBeenChanged Then
            ScanSettingsChanged = True
            If Not rule.ShouldRepair Then repairAll = False
        End If
        If rule.ShouldRepair Then repairAny = True
    Next
    If Not repairAll And repairAny Then
        RepairErrsFound = False
        RepairSomeErrsFound = True
    End If
End Sub
```

# Undoing changes

## resetScanSettings

### Resets the individual scan/repair settings to their defaults

```vb
Public Sub resetScanSettings()
    Rules.ForEach(Sub(rule) rule.resetParams())
    ScanSettingsChanged = False
    RepairSomeErrsFound = False
End Sub
```
