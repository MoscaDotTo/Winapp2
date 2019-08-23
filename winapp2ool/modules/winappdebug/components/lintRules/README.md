
# lintRule

lintRule is a component of **WinappDebug** and holds information about whether or not individual types of scans and/or repairs should run

# Properties

*italicized* properties are *private*

| Name|Type|Default Value|Description|
|:-|:-|:-|:-|
|ShouldScan|`Boolean`|`Nothing`|Indicates whether or not scans for this rule should run
|ShouldRepair|`Boolean`|`Nothing`|Indicates whether or not repairs for this rule should run
|ScanText|`String`|`Nothing`|Describes which type of scan routines this rule controls
|RepairText|`String`|`Nothing`|Describes which type of repairs this rule controls
|LintName|`String`|`Nothing`|The name of the rule as it will appear in menus
|*initScanState*|`Boolean`|`Nothing`|The instantiating (default) value of ShouldScan
|*initRepairState*|`Boolean`|`Nothing`|The instantiating (default) value of ShouldRepair
Lint rules are not instantiated with anything less than *all* the parameters that they require, so they do not carry default values.

# Creating a lint rule

## New

### Creates a new rule for the linter, retains the inital given parameters for later restoration

```vb
Public Sub New(scan As Boolean, repair As Boolean, name As String, scTxt As String, rpTxt As String)
    ShouldScan = scan
    initScanState = scan
    ShouldRepair = repair
    initRepairState = repair
    LintName = name
    ScanText = $"detecting {scTxt}"
    RepairText = rpTxt
End Sub
```

|Parameter|Type|Description|
|:-|:-|:-
|scan|`Boolean`|The default state for scans \*
|repair|`Boolean`|The default state for repairs \*
|name|`String`|The name of the rule as it will appear in menus
|scTxt|`String`|The description of what the rule scans for as it will appear in menus
|rpTxt|`String`|The description of what the the rule repairs as it will appear in menus

# Determining when a rule's repairs should run

## fixFormat

### Returns a Boolean indicating whether or not the repairs gated by this rule should be run

```vb
Public Function fixFormat() As Boolean
    Return RepairErrsFound Or (RepairSomeErrsFound And ShouldRepair)
End Function
```

# Determining when a rules current state is has been altered from its initial state

## hasBeenChanged

### Returns `True` if the current scan/repair settings do not match their inital states

```vb
Public Function hasBeenChanged() As Boolean
    Return Not ShouldScan = initScanState Or Not ShouldRepair = initRepairState
End Function
```

# Turning rules on or off

## turnOn

### Enables both the scan and repair for the rule

```vb
Public Sub turnOn()
    ShouldScan = True
    ShouldRepair = True
End Sub
```

## turnOff

### Disables both the scan and repair for the rule

```vb
Public Sub turnOff()
    ShouldScan = False
    ShouldRepair = False
End Sub
```
