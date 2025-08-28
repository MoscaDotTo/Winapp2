Public Module lintsettings

    ''' <summary> 
    ''' The winapp2.ini file that will be linted 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property winappDebugFile1 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", mExist:=True)

    ''' <summary> 
    ''' The save path for the linted file. Overwrites the input file by default 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property winappDebugFile3 As New iniFile(Environment.CurrentDirectory, "winapp2.ini", "winapp2-debugged.ini")

    ''' <summary> 
    ''' Indicates that some but not all repairs will run 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property RepairSomeErrsFound As Boolean = False

    ''' <summary>
    ''' Indicates that the scan settings have been modified from their defaults 
    ''' <br /> Default: <c> False </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property ScanSettingsChanged As Boolean = False

    ''' <summary>
    ''' Indicates that the module settings have been modified from their defaults 
    ''' <br/> Default: <c> False </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property LintModuleSettingsChanged As Boolean = False

    ''' <summary> 
    ''' Indicates that the any changes made by the linter should be saved back to disk 
    ''' <br/> Default: <c> False </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property SaveChanges As Boolean = False

    ''' <summary> 
    ''' Indicates that the linter should attempt to repair errors it finds 
    ''' <br/> Default: <c> True </c>
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property RepairErrsFound As Boolean = True

    ''' <summary> 
    ''' Indicates that Default keys should have their values auited instead of being considered invalid for existing 
    ''' <br /> Default: <c> False </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property overrideDefaultVal As Boolean = False

    ''' <summary> 
    ''' The expected value for Default keys when auditing their values 
    ''' <br/> Default: <c> Faalse </c> 
    ''' </summary>
    ''' 
    ''' Docs last updated: 2021-11-13 | Code last updated: 2021-11-13
    Public Property expectedDefaultValue As Boolean = False

End Module
