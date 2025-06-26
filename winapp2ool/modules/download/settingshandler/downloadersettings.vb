Public Module downloadersettings

    '''<summary> 
    '''Holds the path of any files to be saved by the Downloader 
    '''</summary>
    '''
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Property downloadFile As iniFile = New iniFile(Environment.CurrentDirectory, "")

    ''' <summary> 
    ''' Indicates that the Downloader module's settings have been changed from their defaults 
    '''</summary>
    '''
    ''' Docs last updated: 2020-09-14 | Code last updated: 2020-09-14
    Public Property DownloadModuleSettingsChanged As Boolean = False

End Module
