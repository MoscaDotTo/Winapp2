Option Strict On
Imports Microsoft.Win32
Module RegistryHelper
    Public Function getLMKey(Optional subkey As String = "") As RegistryKey
        Return Registry.LocalMachine.OpenSubKey(subkey)
    End Function

    Public Function getCRKey(Optional subkey As String = "") As RegistryKey
        Return Registry.ClassesRoot.OpenSubKey(subkey)
    End Function

    Public Function getCUKey(Optional subkey As String = "") As RegistryKey
        Return Registry.CurrentUser.OpenSubKey(subkey)
    End Function

    Public Function getUserKey(Optional subkey As String = "") As RegistryKey
        Return Registry.Users.OpenSubKey(subkey)
    End Function
End Module
