'    Copyright (C) 2018-2026 Hazel Ward
'
'    This file is a part of Winapp2ool
'
'    Winapp2ool is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Winapp2ool is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Winapp2ool.  If not, see <http://www.gnu.org/licenses/>.

Option Strict On

Imports System.IO
Imports System.Reflection
Imports System.Runtime.Versioning
Imports Microsoft.Win32

''' <summary>
''' Works out which .NET Framework release a winapp2ool build needs, and whether this PC has it.
''' <br /><br />
'''
''' <c> Environment.Version </c> can't answer that: every Framework from 4.6 on reports the same runtime
''' version, 4.0.30319.42000. Windows records the installed Framework as a <c> Release </c> number under the
''' <c> NDP\v4\Full </c> registry key, and each Framework version has a documented minimum for it.
''' </summary>
Module FrameworkCheck

    ''' <summary>
    ''' The smallest <c> Release </c> value for each .NET Framework 4.x version, from Microsoft's
    ''' "determine which .NET Framework versions are installed" documentation. Lookups match the
    ''' version exactly, so a version missing here, such as <c> v4.8.0 </c> or a later 4.8.x, is unknown.
    ''' </summary>
    Private ReadOnly MinimumReleases As New Dictionary(Of Version, Integer) From {
        {New Version(4, 5), 378389}, {New Version(4, 5, 1), 378675}, {New Version(4, 5, 2), 379893},
        {New Version(4, 6), 393295}, {New Version(4, 6, 1), 394254}, {New Version(4, 6, 2), 394802},
        {New Version(4, 7), 460798}, {New Version(4, 7, 1), 461308}, {New Version(4, 7, 2), 461808},
        {New Version(4, 8), 528040}, {New Version(4, 8, 1), 533320}
    }

    ''' <summary>
    ''' Returns the installed .NET Framework 4.x <c> Release </c> number, read from
    ''' <c> HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full </c> through the 32-bit registry
    ''' view, or 0 if the key is missing or its <c> Release </c> value isn't a DWORD
    ''' </summary>
    Friend Function installedFrameworkRelease() As Integer

        Using hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)

            Using ndp = hklm.OpenSubKey("SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full")

                If ndp Is Nothing Then Return 0

                Dim release = ndp.GetValue("Release")
                Return If(TypeOf release Is Integer, CInt(release), 0)

            End Using

        End Using

    End Function

    ''' <summary>
    ''' Returns the <c> Release </c> number a target framework needs
    ''' </summary>
    '''
    ''' <param name="frameworkName">
    ''' A target framework name, such as <c> .NETFramework,Version=v4.8 </c>
    ''' </param>
    '''
    ''' <returns>
    ''' The minimum <c> Release </c> number, <br />
    ''' <c> Nothing </c> if <paramref name="frameworkName"/> isn't a .NET Framework version listed in
    ''' <see cref="MinimumReleases"/>
    ''' </returns>
    Friend Function requiredFrameworkRelease(frameworkName As String) As Integer?

        If String.IsNullOrWhiteSpace(frameworkName) Then Return Nothing

        Dim parsed As FrameworkName

        Try

            parsed = New FrameworkName(frameworkName)

        Catch ex As ArgumentException

            Return Nothing

        End Try

        If Not parsed.Identifier.Equals(".NETFramework", StringComparison.OrdinalIgnoreCase) Then Return Nothing

        Dim minimum As Integer
        If Not MinimumReleases.TryGetValue(parsed.Version, minimum) Then Return Nothing

        Return minimum

    End Function

    ''' <summary>
    ''' Returns whether this PC has the .NET Framework that <paramref name="frameworkName"/> needs
    ''' </summary>
    '''
    ''' <param name="frameworkName">
    ''' A target framework name, such as <c> .NETFramework,Version=v4.8 </c>
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the installed <c> Release </c> number is at least the one
    ''' <paramref name="frameworkName"/> needs, <br />
    ''' <c> False </c> if it is lower, <br />
    ''' <c> Nothing </c> when <see cref="requiredFrameworkRelease"/> doesn't know the version,
    ''' including for anything that isn't .NET Framework 4.x
    ''' </returns>
    Friend Function frameworkRequirementMet(frameworkName As String) As Boolean?

        Dim needed = requiredFrameworkRelease(frameworkName)
        If Not needed.HasValue Then Return Nothing

        Return installedFrameworkRelease() >= needed.Value

    End Function

    ''' <summary>
    ''' Returns the target framework compiled into an assembly's <see cref="TargetFrameworkAttribute"/>.
    ''' We load the bytes reflection-only, so none of the assembly's code runs, but the assembly stays
    ''' loaded in this process until it exits. Loading the same assembly a second time throws
    ''' <see cref="FileLoadException"/>, so a repeat call in the same process returns <c> Nothing </c>.
    ''' </summary>
    '''
    ''' <param name="assemblyBytes">
    ''' The assembly's file contents
    ''' </param>
    '''
    ''' <returns>
    ''' The target framework name, such as <c> .NETFramework,Version=v4.8 </c>, <br />
    ''' <c> Nothing </c> if the bytes aren't a .NET assembly, carry no target framework, or were
    ''' already loaded by an earlier call
    ''' </returns>
    Friend Function targetFrameworkOf(assemblyBytes As Byte()) As String

        Try

            Dim inspected = Assembly.ReflectionOnlyLoad(assemblyBytes)

            For Each attribute In CustomAttributeData.GetCustomAttributes(inspected)

                If attribute.AttributeType.FullName = GetType(TargetFrameworkAttribute).FullName AndAlso attribute.ConstructorArguments.Count > 0 Then Return TryCast(attribute.ConstructorArguments(0).Value, String)

            Next

        Catch ex As BadImageFormatException

            Return Nothing

        Catch ex As FileLoadException

            Return Nothing

        End Try

        Return Nothing

    End Function

    ''' <summary>
    ''' Returns the target framework of the running winapp2ool, or <c> Nothing </c> if there's no
    ''' entry assembly or it carries no <see cref="TargetFrameworkAttribute"/>
    ''' </summary>
    Friend Function runningTargetFramework() As String

        Return Assembly.GetEntryAssembly()?.GetCustomAttribute(Of TargetFrameworkAttribute)()?.FrameworkName

    End Function

    ''' <summary>
    ''' Returns a readable name for a target framework, such as <c> .NET Framework 4.8 </c>
    ''' </summary>
    '''
    ''' <param name="frameworkName">
    ''' A target framework name, such as <c> .NETFramework,Version=v4.8 </c>
    ''' </param>
    '''
    ''' <returns>
    ''' <c> .NET Framework </c> and the version for a .NET Framework name, <br />
    ''' any other parseable name unchanged, <br />
    ''' <c> "a newer version of .NET" </c> if <paramref name="frameworkName"/> is <c> Nothing </c> or can't be parsed
    ''' </returns>
    Friend Function describeFramework(frameworkName As String) As String

        Try

            Dim parsed As New FrameworkName(frameworkName)
            If parsed.Identifier.Equals(".NETFramework", StringComparison.OrdinalIgnoreCase) Then Return $".NET Framework {parsed.Version}"

            Return frameworkName

        Catch ex As ArgumentException

            Return "a newer version of .NET"

        End Try

    End Function

End Module
