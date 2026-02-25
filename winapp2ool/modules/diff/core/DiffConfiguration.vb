'    Copyright (C) 2018-2025 Hazel Ward
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

''' <summary>
''' Shared constants and lookup tables used across the Diff module's detection and comparison logic
''' </summary>
Public Module DiffConfiguration

    ''' <summary>
    ''' A set of Regex characters that need to be escaped when using Regex.Match
    ''' </summary>
    Public ReadOnly Property RegexCharsToEscape As New Dictionary(Of String, String) From {
        {"*", ".*"}, {"+", "\+"}, {"{", "\{"}, {"}", "\}"},
        {"[", "\["}, {"]", "\]"}, {"$", "\$"}, {"(", "\("}, {")", "\)"}
    }

    ''' <summary>
    ''' Deprecated values that are no longer used or being 
    ''' phased out of winapp2.ini and their replacements. 
    ''' 
    ''' Keys will be replaced by their value for the purposes of not triggering a 
    ''' "false positive" when diffing, particularly in the case of V22XXXX to V23XXXX or newer 
    ''' </summary>
    Public ReadOnly Property PathReplacements As New Dictionary(Of String, String) From {
        {"%CommonAppData%", "%ProgramData%"},
        {"%LocalLowAppData%", "%UserProfile%\AppData\LocalLow"},
        {"*.*", "*"}, {"%Pictures%", "%UserProfile%\Pictures"},
        {"%Videos%", "%UserProfile%\Videos"},
        {"%Music%", "%UserProfile%\Music"},
        {"%Documents%", "%UserProfile%\Documents"}
    }

    ''' <summary>
    ''' File system and registry locations which are considered too vague to be used to establish matching key content across entries on their own
    ''' </summary>

    Public ReadOnly Property DisallowedPaths As New HashSet(Of String) From {
                "%Documents%\Add-in Express", "%UserProfile%\Desktop", "%LocalAppData%",
                "%WinDir%\System32", "%SystemDrive%", "%WinDir%", "%UserProfile%", "%Documents%",
                "%CommonAppData%", "%AppData%", "%Pictures%", "%Public%", "%Music%", "%Video%",
                "HKCU\Software\Microsoft\Windows", "HKLM\Software\Microsoft\Windows",
                "HKCU\Software\Microsoft\VisualStudio", "%LocalAppData%\Microsoft\Edge*",
                "HKCU\Software\Opera Software", "HKCU\Software\Vivaldi", "HKCU\Software\BraveSoftware",
                "%LocalAppData%\Packages\*\AC\Microsoft\CLR_v4.0*"
    }

    ''' <summary>
    ''' LangSecRef values associated with browsers
    ''' </summary>
    Public ReadOnly Property BrowserSecRefs As String() = {
        "3029", "3006", "3033", "3034", "3027", "3026", "3030", "3001"
    }

    ''' <summary>
    ''' Separator character used in movement key signatures stored in <c>KeyMovementTracker.MovedKeys</c>.
    ''' Null character (Chr(0)) cannot appear in ini file key values, making it unambiguous as a delimiter.
    ''' </summary>
    Public ReadOnly Property MovementKeySeparator As Char
        Get
            Return Chr(0)
        End Get
    End Property

End Module