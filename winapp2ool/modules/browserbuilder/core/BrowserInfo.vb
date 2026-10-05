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

''' <summary>
''' Stores the information parsed from one <c> BrowserInfo: </c> section, which the
''' EntryScaffold sections need to generate that browser's entries
''' </summary>
Friend Structure BrowserInfo

    ''' <summary>
    ''' The name of the browser. Each generated entry is named
    ''' <c> {Name} {scaffold name} * </c>.
    ''' </summary>
    Public Name As String

    ''' <summary>
    ''' The set of provided user data (chromium) or profiles (gecko) paths, substituted for
    ''' <c> %UserDataPath% </c>
    ''' </summary>
    Public UserDataPaths As List(Of String)

    ''' <summary>
    ''' The parent of each path in <see cref="UserDataPaths"/> (everything before its last
    ''' backslash), at the same index. Substituted for <c> %BrowserPath% </c>, and used for
    ''' the DetectFile keys when <see cref="TruncateDetect"/> is set.
    ''' </summary>
    Public UserDataParentPaths As List(Of String)

    ''' <summary>
    ''' The <c> Section= </c> value that all entries for this browser will be grouped into,
    ''' or empty when the BrowserInfo has none
    ''' </summary>
    Public SectionName As String

    ''' <summary>
    ''' Indicates whether the DetectFile keys use <see cref="UserDataParentPaths"/> instead
    ''' of <see cref="UserDataPaths"/>. Useful for easily supporting multiple versions of a
    ''' single browser.
    ''' </summary>
    Public TruncateDetect As Boolean

    ''' <summary>
    ''' The set of parent paths in the registry for the browser, substituted for
    ''' <c> %RegistryRoot% </c> to generate RegKeys
    ''' </summary>
    Public RegistryRoots As List(Of String)

    ''' <summary>
    ''' Indicates whether the current browser should be omitted from the generation
    ''' process <br /><br />
    ''' Allows the easy enabling and disabling of browser support over time without requiring
    ''' any information to be truly lost
    ''' </summary>
    Public ShouldSkip As Boolean

    ''' <summary>
    ''' Creates a new <c> BrowserInfo </c> for a particular browser, with empty lists and
    ''' no section
    ''' </summary>
    '''
    ''' <param name="name">
    ''' The name of the web browser as it appears in the BrowserInfo section name
    ''' </param>
    Public Sub New(name As String)

        Me.Name = name
        UserDataPaths = New List(Of String)
        UserDataParentPaths = New List(Of String)
        SectionName = ""
        TruncateDetect = False
        RegistryRoots = New List(Of String)
        ShouldSkip = False

    End Sub

End Structure
