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
''' Stores the information required by EntryScaffold sections to
''' generate individual browser entries
''' </summary>
Friend Structure BrowserInfo

    ''' <summary>
    ''' The name of the browser, will be prepended to the name of each entry scaffold
    ''' </summary>
    Public Name As String

    ''' <summary>
    ''' The set of provided user data (chromium) or profiles (gecko) paths
    ''' </summary>
    Public UserDataPaths As List(Of String)

    ''' <summary>
    ''' The set of parent paths to the <c>UserDataPaths</c>
    ''' </summary>
    Public UserDataParentPaths As List(Of String)

    ''' <summary>
    ''' The name of the CCleaner "Section" that all entries for this browser will be grouped into
    ''' </summary>
    Public SectionName As String

    ''' <summary>
    ''' Indicates that the User Data path should be truncated off for the DetectFile <br />
    ''' Useful for easily supporting multiple versions of a single browser
    ''' </summary>
    Public TruncateDetect As Boolean

    ''' <summary>
    ''' The set of parent paths in the registry for the browser, used to generate RegKeys
    ''' </summary>
    Public RegistryRoots As List(Of String)

    ''' <summary>
    ''' Indicates whether or not the current browser should be omitted from the generation
    ''' process <br /><br />
    ''' Allows the easy enabling and disabling of browser support over time without requiring
    ''' any information to be truly lost
    ''' </summary>
    Public ShouldSkip As Boolean

    ''' <summary>
    ''' Creates a new BrowserInfo object for a particular browser
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
