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
''' The include and exclude lists a single trim runs with. <see cref="Trim.trimFile"/> reads
''' these instead of the Trim module's settings, so another module can trim with its own lists,
''' or with none, without touching what the user saved for Trim. A new instance uses no lists.
''' </summary>
Public Class TrimOptions

    ''' <summary>
    ''' The file naming entries we always keep, or <c> Nothing </c> to keep none regardless of
    ''' detection
    ''' </summary>
    Public Property IncludesFile As iniFileChooser = Nothing

    ''' <summary>
    ''' The file naming entries we always discard, or <c> Nothing </c> to discard none
    ''' regardless of detection
    ''' </summary>
    Public Property ExcludesFile As iniFileChooser = Nothing

    ''' <summary>
    ''' Returns the lists currently chosen in the Trim module's settings
    ''' </summary>
    Public Shared Function FromSettings() As TrimOptions

        Return New TrimOptions With {
            .IncludesFile = If(UseTrimIncludes, TrimFile2, Nothing),
            .ExcludesFile = If(UseTrimExcludes, TrimFile4, Nothing)
        }

    End Function

End Class
