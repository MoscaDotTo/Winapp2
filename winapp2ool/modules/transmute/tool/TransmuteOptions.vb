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
''' The modes a single transmutation runs with. The engine reads these instead of the module's
''' settings, so a caller from another module builds its own and never touches what the user
''' saved. A new instance holds the defaults, and <see cref="FromSettings"/> copies the
''' user's current choices.
''' </summary>
Public Class TransmuteOptions

    ''' <summary>
    ''' The primary transmutation mode
    ''' <br /> Default: <c> TransmuteMode.Add </c>
    ''' </summary>
    Public Property Mode As TransmuteMode = TransmuteMode.Add

    ''' <summary>
    ''' The granularity of Replace operations
    ''' <br /> Default: <c> ReplaceMode.ByKey </c>
    ''' </summary>
    Public Property ReplaceBy As ReplaceMode = ReplaceMode.ByKey

    ''' <summary>
    ''' The granularity of Remove operations
    ''' <br /> Default: <c> RemoveMode.ByKey </c>
    ''' </summary>
    Public Property RemoveBy As RemoveMode = RemoveMode.ByKey

    ''' <summary>
    ''' What counts as the same key when Remove works by key
    ''' <br /> Default: <c> RemoveKeyMode.ByName </c>
    ''' </summary>
    Public Property RemoveKeysBy As RemoveKeyMode = RemoveKeyMode.ByName

    ''' <summary>
    ''' Indicates whether <c> [*] </c>, <c> [*Map:] </c> and <c> [*Name:] </c> source sections
    ''' are global operations rather than ordinary section names
    ''' <br /> Default: <c> True </c>
    ''' </summary>
    Public Property RecognizeGlobalSections As Boolean = True

    ''' <summary>
    ''' Returns the modes currently held in the Transmute module's settings
    ''' </summary>
    Public Shared Function FromSettings() As TransmuteOptions

        Return New TransmuteOptions With {
            .Mode = Transmutator,
            .ReplaceBy = TransmuteReplaceMode,
            .RemoveBy = TransmuteRemoveMode,
            .RemoveKeysBy = TransmuteRemoveKeyMode,
            .RecognizeGlobalSections = transmuteSettings.RecognizeGlobalSections
        }

    End Function

End Class
