'    Copyright (C) 2018-2021 Hazel Ward
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
''' <summary>Provides some tools to simplify unit testing winapp2ool</summary>
Module UnitTools
    ''' <summary>
    ''' Replaces <see cref="winapp2ool.commandLineHandler.cmdargs"/> with <paramref name="args"/>
    ''' and runs <paramref name="handler"/> against them
    ''' </summary>
    '''
    ''' <param name="handler">A module's command line handler, such as <c> HandleLintCmdLine </c></param>
    '''
    ''' <param name="args">The arguments to hand the handler, without the executable's path</param>
    '''
    ''' <param name="addHalt">
    ''' Indicates whether to append <c> UNIT_TESTING_HALT </c> to the arguments. A handler that
    ''' checks for it returns after parsing instead of running its module. Only WinappDebug's
    ''' handler checks for it. <br /><br />
    ''' Optional, Default: <c> False </c>
    ''' </param>
    Public Sub setCmdLineArgs(handler As Action, args As String(), Optional addHalt As Boolean = False)
        winapp2ool.commandLineHandler.cmdargs = args.ToList
        If addHalt Then winapp2ool.commandLineHandler.cmdargs.Add("UNIT_TESTING_HALT")
        handler()
    End Sub
End Module