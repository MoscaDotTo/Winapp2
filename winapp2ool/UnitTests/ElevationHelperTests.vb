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

''' <summary>
''' Tests for the startup check that decides whether winapp2ool relaunches with administrator rights
''' </summary>
<TestClass()> Public Class ElevationHelperTests

    ''' <summary>
    ''' A folder we can write to never triggers elevation, and the probe file doesn't linger
    ''' </summary>
    <TestMethod()> Public Sub WritableFolder_NeedsNoElevation()

        Dim folder = Path.Combine(Path.GetTempPath(), $"w2e-{Guid.NewGuid():N}")
        Directory.CreateDirectory(folder)

        Try

            Assert.IsFalse(winapp2ool.ElevationHelper.needsElevationToWrite(folder))
            Assert.AreEqual(0, Directory.GetFiles(folder).Length)

        Finally

            Directory.Delete(folder, True)

        End Try

    End Sub

    ''' <summary>
    ''' The Windows folder refuses writes to an unelevated process, so it needs elevation. Skipped when
    ''' the tests themselves run elevated, where the folder is writable
    ''' </summary>
    <TestMethod()> Public Sub ProtectedFolder_NeedsElevation()

        If winapp2ool.ElevationHelper.isElevated() Then Assert.Inconclusive("Tests are running elevated")

        Assert.IsTrue(winapp2ool.ElevationHelper.needsElevationToWrite(Environment.GetFolderPath(Environment.SpecialFolder.Windows)))

    End Sub

    ''' <summary>
    ''' A folder that doesn't exist yet is judged by the folder it would be created in, which here is writable
    ''' </summary>
    <TestMethod()> Public Sub MissingFolderUnderWritableParent_NeedsNoElevation()

        Assert.IsFalse(winapp2ool.ElevationHelper.needsElevationToWrite(Path.Combine(Path.GetTempPath(), $"w2e-missing-{Guid.NewGuid():N}", "deeper")))

    End Sub

    ''' <summary>
    ''' A folder that doesn't exist yet under a protected one needs elevation, since creating it does
    ''' </summary>
    <TestMethod()> Public Sub MissingFolderUnderProtectedParent_NeedsElevation()

        If winapp2ool.ElevationHelper.isElevated() Then Assert.Inconclusive("Tests are running elevated")

        Dim windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
        Assert.IsTrue(winapp2ool.ElevationHelper.needsElevationToWrite(Path.Combine(windows, $"w2e-missing-{Guid.NewGuid():N}")))

    End Sub

    ''' <summary>
    ''' Only -Nd values count, resolved the way the command line handler resolves them
    ''' </summary>
    <TestMethod()> Public Sub CommandLineFolders_ResolvesDirectoryArgs()

        Dim args = {"-s", "trim", "-1d", "C:\Program Files\CCleaner", "-1f", "winapp2.ini",
                    "-3d", "C:\Out\trimmed.ini", "-2d", "\Sub\Folder", "-4d"}

        Dim folders = winapp2ool.ElevationHelper.commandLineFolders(args)

        CollectionAssert.AreEqual({"C:\Program Files\CCleaner",
                                   "C:\Out\",
                                   Environment.CurrentDirectory & "\Sub\Folder"}, folders)

    End Sub

End Class
