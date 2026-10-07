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
''' Tests for the two inputs to the startup elevation check: <see cref="winapp2ool.ElevationHelper.needsElevationToWrite"/>
''' and <see cref="winapp2ool.ElevationHelper.commandLineFolders"/>. Nothing here runs the relaunch itself.
''' </summary>
<TestClass()> Public Class ElevationHelperTests

    ''' <summary>
    ''' A fresh folder under the temp directory needs no elevation, and the probe file is gone afterward
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
    ''' The Windows folder refuses writes to an unelevated process, so it needs elevation. Inconclusive
    ''' when the tests themselves run elevated, where the folder is writable
    ''' </summary>
    <TestMethod()> Public Sub ProtectedFolder_NeedsElevation()

        If winapp2ool.ElevationHelper.isElevated() Then Assert.Inconclusive("Tests are running elevated")

        Assert.IsTrue(winapp2ool.ElevationHelper.needsElevationToWrite(Environment.GetFolderPath(Environment.SpecialFolder.Windows)))

    End Sub

    ''' <summary>
    ''' A folder whose parent is missing too is judged by the nearest folder above it that exists, which
    ''' here is the writable temp directory two levels up
    ''' </summary>
    <TestMethod()> Public Sub MissingFolderUnderWritableParent_NeedsNoElevation()

        Assert.IsFalse(winapp2ool.ElevationHelper.needsElevationToWrite(Path.Combine(Path.GetTempPath(), $"w2e-missing-{Guid.NewGuid():N}", "deeper")))

    End Sub

    ''' <summary>
    ''' A folder that doesn't exist yet under a protected one needs elevation, since creating it does.
    ''' Inconclusive when the tests run elevated.
    ''' </summary>
    <TestMethod()> Public Sub MissingFolderUnderProtectedParent_NeedsElevation()

        If winapp2ool.ElevationHelper.isElevated() Then Assert.Inconclusive("Tests are running elevated")

        Dim windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows)
        Assert.IsTrue(winapp2ool.ElevationHelper.needsElevationToWrite(Path.Combine(windows, $"w2e-missing-{Guid.NewGuid():N}")))

    End Sub

    ''' <summary>
    ''' Only <c> -Nd </c> values count: <c> -1f </c> and a trailing <c> -4d </c> with no value are
    ''' ignored, a final segment with a dot is cut off as a file name, and a leading backslash is
    ''' resolved against the working folder
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
