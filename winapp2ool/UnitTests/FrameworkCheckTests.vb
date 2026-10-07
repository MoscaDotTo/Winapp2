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
''' Tests for reading which .NET Framework a build needs and whether this PC has it
''' </summary>
<TestClass()> Public Class FrameworkCheckTests

    ''' <summary>
    ''' Known Framework versions map to Microsoft's documented minimum Release numbers
    ''' </summary>
    <TestMethod()> Public Sub RequiredRelease_KnownVersions()

        Assert.AreEqual(528040, winapp2ool.FrameworkCheck.requiredFrameworkRelease(".NETFramework,Version=v4.8"))
        Assert.AreEqual(394802, winapp2ool.FrameworkCheck.requiredFrameworkRelease(".NETFramework,Version=v4.6.2"))
        Assert.AreEqual(533320, winapp2ool.FrameworkCheck.requiredFrameworkRelease(".NETFramework,Version=v4.8.1"))

    End Sub

    ''' <summary>
    ''' A .NET Core name, an unlisted 4.x version, a non-framework string and <c> Nothing </c> give no
    ''' required release rather than a guess, and an unlisted version gives no answer on whether it's met
    ''' </summary>
    <TestMethod()> Public Sub RequiredRelease_UnknownGivesNothing()

        Assert.IsFalse(winapp2ool.FrameworkCheck.requiredFrameworkRelease(".NETCoreApp,Version=v8.0").HasValue)
        Assert.IsFalse(winapp2ool.FrameworkCheck.requiredFrameworkRelease(".NETFramework,Version=v4.9").HasValue)
        Assert.IsFalse(winapp2ool.FrameworkCheck.requiredFrameworkRelease("not a framework").HasValue)
        Assert.IsFalse(winapp2ool.FrameworkCheck.requiredFrameworkRelease(Nothing).HasValue)
        Assert.IsFalse(winapp2ool.FrameworkCheck.frameworkRequirementMet(".NETFramework,Version=v4.9").HasValue)

    End Sub

    ''' <summary>
    ''' The tests themselves run on 4.8, so the registry must report at least 4.8, and both a 4.8 and a
    ''' 4.6 target are met
    ''' </summary>
    <TestMethod()> Public Sub InstalledRelease_IsAtLeast48()

        Assert.IsTrue(winapp2ool.FrameworkCheck.installedFrameworkRelease() >= 528040)
        Assert.AreEqual(True, winapp2ool.FrameworkCheck.frameworkRequirementMet(".NETFramework,Version=v4.8"))
        Assert.AreEqual(True, winapp2ool.FrameworkCheck.frameworkRequirementMet(".NETFramework,Version=v4.6"))

    End Sub

    ''' <summary>
    ''' The target framework is read from winapp2ool's own bytes, and junk bytes give Nothing without throwing
    ''' </summary>
    <TestMethod()> Public Sub TargetFramework_ReadFromBytes()

        Dim exeBytes = File.ReadAllBytes(GetType(winapp2ool.iniFile).Assembly.Location)

        Assert.AreEqual(".NETFramework,Version=v4.8", winapp2ool.FrameworkCheck.targetFrameworkOf(exeBytes))
        Assert.IsNull(winapp2ool.FrameworkCheck.targetFrameworkOf({1, 2, 3, 4}))

    End Sub

    ''' <summary>
    ''' Target framework names read as a person would write them, and a missing name reads as
    ''' <c> a newer version of .NET </c>
    ''' </summary>
    <TestMethod()> Public Sub DescribeFramework_IsReadable()

        Assert.AreEqual(".NET Framework 4.8", winapp2ool.FrameworkCheck.describeFramework(".NETFramework,Version=v4.8"))
        Assert.AreEqual("a newer version of .NET", winapp2ool.FrameworkCheck.describeFramework(Nothing))

    End Sub

End Class
