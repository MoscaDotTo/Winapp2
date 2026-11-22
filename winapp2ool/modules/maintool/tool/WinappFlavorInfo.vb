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

Imports System.Text.RegularExpressions

''' <summary>
''' What a published winapp2.ini says about its own flavor: the name on its <c> ; Flavor: </c>
''' line, the description under it, and the release assets its links point at. It also holds
''' the flavor's command line flags and its download menu option. One instance per
''' <see cref="WinappFlavor"/>.
''' <br /><br />
'''
''' The asset names are the file names the build publishes in <c> Releases\ </c> and uploads to
''' each GitHub release, and must match <c> .github/scripts/flavors.csv </c>, which
''' <c> FlavorManifestTests </c> checks.
''' </summary>
Public NotInheritable Class WinappFlavorInfo

    ''' <summary>The flavor this instance describes</summary>
    Public ReadOnly Property Flavor As WinappFlavor

    ''' <summary>The name written on a file's <c> ; Flavor: </c> line</summary>
    Public ReadOnly Property DisplayName As String

    ''' <summary>The description lines written under the flavor line, without their leading <c> ; </c></summary>
    Public ReadOnly Property Description As IReadOnlyList(Of String)

    ''' <summary>The flavor's file name among the release assets</summary>
    Public ReadOnly Property AssetName As String

    ''' <summary>The flavor's changelog name among the release assets</summary>
    Public ReadOnly Property ChangelogAssetName As String

    ''' <summary>The command line flags that select the flavor, such as <c> -ccleaner </c> and <c> -cc </c></summary>
    Public ReadOnly Property CliFlags As IReadOnlyList(Of String)

    ''' <summary>The color of the flavor's option in the download menu</summary>
    Public ReadOnly Property MenuColor As ConsoleColor

    ''' <summary>The name of the flavor's option in the download menu, such as <c> System Ninja winapp2.rules </c></summary>
    Public ReadOnly Property MenuLabel As String
        Get
            Return $"{DisplayName} {LocalFileName}"
        End Get
    End Property

    ''' <summary>The description of the flavor's option in the download menu</summary>
    Public ReadOnly Property MenuDescription As String
        Get
            Return $"Download the latest {DisplayName} flavor of winapp2.ini"
        End Get
    End Property

    ''' <summary>The name the flavor's file has on a user's disk: <c> winapp2.rules </c> for System Ninja, <c> winapp2.ini </c> otherwise</summary>
    Public ReadOnly Property LocalFileName As String
        Get
            Return If(AssetName.EndsWith(".rules", StringComparison.OrdinalIgnoreCase), "winapp2.rules", "winapp2.ini")
        End Get
    End Property

    ' A header line naming the file's flavor, such as "; Flavor: BleachBit"
    Private Shared ReadOnly FlavorLine As New Regex("^;\s*Flavor:\s*(.+?)\s*$", RegexOptions.IgnoreCase)

    Private Sub New(flavor As WinappFlavor,
                    displayName As String,
                    assetName As String,
                    changelogAssetName As String,
                    menuColor As ConsoleColor,
                    cliFlags As String(),
                    ParamArray description As String())

        Me.Flavor = flavor
        Me.DisplayName = displayName
        Me.MenuColor = menuColor
        Me.CliFlags = cliFlags
        Me.AssetName = assetName
        Me.ChangelogAssetName = changelogAssetName
        Me.Description = description

    End Sub

    ''' <summary>Every flavor, in download menu order</summary>
    Public Shared ReadOnly Property All As IReadOnlyList(Of WinappFlavorInfo) = New List(Of WinappFlavorInfo) From {
        New WinappFlavorInfo(WinappFlavor.Base, "Base", "winapp2.ini", "changelog-base.txt", ConsoleColor.Cyan, {"-base", "-ncc"},
            "This is the base flavor of winapp2.ini, with no entries removed or changed. Use it with",
            "BleachBit 6+, and with any program that has no flavor of its own. This file will not",
            "work with System Ninja or CCleaner 7+, use those applications' respective flavors with them."),
        New WinappFlavorInfo(WinappFlavor.CCleaner, "CCleaner", "winapp2-ccleaner.ini", "changelog-ccleaner.txt", ConsoleColor.DarkRed, {"-ccleaner", "-cc"},
            "This is the CCleaner flavor of winapp2.ini, for CCleaner 6.41 and earlier. It re-categorizes entries",
            "to better fit the CCleaner Classic UI. CCleaner 7 is not compatible with this file. To use winapp2.ini",
            "with CCleaner 7+, use winapp2ool's CC7Patcher."),
        New WinappFlavorInfo(WinappFlavor.CCleaner7, "CCleaner 7", "winapp2-ccleaner7.ini", "changelog-ccleaner7.txt", ConsoleColor.Magenta, {"-ccleaner7", "-cc7"},
            "This is the CCleaner 7 flavor of winapp2.ini, converted to CCleaner 7's entry format.",
            "Don't install it by hand: winapp2ool's CC7Patcher patches it into ccleaner.ini.",
            "Do not use this file with applications other than CCleaner 7"),
        New WinappFlavorInfo(WinappFlavor.BleachBit, "BleachBit", "winapp2-bleachbit.ini", "changelog-bleachbit.txt", ConsoleColor.DarkCyan, {"-bleachbit", "-bb"},
            "This is the BleachBit flavor of winapp2.ini, for BleachBit 5 and earlier. It removes the",
            "registry exclusions those versions reject. BleachBit 6 and newer should use the base flavor."),
        New WinappFlavorInfo(WinappFlavor.FluentCleaner, "FluentCleaner", "winapp2-fluentcleaner.ini", "changelog-fluentcleaner.txt", ConsoleColor.Green, {"-fluentcleaner", "-fc"},
            "This is the FluentCleaner flavor of winapp2.ini, which carries the changes FluentCleaner",
            "makes to winapp2.ini downstream."),
        New WinappFlavorInfo(WinappFlavor.SystemNinja, "System Ninja", "winapp2-systemninja.rules", "changelog-systemninja.txt", ConsoleColor.Blue, {"-systemninja", "-sn"},
            "This is the System Ninja flavor of winapp2.ini. It replaces keys System Ninja can't run",
            "with ones it can."),
        New WinappFlavorInfo(WinappFlavor.Tron, "Tron", "winapp2-tron.ini", "changelog-tron.txt", ConsoleColor.Red, {"-tron"},
            "This is the Tron flavor of winapp2.ini, which carries the changes Tron makes to",
            "winapp2.ini downstream.")
    }

    Private Shared ReadOnly Table As Dictionary(Of WinappFlavor, WinappFlavorInfo) = All.ToDictionary(Function(info) info.Flavor)

    ''' <summary>Returns the header information for <paramref name="flavor"/></summary>
    '''
    ''' <param name="flavor">A winapp2.ini flavor</param>
    Public Shared Function Lookup(flavor As WinappFlavor) As WinappFlavorInfo

        Return Table(flavor)

    End Function

    ''' <summary>
    ''' Returns the flavor a <c> ; Flavor: </c> line names. Case and spaces don't matter, and
    ''' the enum name works as well as the display name, so <c> Base </c>, <c> CCleaner7 </c>
    ''' and <c> system ninja </c> all parse.
    ''' </summary>
    '''
    ''' <param name="name">The text after <c> Flavor: </c></param>
    '''
    ''' <returns>
    ''' The flavor, <br />
    ''' <c> Nothing </c> if <paramref name="name"/> names none
    ''' </returns>
    Public Shared Function ParseName(name As String) As WinappFlavor?

        If name Is Nothing Then Return Nothing

        Dim wanted = name.Replace(" ", "")

        For Each info In All

            If wanted.Equals(info.DisplayName.Replace(" ", ""), StringComparison.OrdinalIgnoreCase) OrElse
               wanted.Equals(info.Flavor.ToString(), StringComparison.OrdinalIgnoreCase) Then Return info.Flavor

        Next

        Return Nothing

    End Function

    ''' <summary>
    ''' Returns the flavor a <c> ; Flavor: </c> header line names. A flavor line naming an
    ''' unknown flavor is logged.
    ''' </summary>
    '''
    ''' <param name="line">A line of a winapp2.ini header, leading <c> ; </c> included</param>
    '''
    ''' <returns>
    ''' The flavor, <br />
    ''' <c> Nothing </c> if <paramref name="line"/> isn't a flavor line or names an unknown flavor
    ''' </returns>
    Public Shared Function ParseFlavorLine(line As String) As WinappFlavor?

        If line Is Nothing Then Return Nothing

        Dim m = FlavorLine.Match(line)
        If Not m.Success Then Return Nothing

        Dim flavor = ParseName(m.Groups(1).Value)
        If Not flavor.HasValue Then gLog($"Ignoring an unknown flavor: {line}")

        Return flavor

    End Function

End Class
