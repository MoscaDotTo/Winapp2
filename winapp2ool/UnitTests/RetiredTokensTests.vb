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
''' Tests for <see cref="winapp2ool.RetiredTokens.FindRetiredToken"/>: prefix matching, the new
''' spelling each shape is told to use, case-insensitivity, and the values it must leave alone
''' </summary>
<TestClass()> Public Class RetiredTokensTests

    ''' <summary>The stems UWPBuilder refuses today</summary>
    Private Shared ReadOnly PackageStem As String() = {"%Package"}

    ''' <summary>Returns the refusal message for <paramref name="template"/> against <see cref="PackageStem"/></summary>
    ''' <param name="template">The key value to scan</param>
    Private Shared Function Find(template As String) As String

        Return winapp2ool.RetiredTokens.FindRetiredToken(template, PackageStem, "[Test *].FileKey1")

    End Function

    ''' <summary>
    ''' A closed <c> %Package% </c> is told its new spelling, with the context named
    ''' </summary>
    <TestMethod()> Public Sub ClosedToken_NamesNewSpelling()

        Assert.AreEqual("`%Package%` in [Test *].FileKey1 is now spelled `<Package>`", Find("%Package%\LocalState|*.log"))

    End Sub

    ''' <summary>
    ''' A numbered token becomes an indexed reference
    ''' </summary>
    <TestMethod()> Public Sub NumberedToken_BecomesIndexedReference()

        Assert.AreEqual("`%Package2%` in [Test *].FileKey1 is now spelled `<Package[2]>`", Find("%Package2%\LocalState|*.log"))

    End Sub

    ''' <summary>
    ''' A numbered token with its closing <c> % </c> missing still gets the indexed spelling
    ''' </summary>
    <TestMethod()> Public Sub UnclosedNumberedToken_GetsIndexedHint()

        Assert.AreEqual("`%Package2` in [Test *].FileKey1 looks like the retired `%Package2%` token, which is now spelled `<Package[2]>`", Find("%Package2|*.log"))

    End Sub

    ''' <summary>
    ''' The stem matches as a prefix, so a token with its closing <c> % </c> missing is caught.
    ''' That's the shape that once published <c> %Package\LocalCache|*.log </c>.
    ''' </summary>
    <TestMethod()> Public Sub UnclosedToken_IsCaughtWithHint()

        Dim msg = Find("%Package\LocalCache|*.log")

        StringAssert.Contains(msg, "`%Package`")
        StringAssert.Contains(msg, "retired `%Package%` token")
        StringAssert.Contains(msg, "`<Package>`")

    End Sub

    ''' <summary>
    ''' Matching ignores case, and the new spelling takes the stem's capitalization
    ''' </summary>
    <TestMethod()> Public Sub Matching_IsCaseInsensitive()

        Assert.AreEqual("`%package%` in [Test *].FileKey1 is now spelled `<Package>`", Find("%package%\AC|*"))

    End Sub

    ''' <summary>
    ''' Extra letters after the stem get the generic hint rather than an exact respelling
    ''' </summary>
    <TestMethod()> Public Sub ExtraLetters_GetGenericHint()

        StringAssert.Contains(Find("%PackageRoot%\AC|*"), "looks like the retired `%Package%` token")

    End Sub

    ''' <summary>
    ''' Real environment variables that share a first letter, the new spelling, and a value
    ''' with no token at all all pass
    ''' </summary>
    <TestMethod()> Public Sub UnrelatedValues_Pass()

        For Each value In {"%ProgramFiles%\Contoso|*.log", "%ProgramData%\Contoso|*", "<Package>\AC|*",
                           "<Package[2]>\LocalState|*", "%LocalAppData%\Packages\Contoso|*", "Packages|*"}

            Assert.AreEqual("", Find(value), value)

        Next

    End Sub

    ''' <summary>
    ''' Several stems can be passed at once, and a later stem's hit names its own spelling
    ''' </summary>
    <TestMethod()> Public Sub SeveralStems_EachMatches()

        Dim msg = winapp2ool.RetiredTokens.FindRetiredToken("%WebViewRoot%\Default|*", {"%Package", "%WebViewRoot"}, "ctx")

        Assert.AreEqual("`%WebViewRoot%` in ctx is now spelled `<WebViewRoot>`", msg)

    End Sub

    ' ----- Renamed tokens -----

    ''' <summary>A rename mapping like BrowserBuilder's, where one token also changed its name</summary>
    Private Shared ReadOnly Renames As KeyValuePair(Of String, String)() = {
        New KeyValuePair(Of String, String)("%UserDataPath", "UserDataPath"),
        New KeyValuePair(Of String, String)("%BrowserPath", "UserDataParent")
    }

    ''' <summary>
    ''' A renamed token is told its new name, not a respelling of its old one
    ''' </summary>
    <TestMethod()> Public Sub RenamedToken_NamesNewName()

        Assert.AreEqual("`%BrowserPath%` in ctx is now spelled `<UserDataParent>`",
                        winapp2ool.RetiredTokens.FindRetiredToken("%BrowserPath%\Application|debug.log", Renames, "ctx"))

    End Sub

    ''' <summary>
    ''' A renamed token with its closing <c> % </c> missing gets the hint with the new name
    ''' </summary>
    <TestMethod()> Public Sub RenamedToken_Unclosed_NamesNewName()

        Assert.AreEqual("`%BrowserPath` in ctx looks like the retired `%BrowserPath%` token, which is now spelled `<UserDataParent>`",
                        winapp2ool.RetiredTokens.FindRetiredToken("%BrowserPath\Application|x", Renames, "ctx"))

    End Sub

    ''' <summary>
    ''' Extra letters after a renamed stem get the generic hint with the new name
    ''' </summary>
    <TestMethod()> Public Sub RenamedToken_ExtraLetters_NamesNewName()

        Assert.AreEqual("`%BrowserPaths%` in ctx looks like the retired `%BrowserPath%` token, which is now spelled `<UserDataParent>`",
                        winapp2ool.RetiredTokens.FindRetiredToken("%BrowserPaths%\x|*", Renames, "ctx"))

    End Sub

    ''' <summary>
    ''' When the new tokens take no index, a numbered token gets the generic hint instead of
    ''' an indexed spelling, closed or not
    ''' </summary>
    <TestMethod()> Public Sub NumberedToken_NotIndexable_GetsGenericHint()

        Assert.AreEqual("`%BrowserPath2%` in ctx looks like the retired `%BrowserPath%` token, which is now spelled `<UserDataParent>`",
                        winapp2ool.RetiredTokens.FindRetiredToken("%BrowserPath2%\x|*", Renames, "ctx", indexable:=False))
        Assert.AreEqual("`%BrowserPath2` in ctx looks like the retired `%BrowserPath%` token, which is now spelled `<UserDataParent>`",
                        winapp2ool.RetiredTokens.FindRetiredToken("%BrowserPath2\x|*", Renames, "ctx", indexable:=False))
        Assert.AreEqual("`%BrowserPath2%` in ctx is now spelled `<UserDataParent[2]>`",
                        winapp2ool.RetiredTokens.FindRetiredToken("%BrowserPath2%\x|*", Renames, "ctx"))

    End Sub

    ''' <summary>
    ''' Matching a renamed stem ignores case, and an unrenamed entry in the same mapping keeps its name
    ''' </summary>
    <TestMethod()> Public Sub RenameMapping_IgnoresCase_AndKeepsIdentityEntries()

        Assert.AreEqual("`%browserpath%` in ctx is now spelled `<UserDataParent>`",
                        winapp2ool.RetiredTokens.FindRetiredToken("%browserpath%\x|*", Renames, "ctx"))
        Assert.AreEqual("`%UserDataPath%` in ctx is now spelled `<UserDataPath>`",
                        winapp2ool.RetiredTokens.FindRetiredToken("%UserDataPath%\*|*", Renames, "ctx"))
        Assert.AreEqual("", winapp2ool.RetiredTokens.FindRetiredToken("<UserDataParent>\x|*", Renames, "ctx"))

    End Sub

End Class
