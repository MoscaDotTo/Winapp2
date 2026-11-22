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

Imports System.Text
Imports winapp2ool

''' <summary>
''' Tests for the key-type table (<see cref="KeyTypes"/>): every fact in it, its two lookups,
''' <c> winapp2entry </c>'s buckets and the properties it derives from the table, and that
''' EntryBuilder reserves every key type
''' </summary>
<TestClass()> Public Class KeyTypesTests

    ''' <summary>Returns the first section parsed from <paramref name="text"/></summary>
    ''' <param name="text">Literal ini text</param>
    Private Shared Function FirstSection(text As String) As iniSection

        Using ms As New IO.MemoryStream(Encoding.UTF8.GetBytes(text))
            Using reader As New IO.StreamReader(ms)
                For Each section In iniFile.FromStream(reader, "", "test.ini")
                    Return section
                Next
            End Using
        End Using

        Throw New InvalidOperationException("No section parsed from test input")

    End Function

    ''' <summary>Returns a <c> winapp2entry </c> holding <paramref name="keyLines"/></summary>
    ''' <param name="keyLines">The entry's key lines</param>
    Private Shared Function EntryOf(ParamArray keyLines As String()) As winapp2entry

        Return New winapp2entry(FirstSection("[Test *]" & vbCrLf & String.Join(vbCrLf, keyLines) & vbCrLf))

    End Function

    ''' <summary>
    ''' Every cell of the table, in order. A change here is a change to winapp2.ini's key
    ''' vocabulary and should be deliberate.
    ''' </summary>
    <TestMethod()> Public Sub Table_PinsEveryFact()

        Dim expected = {
            "0|DetectOS|Never|Detection|None|False",
            "1|LangSecRef|Never|Category|None|False",
            "2|Section|Never|Category|None|False",
            "3|SpecialDetect|Never|Detection|None|True",
            "4|Detect|WhenSeveral|Detection|Registry|False",
            "5|DetectFile|WhenSeveral|Detection|Filesystem|False",
            "6|Default|Never|Setting|None|False",
            "7|Warning|Never|Warning|None|False",
            "8|FileKey|Always|Deletion|Filesystem|False",
            "9|RegKey|Always|Deletion|Registry|False",
            "10|ExcludeKey|Always|Exclusion|ByFlag|False"
        }

        Dim actual = KeyTypes.All.Select(Function(t) $"{t.Order}|{t.Name}|{t.Numbering}|{t.Role}|{t.Domain}|{t.Deprecated}").ToArray()

        CollectionAssert.AreEqual(expected, actual)

    End Sub

    ''' <summary><c> Find </c> ignores case and refuses a numbered name, <c> Error </c> and <c> Nothing </c></summary>
    <TestMethod()> Public Sub Find_IgnoresCaseAndRefusesNonTypes()

        Assert.AreEqual("FileKey", KeyTypes.Find("fileKEY").Name)
        Assert.AreEqual("DetectOS", KeyTypes.Find("DETECTOS").Name)
        Assert.IsNull(KeyTypes.Find("FileKey1"))
        Assert.IsNull(KeyTypes.Find("Error"))
        Assert.IsNull(KeyTypes.Find(""))
        Assert.IsNull(KeyTypes.Find(Nothing))

    End Sub

    ''' <summary>
    ''' <c> FindPrefix </c> takes the longest name the text starts with, so <c> DetectFile </c> and
    ''' <c> DetectOS </c> beat <c> Detect </c> in any case
    ''' </summary>
    <TestMethod()> Public Sub FindPrefix_TakesTheLongestMatch()

        Assert.AreEqual("DetectFile", KeyTypes.FindPrefix("DetectFile%AppData%\x").Name)
        Assert.AreEqual("DetectFile", KeyTypes.FindPrefix("detectfile%AppData%\x").Name)
        Assert.AreEqual("DetectOS", KeyTypes.FindPrefix("DetectOS5.1").Name)
        Assert.AreEqual("DetectOS", KeyTypes.FindPrefix("detectos5.1").Name)
        Assert.AreEqual("Detect", KeyTypes.FindPrefix("Detect2HKCU\Software").Name)
        Assert.AreEqual("ExcludeKey", KeyTypes.FindPrefix("ExcludeKey1FILE|x").Name)
        Assert.AreEqual("DetectOS", KeyTypes.FindPrefix("Detect" & ChrW(&HAD) & "OS10.0|").Name, "a soft hyphen is skipped, as the repair always skipped it")
        Assert.IsNull(KeyTypes.FindPrefix("Detec"))
        Assert.IsNull(KeyTypes.FindPrefix(Nothing))

    End Sub

    ''' <summary>
    ''' <c> GetBucketIndex </c> gives the index each key type always had, <c> Error </c> the error
    ''' bucket, and anything else -1
    ''' </summary>
    <TestMethod()> Public Sub GetBucketIndex_MatchesTheBucketOrder()

        Dim expected As New Dictionary(Of String, Integer) From {
            {"DETECTOS", 0}, {"LANGSECREF", 1}, {"SECTION", 2}, {"SPECIALDETECT", 3}, {"DETECT", 4},
            {"DETECTFILE", 5}, {"DEFAULT", 6}, {"WARNING", 7}, {"FILEKEY", 8}, {"REGKEY", 9},
            {"EXCLUDEKEY", 10}, {"ERROR", 11}
        }

        For Each kvp In expected
            Assert.AreEqual(kvp.Value, winapp2entry.GetBucketIndex(kvp.Key), kvp.Key)
            Assert.AreEqual(kvp.Value, winapp2entry.GetBucketIndex(kvp.Key.ToLowerInvariant()), kvp.Key)
        Next

        Assert.AreEqual(11, winapp2entry.ErrorBucket)
        Assert.AreEqual(-1, winapp2entry.GetBucketIndex("FileKey1"))
        Assert.AreEqual(-1, winapp2entry.GetBucketIndex("DeleteMe"))

    End Sub

    ''' <summary>
    ''' Keys land in their type's bucket whatever their case, an unknown type in the error bucket,
    ''' and <c> ToIniSection </c> writes them back in table order
    ''' </summary>
    <TestMethod()> Public Sub Entry_BucketsKeysByType()

        Dim entry = EntryOf("filekey1=%AppData%\x|*.log", "Bogus=1", "Warning=w", "detect=HKCU\Software\x", "LangSecRef=3021")

        Assert.AreEqual(12, entry.KeyLists.Count)
        Assert.AreEqual(1, entry.FileKeys.Count)
        Assert.AreEqual(1, entry.Detects.Count)
        Assert.AreEqual(1, entry.WarningKey.Count)
        Assert.AreEqual(1, entry.LangSecRef.Count)
        Assert.AreEqual(1, entry.ErrorKeys.Count)
        Assert.AreSame(entry.ErrorKeys, entry.KeyLists(winapp2entry.ErrorBucket))

        Dim order = entry.ToIniSection().Keys.Select(Function(k) k.Name).ToArray()
        CollectionAssert.AreEqual({"LangSecRef", "detect", "Warning", "filekey1", "Bogus"}, order)

        entry.RemoveKey(entry.FileKeys(0))
        Assert.AreEqual(0, entry.FileKeys.Count)

    End Sub

    ''' <summary>An error key whose type was repaired moves to its bucket; an <c> Error </c> key stays</summary>
    <TestMethod()> Public Sub ReclassifyErrorKeys_MovesOnlyRepairedKeys()

        Dim entry = EntryOf("Error=x", "Bogus=1")
        Dim bogus = entry.ErrorKeys.Single(Function(k) k.Name = "Bogus")
        bogus.Name = "FileKey1"

        entry.ReclassifyErrorKeys()

        Assert.AreEqual(1, entry.FileKeys.Count)
        Assert.AreEqual(1, entry.ErrorKeys.Count)
        Assert.AreEqual("Error", entry.ErrorKeys(0).Name)

    End Sub

    ''' <summary>
    ''' Every detection type counts as detection, SpecialDetect included, and DetectOS alone is
    ''' only DetectOS
    ''' </summary>
    <TestMethod()> Public Sub DetectionProperties_ReadTheDetectionRole()

        Assert.IsFalse(EntryOf("LangSecRef=3021", "FileKey1=%AppData%\x|*.log").HasDetectionKey)

        For Each line In {"DetectOS=10.0|", "SpecialDetect=DET_CHROME", "Detect=HKCU\Software\x", "DetectFile=%AppData%\x"}
            Assert.IsTrue(EntryOf("LangSecRef=3021", line).HasDetectionKey, line)
        Next

        Assert.IsTrue(EntryOf("DetectOS=10.0|", "Warning=w").HasOnlyDetectOS)
        Assert.IsFalse(EntryOf("DetectOS=10.0|", "SpecialDetect=DET_CHROME").HasOnlyDetectOS)
        Assert.IsFalse(EntryOf("DetectOS=10.0|", "Detect=HKCU\Software\x").HasOnlyDetectOS)
        Assert.IsFalse(EntryOf("Detect=HKCU\Software\x").HasOnlyDetectOS)

    End Sub

    ''' <summary><c> RenumberKeys </c> sorts and renumbers the always-numbered types and leaves Detect alone</summary>
    <TestMethod()> Public Sub RenumberKeys_RenumbersOnlyAlwaysNumberedTypes()

        Dim entry = EntryOf("Detect3=HKCU\Software\b", "Detect7=HKCU\Software\a",
                            "FileKey4=%AppData%\b|*", "FileKey9=%AppData%\a|*",
                            "RegKey5=HKCU\Software\z", "ExcludeKey2=FILE|%AppData%\|x")

        entry.RenumberKeys()

        CollectionAssert.AreEqual({"Detect3", "Detect7"}, entry.Detects.Select(Function(k) k.Name).ToArray())
        CollectionAssert.AreEqual({"FileKey1", "FileKey2"}, entry.FileKeys.Select(Function(k) k.Name).ToArray())
        Assert.AreEqual("%AppData%\a|*", entry.FileKeys(0).Value)
        Assert.AreEqual("RegKey1", entry.RegKeys(0).Name)
        Assert.AreEqual("ExcludeKey1", entry.ExcludeKeys(0).Name)

    End Sub

    ''' <summary>
    ''' An entry holding one key of every type goes through EntryBuilder's parser without any key
    ''' being read as a variable, and comes out with every type EntryBuilder passes through.
    ''' SpecialDetect is refused and Default ignored by design, and Section gives way to LangSecRef.
    ''' A type added to the table without a reserved name and a parser branch shows up here as a
    ''' variable.
    ''' </summary>
    <TestMethod()> Public Sub EntryBuilder_ReservesEveryKeyType()

        Dim samples As New Dictionary(Of String, String) From {
            {"DetectOS", "10.0|"},
            {"LangSecRef", "3021"},
            {"Section", "Key Types"},
            {"SpecialDetect", "DET_CHROME"},
            {"Detect", "HKCU\Software\KeyTypesTest"},
            {"DetectFile", "%LocalAppData%\KeyTypesTest"},
            {"Default", "False"},
            {"Warning", "Key types test"},
            {"FileKey", "%LocalAppData%\KeyTypesTest|*.log"},
            {"RegKey", "HKCU\Software\KeyTypesTest\Cache"},
            {"ExcludeKey", "FILE|%LocalAppData%\KeyTypesTest\|keep.log"}
        }

        Dim text As New StringBuilder("[Key Types Test *]" & vbCrLf)
        For Each t In KeyTypes.All
            Assert.IsTrue(samples.ContainsKey(t.Name), $"No sample value for {t.Name}")
            Dim suffix = If(t.Numbering = KeyNumbering.Always, "1", "")
            text.Append($"{t.Name}{suffix}={samples(t.Name)}{vbCrLf}")
        Next

        Dim menu As New MenuSection
        Dim spec As EntryBuilder.EntrySpec
        Dim section As iniSection
        Dim log As List(Of String)

        Using cap = gLogCapture()
            spec = EntryBuilder.parseEntrySpec(FirstSection(text.ToString()), menu)
            section = EntryBuilder.generateEntry(spec,
                New Dictionary(Of String, List(Of String)),
                New Dictionary(Of String, List(Of String)),
                New Dictionary(Of String, List(Of String)),
                New Dictionary(Of String, List(Of String)),
                New EntryBuilder.EntryBuilderStats,
                menu)
            log = cap.Lines.ToList()
        End Using

        CollectionAssert.AreEqual(New List(Of String), spec.Variables.DeclaredNames(),
                                  "Key types read as variables: " & String.Join(", ", spec.Variables.DeclaredNames()))
        Assert.IsFalse(log.Any(Function(l) l.Contains("shadows reserved key name")), String.Join(vbCrLf, log))

        Dim notEmitted = {"SpecialDetect", "Default", "Section"}
        For Each t In KeyTypes.All.Where(Function(kt) Not notEmitted.Contains(kt.Name))
            Assert.IsTrue(section.Keys.Any(Function(k) k.KeyType = t.Name), $"{t.Name} not emitted")
        Next
        For Each name In notEmitted
            Assert.IsFalse(section.Keys.Any(Function(k) k.KeyType = name), $"{name} emitted")
        Next

    End Sub

End Class
