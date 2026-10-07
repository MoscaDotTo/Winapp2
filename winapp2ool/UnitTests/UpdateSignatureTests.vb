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

Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text

''' <summary>
''' Tests for the self-updater's pure parts: signature verification against a given key list,
''' the well-formedness of the embedded trusted keys, the anti-downgrade version comparison, and
''' the quoting of relaunch arguments. None of these touch the network or the file system. Each
''' rejection test covers only the malformed or tampered inputs it builds, so together they don't
''' show that every tampered file is rejected.
''' </summary>
<TestClass()> Public Class UpdateSignatureTests

    ''' <summary>
    ''' A payload signed by <c> scripts\Sign-Winapp2oolRelease.ps1 </c> with a throwaway key made by
    ''' <c> scripts\New-UpdateSigningKey.ps1 </c>. The key was discarded, so this vector only proves that
    ''' the script and the client agree on the signature and key formats
    ''' </summary>
    Private Const VectorPayload As String = "winapp2ool cross-implementation test vector"

    ''' <summary>
    ''' The <c> .sig </c> file contents the signing script wrote for <see cref="VectorPayload"/>
    ''' </summary>
    Private Const VectorSignature As String = "u21rb9P7UC4ouHeKJU40SIhO9K0sWM9TGCC+UUQ7x38Z8lH+ckfzeWWrEfHWlEPKtYXvSyiyP6UpfHKkF6txWQ=="

    ''' <summary>
    ''' The public key the key generation script printed for the throwaway key
    ''' </summary>
    Private Const VectorPublicKey As String = "uJURBShNRWqwITy/KorUbf+O9HB6xmU0vvpdbi9mkO+yAz+ZX+BJNG2T+kz2hOjEnAuxrMQzPNPxaUJNKosuIw=="

    Private Shared ReadOnly Payload As Byte() = Encoding.ASCII.GetBytes("stand-in for the bytes of winapp2ool.exe")

    Private Shared Function NewKey() As ECDsa

        Return ECDsa.Create(ECCurve.NamedCurves.nistP256)

    End Function

    ''' <summary>
    ''' Returns a key's public point in the form the updater stores: base64 of X followed by Y
    ''' </summary>
    Private Shared Function PublicKeyOf(key As ECDsa) As String

        Dim q = key.ExportParameters(False).Q
        Return Convert.ToBase64String(q.X.Concat(q.Y).ToArray())

    End Function

    ''' <summary>
    ''' Returns the base64 P1363 signature of <paramref name="data"/>, as a <c> .sig </c> file holds it
    ''' </summary>
    Private Shared Function SignatureOf(key As ECDsa, data As Byte()) As String

        Return Convert.ToBase64String(key.SignData(data, HashAlgorithmName.SHA256))

    End Function

    Private Shared Function Verify(data As Byte(), signature As String, ParamArray keys As String()) As Boolean

        Return winapp2ool.UpdateSignature.VerifyUpdateSignature(data, signature, keys)

    End Function

    <TestMethod()> Public Sub ValidSignature_Accepted()

        Using key = NewKey()

            Assert.IsTrue(Verify(Payload, SignatureOf(key, Payload), PublicKeyOf(key)))

        End Using

    End Sub

    ''' <summary>
    ''' The published <c> .sig </c> may pick up a trailing newline in transit, so the client trims before decoding
    ''' </summary>
    <TestMethod()> Public Sub SignatureWithSurroundingWhitespace_Accepted()

        Using key = NewKey()

            Assert.IsTrue(Verify(Payload, "  " & SignatureOf(key, Payload) & vbCrLf, PublicKeyOf(key)))

        End Using

    End Sub

    <TestMethod()> Public Sub FlippedDataByte_Rejected()

        Using key = NewKey()

            Dim signature = SignatureOf(key, Payload)
            Dim tampered = CType(Payload.Clone(), Byte())
            tampered(tampered.Length \ 2) = tampered(tampered.Length \ 2) Xor CByte(1)

            Assert.IsFalse(Verify(tampered, signature, PublicKeyOf(key)))

        End Using

    End Sub

    <TestMethod()> Public Sub FlippedSignatureByte_Rejected()

        Using key = NewKey()

            Dim signature = Convert.FromBase64String(SignatureOf(key, Payload))
            signature(10) = signature(10) Xor CByte(&H80)

            Assert.IsFalse(Verify(Payload, Convert.ToBase64String(signature), PublicKeyOf(key)))

        End Using

    End Sub

    <TestMethod()> Public Sub WrongKey_Rejected()

        Using signer = NewKey(), other = NewKey()

            Assert.IsFalse(Verify(Payload, SignatureOf(signer, Payload), PublicKeyOf(other)))

        End Using

    End Sub

    ''' <summary>
    ''' A release signed with the backup key must pass while the primary key is still listed first
    ''' </summary>
    <TestMethod()> Public Sub BackupKey_Accepted()

        Using primary = NewKey(), backup = NewKey()

            Assert.IsTrue(Verify(Payload, SignatureOf(backup, Payload), PublicKeyOf(primary), PublicKeyOf(backup)))

        End Using

    End Sub

    ''' <summary>
    ''' A malformed entry in the key list must not stop a later, valid key from verifying
    ''' </summary>
    <TestMethod()> Public Sub MalformedKeyBeforeValidKey_StillAccepted()

        Using key = NewKey()

            Assert.IsTrue(Verify(Payload, SignatureOf(key, Payload), "not base64!", Nothing, PublicKeyOf(key)))

        End Using

    End Sub

    <TestMethod()> Public Sub MalformedBase64Signature_Rejected()

        Using key = NewKey()

            Assert.IsFalse(Verify(Payload, "this is not base64!", PublicKeyOf(key)))
            Assert.IsFalse(Verify(Payload, "", PublicKeyOf(key)))
            Assert.IsFalse(Verify(Payload, Nothing, PublicKeyOf(key)))

        End Using

    End Sub

    ''' <summary>
    ''' Rejects a signature one byte short, and one of DER length, which is the other common encoding
    ''' </summary>
    <TestMethod()> Public Sub WrongLengthSignature_Rejected()

        Using key = NewKey()

            Dim signature = Convert.FromBase64String(SignatureOf(key, Payload))

            Assert.IsFalse(Verify(Payload, Convert.ToBase64String(signature.Take(63).ToArray()), PublicKeyOf(key)))
            Assert.IsFalse(Verify(Payload, Convert.ToBase64String(signature.Concat(New Byte(7) {}).ToArray()), PublicKeyOf(key)))

        End Using

    End Sub

    ''' <summary>
    ''' Rejects a key carrying the <c> 0x04 </c> uncompressed-point prefix, the easiest paste mistake to make,
    ''' and a key cut to its X coordinate
    ''' </summary>
    <TestMethod()> Public Sub WrongLengthKey_Rejected()

        Using key = NewKey()

            Dim signature = SignatureOf(key, Payload)
            Dim point = Convert.FromBase64String(PublicKeyOf(key))
            Dim prefixed = (New Byte() {4}).Concat(point).ToArray()

            Assert.IsFalse(Verify(Payload, signature, Convert.ToBase64String(prefixed)))
            Assert.IsFalse(Verify(Payload, signature, Convert.ToBase64String(point.Take(32).ToArray())))

        End Using

    End Sub

    <TestMethod()> Public Sub MalformedBase64Key_Rejected()

        Using key = NewKey()

            Assert.IsFalse(Verify(Payload, SignatureOf(key, Payload), "this is not base64!"))

        End Using

    End Sub

    <TestMethod()> Public Sub PointNotOnCurve_Rejected()

        Using key = NewKey()

            Dim signature = SignatureOf(key, Payload)
            Dim ones = Enumerable.Repeat(CByte(1), 64).ToArray()

            Assert.IsFalse(Verify(Payload, signature, Convert.ToBase64String(ones)))
            Assert.IsFalse(Verify(Payload, signature, Convert.ToBase64String(New Byte(63) {})))

        End Using

    End Sub

    <TestMethod()> Public Sub EmptyKeyList_Rejected()

        Using key = NewKey()

            Dim signature = SignatureOf(key, Payload)

            Assert.IsFalse(Verify(Payload, signature))
            Assert.IsFalse(winapp2ool.UpdateSignature.VerifyUpdateSignature(Payload, signature, Nothing))

        End Using

    End Sub

    <TestMethod()> Public Sub MissingData_Rejected()

        Using key = NewKey()

            Assert.IsFalse(Verify(Nothing, SignatureOf(key, Payload), PublicKeyOf(key)))

        End Using

    End Sub

    ''' <summary>
    ''' The VB verifier must accept what the PowerShell signing script produced
    ''' </summary>
    <TestMethod()> Public Sub CrossImplementationVector_Accepted()

        Assert.IsTrue(Verify(Encoding.ASCII.GetBytes(VectorPayload), VectorSignature, VectorPublicKey))

    End Sub

    <TestMethod()> Public Sub CrossImplementationVector_OtherPayloadRejected()

        Assert.IsFalse(Verify(Encoding.ASCII.GetBytes(VectorPayload & "!"), VectorSignature, VectorPublicKey))

    End Sub

    ''' <summary>
    ''' Every embedded trusted key must be a 64-byte point on P-256, so a bad paste fails this test
    ''' rather than every user's update
    ''' </summary>
    <TestMethod()> Public Sub TrustedUpdateKeys_AreWellFormedPoints()

        For Each key In winapp2ool.UpdateSignature.TrustedUpdateKeys

            Dim point = Convert.FromBase64String(key)
            Assert.AreEqual(64, point.Length, $"Trusted key {key} is not 64 bytes")

            Dim parameters As New ECParameters With {
                .Curve = ECCurve.NamedCurves.nistP256,
                .Q = New ECPoint With {.X = point.Take(32).ToArray(), .Y = point.Skip(32).ToArray()}
            }

            Using ECDsa.Create(parameters)
            End Using

        Next

    End Sub

    <TestMethod()> Public Sub NewerVersion_Accepted()

        Assert.IsTrue(winapp2ool.updater.IsNewerToolVersion("1.7.9767.27789", "1.7.9767.27788"))
        Assert.IsTrue(winapp2ool.updater.IsNewerToolVersion("1.8.0.0", "1.7.9767.27788"))
        Assert.IsTrue(winapp2ool.updater.IsNewerToolVersion(" 1.7.9767.27789 ", "1.7.9767.27788"))

    End Sub

    ''' <summary>
    ''' Segments compare as numbers, so a build just after midnight, whose last segment is short, still orders correctly
    ''' </summary>
    <TestMethod()> Public Sub ShortLastSegment_ComparedNumerically()

        Assert.IsTrue(winapp2ool.updater.IsNewerToolVersion("1.7.9768.5", "1.7.9767.27788"))
        Assert.IsFalse(winapp2ool.updater.IsNewerToolVersion("1.7.9767.900", "1.7.9767.27788"))

    End Sub

    <TestMethod()> Public Sub EqualOrLowerVersion_Refused()

        Assert.IsFalse(winapp2ool.updater.IsNewerToolVersion("1.7.9767.27788", "1.7.9767.27788"))
        Assert.IsFalse(winapp2ool.updater.IsNewerToolVersion("1.7.9767.27787", "1.7.9767.27788"))
        Assert.IsFalse(winapp2ool.updater.IsNewerToolVersion("1.6.9999.0", "1.7.0.0"))

    End Sub

    <TestMethod()> Public Sub UnparseableVersion_Refused()

        Assert.IsFalse(winapp2ool.updater.IsNewerToolVersion("", "1.7.9767.27788"))
        Assert.IsFalse(winapp2ool.updater.IsNewerToolVersion(Nothing, "1.7.9767.27788"))
        Assert.IsFalse(winapp2ool.updater.IsNewerToolVersion("404: Not Found", "1.7.9767.27788"))
        Assert.IsFalse(winapp2ool.updater.IsNewerToolVersion("9.9.9.9", "000000 (file not found)"))
        Assert.IsNull(winapp2ool.updater.parseToolVersion("1.7.x.0"))

    End Sub

    <TestMethod()> Public Sub QuoteArgument_Cases()

        Assert.AreEqual("-s", winapp2ool.updater.QuoteArgument("-s"))
        Assert.AreEqual("C:\no\spaces\", winapp2ool.updater.QuoteArgument("C:\no\spaces\"))
        Assert.AreEqual("""""", winapp2ool.updater.QuoteArgument(""))
        Assert.AreEqual("""has space""", winapp2ool.updater.QuoteArgument("has space"))
        Assert.AreEqual("""C:\Program Files\x\\""", winapp2ool.updater.QuoteArgument("C:\Program Files\x\"))
        Assert.AreEqual("""a\""b""", winapp2ool.updater.QuoteArgument("a""b"))
        Assert.AreEqual("""a\\\""b""", winapp2ool.updater.QuoteArgument("a\""b"))
        Assert.AreEqual("""a\\b c""", winapp2ool.updater.QuoteArgument("a\\b c"))

    End Sub

    <TestMethod()> Public Sub BuildArgumentString_JoinsWithSpaces()

        Assert.AreEqual("-s -1d ""C:\My Files\\"" """"", winapp2ool.updater.BuildArgumentString({"-s", "-1d", "C:\My Files\", ""}))
        Assert.AreEqual("", winapp2ool.updater.BuildArgumentString(New String() {}))

    End Sub

    ''' <summary>
    ''' Windows must split the built command line back into exactly the original arguments
    ''' </summary>
    <TestMethod()> Public Sub BuildArgumentString_RoundTripsThroughCommandLineToArgvW()

        Dim args = {"-autoupdate", "-s", "has space", "", "C:\Program Files\x\", "C:\no\spaces\", "embedded ""quotes"" here",
                    "back\""slash", "trailing\\", "\\server\share\", "tab" & vbTab & "separated", "a\\b c", """", "\"}

        CollectionAssert.AreEqual(args, SplitLikeWindows(winapp2ool.updater.BuildArgumentString(args)))

    End Sub

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Private Shared Function CommandLineToArgvW(lpCmdLine As String, ByRef pNumArgs As Integer) As IntPtr
    End Function

    <DllImport("kernel32.dll")>
    Private Shared Function LocalFree(hMem As IntPtr) As IntPtr
    End Function

    ''' <summary>
    ''' Splits a command line the way Windows does for a process, dropping the program name
    ''' </summary>
    Private Shared Function SplitLikeWindows(arguments As String) As String()

        Dim count = 0
        Dim argv = CommandLineToArgvW("winapp2ool.exe " & arguments, count)
        If argv = IntPtr.Zero Then Throw New ComponentModel.Win32Exception()

        Try

            Dim result(count - 2) As String

            For i = 1 To count - 1

                result(i - 1) = Marshal.PtrToStringUni(Marshal.ReadIntPtr(argv, i * IntPtr.Size))

            Next

            Return result

        Finally

            LocalFree(argv)

        End Try

    End Function

End Class
