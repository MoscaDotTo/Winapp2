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

Imports System.Security.Cryptography

''' <summary>
''' Verifies the signature published beside winapp2ool.exe before the self-updater will install it.
''' <br /><br />
'''
''' Releases are signed with ECDSA on the P-256 curve over a SHA-256 hash of the exe's bytes.
''' The signature is the 64-byte IEEE P1363 form (r followed by s), stored base64 on a single line in
''' <c> winapp2ool.exe.sig </c>. <c> scripts\New-UpdateSigningKey.ps1 </c> makes a signing key and
''' <c> scripts\Sign-Winapp2oolRelease.ps1 </c> writes the <c> .sig </c> file.
''' <br /><br />
'''
''' The signature covers the exe's bytes and nothing else. It doesn't name a version, branch or file,
''' so an older signed build passes this check too, and <see cref="autoUpdate"/> refuses it with a
''' separate version check.
''' </summary>
Module UpdateSignature

    ''' <summary>
    ''' The length in bytes of a P-256 signature in IEEE P1363 form
    ''' </summary>
    Private Const SignatureLength As Integer = 64

    ''' <summary>
    ''' The length in bytes of a raw uncompressed P-256 public key point without its leading
    ''' <c> 0x04 </c> marker
    ''' </summary>
    Private Const PublicKeyLength As Integer = 64

    ''' <summary>
    ''' The public keys trusted to sign winapp2ool releases, each the base64 of the raw point
    ''' X followed by Y. The updater accepts a signature from any key in the list.
    ''' The list is compiled in, so removing a key protects only the builds released after its removal.
    ''' A build with an empty list refuses every self-update.
    ''' </summary>
    Friend ReadOnly Property TrustedUpdateKeys As IReadOnlyList(Of String) = New String() {
                             "gkxCzJm0znpBIcJ5fNR2u7AZRIUEbNs0zuPPJVMvBD7ViNrxDZyRKRX0XyuQACt4UqjsqiotqUafAISkrJaYBw==",
                             "s1E/+J0tYUGESrS0tclac8IIUACbXn+KWdxAs0rSCMIwFNOrhMu1hfDFACT8LGZpR/YbbWfdsA78BjgVCbLLPA=="
                            }

    ''' <summary>
    ''' Returns whether any of <paramref name="trustedKeys"/> verifies <paramref name="signatureText"/>
    ''' as an ECDSA P-256 signature over the SHA-256 hash of <paramref name="data"/>
    ''' </summary>
    '''
    ''' <param name="data">
    ''' The signed bytes
    ''' </param>
    '''
    ''' <param name="signatureText">
    ''' The base64 signature as published. We ignore spaces, tabs and line breaks anywhere in it, and other whitespace only at the ends.
    ''' </param>
    '''
    ''' <param name="trustedKeys">
    ''' The base64 public keys to accept a signature from. A malformed key is skipped and the
    ''' rest are still tried.
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if any key in <paramref name="trustedKeys"/> produced the signature, <br />
    ''' <c> False </c> otherwise, including for any malformed input
    ''' </returns>
    Friend Function VerifyUpdateSignature(data As Byte(),
                                          signatureText As String,
                                          trustedKeys As IEnumerable(Of String)) As Boolean

        If data Is Nothing OrElse trustedKeys Is Nothing Then Return False

        Dim signature = decodeFixedLength(signatureText, SignatureLength)
        If signature Is Nothing Then Return False

        For Each key In trustedKeys

            If verifyWithKey(data, signature, key) Then Return True

        Next

        Return False

    End Function

    ''' <summary>
    ''' Returns whether a single base64 public key verifies a decoded signature
    ''' </summary>
    '''
    ''' <param name="data">
    ''' The signed bytes
    ''' </param>
    '''
    ''' <param name="signature">
    ''' The 64-byte P1363 signature
    ''' </param>
    '''
    ''' <param name="keyText">
    ''' The base64 of the key's raw X and Y coordinates
    ''' </param>
    '''
    ''' <returns>
    ''' <c> True </c> if the key is well formed and verifies the signature, <br />
    ''' <c> False </c> otherwise
    ''' </returns>
    Private Function verifyWithKey(data As Byte(),
                                   signature As Byte(),
                                   keyText As String) As Boolean

        Dim point = decodeFixedLength(keyText, PublicKeyLength)
        If point Is Nothing Then Return False

        Dim x(31) As Byte
        Dim y(31) As Byte
        Buffer.BlockCopy(point, 0, x, 0, 32)
        Buffer.BlockCopy(point, 32, y, 0, 32)

        Dim parameters As New ECParameters With {
            .Curve = ECCurve.NamedCurves.nistP256,
            .Q = New ECPoint With {.X = x, .Y = y}
        }

        ' Windows refuses to import a point that isn't on the curve, and .NET Framework reports that
        ' as PlatformNotSupportedException wrapping the CryptographicException
        Try

            Using verifier = ECDsa.Create(parameters)

                Return verifier.VerifyData(data, signature, HashAlgorithmName.SHA256)

            End Using

        Catch ex As CryptographicException

            Return False

        Catch ex As PlatformNotSupportedException

            Return False

        End Try

    End Function

    ''' <summary>
    ''' Decodes base64 text that must hold exactly <paramref name="length"/> bytes
    ''' </summary>
    '''
    ''' <param name="text">
    ''' The base64 text. We ignore spaces, tabs and line breaks anywhere in it, and other whitespace only at the ends.
    ''' </param>
    '''
    ''' <param name="length">
    ''' The required number of decoded bytes
    ''' </param>
    '''
    ''' <returns>
    ''' The decoded bytes, <br />
    ''' <c> Nothing </c> if <paramref name="text"/> is missing, isn't base64, or decodes to the wrong length
    ''' </returns>
    Private Function decodeFixedLength(text As String,
                                       length As Integer) As Byte()

        If text Is Nothing Then Return Nothing

        Try

            Dim bytes = Convert.FromBase64String(text.Trim())
            Return If(bytes.Length = length, bytes, Nothing)

        Catch ex As FormatException

            Return Nothing

        End Try

    End Function

End Module
