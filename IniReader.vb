'
'    IniReader class for C#
'		Version: 1.0		Date: 2002/04/24
'
'
'    Copyright © 2002, The KPD-Team
'    All rights reserved.
'    http://www.mentalis.org/
'
'  Redistribution and use in source and binary forms, with or without
'  modification, are permitted provided that the following conditions
'  are met:
'
'    - Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer. 
'
'    - Neither the name of the KPD-Team, nor the names of its contributors
'       may be used to endorse or promote products derived from this
'       software without specific prior written permission. 
'
'  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'  "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'  LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS
'  FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
'  THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT,
'  INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
'  (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
'  SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION)
'  HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT,
'  STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
'  ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED
'  OF THE POSSIBILITY OF SUCH DAMAGE.
'

Imports System
Imports System.Text
Imports System.Collections
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic

' The Org.Mentalis.Files contains a number of classes that read from or write to special files.
Namespace Org.Mentalis.Files
    '/// <summary>
    '/// The INIReader class can read keys from and write keys to an INI file.
    '/// </summary>
    '/// <remarks>
    '/// This class uses several Win32 API functions to read from and write to INI files. It will not work on Linux or FreeBSD.
    '/// </remarks>
    Public Class IniReader
        '/// <summary>
        '/// The GetPrivateProfileInt function retrieves an integer associated with a key in the specified section of an initialization file.
        '/// </summary>
        '/// <param name="lpApplicationName">Pointer to a null-terminated string specifying the name of the section in the initialization file.</param>
        '/// <param name="lpKeyName">Pointer to the null-terminated string specifying the name of the key whose value is to be retrieved. This value is in the form of a string; the GetPrivateProfileInt function converts the string into an integer and returns the integer.</param>
        '/// <param name="nDefault">Specifies the default value to return if the key name cannot be found in the initialization file.</param>
        '/// <param name="lpFileName">Pointer to a null-terminated string that specifies the name of the initialization file. If this parameter does not contain a full path to the file, the system searches for the file in the Windows directory.</param>
        '/// <returns>The return value is the integer equivalent of the string following the specified key name in the specified initialization file. If the key is not found, the return value is the specified default value. If the value of the key is less than zero, the return value is zero.</returns>
        Private Declare Ansi Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
        '/// <summary>
        '/// The WritePrivateProfileString function copies a string into the specified section of an initialization file.
        '/// </summary>
        '/// <param name="lpApplicationName">Pointer to a null-terminated string containing the name of the section to which the string will be copied. If the section does not exist, it is created. The name of the section is case-independent; the string can be any combination of uppercase and lowercase letters.</param>
        '/// <param name="lpKeyName">Pointer to the null-terminated string containing the name of the key to be associated with a string. If the key does not exist in the specified section, it is created. If this parameter is NULL, the entire section, including all entries within the section, is deleted.</param>
        '/// <param name="lpString">Pointer to a null-terminated string to be written to the file. If this parameter is NULL, the key pointed to by the lpKeyName parameter is deleted.</param>
        '/// <param name="lpFileName">Pointer to a null-terminated string that specifies the name of the initialization file.</param>
        '/// <returns>If the function successfully copies the string to the initialization file, the return value is nonzero; if the function fails, or if it flushes the cached version of the most recently accessed initialization file, the return value is zero.</returns>
        Private Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
        '/// <summary>
        '/// The GetPrivateProfileString function retrieves a string from the specified section in an initialization file.
        '/// </summary>
        '/// <param name="lpApplicationName">Pointer to a null-terminated string that specifies the name of the section containing the key name. If this parameter is NULL, the GetPrivateProfileString function copies all section names in the file to the supplied buffer.</param>
        '/// <param name="lpKeyName">Pointer to the null-terminated string specifying the name of the key whose associated string is to be retrieved. If this parameter is NULL, all key names in the section specified by the lpAppName parameter are copied to the buffer specified by the lpReturnedString parameter.</param>
        '/// <param name="lpDefault">Pointer to a null-terminated default string. If the lpKeyName key cannot be found in the initialization file, GetPrivateProfileString copies the default string to the lpReturnedString buffer. This parameter cannot be NULL. <br>Avoid specifying a default string with trailing blank characters. The function inserts a null character in the lpReturnedString buffer to strip any trailing blanks.</br></param>
        '/// <param name="lpReturnedString">Pointer to the buffer that receives the retrieved string.</param>
        '/// <param name="nSize">Specifies the size, in TCHARs, of the buffer pointed to by the lpReturnedString parameter.</param>
        '/// <param name="lpFileName">Pointer to a null-terminated string that specifies the name of the initialization file. If this parameter does not contain a full path to the file, the system searches for the file in the Windows directory.</param>
        '/// <returns>The return value is the number of characters copied to the buffer, not including the terminating null character.</returns>
        Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
        '/// <summary>
        '/// The GetPrivateProfileSectionNames function retrieves the names of all sections in an initialization file.
        '/// </summary>
        '/// <param name="lpszReturnBuffer">Pointer to a buffer that receives the section names associated with the named file. The buffer is filled with one or more null-terminated strings; the last string is followed by a second null character.</param>
        '/// <param name="nSize">Specifies the size, in TCHARs, of the buffer pointed to by the lpszReturnBuffer parameter.</param>
        '/// <param name="lpFileName">Pointer to a null-terminated string that specifies the name of the initialization file. If this parameter is NULL, the function searches the Win.ini file. If this parameter does not contain a full path to the file, the system searches for the file in the Windows directory.</param>
        '/// <returns>The return value specifies the number of characters copied to the specified buffer, not including the terminating null character. If the buffer is not large enough to contain all the section names associated with the specified initialization file, the return value is equal to the length specified by nSize minus two.</returns>
        Private Declare Ansi Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer() As Byte, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
        '/// <summary>
        '/// The WritePrivateProfileSection function replaces the keys and values for the specified section in an initialization file.
        '/// </summary>
        '/// <param name="lpAppName">Pointer to a null-terminated string specifying the name of the section in which data is written. This section name is typically the name of the calling application.</param>
        '/// <param name="lpString">Pointer to a buffer containing the new key names and associated values that are to be written to the named section.</param>
        '/// <param name="lpFileName">Pointer to a null-terminated string containing the name of the initialization file. If this parameter does not contain a full path for the file, the function searches the Windows directory for the file. If the file does not exist and lpFileName does not contain a full path, the function creates the file in the Windows directory. The function does not create a file if lpFileName contains the full path and file name of a file that does not exist.</param>
        '/// <returns>If the function succeeds, the return value is nonzero.<br>If the function fails, the return value is zero.</br></returns>
		Private Declare Ansi Function WritePrivateProfileSection Lib "kernel32.dll" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
        '/// <summary>Constructs a new IniReader instance.</summary>
        '/// <param name="file">Specifies the full path to the INI file (the file doesn't have to exist).</param>
        Public Sub New(ByVal file As String)
            Filename = file
        End Sub
        '/// <summary>Gets or sets the full path to the INI file.</summary>
        '/// <value>A String representing the full path to the INI file.</value>
        Public Property Filename() As String
            Get
                Return m_Filename
            End Get
            Set(ByVal Value As String)
                m_Filename = Value
            End Set
        End Property
        '/// <summary>Gets or sets the section you're working in. (aka 'the active section')</summary>
        '/// <value>A String representing the section you're working in.</value>
        Public Property Section() As String
            Get
                Return m_Section
            End Get
            Set(ByVal Value As String)
                m_Section = Value
            End Set
        End Property
        '/// <summary>Reads an Integer from the specified key of the specified section.</summary>
        '/// <param name="section">The section to search in.</param>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <param name="defVal">The value to return if the specified key isn't found.</param>
        '/// <returns>Returns the value of the specified section/key pair, or returns the default value if the specified section/key pair isn't found in the INI file.</returns>
        Public Function ReadInteger(ByVal section As String, ByVal key As String, ByVal defVal As Integer) As Integer
            Return GetPrivateProfileInt(section, key, defVal, Filename)
        End Function
        '/// <summary>Reads an Integer from the specified key of the specified section.</summary>
        '/// <param name="section">The section to search in.</param>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <returns>Returns the value of the specified section/key pair, or returns 0 if the specified section/key pair isn't found in the INI file.</returns>
        Public Function ReadInteger(ByVal section As String, ByVal key As String) As Integer
            Return ReadInteger(section, key, 0)
        End Function
        '/// <summary>Reads an Integer from the specified key of the active section.</summary>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <param name="defVal">The section to search in.</param>
        '/// <returns>Returns the value of the specified Key, or returns the default value if the specified Key isn't found in the active section of the INI file.</returns>
        Public Function ReadInteger(ByVal key As String, ByVal defVal As Integer) As Integer
            Return ReadInteger(Section, key, defVal)
        End Function
        '/// <summary>Reads an Integer from the specified key of the active section.</summary>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <returns>Returns the value of the specified key, or returns 0 if the specified key isn't found in the active section of the INI file.</returns>
        Public Function ReadInteger(ByVal key As String) As Integer
            Return ReadInteger(key, 0)
        End Function
        '/// <summary>Reads a String from the specified key of the specified section.</summary>
        '/// <param name="section">The section to search in.</param>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <param name="defVal">The value to return if the specified key isn't found.</param>
        '/// <returns>Returns the value of the specified section/key pair, or returns the default value if the specified section/key pair isn't found in the INI file.</returns>
        Public Function ReadString(ByVal section As String, ByVal key As String, ByVal defVal As String) As String
            Dim sb As New StringBuilder(MAX_ENTRY)
            Dim Ret As Integer = GetPrivateProfileString(section, key, defVal, sb, MAX_ENTRY, Filename)
            Return sb.ToString()
        End Function
        '/// <summary>Reads a String from the specified key of the specified section.</summary>
        '/// <param name="section">The section to search in.</param>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <returns>Returns the value of the specified section/key pair, or returns an empty String if the specified section/key pair isn't found in the INI file.</returns>
        Public Function ReadString(ByVal section As String, ByVal key As String) As String
            Return ReadString(section, key, "")
        End Function
        '/// <summary>Reads a String from the specified key of the active section.</summary>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <returns>Returns the value of the specified key, or returns an empty String if the specified key isn't found in the active section of the INI file.</returns>
        Public Function ReadString(ByVal key As String) As String
            Return ReadString(Section, key)
        End Function
        '/// <summary>Reads a Long from the specified key of the specified section.</summary>
        '/// <param name="section">The section to search in.</param>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <param name="defVal">The value to return if the specified key isn't found.</param>
        '/// <returns>Returns the value of the specified section/key pair, or returns the default value if the specified section/key pair isn't found in the INI file.</returns>
        Public Function ReadLong(ByVal section As String, ByVal key As String, ByVal defVal As Long) As Long
            Return Long.Parse(ReadString(section, key, defVal.ToString()))
        End Function
        '/// <summary>Reads a Long from the specified key of the specified section.</summary>
        '/// <param name="section">The section to search in.</param>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <returns>Returns the value of the specified section/key pair, or returns 0 if the specified section/key pair isn't found in the INI file.</returns>
        Public Function ReadLong(ByVal section As String, ByVal key As String) As Long
            Return ReadLong(section, key, 0)
        End Function
        '/// <summary>Reads a Long from the specified key of the active section.</summary>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <param name="defVal">The section to search in.</param>
        '/// <returns>Returns the value of the specified key, or returns the default value if the specified key isn't found in the active section of the INI file.</returns>
        Public Function ReadLong(ByVal key As String, ByVal defVal As Long) As Long
            Return ReadLong(Section, key, defVal)
        End Function
        '/// <summary>Reads a Long from the specified key of the active section.</summary>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <returns>Returns the value of the specified Key, or returns 0 if the specified Key isn't found in the active section of the INI file.</returns>
        Public Function ReadLong(ByVal key As String) As Long
            Return ReadLong(key, 0)
        End Function
        '/// <summary>Reads a Byte array from the specified key of the specified section.</summary>
        '/// <param name="section">The section to search in.</param>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <returns>Returns the value of the specified section/key pair, or returns null (Nothing in VB.NET) if the specified section/key pair isn't found in the INI file.</returns>
        Public Function ReadByteArray(ByVal section As String, ByVal key As String) As Byte()
            Try
                Return Convert.FromBase64String(ReadString(section, key))
            Catch
            End Try
            Return Nothing
        End Function
        '/// <summary>Reads a Byte array from the specified key of the active section.</summary>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <returns>Returns the value of the specified key, or returns null (Nothing in VB.NET) if the specified key pair isn't found in the active section of the INI file.</returns>
        Public Function ReadByteArray(ByVal key As String) As Byte()
            Return ReadByteArray(Section, key)
        End Function
        '/// <summary>Reads a Boolean from the specified key of the specified section.</summary>
        '/// <param name="section">The section to search in.</param>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <param name="defVal">The value to return if the specified key isn't found.</param>
        '/// <returns>Returns the value of the specified section/key pair, or returns the default value if the specified section/key pair isn't found in the INI file.</returns>
        Public Function ReadBoolean(ByVal section As String, ByVal key As String, ByVal defVal As Boolean) As Boolean
            Return Boolean.Parse(ReadString(section, key, defVal.ToString()))
        End Function
        '/// <summary>Reads a Boolean from the specified key of the specified section.</summary>
        '/// <param name="section">The section to search in.</param>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <returns>Returns the value of the specified section/key pair, or returns false if the specified section/key pair isn't found in the INI file.</returns>
        Public Function ReadBoolean(ByVal section As String, ByVal key As String) As Boolean
            Return ReadBoolean(section, key, False)
        End Function
        '/// <summary>Reads a Boolean from the specified key of the specified section.</summary>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <param name="defVal">The value to return if the specified key isn't found.</param>
        '/// <returns>Returns the value of the specified key pair, or returns the default value if the specified key isn't found in the active section of the INI file.</returns>
        Public Function ReadBoolean(ByVal key As String, ByVal defVal As Boolean) As Boolean
            Return ReadBoolean(Section, key, defVal)
        End Function
        '/// <summary>Reads a Boolean from the specified key of the specified section.</summary>
        '/// <param name="key">The key from which to return the value.</param>
        '/// <returns>Returns the value of the specified key, or returns false if the specified key isn't found in the active section of the INI file.</returns>
        Public Function ReadBoolean(ByVal key As String) As Boolean
            Return ReadBoolean(Section, key)
        End Function
        '/// <summary>Writes an Integer to the specified key in the specified section.</summary>
        '/// <param name="section">The section to write in.</param>
        '/// <param name="key">The key to write to.</param>
        '/// <param name="value">The value to write.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal section As String, ByVal key As String, ByVal value As Integer) As Boolean
            Return Write(Section, key, value.ToString())
        End Function
        '/// <summary>Writes an Integer to the specified key in the active section.</summary>
        '/// <param name="key">The key to write to.</param>
        '/// <param name="value">The value to write.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal key As String, ByVal value As Integer) As Boolean
            Return Write(Section, key, value)
        End Function
        '/// <summary>Writes a String to the specified key in the specified section.</summary>
        '/// <param name="section">Specifies the section to write in.</param>
        '/// <param name="key">Specifies the key to write to.</param>
        '/// <param name="value">Specifies the value to write.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal section As String, ByVal key As String, ByVal value As String) As Boolean
            Return (WritePrivateProfileString(section, key, value, Filename) <> 0)
        End Function
        '/// <summary>Writes a String to the specified key in the active section.</summary>
        '///	<param name="key">The key to write to.</param>
        '/// <param name="value">The value to write.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal key As String, ByVal value As String) As Boolean
            Return Write(Section, key, value)
        End Function
        '/// <summary>Writes a Long to the specified key in the specified section.</summary>
        '/// <param name="section">The section to write in.</param>
        '/// <param name="key">The key to write to.</param>
        '/// <param name="value">The value to write.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal section As String, ByVal key As String, ByVal value As Long) As Boolean
            Return Write(section, key, value.ToString())
        End Function
        '/// <summary>Writes a Long to the specified key in the active section.</summary>
        '/// <param name="key">The key to write to.</param>
        '/// <param name="value">The value to write.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal key As String, ByVal value As Long) As Boolean
            Return Write(Section, key, value)
        End Function
        '/// <summary>Writes a Byte array to the specified key in the specified section.</summary>
        '/// <param name="section">The section to write in.</param>
        '/// <param name="key">The key to write to.</param>
        '/// <param name="value">The value to write.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal section As String, ByVal key As String, ByVal value() As Byte) As Boolean
            If value Is Nothing Then
                Return Write(section, key, CType(Nothing, String))
            Else
                Return Write(section, key, value, 0, value.Length)
            End If
        End Function
        '/// <summary>Writes a Byte array to the specified key in the active section.</summary>
        '/// <param name="key">The key to write to.</param>
        '/// <param name="value">The value to write.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal key As String, ByVal value() As Byte) As Boolean
            Return Write(Section, key, value)
        End Function
        '/// <summary>Writes a Byte array to the specified key in the specified section.</summary>
        '/// <param name="section">The section to write in.</param>
        '/// <param name="key">The key to write to.</param>
        '/// <param name="value">The value to write.</param>
        '/// <param name="offset">An offset in <i>value</i>.</param>
        '/// <param name="length">The number of elements of <i>value</i> to convert.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal section As String, ByVal key As String, ByVal value() As Byte, ByVal offset As Integer, ByVal length As Integer) As Boolean
            If value Is Nothing Then
                Return Write(section, key, CType(Nothing, String))
            Else
                Return Write(section, key, Convert.ToBase64String(value, offset, length))
            End If
        End Function
        '/// <summary>Writes a Boolean to the specified key in the specified section.</summary>
        '/// <param name="section">The section to write in.</param>
        '/// <param name="key">The key to write to.</param>
        '/// <param name="value">The value to write.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal section As String, ByVal key As String, ByVal value As Boolean) As Boolean
            Return Write(section, key, value.ToString())
        End Function
        '/// <summary>Writes a Boolean to the specified key in the active section.</summary>
        '/// <param name="key">The key to write to.</param>
        '/// <param name="value">The value to write.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function Write(ByVal key As String, ByVal value As Boolean) As Boolean
            Return Write(Section, key, value)
        End Function
        '/// <summary>Deletes a key from the specified section.</summary>
        '/// <param name="section">The section to delete from.</param>
        '/// <param name="key">The key to delete.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function DeleteKey(ByVal section As String, ByVal key As String) As Boolean
            Return (WritePrivateProfileString(Section, key, Nothing, Filename) <> 0)
        End Function
        '/// <summary>Deletes a key from the active section.</summary>
        '/// <param name="key">The key to delete.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function DeleteKey(ByVal key As String) As Boolean
            Return (WritePrivateProfileString(Section, key, Nothing, Filename) <> 0)
        End Function
        '/// <summary>Deletes a section from an INI file.</summary>
        '/// <param name="section">The section to delete.</param>
        '/// <returns>Returns true if the function succeeds, false otherwise.</returns>
        Public Function DeleteSection(ByVal section As String) As Boolean
            Return WritePrivateProfileSection(section, Nothing, Filename) <> 0
        End Function
        '/// <summary>Retrieves a list of all available sections in the INI file.</summary>
        '/// <returns>Returns an ArrayList with all available sections.</returns>
        Public Function GetSectionNames() As ArrayList
            Try
                Dim buffer(MAX_ENTRY) As Byte
                GetPrivateProfileSectionNames(buffer, MAX_ENTRY, Filename)
                Dim parts() As String = Encoding.ASCII.GetString(buffer).Trim(ControlChars.NullChar).Split(ControlChars.NullChar)
                Return New ArrayList(parts)
            Catch
            End Try
            Return Nothing
        End Function
        'Private variables and constants
        '/// <summary>
        '/// Holds the full path to the INI file.
        '/// </summary>
        Private m_Filename As String
        '/// <summary>
        '/// Holds the active section name
        '/// </summary>
        Private m_Section As String
        '/// <summary>
        '/// The maximum number of bytes in a section buffer.
        '/// </summary>
        Private Const MAX_ENTRY As Integer = 32768
    End Class
End Namespace
