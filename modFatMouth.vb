Imports System.Runtime.InteropServices 'Import
Imports System.IO
Imports FatMouth.Org.Mentalis.Files

Module modFatMouth
    'Win32 API's via (p)invoke substitutes long datatype for Int
    Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer) As Integer
    Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Integer, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    Declare Function GetWindow Lib "user32" (ByVal hWnd As Integer, ByVal uCmd As UInt32) As Integer
    Declare Function SetParent Lib "user32" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
    Declare Function GetActiveWindow Lib "user32" () As Long
    Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
    Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
    Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Declare Function SendMessageByString Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
    Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Integer) As Integer
    Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Integer, ByVal lpString As String, ByVal cch As Integer) As Integer
    Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer

    Public Const GWL_STYLE = (-16)
    Public Const WS_POPUP = &H80000000
    Public Const WM_SETTEXT As Integer = &HC
    Public Const WM_GETTEXT As Integer = &HD
    Public Const WM_GETTEXTLENGTH As Integer = &HE
    Public Const WM_KEYDOWN As Integer = &H100
    Public Const VK_RETURN As Integer = &HD
    Public Const CC_COLOR As String = ""
    Public Const CC_BOLD As String = ""
    Public Const GW_HWNDFIRST = 0
    Public Const GW_HWNDLAST = 1
    Public Const GW_HWNDNEXT = 2
    Public Const GW_HWNDPREV = 3
    Public Const GW_OWNER = 4
    Public Const GW_CHILD = 5
    Private Const WM_SETREDRAW As Integer = 11
    Public colTable(15) As Color

    Public Structure colElements ' Voodoo2 Elements
        Shared _colDefault As String = 6
        Shared _colAction As String = 6
        Shared _colDialog As String = 1
        Shared _colTelepathy As String = 14
        Shared _colForeign As String = 2
        Shared _colNext As String = 14
        Shared _colCont As String = 14
        Shared _colBegin As String = 14
        Shared _colDone As String = 14
        Shared _txtNext As String = "--"
        Shared _txtCont As String = "--"
        Shared _txtBegin As String = vbNullString
        Shared _txtDone As String = vbNullString
        Shared _txtDialog As String = Chr(34)
        Shared _txtForeign As String = Chr(42)
        Shared _txtTelepathy As String = Chr(126)
    End Structure

    Public Sub ResumeDrawing(ByVal Target As Control, ByVal Redraw As Boolean)
        SendMessage(Target.Handle, WM_SETREDRAW, 1, 0)

        If Redraw Then
            Target.Refresh()
        End If

    End Sub

    Public Sub SuspendDrawing(ByVal Target As Control)
        SendMessage(Target.Handle, WM_SETREDRAW, 0, 0)
    End Sub


    Public Sub ResumeDrawing(ByVal Target As Control)
        ResumeDrawing(Target, True)
    End Sub
    Sub mircSendText(ByVal txtToSend As String, ByVal hwnd As Integer)

        Dim richeditw As Long
        Dim FullText As String = "/say " + txtToSend

        richeditw = FindWindowEx(hwnd, 0&, "richedit20w", vbNullString)

        Call SendMessageByString(richeditw, WM_SETTEXT, 0&, FullText)
        Call PostMessage(richeditw, WM_KEYDOWN, VK_RETURN, 0)
    End Sub
    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Dim cnt As Integer = 0

        For Each test As Char In value
            If test = ch Then cnt += 1
        Next

        Return cnt
    End Function
    Public Function StrToByteArray(ByVal str As String) As Byte()
        Dim encoding As New System.Text.UTF8Encoding()
        Return encoding.GetBytes(str)
    End Function 'StrToByteArray
    Public Function GetChannelNames() As String()
        Dim regex As String = "([#&][^\x07\x2C\s]{0,200})"
        Dim reg As New System.Text.RegularExpressions.Regex(regex)
        Dim mirc As Integer
        Dim mdiclient As Integer
        Dim mircchannel As Integer
        Dim cName(1) As String
        Dim i As Integer = 0
        Dim TheText As String, TL As Long

        ' First Channel
        mirc = FindWindow("mirc", vbNullString)
        mdiclient = FindWindowEx(mirc, 0, "mdiclient", vbNullString)
        mircchannel = FindWindowEx(mdiclient, 0, "mirc_channel", vbNullString)
        'Create buffer
        TL = GetWindowTextLength(mircchannel)
        TheText = Space(TL + 1)
        ' Fill with Text and remove null byte
        GetWindowText(mircchannel, TheText, TL)
        TheText = Left(TheText, TL)
        Dim mat As System.Text.RegularExpressions.Match = reg.Match(TheText)

        If mat.Success Then cName(i) = mat.ToString

        ' Next Channels
        Do Until mircchannel = 0
            mircchannel = GetWindow(mircchannel, GW_HWNDNEXT)
            TL = GetWindowTextLength(mircchannel)
            TheText = Space(TL + 1)
            GetWindowText(mircchannel, TheText, TL)
            TheText = Left(TheText, TL)
            Dim mat2 As System.Text.RegularExpressions.Match = reg.Match(TheText)

            If mat2.Success Then
                i = i + 1
                ReDim Preserve cName(i)
                cName(i) = mat2.ToString
            End If

        Loop

        Return cName ' Array of Open Channel Names

    End Function
    Public Function GetChannelHandle() As Integer()
        ' Get Channel Handles
        Dim regex As String = "([#&][^\x07\x2C\s]{0,200})"
        Dim reg As New System.Text.RegularExpressions.Regex(regex)
        Dim mirc As Integer
        Dim mdiclient As Integer
        Dim mircchannel As Integer
        Dim cclass(1) As Integer
        Dim cName(1) As String
        Dim i As Integer = 0
        Dim TheText As String, TL As Long

        ' First Channel
        mirc = FindWindow("mirc", vbNullString)
        mirc = FindWindow("mirc", vbNullString)
        mdiclient = FindWindowEx(mirc, 0&, "mdiclient", vbNullString)
        mircchannel = FindWindowEx(mdiclient, 0&, "mirc_channel", vbNullString)

        TL = GetWindowTextLength(mircchannel)
        TheText = Space(TL + 1)
        GetWindowText(mircchannel, TheText, TL)
        TheText = Left(TheText, TL)
        Dim mat As System.Text.RegularExpressions.Match = reg.Match(TheText)
        If mat.Success Then cclass(i) = mircchannel
        ' Next Channels
        Do Until mircchannel = 0
            mircchannel = GetWindow(mircchannel, GW_HWNDNEXT)
            TL = GetWindowTextLength(mircchannel)
            TheText = Space(TL + 1)
            GetWindowText(mircchannel, TheText, TL)
            TheText = Left(TheText, TL)
            Dim mat2 As System.Text.RegularExpressions.Match = reg.Match(TheText)
            If mat2.Success And mat2.Length > 0 Then
                i = i + 1
                ReDim Preserve cclass(i)
                cclass(i) = mircchannel
            End If
        Loop

        Return cclass ' Array of Open Channel Handles
    End Function
    Public Sub FillForm()

        With My.Settings

            If .Standalone = True Then
                colTable(0) = .Color0
                colTable(1) = .Color1
                colTable(2) = .Color2
                colTable(3) = .Color3
                colTable(4) = .Color4
                colTable(5) = .Color5
                colTable(6) = .Color6
                colTable(7) = .Color7
                colTable(8) = .Color8
                colTable(9) = .Color9
                colTable(10) = .Color10
                colTable(11) = .Color11
                colTable(12) = .Color12
                colTable(13) = .Color13
                colTable(14) = .Color14
                colTable(15) = .Color15
                colElements._txtBegin = .txtBegin
                colElements._txtCont = .txtContinue
                colElements._txtNext = .txtNext
                colElements._txtDone = .txtDone
                colElements._txtDialog = .txtDialog
                colElements._txtForeign = .txtForeign
                colElements._txtTelepathy = .txtTelepathy
                colElements._txtBegin = .txtBegin
                colElements._colDefault = .colDefault
                colElements._colDialog = .colDialog
                colElements._colForeign = .colForeign
                colElements._colTelepathy = .colTelepathy
                colElements._colBegin = .colBegin
                colElements._colNext = .colNext
                colElements._colDone = .colDone
                colElements._colCont = .colContinue
                frmMain.cmbProfiles.Enabled = False
                frmMain.LoadProfile("None")
            Else
                Dim searchPattern As String = "*.clr"
                Dim ndi As DirectoryInfo = New DirectoryInfo(.VPath)
                Dim fnames() As FileInfo = ndi.GetFiles(searchPattern, SearchOption.AllDirectories)
                Dim pFile As String
                Dim iniSettings As New IniReader(fnames(0).FullName)
                frmMain.cmbProfiles.Items.Clear()

                For Each cntIndex In fnames
                    pFile = Path.GetFileNameWithoutExtension(cntIndex.FullName.ToString)
                    frmMain.cmbProfiles.Items.Add(pFile)
                Next

                frmMain.cmbProfiles.Enabled = True
                frmMain.cmbProfiles.SelectedIndex = 0
                frmMain.LoadProfile(fnames(0).FullName.ToString)

            End If

            frmMain.rchBoxMain.ForeColor = colTable(colElements._colDefault)
            frmMain.rchBoxMain.BackColor = .CanvasColor
            frmMain.rchBoxMain.Font = .CanvasFont
            frmMain.ToolStripMenuItem31.Checked = .AlwaysTop
            frmMain.ToolStripMenuItem32.Checked = .ClearSend
            frmMain.CheckBox1.Checked = .AutoColor
            frmMain.CheckBox2.Checked = .CutPreview

        End With
    End Sub
End Module