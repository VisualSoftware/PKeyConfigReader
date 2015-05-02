Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Imports System.Text
Imports System.Xml
Imports System.Net
Imports System.IO
Imports System.Security.Cryptography

Public Class Form2

    Private DicPos As New Dictionary(Of Integer, Integer)
    Dim XML As String
    Dim RandomClass As New Random()
    Dim PartNumberFinder As String
    Dim newkey As String
    Dim genkeys As Integer = 0
    Dim IsBruteForcerRunning As Boolean = False
    Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
    Declare Auto Function PidGenX Lib "pidgenx.dll" (ByVal one As String, ByVal two As String, ByVal three As String, ByVal four As Integer, ByVal five As IntPtr, ByVal six As IntPtr, ByVal seven As IntPtr) As Integer

    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Form1.PictureBox2.Enabled = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim serial As String
        serial = TextBox2.Text
        'This Function will accept a product key and check it to ensure validity...
        'Basic error checking to ensure 'serial' is not an empty value.
        If serial = Nothing Then Exit Sub

        'Create memory spaces to pass to, and accept data from pidgenx.dll
        Dim genPID As IntPtr = Marshal.AllocHGlobal(100)
        Marshal.WriteByte(genPID, 0, &H32)
        Dim clearGenPID As Integer = 0
        For clearGenPID = 1 To 99 'Clear out memory space...
            Marshal.WriteByte(genPID, clearGenPID, &H0)
        Next clearGenPID

        Dim oldPID As IntPtr = Marshal.AllocHGlobal(164)
        Marshal.WriteByte(oldPID, 0, &HA4)
        Dim clearOldPID As Integer = 0
        For clearOldPID = 1 To 163 'Clear out memory space...
            Marshal.WriteByte(oldPID, clearOldPID, &H0)
        Next clearOldPID

        Dim DPID4 As IntPtr = Marshal.AllocHGlobal(1272)
        Marshal.WriteByte(DPID4, 0, &HF8)
        Marshal.WriteByte(DPID4, 1, &H4)
        Dim clearDPID4 As Integer = 0
        For clearDPID4 = 2 To 1271 'Clear out memory space...
            Marshal.WriteByte(DPID4, clearDPID4, &H0)
        Next clearDPID4

        'Set location of pkeyconfig.xrm-ms (needed by pidgenx.dll to verify key)...
        Dim pkeyconfig As String = XML 'Environment.GetFolderPath(Environment.SpecialFolder.System) + "\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms"
        'Call PidGenX() to determine if key is valid...
        Dim RetID As Integer = PidGenX(serial, pkeyconfig, "XXXXX", 0, genPID, oldPID, DPID4)

        'Check returned value 'RetID' for valid key...
        If RetID = 0 Then

        ElseIf RetID = -2147024893 Then
            Panel1.Visible = True
            Label3.Text = "Missing or corrupted file: PIDGENX.DLL"
        ElseIf RetID = -2147024894 Then
            Panel1.Visible = True
            Label3.Text = "Missing or corrupted file: pkeyconfig.xrm-ms"
        ElseIf RetID = -2147024809 Then
            Panel1.Visible = True
            Label3.Text = "Invalid Key"
        ElseIf RetID = -1979645695 Then
            Panel1.Visible = True
            Label3.Text = "The specified key does not work with the loaded PKeyConfig file"
        Else
            Panel1.Visible = True
            Label3.Text = "Unsupported PKeyConfig file or Invalid Key"
        End If

        'Parse out pertinent information...
        If RetID = 0 Then
            Dim pidb As Byte() = New Byte(99) {}
            For i As Integer = 0 To pidb.Length - 1
                pidb(i) = Marshal.ReadByte(genPID, i)
            Next
            Dim core As Byte() = New Byte(1271) {}
            For i As Integer = 0 To core.Length - 1
                core(i) = Marshal.ReadByte(DPID4, i)
            Next

            'Display the parsed information...
            Dim enc As System.Text.Encoding = System.Text.Encoding.ASCII
            Dim NEWKEY As New ListViewItem
            NEWKEY.Text = TextBox2.Text

            Dim pid As String = enc.GetString(pidb).Replace(vbNullChar, "") 'PID
            NEWKEY.SubItems.Add(pid)
            Dim epid As String = enc.GetString(core, 8, 96).Replace(vbNullChar, "") 'Extendend PID 
            NEWKEY.SubItems.Add(epid)
            Dim aid As String = enc.GetString(core, 136, 72).Replace(vbNullChar, "") 'Activation ID 
            NEWKEY.SubItems.Add(aid)
            Dim edi As String = enc.GetString(core, 280, 55).Replace(vbNullChar, "") 'Edition
            If edi = "" Then
                Try
                    Dim lvItem As ListViewItem = Form1.ListView1.FindItemWithText("{" & aid & "}", True, 0, True)
                    If (lvItem IsNot Nothing) Then
                        NEWKEY.SubItems.Add(lvItem.SubItems(2).Text)
                    End If
                Catch EX As Exception

                End Try
            Else
                NEWKEY.SubItems.Add(edi)
            End If
            Dim [sub] As String = enc.GetString(core, 888, 30).Replace(vbNullChar, "") 'SubType
            NEWKEY.SubItems.Add([sub])
            Dim lit As String = enc.GetString(core, 1016, 25).Replace(vbNullChar, "") 'License Type
            NEWKEY.SubItems.Add(lit)
            Dim lic As String = enc.GetString(core, 1144, 20).Replace(vbNullChar, "") 'License Channel
            NEWKEY.SubItems.Add(lic)
            Dim cid As String = epid.Substring(6, 5) 'Crypto ID
            If cid.Chars(0).ToString = 0 Then
                cid = cid.Remove(0, 1)
                If cid.Chars(0).ToString = 0 Then
                    cid = cid.Remove(0, 1)
                    If cid.Chars(0).ToString = 0 Then
                        cid = cid.Remove(0, 1)
                    End If
                End If
            End If
            NEWKEY.SubItems.Add(cid)
            If lit.Contains("MAK") = True Then
                NEWKEY.SubItems.Add(GetRemainingActivations(epid))
            Else
                NEWKEY.SubItems.Add("Unknown")
            End If
            ListView1.Items.Add(NEWKEY)
        End If
        'Clean up memory used and return it to Windows...
        Marshal.FreeHGlobal(genPID)
        Marshal.FreeHGlobal(oldPID)
        Marshal.FreeHGlobal(DPID4)
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        XML = Form1.lastloadedxrm.Text
        Dim Allargs As String
        If My.Application.CommandLineArgs.Count = 0 Then
            'No args
        Else
            'Args
            For i As Integer = 0 To CommandLineArgs.Count - 1
                Allargs = Allargs & " " & CommandLineArgs(i)
            Next
            If Allargs.Contains("-dev") Then
                Panel2.Visible = True
            End If
        End If
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        XML = Form1.lastloadedxrm.Text
    End Sub

    Function ExportToExcel(ByVal FileName As String, SelectedLV As ListView, WorkShetNumber As Integer)
        Try
            Dim xls As New Excel.Application
            Dim sheet As Excel.Worksheet
            Dim i As Integer
            xls.Workbooks.Add()
            sheet = xls.ActiveWorkbook.Worksheets(WorkShetNumber)
            Dim col As Integer = 1
            For j As Integer = 0 To SelectedLV.Columns.Count - 1
                sheet.Cells(1, col) = SelectedLV.Columns(j).Text.ToString
                col = col + 1
            Next
            For i = 0 To SelectedLV.Items.Count - 1
                Dim subitemscount As String = ""
                Dim columnscount As String = SelectedLV.Columns.Count
                Dim currentccount As Object = 1
                Dim currentsubcount As Integer = 0
                While currentccount <= columnscount
                    sheet.Cells(i + 2, currentccount) = SelectedLV.Items.Item(i).SubItems(currentsubcount).Text
                    currentccount = Val(currentccount) + 1
                    currentsubcount = Val(currentsubcount) + 1
                End While
            Next
            xls.ActiveWorkbook.SaveAs(FileName)
            xls.Workbooks.Close()
            xls.Quit()
        Catch ex As Exception
            MetroFramework.MetroMessageBox.Show(Me, "Error saving the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Panel1.Visible = False
        If IsBruteForcerRunning = True Then

        Else
            If TextBox2.Text.Length = 29 Then
                Try
                    Dim lvItem As ListViewItem = _ListView1.FindItemWithText(TextBox2.Text, False, 0, True)
                    If (lvItem IsNot Nothing) Then

                    Else
                        Button1_Click(sender, e)
                    End If
                Catch EX As Exception
                    Button1_Click(sender, e)
                End Try
            Else
            End If
        End If
    End Sub

    Function ConvertToKey(ByVal KeyPath As String, ByVal ValueName As String)
        Dim Key As Object = My.Computer.Registry.GetValue(KeyPath, ValueName, 0)
        Dim KeyOutput As String
        Dim Cur As Integer
        Dim Last As Integer
        Dim keypart1 As String
        Dim insert As String
        Const KeyOffset = 52 ' Offset of the first byte of key in DigitalProductId - helps in loops
        Dim isWin8 As Integer = (Key(66) \ 8) And 1 ' Check if it's Windows 8 here...
        Key(66) = (Key(66) And &HF7) Or ((isWin8 And 2) * 4) ' Replace 66 byte with logical result
        Dim Chars As String = "BCDFGHJKMPQRTVWXY2346789" ' Characters used in Windows key
        ' Standard Base24 decoding...
        For i = 24 To 0 Step -1
            Cur = 0
            For X = 14 To 0 Step -1
                Cur = Cur * 256
                Cur = Key(X + KeyOffset) + Cur
                Key(X + KeyOffset) = (Cur \ 24)
                Cur = Cur Mod 24
            Next
            KeyOutput = Mid(Chars, Cur + 1, 1) & KeyOutput
            Last = Cur
        Next
        ' If it's Windows 8, put "N" in the right place
        If (isWin8 = 1) Then
            keypart1 = Mid(KeyOutput, 2, Cur)
            insert = "N"
            KeyOutput = keypart1 & insert & Mid(KeyOutput, Cur + 2)
        End If
        ' Divide keys to 5-character parts
        Dim a = Mid(KeyOutput, 1, 5)
        Dim b = Mid(KeyOutput, 6, 5)
        Dim c = Mid(KeyOutput, 11, 5)
        Dim d = Mid(KeyOutput, 16, 5)
        Dim e = Mid(KeyOutput, 21, 5)
        ' And join them again adding dashes
        ConvertToKey = a & "-" & b & "-" & c & "-" & d & "-" & e
        ' The result of this function is now the actual product key
        Return ConvertToKey
    End Function

    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        TextBox2.Text = ConvertToKey("HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS NT\CURRENTVERSION\", "DigitalProductID")
    End Sub

    Private Sub ClearListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearListToolStripMenuItem.Click
        ListView1.Items.Clear()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim lvItem As ListViewItem = Form1.ListView1.FindItemWithText(PartNumberFinder, True, 0, True)
            If (lvItem IsNot Nothing) Then
                Dim CurrentColor As Color = arsenyisafk()
                lvItem.BackColor = CurrentColor
                Dim lvItem2 As ListViewItem = _ListView1.FindItemWithText(PartNumberFinder, True, 0, True)
                If (lvItem2 IsNot Nothing) Then
                    lvItem2.BackColor = CurrentColor

                Else

                End If
                Form1.Focus()
            Else

            End If
        Catch EX As Exception

        End Try
    End Sub

    Function arsenyisafk() As Color
        Randomize()
        Dim RColor As Integer = RandomClass.Next("0", "20")
        If RColor = 0 Then
            arsenyisafk = Color.Red
        ElseIf RColor = 1 Then
            arsenyisafk = Color.Orange
        ElseIf RColor = 2 Then
            arsenyisafk = Color.Green
        ElseIf RColor = 3 Then
            arsenyisafk = Color.Blue
        ElseIf RColor = 4 Then
            arsenyisafk = Color.Yellow
        ElseIf RColor = 5 Then
            arsenyisafk = Color.Violet
        ElseIf RColor = 6 Then
            arsenyisafk = Color.Pink
        ElseIf RColor = 7 Then
            arsenyisafk = Color.Brown
        ElseIf RColor = 8 Then
            arsenyisafk = Color.OrangeRed
        ElseIf RColor = 9 Then
            arsenyisafk = Color.YellowGreen
        ElseIf RColor = 10 Then
            arsenyisafk = Color.BlueViolet
        ElseIf RColor = 11 Then
            arsenyisafk = Color.Coral
        ElseIf RColor = 12 Then
            arsenyisafk = Color.Azure
        ElseIf RColor = 13 Then
            arsenyisafk = Color.DimGray
        ElseIf RColor = 14 Then
            arsenyisafk = Color.Lime
        ElseIf RColor = 15 Then
            arsenyisafk = Color.Peru
        ElseIf RColor = 16 Then
            arsenyisafk = Color.Salmon
        ElseIf RColor = 17 Then
            arsenyisafk = Color.Silver
        ElseIf RColor = 18 Then
            arsenyisafk = Color.Tomato
        ElseIf RColor = 19 Then
            arsenyisafk = Color.Teal
        ElseIf RColor = 20 Then
            arsenyisafk = Color.Olive
        End If

        Return arsenyisafk
    End Function
    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        Try
            For Each item As ListViewItem In ListView1.SelectedItems
                If item.Selected = True Then
                    PartNumberFinder = item.SubItems(5).Text
                Else

                End If
                Exit For
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FindPartNumberInPKeyConfigReaderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindPartNumberInPKeyConfigReaderToolStripMenuItem.Click
        Button3_Click(sender, e)
    End Sub

    Private Sub Form2_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer
            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                For Each line In IO.File.ReadAllLines(MyFiles(i))
                    TextBox2.Text = line
                Next
            Next
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub Form2_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub ExportToExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToExcelToolStripMenuItem.Click
        SaveFileDialog1.FileName = ""
        SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.FileName = "" Then

        Else
            ExportToExcel(SaveFileDialog1.FileName, ListView1, 1)
        End If
    End Sub

    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        MetroButton3.Enabled = True
        MetroButton2.Enabled = False
        genkeys = 0
        Label4.Text = "Generated: 0 keys."
        IsBruteForcerRunning = True
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        If IsBruteForcerRunning = True Then
            Dim KeyGen As RandomKeyGenerator
            Dim NumKeys As Integer
            Dim i_Keys As Integer
            Dim KeyOne As String
            Dim KeyTwo As String
            Dim KeyThree As String
            Dim KeyFour As String
            Dim KeyFive As String

            ' MODIFY THIS TO GET MORE KEYS
            NumKeys = 1
            KeyGen = New RandomKeyGenerator
            If MetroCheckBox1.Checked = True Then
                KeyGen.KeyLetters = "BCDFNGHJKMNPQRTVWNXY"
            Else
                KeyGen.KeyLetters = "BCDFGHJKMPQRTVWXY"
            End If
            KeyGen.KeyNumbers = "2346789"
            KeyGen.KeyChars = 5
            KeyOne = KeyGen.Generate()
            KeyTwo = KeyGen.Generate()
            KeyThree = KeyGen.Generate()
            KeyFour = KeyGen.Generate()
            KeyFive = KeyGen.Generate()
            newkey = KeyOne & "-" & KeyTwo & "-" & KeyThree & "-" & KeyFour & "-" & KeyFive
        Else

        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If IsBruteForcerRunning = True Then
            If MetroCheckBox1.Checked = True Then
                'WIN8 KEYZ
                If newkey.Contains("N") = True Then
                    'VALID KEY
                    TextBox2.Text = newkey
                    genkeys += 1
                    Label4.Text = "Generated: " & genkeys & " keys."
                    BackgroundWorker2.RunWorkerAsync()
                Else
                    'INVALID KEY, TRASH
                    BackgroundWorker1.RunWorkerAsync()
                End If
            Else
                'WIN7 OR EARLIER KEYZ
                TextBox2.Text = newkey
                genkeys += 1
                Label4.Text = "Generated: " & genkeys & " keys."
                BackgroundWorker2.RunWorkerAsync()
            End If
        Else
        End If
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        Try
            IsBruteForcerRunning = False
            MetroButton3.Enabled = False
            MetroButton2.Enabled = True
            BackgroundWorker1.CancelAsync()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub BackgroundWorker2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker2.DoWork
        Dim serial As String
        serial = TextBox2.Text
        'This Function will accept a product key and check it to ensure validity...
        'Basic error checking to ensure 'serial' is not an empty value.
        If serial = Nothing Then Exit Sub

        'Create memory spaces to pass to, and accept data from pidgenx.dll
        Dim genPID As IntPtr = Marshal.AllocHGlobal(100)
        Marshal.WriteByte(genPID, 0, &H32)
        Dim clearGenPID As Integer = 0
        For clearGenPID = 1 To 99 'Clear out memory space...
            Marshal.WriteByte(genPID, clearGenPID, &H0)
        Next clearGenPID

        Dim oldPID As IntPtr = Marshal.AllocHGlobal(164)
        Marshal.WriteByte(oldPID, 0, &HA4)
        Dim clearOldPID As Integer = 0
        For clearOldPID = 1 To 163 'Clear out memory space...
            Marshal.WriteByte(oldPID, clearOldPID, &H0)
        Next clearOldPID

        Dim DPID4 As IntPtr = Marshal.AllocHGlobal(1272)
        Marshal.WriteByte(DPID4, 0, &HF8)
        Marshal.WriteByte(DPID4, 1, &H4)
        Dim clearDPID4 As Integer = 0
        For clearDPID4 = 2 To 1271 'Clear out memory space...
            Marshal.WriteByte(DPID4, clearDPID4, &H0)
        Next clearDPID4

        'Set location of pkeyconfig.xrm-ms (needed by pidgenx.dll to verify key)...
        Dim pkeyconfig As String = XML 'Environment.GetFolderPath(Environment.SpecialFolder.System) + "\spp\tokens\pkeyconfig\pkeyconfig.xrm-ms"
        'Call PidGenX() to determine if key is valid...
        Dim RetID As Integer = PidGenX(serial, pkeyconfig, "XXXXX", 0, genPID, oldPID, DPID4)

        'Check returned value 'RetID' for valid key...
        If RetID = 0 Then

        ElseIf RetID = -2147024893 Then
            Panel1.Visible = True
            Label3.Text = "Missing or corrupted file: PIDGENX.DLL"
        ElseIf RetID = -2147024894 Then
            Panel1.Visible = True
            Label3.Text = "Missing or corrupted file: pkeyconfig.xrm-ms"
        ElseIf RetID = -2147024809 Then
            Panel1.Visible = True
            Label3.Text = "Invalid Key"
        ElseIf RetID = -1979645695 Then
            Panel1.Visible = True
            Label3.Text = "The specified key does not work with the loaded PKeyConfig file"
        Else
            Panel1.Visible = True
            Label3.Text = "Unsupported PKeyConfig file or Invalid Key"
        End If

        'Parse out pertinent information...
        If RetID = 0 Then
            Dim pidb As Byte() = New Byte(99) {}
            For i As Integer = 0 To pidb.Length - 1
                pidb(i) = Marshal.ReadByte(genPID, i)
            Next
            Dim core As Byte() = New Byte(1271) {}
            For i As Integer = 0 To core.Length - 1
                core(i) = Marshal.ReadByte(DPID4, i)
            Next

            'Display the parsed information...
            Dim enc As System.Text.Encoding = System.Text.Encoding.ASCII
            Dim NEWKEY As New ListViewItem
            NEWKEY.Text = TextBox2.Text

            Dim pid As String = enc.GetString(pidb).Replace(vbNullChar, "") 'PID
            NEWKEY.SubItems.Add(pid)
            Dim epid As String = enc.GetString(core, 8, 96).Replace(vbNullChar, "") 'Extendend PID 
            NEWKEY.SubItems.Add(epid)
            Dim aid As String = enc.GetString(core, 136, 72).Replace(vbNullChar, "") 'Activation ID 
            NEWKEY.SubItems.Add(aid)
            Dim edi As String = enc.GetString(core, 280, 40).Replace(vbNullChar, "") 'Edition
            NEWKEY.SubItems.Add(edi)
            Dim [sub] As String = enc.GetString(core, 888, 30).Replace(vbNullChar, "") 'SubType
            NEWKEY.SubItems.Add([sub])
            Dim lit As String = enc.GetString(core, 1016, 25).Replace(vbNullChar, "") 'License Type
            NEWKEY.SubItems.Add(lit)
            Dim lic As String = enc.GetString(core, 1144, 20).Replace(vbNullChar, "") 'License Channel
            NEWKEY.SubItems.Add(lic)
            Dim cid As String = epid.Substring(6, 5) 'Crypto ID
            NEWKEY.SubItems.Add(cid)
            NEWKEY.SubItems.Add("Not Checked")
            ListView1.Items.Add(NEWKEY)
        End If
        'Clean up memory used and return it to Windows...
        Marshal.FreeHGlobal(genPID)
        Marshal.FreeHGlobal(oldPID)
        Marshal.FreeHGlobal(DPID4)
    End Sub

    Private Sub BackgroundWorker2_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker2.RunWorkerCompleted
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Public Shared Function GetRemainingActivations(pid As String) As String
        ' Microsoft's PRIVATE KEY for HMAC-SHA256 encoding
        Dim bPrivateKey As Byte() = New Byte() {&HFE, &H31, &H98, &H75, &HFB, &H48, _
            &H84, &H86, &H9C, &HF3, &HF1, &HCE, _
            &H99, &HA8, &H90, &H64, &HAB, &H57, _
            &H1F, &HCA, &H47, &H4, &H50, &H58, _
            &H30, &H24, &HE2, &H14, &H62, &H87, _
            &H79, &HA0}

        ' XML Namespace
        Const uri As String = "http://www.microsoft.com/DRM/SL/BatchActivationRequest/1.0"

        ' Create new XML Document
        Dim xmlDoc As New XmlDocument()

        ' Create Root Element
        Dim rootElement As XmlElement = xmlDoc.CreateElement("ActivationRequest", uri)
        xmlDoc.AppendChild(rootElement)

        ' Create VersionNumber Element
        Dim versionNumber As XmlElement = xmlDoc.CreateElement("VersionNumber", rootElement.NamespaceURI)
        versionNumber.InnerText = "2.0"
        rootElement.AppendChild(versionNumber)

        ' Create RequestType Element
        Dim requestType As XmlElement = xmlDoc.CreateElement("RequestType", rootElement.NamespaceURI)
        requestType.InnerText = "2"
        rootElement.AppendChild(requestType)

        ' Create Requests Group Element
        Dim requestsGroupElement As XmlElement = xmlDoc.CreateElement("Requests", rootElement.NamespaceURI)

        ' Create Request Element
        Dim requestElement As XmlElement = xmlDoc.CreateElement("Request", requestsGroupElement.NamespaceURI)

        ' Add PID as Request Element
        Dim pidEntry As XmlElement = xmlDoc.CreateElement("PID", requestElement.NamespaceURI)
        pidEntry.InnerText = pid.Replace("XXXXX", "55041")
        requestElement.AppendChild(pidEntry)

        ' Add Request Element to Requests Group Element
        requestsGroupElement.AppendChild(requestElement)

        ' Add Requests and Request to XML Document
        rootElement.AppendChild(requestsGroupElement)

        ' Get Unicode Byte Array of XML Document
        Dim byteXml As Byte() = Encoding.Unicode.GetBytes(xmlDoc.InnerXml)

        ' Convert Byte Array to Base64
        Dim base64Xml As String = Convert.ToBase64String(byteXml)

        ' Compute Digest of the Base 64 XML Bytes
        Dim digest As String
        Using hmacsha256 As New HMACSHA256() With { _
            .Key = bPrivateKey _
        }
            digest = Convert.ToBase64String(hmacsha256.ComputeHash(byteXml))
        End Using

        ' Create SOAP Envelope for Web Request
        Dim form As String = "<?xml version=""1.0"" encoding=""utf-8""?><soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema""><soap:Body><BatchActivate xmlns=""http://www.microsoft.com/BatchActivationService""><request><Digest>REPLACEME1</Digest><RequestXml>REPLACEME2</RequestXml></request></BatchActivate></soap:Body></soap:Envelope>"
        form = form.Replace("REPLACEME1", digest)
        ' Put your Digest value (BASE64 encoded)
        form = form.Replace("REPLACEME2", base64Xml)
        ' Put your Base64 XML value (BASE64 encoded)
        Dim soapEnvelopeXml As New XmlDocument()
        soapEnvelopeXml.LoadXml(form)

        ' Create Web Request
        Dim webRequest__1 As HttpWebRequest = DirectCast(WebRequest.Create("https://activation.sls.microsoft.com/BatchActivation/BatchActivation.asmx"), HttpWebRequest)
        webRequest__1.Method = "POST"
        webRequest__1.ContentType = "text/xml; charset=""utf-8"""
        webRequest__1.Headers.Add("SOAPAction", "http://www.microsoft.com/BatchActivationService/BatchActivate")

        ' Insert SOAP Envelope into Web Request
        Using stream As Stream = webRequest__1.GetRequestStream()
            soapEnvelopeXml.Save(stream)
        End Using

        ' Begin Async call to Web Request
        Dim asyncResult As IAsyncResult = webRequest__1.BeginGetResponse(Nothing, Nothing)

        ' Suspend Thread until call is complete
        asyncResult.AsyncWaitHandle.WaitOne()

        ' Get the Response from the completed Web Request
        Dim soapResult As String
        Using webResponse As WebResponse = webRequest__1.EndGetResponse(asyncResult)

            ' ReSharper disable AssignNullToNotNullAttribute
            Using rd As New StreamReader(webResponse.GetResponseStream())
                ' ReSharper restore AssignNullToNotNullAttribute
                soapResult = rd.ReadToEnd()
            End Using
        End Using

        ' Parse the ResponseXML from the Response
        Using soapReader As XmlReader = XmlReader.Create(New StringReader(soapResult))
            ' Read ResponseXML Value
            soapReader.ReadToFollowing("ResponseXml")
            Dim responseXml As String = soapReader.ReadElementContentAsString()

            ' Remove HTML Entities from ResponseXML
            responseXml = responseXml.Replace("&gt;", ">")
            responseXml = responseXml.Replace("&lt;", "<")

            ' Change Encoding Value in ResponseXML
            responseXml = responseXml.Replace("utf-16", "utf-8")

            ' Read Fixed ResponseXML Value as XML
            Using reader As XmlReader = XmlReader.Create(New StringReader(responseXml))
                reader.ReadToFollowing("ActivationRemaining")
                Dim count As String = reader.ReadElementContentAsString()

                If Convert.ToInt32(count) < 0 Then
                    reader.ReadToFollowing("ErrorCode")
                    Dim [error] As String = reader.ReadElementContentAsString()

                    If [error] = "0x67" Then
                        Return "0 (Blocked)"
                    End If
                End If
                Return count
            End Using
        End Using
    End Function

    Private Sub CopyKeyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyKeyToolStripMenuItem.Click
        Clipboard.SetText(TextBox2.Text)
    End Sub

    Private Sub PasteKeyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteKeyToolStripMenuItem.Click
        Try
            TextBox2.Text = Clipboard.GetText()
        Catch ex As Exception
        End Try
    End Sub
End Class

Public Class RandomKeyGenerator
    Dim Key_Letters As String
    Dim Key_Numbers As String
    Dim Key_Chars As Integer
    Dim LettersArray As Char()
    Dim NumbersArray As Char()

    ''' <date>27072005</date><time>071924</time>
    ''' <type>property</type>
    ''' <summary>
    ''' WRITE ONLY PROPERTY. HAS TO BE SET BEFORE CALLING GENERATE()
    ''' </summary>
    Protected Friend WriteOnly Property KeyLetters() As String
        Set(ByVal Value As String)
            Key_Letters = Value
        End Set
    End Property

    ''' <date>27072005</date><time>071924</time>
    ''' <type>property</type>
    ''' <summary>
    ''' WRITE ONLY PROPERTY. HAS TO BE SET BEFORE CALLING GENERATE()
    ''' </summary>
    Protected Friend WriteOnly Property KeyNumbers() As String
        Set(ByVal Value As String)
            Key_Numbers = Value
        End Set
    End Property

    ''' <date>27072005</date><time>071924</time>
    ''' <type>property</type>
    ''' <summary>
    ''' WRITE ONLY PROPERTY. HAS TO BE SET BEFORE CALLING GENERATE()
    ''' </summary>
    Protected Friend WriteOnly Property KeyChars() As Integer
        Set(ByVal Value As Integer)
            Key_Chars = Value
        End Set
    End Property

    ''' <date>27072005</date><time>072344</time>
    ''' <type>function</type>
    ''' <summary>
    ''' GENERATES A RANDOM STRING OF LETTERS AND NUMBERS.
    ''' LETTERS CAN BE RANDOMLY CAPITAL OR SMALL.
    ''' </summary>
    ''' <returns type="String">RETURNS THE
    '''         RANDOMLY GENERATED KEY</returns>
    Function Generate() As String
        Dim i_key As Integer
        Dim Random1 As Single
        Dim arrIndex As Int16
        Dim sb As New StringBuilder
        Dim RandomLetter As String

        ' CONVERT LettersArray & NumbersArray TO CHARACTR ARRAYS
        LettersArray = Key_Letters.ToCharArray
        NumbersArray = Key_Numbers.ToCharArray

        For i_key = 1 To Key_Chars
            ' START THE CLOCK
            Randomize()
            Random1 = Rnd()
            arrIndex = -1
            ' IF THE VALUE IS AN EVEN NUMBER WE GENERATE A LETTER,
            ' OTHERWISE WE GENERATE A NUMBER  
            ' THE NUMBER '111' WAS RANDOMLY CHOSEN. ANY NUMBER
            ' WILL DO, WE JUST NEED TO BRING THE VALUE
            ' ABOVE '0'
            If (CType(Random1 * 111, Integer)) Mod 2 = 0 Then
                ' GENERATE A RANDOM INDEX IN THE LETTERS
                ' CHARACTER ARRAY
                Do While arrIndex < 0
                    arrIndex = _
                     Convert.ToInt16(LettersArray.GetUpperBound(0) _
                     * Random1)
                Loop
                RandomLetter = LettersArray(arrIndex)
                ' CREATE ANOTHER RANDOM NUMBER. IF IT IS ODD,
                ' WE CAPITALIZE THE LETTER
                If (CType(arrIndex * Random1 * 99, Integer)) Mod 2 <> 0 Then
                    RandomLetter = LettersArray(arrIndex).ToString
                    RandomLetter = RandomLetter.ToUpper
                End If
                sb.Append(RandomLetter)
            Else
                'GENERATE A RANDOM INDEX IN THE NUMBERS
                'CHARACTER ARRAY
                Do While arrIndex < 0
                    arrIndex = _
                      Convert.ToInt16(NumbersArray.GetUpperBound(0) _
                      * Random1)
                Loop
                sb.Append(NumbersArray(arrIndex))
            End If
        Next
        Return sb.ToString
    End Function
End Class