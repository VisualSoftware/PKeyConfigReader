'============================================================================
'
'    PKeyConfigReader
'    Copyright (C) 2013 - 2015 Visual Software Corporation
'
'    Author: ASV93
'    File: Form1.vb
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License along
'    with this program; if not, write to the Free Software Foundation, Inc.,
'    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
'============================================================================

Imports System.Xml
Imports Microsoft.Office.Interop
Imports System.Text
Imports System.Reflection

Public Class Form1
    Dim VSTools As VSSharedSource = New VSSharedSource
    Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
    Dim ShellEXE As String
    Dim AutoOpen As String
    Dim LongMode As Integer
    Dim BGColor As Color = Color.FromArgb(0, 174, 219)
    Dim HoverColor As Color = Color.FromArgb(0, 204, 219)
    Dim PressedColor As Color = Color.FromArgb(0, 144, 219)
    Dim pidconfigdata As String
    Dim loadfromstring As Integer = 0
    Dim IgnoreReservePN As Integer
    Dim FileNameXML As String
    Dim SelectedLV As ListView
    Dim WorkShetNumber As Integer = 1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
        Button5_Click(sender, e)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SaveFileDialog1.FileName = ""
        SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.FileName = "" Then

        Else
            'SAVE
            If MetroTabControl1.SelectedIndex = 0 Then
                'SAVE PKEYCONFIG PAGE
                FileNameXML = SaveFileDialog1.FileName
                Panel3.Visible = True
                Timer3.Enabled = True
                MetroTabControl1.Enabled = False
                Panel2.Enabled = False
                exporttoxlsworker.RunWorkerAsync(ListView1)
            ElseIf MetroTabControl1.SelectedIndex = 1 Then
                'SAVE INFORMATION PAGE
                FileNameXML = SaveFileDialog1.FileName
                Panel3.Visible = True
                Timer3.Enabled = True
                MetroTabControl1.Enabled = False
                Panel2.Enabled = False
                exporttoxlsworker.RunWorkerAsync(ListView3)
            ElseIf MetroTabControl1.SelectedIndex = 2 Then
                'SAVE POLICIES PAGE
                FileNameXML = SaveFileDialog1.FileName
                Panel3.Visible = True
                Timer3.Enabled = True
                MetroTabControl1.Enabled = False
                Panel2.Enabled = False
                exporttoxlsworker.RunWorkerAsync(ListView2)
            ElseIf MetroTabControl1.SelectedIndex = 3 Then
                'SAVE EDITIONMATRIX PAGE
                FileNameXML = SaveFileDialog1.FileName
                Panel3.Visible = True
                Timer3.Enabled = True
                MetroTabControl1.Enabled = False
                Panel2.Enabled = False
                exporttoxlsworker.RunWorkerAsync(ListView4)
            ElseIf MetroTabControl1.SelectedIndex = 4 Then
                'SAVE UPGRADEMATRIX PAGE
                Try
                    Dim xls As New Excel.Application
                    Dim sheet As Excel.Worksheet
                    Dim i As Integer
                    xls.Workbooks.Add()
                    sheet = xls.ActiveWorkbook.Worksheets(1)
                    Dim col As Integer = 1
                    For j As Integer = 0 To ListView5.Columns.Count - 1
                        sheet.Cells(1, col) = ListView5.Columns(j).Text.ToString
                        col = col + 1
                    Next
                    For i = 0 To ListView5.Items.Count - 1
                        Dim subitemscount As String = ""
                        Dim columnscount As String = ListView5.Columns.Count
                        Dim currentccount As Object = 1
                        Dim currentsubcount As Integer = 0
                        While currentccount <= columnscount
                            sheet.Cells(i + 2, currentccount) = ListView5.Items.Item(i).SubItems(currentsubcount).Text
                            currentccount = Val(currentccount) + 1
                            currentsubcount = Val(currentsubcount) + 1
                        End While
                    Next
                    Dim xlWorkSheet1 As Excel.Worksheet
                    xlWorkSheet1 = CType(xls.ActiveWorkbook.Worksheets.Add(), Excel.Worksheet)
                    xlWorkSheet1.Name = "VersionRanges"
                    xlWorkSheet1.Move(After:=sheet) 'Move the new sheet after the original
                    xlWorkSheet1.Select() 'Select the sheet and enter data
                    sheet = xls.ActiveWorkbook.Worksheets(2)
                    Dim col1 As Integer = 1
                    For j As Integer = 0 To ListView6.Columns.Count - 1
                        sheet.Cells(1, col1) = ListView6.Columns(j).Text.ToString
                        col1 = col1 + 1
                    Next
                    For i = 0 To ListView6.Items.Count - 1
                        Dim subitemscount As String = ""
                        Dim columnscount As String = ListView6.Columns.Count
                        Dim currentccount As Object = 1
                        Dim currentsubcount As Integer = 0
                        While currentccount <= columnscount
                            sheet.Cells(i + 2, currentccount) = ListView6.Items.Item(i).SubItems(currentsubcount).Text
                            currentccount = Val(currentccount) + 1
                            currentsubcount = Val(currentsubcount) + 1
                        End While
                    Next
                    xls.ActiveWorkbook.SaveAs(SaveFileDialog1.FileName)
                    xls.Workbooks.Close()
                    xls.Quit()
                Catch ex As Exception
                    MetroFramework.MetroMessageBox.Show(Me, "Error saving the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            Else
                MetroFramework.MetroMessageBox.Show(Me, "Please select a valid tabpage to export", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        End If
    End Sub

    Function ExportToExcel(ByVal FileName As String, SelectedLV As ListView, WorkShetNumber As Integer)

    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MetroTabControl1.SelectedIndex = 0
        MetroTabControl1.SelectedIndex = 1
        MetroTabControl1.SelectedIndex = 2
        MetroTabControl1.SelectedIndex = 0
        Control.CheckForIllegalCrossThreadCalls = False
        If My.User.IsInRole(ApplicationServices.BuiltInRole.Administrator) = True Then
            Me.Text = Me.Text & " (Administrator)"
        End If
        Label3.Text = "Registered to: " & Environment.UserName
        If My.Application.Info.CompanyName = "Visual Software" Then
            If Now.Year > 2013 Then
                linklabel1.Text = My.Application.Info.AssemblyName & " © 2013-" & Now.Year & " " & My.Application.Info.CompanyName
            Else
                linklabel1.Text = My.Application.Info.AssemblyName & " © 2013" & " " & My.Application.Info.CompanyName
            End If
        Else
            MetroFramework.MetroMessageBox.Show(Me, "Error, This application has been modified", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
        End If
        Dim myFileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo([Assembly].GetExecutingAssembly().Location)
        label8.Text = "Version " & myFileVersionInfo.ProductVersion
        Try
            If IO.File.Exists(My.Application.Info.DirectoryPath & "\Setup.upd") = True Then
                If IO.File.Exists(My.Application.Info.DirectoryPath & "\Setup.exe") = True Then
                    IO.File.Delete(My.Application.Info.DirectoryPath & "\Setup.exe")
                Else

                End If
                My.Computer.FileSystem.RenameFile(My.Application.Info.DirectoryPath & "\Setup.upd", "Setup.exe")
            End If
        Catch ex As Exception

        End Try
        If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF2.dat") = True Then
            Dim reader As String
            reader = (IO.File.ReadAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF2.dat"))
            If reader = "1" Then
                CheckBox2.Checked = True
            End If
        Else

        End If
        If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF3.dat") = True Then
            Dim reader As String
            reader = (IO.File.ReadAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF3.dat"))
            If reader = "1" Then
                CheckBox3.Checked = True
            End If
        Else

        End If
        If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF4.dat") = True Then
            Dim reader As String
            reader = (IO.File.ReadAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF4.dat"))
            If reader = "1" Then
                CheckBox4.Checked = True
            End If
        Else

        End If
        If IO.File.Exists(My.Application.Info.DirectoryPath & "\pidgenx.dll") = False Then
            Dim NewPIDGenX() As Byte = My.Resources.pidgenx
            My.Computer.FileSystem.WriteAllBytes(My.Application.Info.DirectoryPath & "\pidgenx.dll", NewPIDGenX, False)
        End If
        If My.Application.CommandLineArgs.Count = 0 Then
            'No args
        Else
            'Args
            For i As Integer = 0 To CommandLineArgs.Count - 1
                ShellEXE = ShellEXE & " " & CommandLineArgs(i)
            Next
            If ShellEXE.Contains("-dev") Then
                If CommandLineArgs.Count = 1 Then

                Else
                    OpenFileDialog1.FileName = CommandLineArgs(1)
                    Button5_Click(sender, e)
                End If
            Else
                OpenFileDialog1.FileName = ShellEXE
                Button5_Click(sender, e)
            End If
        End If
    End Sub

    Function Base64Decoder(ByVal SourceText As String) As String
        Dim decodedBytes As Byte()
        decodedBytes = Convert.FromBase64String(SourceText)
        Dim decodedText As String
        decodedText = Encoding.UTF8.GetString(decodedBytes)
        Return decodedText
    End Function

    Function getthefuckingoutput(ByVal InputXML As String) As String
        Dim xmlDocument As System.Xml.XmlDocument = New System.Xml.XmlDocument()
        xmlDocument.Load(OpenFileDialog1.FileName)
        Dim innerText As String = xmlDocument.SelectSingleNode("/*[local-name()='licenseGroup']/*[local-name()='license']/*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='pkeyConfigData']").InnerText
        Return Base64Decoder(innerText)
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        End
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Process.Start("http://visualsoftware.wordpress.com")
    End Sub

    Private Sub DonateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DonateToolStripMenuItem.Click
        VSTools.OpenDonationPage()
    End Sub

    Private Sub VisualSoftCorpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VisualSoftCorpToolStripMenuItem.Click
        Process.Start("https://www.twitter.com/VisualSoftCorp")
    End Sub

    Private Sub ASV93ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ASV93ToolStripMenuItem.Click
        Process.Start("https://www.twitter.com/ASV93")
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If BackgroundWorker1.IsBusy = True Then

        Else
            BackgroundWorker1.RunWorkerAsync()
        End If
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Panel1.Visible = True
        Button4.Enabled = False
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If OpenFileDialog1.FileName = "" Then

        Else
            Dim typeofxrm As String = IO.File.ReadAllText(OpenFileDialog1.FileName)
            If typeofxrm.Contains("pkeyConfigData") = True Then
                lastloadedxrm.Text = OpenFileDialog1.FileName
                Form2.Button2_Click(sender, e)
                MetroTabControl1.SelectedIndex = 0
                If CheckBox2.Checked = True Then
                    ListView1.Columns.Clear()
                Else
                    ListView1.Clear()
                End If
                Dim DecryptedXML As String = getthefuckingoutput(OpenFileDialog1.FileName)
                'LOAD
                Try
                    Dim PartNumber As String
                    Dim EulaType As String
                    Dim IsValid As String
                    Dim Start As String
                    Dim EndS As String
                    Dim doc As New XmlDocument
                    DecryptedXML = DecryptedXML.Replace("<09>", "")
                    If DecryptedXML.Contains("ProductFamilyCode") Then
                        'LONGHORN XML
                        ListView1.Columns.Add("ActConfigID")
                        ListView1.Columns.Add("RefGroupID")
                        ListView1.Columns.Add("ProductFamily")
                        ListView1.Columns.Add("ProductFamilyCode")
                        ListView1.Columns.Add("ProductName")
                        ListView1.Columns.Add("ProductVersion")
                        ListView1.Columns.Add("ProductVersionCode")
                        ListView1.Columns.Add("ProductDescription")
                        ListView1.Columns.Add("ProductKeyType")
                        ListView1.Columns.Add("IsRandomized")
                        ListView1.Columns.Add("PartNumber")
                        ListView1.Columns.Add("EULAType")
                        ListView1.Columns.Add("IsValid")
                        ListView1.Columns.Add("Start")
                        ListView1.Columns.Add("End")
                        ListView1.Columns.Add("Total Keys")
                    Else
                        ListView1.Columns.Add("ActConfigID")
                        ListView1.Columns.Add("RefGroupID")
                        ListView1.Columns.Add("EditionID")
                        ListView1.Columns.Add("ProductDescription")
                        ListView1.Columns.Add("ProductKeyType")
                        ListView1.Columns.Add("IsRandomized")
                        ListView1.Columns.Add("PartNumber")
                        ListView1.Columns.Add("EULAType")
                        ListView1.Columns.Add("IsValid")
                        ListView1.Columns.Add("Start")
                        ListView1.Columns.Add("End")
                        ListView1.Columns.Add("Total Keys")
                    End If
                    doc.LoadXml(DecryptedXML)
                    Dim doc1 As New XmlDocument
                    doc1.LoadXml(DecryptedXML)
                    Dim GTotalKeys As String = ""
                    Dim LHXML As Integer = 0
                    Dim nodes As XmlNodeList = doc.SelectNodes("/*[local-name()='ProductKeyConfiguration']/*[local-name()='Configurations']/*[local-name()='Configuration']") '("ProductKeyConfiguration/Configurations/Configuration")
                    For Each node As XmlNode In nodes
                        Dim PKEYRED As New ListViewItem
                        If DecryptedXML.Contains("ProductFamilyCode") Then
                            'LONGHORN XML
                            LHXML = 1
                            Dim ActConfigId As String = node.SelectSingleNode("*[local-name()='ActConfigId']").InnerText
                            Dim RefGroupId As String = node.SelectSingleNode("*[local-name()='RefGroupId']").InnerText
                            Dim ProductFamily As String = node.SelectSingleNode("*[local-name()='ProductFamily']").InnerText
                            Dim ProductFamilyCode As String = node.SelectSingleNode("*[local-name()='ProductFamilyCode']").InnerText
                            Dim ProductName As String = node.SelectSingleNode("*[local-name()='ProductName']").InnerText
                            Dim ProductVersion As String = node.SelectSingleNode("*[local-name()='ProductVersion']").InnerText
                            Dim ProductVersionCode As String = node.SelectSingleNode("*[local-name()='ProductVersionCode']").InnerText
                            Dim ProductDescription As String = node.SelectSingleNode("*[local-name()='ProductDescription']").InnerText
                            Dim ProductKeyType As String = node.SelectSingleNode("*[local-name()='ProductKeyType']").InnerText
                            Dim IsRandomized As String = node.SelectSingleNode("*[local-name()='IsRandomized']").InnerText
                            PKEYRED.Text = ActConfigId
                            PKEYRED.SubItems.Add(RefGroupId)
                            PKEYRED.SubItems.Add(ProductFamily)
                            PKEYRED.SubItems.Add(ProductFamilyCode)
                            PKEYRED.SubItems.Add(ProductName)
                            PKEYRED.SubItems.Add(ProductVersion)
                            PKEYRED.SubItems.Add(ProductVersionCode)
                            PKEYRED.SubItems.Add(ProductDescription)
                            PKEYRED.SubItems.Add(ProductKeyType)
                            PKEYRED.SubItems.Add(IsRandomized)
                        Else
                            Dim ActConfigId As String = node.SelectSingleNode("*[local-name()='ActConfigId']").InnerText
                            Dim RefGroupId As String = node.SelectSingleNode("*[local-name()='RefGroupId']").InnerText
                            Dim EditionId As String = node.SelectSingleNode("*[local-name()='EditionId']").InnerText
                            Dim ProductDescription As String = node.SelectSingleNode("*[local-name()='ProductDescription']").InnerText
                            Dim ProductKeyType As String = node.SelectSingleNode("*[local-name()='ProductKeyType']").InnerText
                            Dim IsRandomized As String = node.SelectSingleNode("*[local-name()='IsRandomized']").InnerText
                            PKEYRED.Text = ActConfigId
                            PKEYRED.SubItems.Add(RefGroupId)
                            PKEYRED.SubItems.Add(EditionId)
                            PKEYRED.SubItems.Add(ProductDescription)
                            PKEYRED.SubItems.Add(ProductKeyType)
                            PKEYRED.SubItems.Add(IsRandomized)
                        End If
                        PartNumber = ""
                        EulaType = ""
                        IsValid = ""
                        Start = ""
                        EndS = ""
                        Dim TotalKeys As String = ""
                        Dim nodes2 As XmlNodeList = doc1.SelectNodes("/*[local-name()='ProductKeyConfiguration']/*[local-name()='KeyRanges']/*[local-name()='KeyRange']")
                        For Each node1 As XmlNode In nodes2
                            Dim newnode As String = node1.SelectSingleNode("*[local-name()='RefActConfigId']").InnerText
                            If newnode = PKEYRED.Text Then
                                If LongMode = 1 Then
                                    'LONG MODE
                                    Try
                                        If PartNumber = "" Then
                                            PartNumber = node1.SelectSingleNode("*[local-name()='PartNumber']").InnerText
                                        Else
                                            PartNumber = PartNumber & vbCrLf & node1.SelectSingleNode("*[local-name()='PartNumber']").InnerText
                                        End If
                                    Catch ex As Exception
                                    End Try
                                    Try
                                        If EulaType = "" Then
                                            EulaType = node1.SelectSingleNode("*[local-name()='EulaType']").InnerText
                                        Else
                                            EulaType = EulaType & vbCrLf & node1.SelectSingleNode("*[local-name()='EulaType']").InnerText
                                        End If
                                    Catch ex As Exception

                                    End Try
                                    Try
                                        If IsValid = "" Then
                                            IsValid = node1.SelectSingleNode("*[local-name()='IsValid']").InnerText
                                        Else
                                            IsValid = IsValid & vbCrLf & node1.SelectSingleNode("*[local-name()='IsValid']").InnerText
                                        End If
                                    Catch ex As Exception

                                    End Try
                                    Try
                                        If Start = "" Then
                                            Start = node1.SelectSingleNode("*[local-name()='Start']").InnerText
                                        Else
                                            Start = Start & vbCrLf & node1.SelectSingleNode("*[local-name()='Start']").InnerText
                                        End If
                                    Catch ex As Exception

                                    End Try
                                    Try
                                        If EndS = "" Then
                                            EndS = node1.SelectSingleNode("*[local-name()='End']").InnerText
                                        Else
                                            EndS = EndS & vbCrLf & node1.SelectSingleNode("*[local-name()='End']").InnerText
                                        End If
                                    Catch ex As Exception

                                    End Try
                                Else
                                    Dim PKEYREDEX As New ListViewItem
                                    Dim PNEX As String
                                    Dim ETEX As String
                                    Dim IVEX As String
                                    Dim STEX As String
                                    Dim EDEX As String
                                    'PARTNUMBER
                                    Try
                                        If PartNumber = "" Then
                                            PartNumber = node1.SelectSingleNode("*[local-name()='PartNumber']").InnerText
                                            PKEYRED.SubItems.Add(PartNumber)
                                        Else
                                            PNEX = node1.SelectSingleNode("*[local-name()='PartNumber']").InnerText
                                            PKEYREDEX.SubItems.Add(" ")
                                            PKEYREDEX.SubItems.Add(" ")
                                            PKEYREDEX.SubItems.Add(" ")
                                            PKEYREDEX.SubItems.Add(" ")
                                            PKEYREDEX.SubItems.Add(" ")
                                            If LHXML = 1 Then
                                                PKEYREDEX.SubItems.Add(" ")
                                                PKEYREDEX.SubItems.Add(" ")
                                                PKEYREDEX.SubItems.Add(" ")
                                                PKEYREDEX.SubItems.Add(" ")
                                            End If
                                            PKEYREDEX.SubItems.Add(PNEX)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                    'EULATYPE
                                    Try
                                        If node1.SelectSingleNode("*[local-name()='EulaType']").InnerText = "" Then
                                            EulaType = "Unknown"
                                            ETEX = "Unknown"
                                        End If
                                        If EulaType = "" Then
                                            EulaType = node1.SelectSingleNode("*[local-name()='EulaType']").InnerText
                                            PKEYRED.SubItems.Add(EulaType)
                                        Else
                                            ETEX = node1.SelectSingleNode("*[local-name()='EulaType']").InnerText
                                            PKEYREDEX.SubItems.Add(ETEX)
                                        End If

                                    Catch ex As Exception
                                        PKEYRED.SubItems.Add(EulaType)
                                        PKEYREDEX.SubItems.Add(ETEX)
                                    End Try
                                    'ISVALID
                                    Try
                                        If IsValid = "" Then
                                            IsValid = node1.SelectSingleNode("*[local-name()='IsValid']").InnerText
                                            PKEYRED.SubItems.Add(IsValid)
                                        Else
                                            IVEX = node1.SelectSingleNode("*[local-name()='IsValid']").InnerText
                                            PKEYREDEX.SubItems.Add(IVEX)
                                        End If
                                    Catch ex As Exception

                                    End Try
                                    Try
                                        If Start = "" Then
                                            Start = node1.SelectSingleNode("*[local-name()='Start']").InnerText
                                            PKEYRED.SubItems.Add(Start)

                                        Else
                                            STEX = node1.SelectSingleNode("*[local-name()='Start']").InnerText
                                            PKEYREDEX.SubItems.Add(STEX)
                                        End If

                                    Catch ex As Exception

                                    End Try
                                    Try
                                        If EndS = "" Then
                                            EndS = node1.SelectSingleNode("*[local-name()='End']").InnerText
                                            PKEYRED.SubItems.Add(EndS)
                                            Dim calc As String = Val(EndS - Start) + 1
                                            PKEYRED.SubItems.Add(calc)
                                            TotalKeys = Val(calc)
                                            ListView1.Items.Add(PKEYRED)
                                        Else
                                            EDEX = node1.SelectSingleNode("*[local-name()='End']").InnerText
                                            PKEYREDEX.SubItems.Add(EDEX)
                                            PKEYRED.SubItems.Add(EndS)
                                            Dim calc As String = Val(EDEX - STEX) + 1
                                            PKEYREDEX.SubItems.Add(calc)
                                            If IgnoreReservePN = 1 Then
                                                If PNEX.Contains("res") = True Then
                                                    ListView1.Items.Add(PKEYREDEX)
                                                Else
                                                    TotalKeys = Val(TotalKeys) + Val(calc)
                                                    ListView1.Items.Add(PKEYREDEX)
                                                End If
                                            Else
                                                TotalKeys = Val(TotalKeys) + Val(calc)
                                                ListView1.Items.Add(PKEYREDEX)
                                            End If
                                        End If

                                    Catch ex As Exception

                                    End Try
                                End If
                            Else
                                'no matches
                            End If
                        Next
                        Dim TotalKeysEntry As New ListViewItem
                        TotalKeysEntry.SubItems.Add(" ")
                        TotalKeysEntry.SubItems.Add(" ")
                        TotalKeysEntry.SubItems.Add(" ")
                        TotalKeysEntry.SubItems.Add(" ")
                        TotalKeysEntry.SubItems.Add(" ")
                        TotalKeysEntry.SubItems.Add(" ")
                        TotalKeysEntry.SubItems.Add(" ")
                        TotalKeysEntry.SubItems.Add(" ")
                        TotalKeysEntry.SubItems.Add(" ")
                        If LHXML = 1 Then
                            TotalKeysEntry.SubItems.Add(" ")
                            TotalKeysEntry.SubItems.Add(" ")
                            TotalKeysEntry.SubItems.Add(" ")
                            TotalKeysEntry.SubItems.Add(" ")
                        End If
                        TotalKeysEntry.SubItems.Add("[SUBTOTAL]")
                        TotalKeysEntry.SubItems.Add(TotalKeys)
                        ListView1.Items.Add(TotalKeysEntry)
                        GTotalKeys = Val(GTotalKeys) + Val(TotalKeys)
                    Next
                    Dim GTotalKeysEntry As New ListViewItem
                    GTotalKeysEntry.SubItems.Add(" ")
                    GTotalKeysEntry.SubItems.Add(" ")
                    GTotalKeysEntry.SubItems.Add(" ")
                    GTotalKeysEntry.SubItems.Add(" ")
                    GTotalKeysEntry.SubItems.Add(" ")
                    GTotalKeysEntry.SubItems.Add(" ")
                    GTotalKeysEntry.SubItems.Add(" ")
                    GTotalKeysEntry.SubItems.Add(" ")
                    GTotalKeysEntry.SubItems.Add(" ")
                    If LHXML = 1 Then
                        GTotalKeysEntry.SubItems.Add(" ")
                        GTotalKeysEntry.SubItems.Add(" ")
                        GTotalKeysEntry.SubItems.Add(" ")
                        GTotalKeysEntry.SubItems.Add(" ")
                    End If
                    GTotalKeysEntry.SubItems.Add("[TOTAL]")
                    GTotalKeysEntry.SubItems.Add(GTotalKeys)
                    ListView1.Items.Add(GTotalKeysEntry)
                Catch ex As Exception
                    MetroFramework.MetroMessageBox.Show(Me, "ERROR: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            ElseIf typeofxrm.Contains("<InvalidRanges") Then
                loadfromstring = 0
                Button12_Click(sender, e)
            ElseIf typeofxrm.Contains("VersionRanges") Then
                'UpgradeMatrix
                loadfromstring = 0
                Button14_Click(sender, e)
            ElseIf typeofxrm.Contains("TmiMatrix") Then
                'EditionMatrix
                loadfromstring = 0
                Button13_Click(sender, e)
            Else
                MetroTabControl1.SelectedIndex = 1
                ListView1.Clear()
            End If
            Button8_Click(sender, e)

        End If
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim MyFiles() As String
            Dim i As Integer
            ' Assign the files to an array.
            MyFiles = e.Data.GetData(DataFormats.FileDrop)
            ' Loop through the array and add the files to the list.
            For i = 0 To MyFiles.Length - 1
                OpenFileDialog1.FileName = MyFiles(i)
            Next
            Button5_Click(sender, e)
        End If
    End Sub

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Panel1.Visible = False
        Button4.Enabled = True
        If CheckBox2.Checked = True Then
            If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF2.dat") = True Then
                IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF2.dat")
            Else

            End If
            IO.File.WriteAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF2.dat", "1")
        Else
            If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF2.dat") = True Then
                IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF2.dat")
            Else

            End If
        End If
        If CheckBox3.Checked = True Then
            If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF3.dat") = True Then
                IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF3.dat")
            Else

            End If
            IO.File.WriteAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF3.dat", "1")
        Else
            If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF3.dat") = True Then
                IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF3.dat")
            Else

            End If
        End If
        If CheckBox4.Checked = True Then
            If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF4.dat") = True Then
                IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF4.dat")
            Else

            End If
            IO.File.WriteAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF4.dat", "1")
        Else
            If IO.File.Exists(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF4.dat") = True Then
                IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "PKEYCONF4.dat")
            Else

            End If
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Timer1.Enabled = False
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If OpenFileDialog1.FileName = "" Then

        Else
            If CheckBox2.Checked = True Then
                ListView3.Columns.Clear()
            Else
                ListView3.Clear()
            End If
            ListView3.Columns.Add("Title")
            ListView3.Columns.Add("LicenseType")
            ListView3.Columns.Add("LicenseCategory")
            ListView3.Columns.Add("LicenseVersion")
            ListView3.Columns.Add("LicensorURL")
            ListView3.Columns.Add("IssuanceCertificateID")
            ListView3.Columns.Add("ProductSKUID")
            ListView3.Columns.Add("ApplicationID")
            ListView3.Columns.Add("PKeyConfigLicenseID")
            ListView3.Columns.Add("ProductName")
            ListView3.Columns.Add("PrivateCertificateID")
            ListView3.Columns.Add("PublicCertificateID")
            ListView3.Columns.Add("WinBranding")
            ListView3.Columns.Add("ProductAuthor")
            ListView3.Columns.Add("ProductDescription")
            ListView3.Columns.Add("PAURL")
            ListView3.Columns.Add("ActivationSequence")
            ListView3.Columns.Add("ValidationTemplateID")
            ListView3.Columns.Add("ValURL")
            ListView3.Columns.Add("UXDifferentiator")
            ListView3.Columns.Add("Family")
            ListView3.Columns.Add("ProductKeyGroupUniqueness")
            ListView3.Columns.Add("EnableNotificationMode")
            ListView3.Columns.Add("GraceTimerUniqueness")
            ListView3.Columns.Add("ValidityTimerUniqueness")
            ListView3.Columns.Add("EnableActivationValidation")
            ListView3.Columns.Add("ApplicationBitmap")
            ListView3.Columns.Add("HWID:ootGrace")
            ListView3.Columns.Add("Migratable")
            ListView3.Columns.Add("ReferralData")
            ListView3.Columns.Add("VLPolicy")
            ListView3.Columns.Add("LicensorKeyIndex")
            ListView3.Columns.Add("BuildVersion")
            ListView3.Columns.Add("EnforceClientClockSync")
            ListView3.Columns.Add("ServerAuthorizationTemplate")
            ListView3.Columns.Add("ClientIssuanceCertificateID")
            ListView3.Columns.Add("DependsOn")
            ListView3.Columns.Add("RuleSetData")
            ListView3.Columns.Add("RuleSetType")
            ListView3.Columns.Add("LicenseNamespace")
            ListView3.Columns.Add("PhonePolicy")
            ListView3.Columns.Add("DecryptionCertificateID")
            ListView3.Columns.Add("AppXLOB")
            ListView3.Columns.Add("ProductInstaller")
            ListView3.Columns.Add("RightsIssuanceCertificateID")
            ListView3.Columns.Add("RightsTemplateID")
            ListView3.Columns.Add("ReferralTag")
            ListView3.Columns.Add("ResellerURL")
            ListView3.Columns.Add("SPCURL")
            ListView3.Columns.Add("RACURL")
            ListView3.Columns.Add("PKCURL")
            ListView3.Columns.Add("EULURL")
            'LOAD
            Dim doc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
            doc.Load(OpenFileDialog1.FileName)
            Dim checkxrmtype As String = IO.File.ReadAllText(OpenFileDialog1.FileName)
            Dim nodes As XmlNodeList
            If checkxrmtype.Contains("<rg:licenseGroup") Then
                nodes = doc.SelectNodes("/*[local-name()='licenseGroup']/*[local-name()='license']")
            Else
                nodes = doc.SelectNodes("/*[local-name()='license']")
                If checkxrmtype.Contains("<r:allConditions") Then

                Else
                    'LONGHORN XML
                End If
            End If
            'MsgBox("fileloaded")
            For Each node As XmlNode In nodes
                Dim NONPKEY As New ListViewItem
                Dim licname As String = ""
                Try
                    licname = node.SelectSingleNode("*[local-name()='title']").InnerText
                Catch ex As Exception

                End Try
                Dim licenseType As String = ""
                Try
                    licenseType = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='licenseType']").InnerText
                Catch ex As Exception

                End Try
                Dim licenseCategory As String = ""
                Try
                    licenseCategory = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='licenseCategory']").InnerText
                Catch ex As Exception

                End Try
                Dim licenseVersion As String = ""
                Try
                    licenseVersion = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='licenseVersion']").InnerText
                Catch ex As Exception

                End Try
                Dim licensorUrl As String = ""
                Try
                    licensorUrl = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='licensorUrl']").InnerText
                Catch ex As Exception

                End Try
                Dim issuanceCertificateId As String = ""
                Try
                    issuanceCertificateId = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='issuanceCertificateId']").InnerText
                Catch ex As Exception

                End Try
                Dim productSkuId As String = ""
                Try
                    productSkuId = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='productSkuId']").InnerText
                Catch ex As Exception

                End Try
                Dim privateCertificateId As String = ""
                Dim applicationId As String = ""
                Dim pkeyConfigLicenseId As String = ""
                Dim productName As String = ""
                Dim publicCertificateId As String = ""
                Dim winbranding As String = ""
                Dim productAuthor As String = ""
                Dim productDescription As String = ""
                Dim PAUrl As String = ""
                Dim ActivationSequence As String = ""
                Dim ValidationTemplateId As String = ""
                Dim ValUrl As String = ""
                Dim UXDifferentiator As String = ""
                Dim Family As String = ""
                Dim ProductKeyGroupUniqueness As String = ""
                Dim EnableNotificationMode As String = ""
                Dim GraceTimerUniqueness As String = ""
                Dim ValidityTimerUniqueness = ""
                Dim EnableActivationValidation = ""
                Try
                    publicCertificateId = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='publicCertificateId']").InnerText
                Catch ex As Exception

                End Try
                Try
                    winbranding = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='win:branding']").InnerText
                Catch ex As Exception

                End Try
                Try
                    privateCertificateId = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='privateCertificateId']").InnerText
                Catch ex As Exception

                End Try
                Try
                    applicationId = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='applicationId']").InnerText
                Catch ex As Exception

                End Try
                Try
                    pkeyConfigLicenseId = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='pkeyConfigLicenseId']").InnerText
                Catch ex As Exception

                End Try
                Try
                    productName = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='productName']").InnerText
                Catch ex As Exception

                End Try
                Try
                    productAuthor = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='productAuthor']").InnerText
                Catch ex As Exception

                End Try
                Try
                    productDescription = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='productDescription']").InnerText
                Catch ex As Exception

                End Try
                Try
                    PAUrl = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='PAUrl']").InnerText
                Catch ex As Exception

                End Try
                Try
                    ActivationSequence = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='ActivationSequence']").InnerText
                Catch ex As Exception

                End Try
                Try
                    ValidationTemplateId = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='ValidationTemplateId']").InnerText
                Catch ex As Exception

                End Try
                Try
                    ValUrl = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='ValUrl']").InnerText
                Catch ex As Exception

                End Try
                Try
                    UXDifferentiator = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='UXDifferentiator']").InnerText
                Catch ex As Exception

                End Try
                Try
                    Family = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='Family']").InnerText
                Catch ex As Exception

                End Try
                Try
                    ProductKeyGroupUniqueness = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='ProductKeyGroupUniqueness']").InnerText
                Catch ex As Exception

                End Try
                Try
                    EnableNotificationMode = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='EnableNotificationMode']").InnerText
                Catch ex As Exception

                End Try
                Try
                    GraceTimerUniqueness = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='GraceTimerUniqueness']").InnerText
                Catch ex As Exception

                End Try
                Try
                    ValidityTimerUniqueness = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='ValidityTimerUniqueness']").InnerText
                Catch ex As Exception

                End Try
                Try
                    EnableActivationValidation = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='EnableActivationValidation']").InnerText
                Catch ex As Exception

                End Try
                Dim ApplicationBitmap As String = ""
                Dim hwidootgrace As String = ""
                Dim migratable As String = ""
                Dim referraldata As String = ""
                Dim vlpolicy As String = ""
                Dim licensorkeyindex As String = ""
                Dim buildversion As String = ""
                Dim enforceclientclocksync As String = ""
                Dim serverauthorizationtemplate As String = ""
                Dim clientissuancecertificateid As String = ""
                Try
                    ApplicationBitmap = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='ApplicationBitmap']").InnerText
                Catch ex As Exception

                End Try
                Try
                    hwidootgrace = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='hwid:ootGrace']").InnerText
                Catch ex As Exception

                End Try
                Try
                    migratable = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='migratable']").InnerText
                Catch ex As Exception

                End Try
                Try
                    referraldata = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='referralData']").InnerText
                Catch ex As Exception

                End Try
                Try
                    vlpolicy = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='vl:policy']").InnerText
                Catch ex As Exception

                End Try
                Try
                    licensorkeyindex = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='licensorKeyIndex']").InnerText
                Catch ex As Exception

                End Try
                Try
                    buildversion = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='BuildVersion']").InnerText
                Catch ex As Exception

                End Try
                Try
                    enforceclientclocksync = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='enforceClientClockSync']").InnerText
                Catch ex As Exception

                End Try
                Try
                    serverauthorizationtemplate = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='serverAuthorizationTemplate']").InnerText
                Catch ex As Exception

                End Try
                Try
                    clientissuancecertificateid = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='clientIssuanceCertificateId']").InnerText
                Catch ex As Exception

                End Try
                Dim dependson As String = ""
                Try
                    dependson = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='DependsOn']").InnerText
                Catch ex As Exception

                End Try
                Dim rulesetdata As String = ""
                Try
                    rulesetdata = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='RuleSetData']").InnerText
                Catch ex As Exception

                End Try
                Dim rulesettype As String = ""
                Try
                    rulesettype = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='RuleSetType']").InnerText
                Catch ex As Exception

                End Try
                Dim licensenamespace As String = ""
                Try
                    licensenamespace = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='licenseNamespace']").InnerText
                Catch ex As Exception

                End Try
                Dim phonepolicy As String = ""
                Try
                    phonepolicy = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='phone:policy']").InnerText
                Catch ex As Exception

                End Try
                Dim decryptioncertificateid As String = ""
                Try
                    decryptioncertificateid = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='decryptionCertificateId']").InnerText
                Catch ex As Exception

                End Try
                Dim appxlob As String = ""
                Try
                    appxlob = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='AppXLOB']").InnerText
                Catch ex As Exception

                End Try
                Dim productinstaller As String = ""
                Dim rightsissuancecertificateid As String = ""
                Dim rightstemplateid As String = ""
                Dim referraltag As String = ""
                Dim resellerurl As String = ""
                Dim spcurl As String = ""
                Dim racurl As String = ""
                Dim pkcurl As String = ""
                Dim eulurl As String = ""
                Try
                    productinstaller = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='productInstaller']").InnerText
                Catch ex As Exception

                End Try
                Try
                    rightsissuancecertificateid = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='rightsIssuanceCertificateId']").InnerText
                Catch ex As Exception

                End Try
                Try
                    rightstemplateid = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='rightsTemplateId']").InnerText
                Catch ex As Exception

                End Try
                Try
                    referraltag = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='referralTag']").InnerText
                Catch ex As Exception

                End Try
                Try
                    resellerurl = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='resellerUrl']").InnerText
                Catch ex As Exception

                End Try
                Try
                    spcurl = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='SPCUrl']").InnerText
                Catch ex As Exception

                End Try
                Try
                    racurl = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='RACUrl']").InnerText
                Catch ex As Exception

                End Try
                Try
                    pkcurl = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='PKCUrl']").InnerText
                Catch ex As Exception

                End Try
                Try
                    eulurl = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='EULUrl']").InnerText
                Catch ex As Exception

                End Try
                Try
                    pidconfigdata = node.SelectSingleNode("*[local-name()='otherInfo']/*[local-name()='infoTables']/*[local-name()='infoList']/*[@name='pidConfigData']").InnerText
                    loadfromstring = 1
                    Button12_Click(sender, e)
                Catch ex As Exception

                End Try
                NONPKEY.Text = licname
                NONPKEY.SubItems.Add(licenseType)
                NONPKEY.SubItems.Add(licenseCategory)
                NONPKEY.SubItems.Add(licenseVersion)
                NONPKEY.SubItems.Add(licensorUrl)
                NONPKEY.SubItems.Add(issuanceCertificateId)
                NONPKEY.SubItems.Add(productSkuId)
                NONPKEY.SubItems.Add(applicationId)
                NONPKEY.SubItems.Add(pkeyConfigLicenseId)
                NONPKEY.SubItems.Add(productName)
                NONPKEY.SubItems.Add(privateCertificateId)
                NONPKEY.SubItems.Add(publicCertificateId)
                NONPKEY.SubItems.Add(winbranding)
                NONPKEY.SubItems.Add(productAuthor)
                NONPKEY.SubItems.Add(productDescription)
                NONPKEY.SubItems.Add(PAUrl)
                NONPKEY.SubItems.Add(ActivationSequence)
                NONPKEY.SubItems.Add(ValidationTemplateId)
                NONPKEY.SubItems.Add(ValUrl)
                NONPKEY.SubItems.Add(UXDifferentiator)
                NONPKEY.SubItems.Add(Family)
                NONPKEY.SubItems.Add(ProductKeyGroupUniqueness)
                NONPKEY.SubItems.Add(EnableNotificationMode)
                NONPKEY.SubItems.Add(GraceTimerUniqueness)
                NONPKEY.SubItems.Add(ValidityTimerUniqueness)
                NONPKEY.SubItems.Add(EnableActivationValidation)
                NONPKEY.SubItems.Add(ApplicationBitmap)
                NONPKEY.SubItems.Add(hwidootgrace)
                NONPKEY.SubItems.Add(migratable)
                NONPKEY.SubItems.Add(referraldata)
                NONPKEY.SubItems.Add(vlpolicy)
                NONPKEY.SubItems.Add(licensorkeyindex)
                NONPKEY.SubItems.Add(buildversion)
                NONPKEY.SubItems.Add(enforceclientclocksync)
                NONPKEY.SubItems.Add(serverauthorizationtemplate)
                NONPKEY.SubItems.Add(clientissuancecertificateid)
                NONPKEY.SubItems.Add(dependson)
                NONPKEY.SubItems.Add(rulesetdata)
                NONPKEY.SubItems.Add(rulesettype)
                NONPKEY.SubItems.Add(licensenamespace)
                NONPKEY.SubItems.Add(phonepolicy)
                NONPKEY.SubItems.Add(decryptioncertificateid)
                NONPKEY.SubItems.Add(appxlob)
                NONPKEY.SubItems.Add(productinstaller)
                NONPKEY.SubItems.Add(rightsissuancecertificateid)
                NONPKEY.SubItems.Add(rightstemplateid)
                NONPKEY.SubItems.Add(referraltag)
                NONPKEY.SubItems.Add(resellerurl)
                NONPKEY.SubItems.Add(spcurl)
                NONPKEY.SubItems.Add(racurl)
                NONPKEY.SubItems.Add(pkcurl)
                NONPKEY.SubItems.Add(eulurl)
                ListView3.Items.Add(NONPKEY)
            Next
            Button10_Click(sender, e)
            If CheckBox3.Checked = True Then

            Else
                Button11_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ListView1.Items.Clear()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox2.Checked = True Then
            Button9.Enabled = True
        Else
            Button9.Enabled = False
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If OpenFileDialog1.FileName = "" Then

        Else
            If CheckBox2.Checked = True Then
                ListView2.Columns.Clear()
            Else
                ListView2.Clear()
            End If
            'LOAD
            Dim policycheck As String = IO.File.ReadAllText(OpenFileDialog1.FileName)
            If policycheck.Contains("<sl:policy") Then
                ListView2.Columns.Add("Name")
                ListView2.Columns.Add("Value")
                ListView2.Columns.Add("Type")
            Else

            End If
            'POLICY INT
            Try
                Dim doc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
                doc.Load(OpenFileDialog1.FileName)
                Dim nodes As XmlNodeList
                Dim checkxrmtype As String = IO.File.ReadAllText(OpenFileDialog1.FileName)
                If checkxrmtype.Contains("<rg:licenseGroup") Then
                    If checkxrmtype.Contains(TextBox1.Text) Then
                        nodes = doc.SelectNodes("/*[local-name()='licenseGroup']/*[local-name()='license']/*[local-name()='grant']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policyInt']")
                    Else
                        nodes = doc.SelectNodes("/*[local-name()='licenseGroup']/*[local-name()='license']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policyInt']")
                    End If
                Else
                    If checkxrmtype.Contains("<r:allConditions") Then
                        nodes = doc.SelectNodes("/*[local-name()='license']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policyInt']")
                    Else
                        'LONGHORN XML
                        nodes = doc.SelectNodes("/*[local-name()='license']/*[local-name()='grant']/*[local-name()='productPolicies']/*[local-name()='policyInt']")
                    End If
                End If

                For Each node As XmlNode In nodes
                    Dim XPLORERKEY As New ListViewItem
                    Dim columntitle As String = node.Attributes(0).InnerText 'node.OuterXml
                    Dim value As String = node.InnerText
                    XPLORERKEY.Text = columntitle
                    XPLORERKEY.SubItems.Add(value)
                    XPLORERKEY.SubItems.Add("INT")
                    ListView2.Items.Add(XPLORERKEY)
                Next
            Catch ex As Exception
            End Try
            'POLICY STR
            Try
                Dim doc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
                doc.Load(OpenFileDialog1.FileName)
                Dim nodes As XmlNodeList
                Dim checkxrmtype As String = IO.File.ReadAllText(OpenFileDialog1.FileName)
                If checkxrmtype.Contains("<rg:licenseGroup") Then
                    If checkxrmtype.Contains(TextBox1.Text) Then
                        nodes = doc.SelectNodes("/*[local-name()='licenseGroup']/*[local-name()='license']/*[local-name()='grant']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policyStr']")
                    Else
                        nodes = doc.SelectNodes("/*[local-name()='licenseGroup']/*[local-name()='license']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policyStr']")
                    End If
                Else
                    If checkxrmtype.Contains("<r:allConditions") Then
                        nodes = doc.SelectNodes("/*[local-name()='license']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policyStr']")
                    Else
                        'LONGHORN XML
                        nodes = doc.SelectNodes("/*[local-name()='license']/*[local-name()='grant']/*[local-name()='productPolicies']/*[local-name()='policyStr']")
                    End If
                End If

                For Each node As XmlNode In nodes
                    Dim XPLORERKEY As New ListViewItem
                    Dim columntitle As String = node.Attributes(0).InnerText
                    Dim value As String = node.InnerText
                    XPLORERKEY.Text = columntitle
                    XPLORERKEY.SubItems.Add(value)
                    XPLORERKEY.SubItems.Add("STR")
                    ListView2.Items.Add(XPLORERKEY)
                Next
            Catch ex As Exception
            End Try
            'POLICY BIN
            Try
                Dim doc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
                doc.Load(OpenFileDialog1.FileName)
                Dim nodes As XmlNodeList
                Dim checkxrmtype As String = IO.File.ReadAllText(OpenFileDialog1.FileName)
                If checkxrmtype.Contains("<rg:licenseGroup") Then
                    If checkxrmtype.Contains(TextBox1.Text) Then
                        nodes = doc.SelectNodes("/*[local-name()='licenseGroup']/*[local-name()='license']/*[local-name()='grant']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policyBin']")
                    Else
                        nodes = doc.SelectNodes("/*[local-name()='licenseGroup']/*[local-name()='license']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policyBin']")
                    End If
                Else
                    If checkxrmtype.Contains("<r:allConditions") Then
                        nodes = doc.SelectNodes("/*[local-name()='license']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policyBin']")
                    Else
                        'LONGHORN XML
                        nodes = doc.SelectNodes("/*[local-name()='license']/*[local-name()='grant']/*[local-name()='productPolicies']/*[local-name()='policyBin']")
                    End If
                End If

                For Each node As XmlNode In nodes
                    Dim XPLORERKEY As New ListViewItem
                    Dim columntitle As String = node.Attributes(0).InnerText
                    Dim value As String = node.InnerText
                    XPLORERKEY.Text = columntitle
                    XPLORERKEY.SubItems.Add(value)
                    XPLORERKEY.SubItems.Add("BIN")
                    ListView2.Items.Add(XPLORERKEY)
                Next
            Catch ex As Exception
            End Try
            'POLICY SUM
            Try
                Dim doc As System.Xml.XmlDocument = New System.Xml.XmlDocument()
                doc.Load(OpenFileDialog1.FileName)
                Dim nodes As XmlNodeList
                Dim checkxrmtype As String = IO.File.ReadAllText(OpenFileDialog1.FileName)
                If checkxrmtype.Contains("<rg:licenseGroup") Then
                    If checkxrmtype.Contains(TextBox1.Text) Then
                        nodes = doc.SelectNodes("/*[local-name()='licenseGroup']/*[local-name()='license']/*[local-name()='grant']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policySum']")
                    Else
                        nodes = doc.SelectNodes("/*[local-name()='licenseGroup']/*[local-name()='license']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policySum']")
                    End If
                Else
                    If checkxrmtype.Contains("<r:allConditions") Then
                        nodes = doc.SelectNodes("/*[local-name()='license']/*[local-name()='grant']/*[local-name()='allConditions']/*[local-name()='productPolicies']/*[local-name()='policySum']")
                    Else
                        'LONGHORN XML
                        nodes = doc.SelectNodes("/*[local-name()='license']/*[local-name()='grant']/*[local-name()='productPolicies']/*[local-name()='policySum']")
                    End If
                End If

                For Each node As XmlNode In nodes
                    Dim XPLORERKEY As New ListViewItem
                    Dim columntitle As String = node.Attributes(0).InnerText
                    Dim value As String = node.InnerText
                    XPLORERKEY.Text = columntitle
                    XPLORERKEY.SubItems.Add(value)
                    XPLORERKEY.SubItems.Add("SUM")
                    ListView2.Items.Add(XPLORERKEY)
                Next
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub ClearItemListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearItemListToolStripMenuItem.Click
        If MetroTabControl1.SelectedIndex = 0 Then
            ListView1.Items.Clear()
        ElseIf MetroTabControl1.SelectedIndex = 1 Then
            ListView3.Items.Clear()
        ElseIf MetroTabControl1.SelectedIndex = 2 Then
            ListView2.Items.Clear()
        ElseIf MetroTabControl1.SelectedIndex = 3 Then
            ListView4.Items.Clear()
        Else
            ListView5.Items.Clear()
            ListView6.Items.Clear()
        End If
    End Sub

    Private Sub PictureBox10_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox10.MouseEnter
        PictureBox10.BackColor = HoverColor
    End Sub

    Private Sub PictureBox10_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox10.MouseLeave
        PictureBox10.BackColor = BGColor
    End Sub

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click
        PictureBox10.BackColor = PressedColor
        With OpenFileDialog1
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                Button5_Click(sender, e)
            End If
        End With
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        PictureBox4.BackColor = PressedColor
        Button2_Click(sender, e)
    End Sub

    Private Sub PictureBox4_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox4.MouseEnter
        PictureBox4.BackColor = HoverColor
    End Sub

    Private Sub PictureBox4_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox4.MouseLeave
        PictureBox4.BackColor = BGColor
    End Sub

    Private Sub linklabel1_Click(sender As Object, e As EventArgs) Handles linklabel1.Click
        Process.Start("http://visualsoftware.wordpress.com")
    End Sub

    Private Sub PictureBox13_Click(sender As Object, e As EventArgs) Handles PictureBox13.Click
        Process.Start("https://www.twitter.com/VisualSoftCorp")
    End Sub

    Private Sub PictureBox14_Click(sender As Object, e As EventArgs) Handles PictureBox14.Click
        VSTools.OpenDonationPage()
    End Sub

    Private Sub MetroButton11_Click(sender As Object, e As EventArgs) Handles MetroButton11.Click
        Try
            If IO.File.Exists(My.Application.Info.DirectoryPath & "\SrvVer.txt") = True Then
                IO.File.Delete(IO.File.Exists(My.Application.Info.DirectoryPath & "\SrvVer.txt"))
            End If
        Catch ex As Exception
        End Try
        MetroButton11.Enabled = False
        Timer2.Enabled = True
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Button6_Click(sender, e)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim ItemsCount As Integer = ListView3.Items.Count
        Dim ColumnsCount As Integer = ListView3.Columns.Count
        Dim CurrentSubItem As Integer = 0
        Dim CurrentItem As Integer = 0
        Dim CurrentColumn As Integer = 0
        Dim SubItemCount As Integer = ListView3.Columns.Count
        Dim RemoveIt As Integer = 0
        While CurrentSubItem < SubItemCount
            RemoveIt = 1
            For Each ListViewItem In ListView3.Items
                If ListView3.Items.Item(CurrentItem).SubItems(CurrentSubItem).Text = "" Then
                Else
                    RemoveIt = 0
                End If
                CurrentItem += 1
            Next
            CurrentItem = 0
            If RemoveIt = 1 Then
                ListView3.Columns.RemoveAt(CurrentColumn)

                CurrentSubItem += 1
            Else
                CurrentSubItem += 1
                CurrentColumn += 1
            End If
        End While
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If OpenFileDialog1.FileName = "" Then

        Else
            MetroTabControl1.SelectedIndex = 0
            If CheckBox2.Checked = True Then
                ListView1.Columns.Clear()
            Else
                ListView1.Clear()
            End If
            'LOAD
            Try
                Dim doc As New XmlDocument
                ListView1.Columns.Add("ProductName")
                ListView1.Columns.Add("ProductVersion")
                ListView1.Columns.Add("PKeyAlgVersion")
                ListView1.Columns.Add("GroupID")
                ListView1.Columns.Add("PubKey")
                ListView1.Columns.Add("SKUID")
                ListView1.Columns.Add("PKeyType")
                ListView1.Columns.Add("OEM")
                ListView1.Columns.Add("InvalidRange-Start")
                ListView1.Columns.Add("InvalidRange-End")
                ListView1.Columns.Add("ValidRange-Start")
                ListView1.Columns.Add("ValidRange-End")
                ListView1.Columns.Add("Randomization")
                If loadfromstring = 1 Then
                    doc.LoadXml(pidconfigdata)
                Else
                    doc.Load(OpenFileDialog1.FileName)
                End If
                Dim nodes As XmlNodeList = doc.SelectNodes("/*[local-name()='ConfigData']/*[local-name()='PubInfoList']/*[local-name()='PubInfo']") '("ProductKeyConfiguration/Configurations/Configuration")
                For Each node As XmlNode In nodes
                    Dim PKEYLH As New ListViewItem
                    Dim ProductName As String = node.SelectSingleNode("*[local-name()='ProductName']").InnerText
                    Dim ProductVersion As String = node.SelectSingleNode("*[local-name()='ProductVersion']").InnerText
                    Dim PKeyAlgVersion As String = node.SelectSingleNode("*[local-name()='PKeyAlgVersion']").InnerText
                    Dim GroupID As String = node.SelectSingleNode("*[local-name()='GroupID']").InnerText
                    Dim PubKey As String = node.SelectSingleNode("*[local-name()='PubKey']").InnerText
                    Dim SKUID As String = node.SelectSingleNode("*[local-name()='SKUID']").InnerText
                    Dim PKeyType As String = node.SelectSingleNode("*[local-name()='PKeyType']").InnerText
                    Dim OEM As String = node.SelectSingleNode("*[local-name()='OEM']").InnerText
                    Dim IRStart As String = ""
                    Try
                        IRStart = node.SelectSingleNode("*[local-name()='InvalidRanges']/*[local-name()='Range']/*[local-name()='Start']").InnerText
                    Catch ex As Exception

                    End Try
                    Dim IREnd As String = ""
                    Try
                        IREnd = node.SelectSingleNode("*[local-name()='InvalidRanges']/*[local-name()='Range']/*[local-name()='End']").InnerText
                    Catch ex As Exception

                    End Try
                    Dim VRStart As String = ""
                    Try
                        VRStart = node.SelectSingleNode("*[local-name()='ValidRanges']/*[local-name()='Range']/*[local-name()='Start']").InnerText
                    Catch ex As Exception

                    End Try
                    Dim VREnd As String = ""
                    Try
                        VREnd = node.SelectSingleNode("*[local-name()='ValidRanges']/*[local-name()='Range']/*[local-name()='End']").InnerText
                    Catch ex As Exception

                    End Try
                    Dim Randomization As String = node.SelectSingleNode("*[local-name()='Randomization']").InnerText
                    PKEYLH.Text = ProductName
                    PKEYLH.SubItems.Add(ProductVersion)
                    PKEYLH.SubItems.Add(PKeyAlgVersion)
                    PKEYLH.SubItems.Add(GroupID)
                    PKEYLH.SubItems.Add(PubKey)
                    PKEYLH.SubItems.Add(SKUID)
                    PKEYLH.SubItems.Add(PKeyType)
                    PKEYLH.SubItems.Add(OEM)
                    PKEYLH.SubItems.Add(IRStart)
                    PKEYLH.SubItems.Add(IREnd)
                    PKEYLH.SubItems.Add(VRStart)
                    PKEYLH.SubItems.Add(VREnd)
                    PKEYLH.SubItems.Add(Randomization)
                    ListView1.Items.Add(PKEYLH)
                Next

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        MetroButton11.Enabled = True
        Timer2.Enabled = False
        Try
            If IO.File.Exists(My.Application.Info.DirectoryPath & "\Setup.exe") = True Then
                If IO.File.Exists(My.Application.Info.DirectoryPath & "\SrvVer.txt") = True Then
                    Dim serverversion As String = IO.File.ReadAllText(My.Application.Info.DirectoryPath & "\SrvVer.txt")
                    If Not serverversion > label8.Text.Replace("Version ", "") Then
                        MetroFramework.MetroMessageBox.Show(Me, "You are running the latest version", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MetroFramework.MetroMessageBox.Show(Me, "You are NOT running the latest version", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End If
            Else
                MetroFramework.MetroMessageBox.Show(Me, "Setup.exe not found", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        If OpenFileDialog1.FileName = "" Then

        Else
            MetroTabControl1.SelectedIndex = 3
            If CheckBox2.Checked = True Then
                ListView4.Columns.Clear()
            Else
                ListView4.Clear()
            End If
            'LOAD
            Try
                Dim doc As New XmlDocument
                ListView4.Columns.Add("EditionID")
                ListView4.Columns.Add("Name")
                ListView4.Columns.Add("ProcessorArchitecture")
                ListView4.Columns.Add("BuildType")
                ListView4.Columns.Add("PublicKeyToken")
                ListView4.Columns.Add("Version")
                ListView4.Columns.Add("TargetID")
                If loadfromstring = 1 Then
                    doc.LoadXml(pidconfigdata)
                Else
                    doc.Load(OpenFileDialog1.FileName)
                End If
                Dim nodes As XmlNodeList = doc.SelectNodes("/*[local-name()='TmiMatrix']/*[local-name()='Edition']")
                For Each node As XmlNode In nodes
                    Dim EDMATRIX As New ListViewItem
                    Dim EditionID As String = node.Attributes(0).InnerText
                    Dim EdName As String = node.Attributes(1).InnerText
                    Dim ProcessorArchitecture As String = node.Attributes(2).InnerText
                    Dim BuildType As String = ""
                    Try
                        BuildType = node.Attributes(3).InnerText
                    Catch ex As Exception

                    End Try
                    Dim PublicKeyToken As String = ""
                    Dim Version As String = ""
                    If Not BuildType.Contains("rel") Or BuildType.Contains("deb") = True Then
                        PublicKeyToken = node.Attributes(3).InnerText
                        Version = node.Attributes(4).InnerText
                        BuildType = ""
                    Else
                        PublicKeyToken = node.Attributes(4).InnerText
                        Version = node.Attributes(5).InnerText
                    End If

                    Dim TargetID As String = ""
                    Dim CurrentCNode As Integer = 0
                    Try
                        If Not node.ChildNodes.Count = 0 Then
                            While CurrentCNode <= node.ChildNodes.Count - 1
                                TargetID = TargetID & node.ChildNodes(CurrentCNode).Attributes(0).InnerText & "; "
                                CurrentCNode += 1
                            End While

                        Else

                        End If
                    Catch ex As Exception

                    End Try
                    If Not TargetID.Length = 0 Then
                        'Remove last ;
                        TargetID = TargetID.Remove(TargetID.Length - 2, 1)
                    End If
                    EDMATRIX.Text = EditionID
                    EDMATRIX.SubItems.Add(EdName)
                    EDMATRIX.SubItems.Add(ProcessorArchitecture)
                    EDMATRIX.SubItems.Add(BuildType)
                    EDMATRIX.SubItems.Add(PublicKeyToken)
                    EDMATRIX.SubItems.Add(Version)
                    EDMATRIX.SubItems.Add(TargetID)
                    ListView4.Items.Add(EDMATRIX)
                Next

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        If OpenFileDialog1.FileName = "" Then

        Else
            MetroTabControl1.SelectedIndex = 4
            If CheckBox2.Checked = True Then
                ListView5.Columns.Clear()
                ListView6.Columns.Clear()
            Else
                ListView5.Clear()
                ListView6.Clear()
            End If
            'LOAD
            Dim altload As String = "0"
            If altload = "1" Then
                Try
                    Dim doc As New XmlDocument
                    ListView5.Columns.Add("TargetEdition")
                    ListView5.Columns.Add("ProcessorArchitecture")
                    ListView5.Columns.Add("Version")
                    ListView5.Columns.Add("Features")
                    ListView5.Columns.Add("SourceEdition")
                    ListView5.Columns.Add("SE-ProcessorArchitecture")
                    ListView5.Columns.Add("SE-VersionRange")
                    ListView5.Columns.Add("SE-DataOnly")
                    ListView5.Columns.Add("SE-FullUpgrade")
                    ListView5.Columns.Add("SE-CleanInstall")
                    If loadfromstring = 1 Then
                        doc.LoadXml(pidconfigdata)
                    Else
                        doc.Load(OpenFileDialog1.FileName)
                    End If
                    Dim nodes As XmlNodeList = doc.SelectNodes("/*[local-name()='TmiMatrix']/*[local-name()='TargetEdition']")
                    For Each node As XmlNode In nodes
                        Dim UPGMATRIX As New ListViewItem
                        Dim TargetEdition As String = node.Attributes(0).InnerText
                        Dim ProcessorArchitecture As String = node.Attributes(1).InnerText
                        Dim Version As String = node.Attributes(2).InnerText
                        Dim Features As String = ""
                        Dim SourceEdition As String = ""
                        Dim SEProcessorArchitecture As String = ""
                        Dim SEVersionRange As String = ""
                        Dim SEDataOnly As String = ""
                        Dim SEFullUpgrade As String = ""
                        Dim SECleanInstall As String = ""
                        Try
                            'features
                            For Each node2 As XmlNode In node.SelectNodes("*[local-name()='Features']/*[local-name()='Feature']")
                                Features = Features & node2.Attributes(0).InnerText & "; "
                            Next
                            For Each node3 As XmlNode In node.SelectNodes("*[local-name()='SourceEdition']")
                                SourceEdition = SourceEdition & node3.Attributes(0).InnerText & vbCrLf
                                SEProcessorArchitecture = SEProcessorArchitecture & node3.Attributes(1).InnerText & vbCrLf
                                SEVersionRange = SEVersionRange & node3.Attributes(2).InnerText & vbCrLf
                                SEDataOnly = SEDataOnly & node3.Attributes(3).InnerText & vbCrLf
                                SEFullUpgrade = SEFullUpgrade & node3.Attributes(4).InnerText & vbCrLf
                                SECleanInstall = SECleanInstall & node3.Attributes(5).InnerText & vbCrLf
                            Next
                        Catch ex As Exception

                        End Try
                        UPGMATRIX.Text = TargetEdition
                        UPGMATRIX.SubItems.Add(ProcessorArchitecture)
                        UPGMATRIX.SubItems.Add(Version)
                        UPGMATRIX.SubItems.Add(Features)
                        UPGMATRIX.SubItems.Add(SourceEdition)
                        UPGMATRIX.SubItems.Add(SEProcessorArchitecture)
                        UPGMATRIX.SubItems.Add(SEVersionRange)
                        UPGMATRIX.SubItems.Add(SEDataOnly)
                        UPGMATRIX.SubItems.Add(SEFullUpgrade)
                        UPGMATRIX.SubItems.Add(SECleanInstall)
                        ListView5.Items.Add(UPGMATRIX)
                    Next

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Else
                '0
                Try
                    Dim doc As New XmlDocument
                    ListView5.Columns.Add("SourceEdition")
                    ListView5.Columns.Add("ProcessorArchitecture")
                    ListView5.Columns.Add("VersionRange")
                    ListView5.Columns.Add("DataOnly")
                    ListView5.Columns.Add("DataSetting")
                    ListView5.Columns.Add("FullUpgrade")
                    ListView5.Columns.Add("CleanInstall")
                    ListView5.Columns.Add("TargetEdition")
                    ListView5.Columns.Add("ProcessorArchitecture")
                    ListView5.Columns.Add("Version")
                    ListView5.Columns.Add("Features")
                    ListView6.Columns.Add("Name")
                    ListView6.Columns.Add("MinVersion")
                    ListView6.Columns.Add("MaxVersion")
                    ListView6.Columns.Add("MinSPVersion")
                    If loadfromstring = 1 Then
                        doc.LoadXml(pidconfigdata)
                    Else
                        doc.Load(OpenFileDialog1.FileName)
                    End If
                    Dim nodes As XmlNodeList = doc.SelectNodes("/*[local-name()='TmiMatrix']/*[local-name()='TargetEdition']")
                    For Each node As XmlNode In nodes
                        Try
                            For Each node3 As XmlNode In node.SelectNodes("*[local-name()='SourceEdition']")
                                Dim UPGMATRIX As New ListViewItem
                                Dim TargetEdition As String = node.Attributes(0).InnerText
                                Dim ProcessorArchitecture As String = node.Attributes(1).InnerText
                                Dim Version As String = node.Attributes(2).InnerText
                                Dim Features As String = ""
                                Dim SourceEdition As String = ""
                                Dim SEProcessorArchitecture As String = ""
                                Dim SEVersionRange As String = ""
                                Dim SEDataOnly As String = ""
                                Dim SEDataSetting As String = ""
                                Dim SEFullUpgrade As String = ""
                                Dim SECleanInstall As String = ""
                                If node3.Attributes(1).InnerText.Contains("win") = True Then
                                    SourceEdition = node3.Attributes(0).InnerText
                                    SEProcessorArchitecture = node3.Attributes(2).InnerText
                                    SEVersionRange = node3.Attributes(1).InnerText
                                    SEDataOnly = node3.Attributes(4).InnerText
                                    SEDataSetting = node3.Attributes(5).InnerText
                                    SEFullUpgrade = node3.Attributes(6).InnerText
                                    SECleanInstall = node3.Attributes(3).InnerText
                                ElseIf node3.Attributes(1).InnerText.Contains("any") = True Then
                                    SourceEdition = node3.Attributes(0).InnerText
                                    SEProcessorArchitecture = node3.Attributes(2).InnerText
                                    SEVersionRange = node3.Attributes(1).InnerText
                                    SEDataOnly = node3.Attributes(4).InnerText
                                    SEDataSetting = node3.Attributes(5).InnerText
                                    SEFullUpgrade = node3.Attributes(6).InnerText
                                    SECleanInstall = node3.Attributes(3).InnerText
                                ElseIf node3.Attributes(1).InnerText.Contains("vista") = True Then
                                    SourceEdition = node3.Attributes(0).InnerText
                                    SEProcessorArchitecture = node3.Attributes(2).InnerText
                                    SEVersionRange = node3.Attributes(1).InnerText
                                    SEDataOnly = node3.Attributes(4).InnerText
                                    SEDataSetting = node3.Attributes(5).InnerText
                                    SEFullUpgrade = node3.Attributes(6).InnerText
                                    SECleanInstall = node3.Attributes(3).InnerText
                                ElseIf node3.Attributes(1).InnerText.Contains("ws") = True Then
                                    SourceEdition = node3.Attributes(0).InnerText
                                    SEProcessorArchitecture = node3.Attributes(2).InnerText
                                    SEVersionRange = node3.Attributes(1).InnerText
                                    SEDataOnly = node3.Attributes(4).InnerText
                                    SEDataSetting = node3.Attributes(5).InnerText
                                    SEFullUpgrade = node3.Attributes(6).InnerText
                                    SECleanInstall = node3.Attributes(3).InnerText
                                ElseIf node3.Attributes(1).InnerText.Contains("threshold") = True Then
                                    SourceEdition = node3.Attributes(0).InnerText
                                    SEProcessorArchitecture = node3.Attributes(2).InnerText
                                    SEVersionRange = node3.Attributes(1).InnerText
                                    SEDataOnly = node3.Attributes(4).InnerText
                                    SEDataSetting = node3.Attributes(5).InnerText
                                    SEFullUpgrade = node3.Attributes(6).InnerText
                                    SECleanInstall = node3.Attributes(3).InnerText
                                Else
                                    SourceEdition = node3.Attributes(0).InnerText
                                    SEProcessorArchitecture = node3.Attributes(1).InnerText
                                    SEVersionRange = node3.Attributes(2).InnerText
                                    SEDataOnly = node3.Attributes(3).InnerText
                                    SEDataSetting = node3.Attributes(4).InnerText
                                    SEFullUpgrade = node3.Attributes(5).InnerText
                                    SECleanInstall = node3.Attributes(6).InnerText
                                End If
                                For Each node2 As XmlNode In node.SelectNodes("*[local-name()='Features']/*[local-name()='Feature']")
                                    Features = Features & node2.Attributes(0).InnerText & "; "
                                Next
                                If Not Features.Length = 0 Then
                                    'Remove last ;
                                    Features = Features.Remove(Features.Length - 2, 1)
                                End If
                                UPGMATRIX.Text = SourceEdition
                                UPGMATRIX.SubItems.Add(SEProcessorArchitecture)
                                UPGMATRIX.SubItems.Add(SEVersionRange)
                                UPGMATRIX.SubItems.Add(SEDataOnly)
                                UPGMATRIX.SubItems.Add(SEDataSetting)
                                UPGMATRIX.SubItems.Add(SEFullUpgrade)
                                UPGMATRIX.SubItems.Add(SECleanInstall)
                                UPGMATRIX.SubItems.Add(TargetEdition)
                                UPGMATRIX.SubItems.Add(ProcessorArchitecture)
                                UPGMATRIX.SubItems.Add(Version)
                                UPGMATRIX.SubItems.Add(Features)
                                ListView5.Items.Add(UPGMATRIX)
                            Next
                        Catch ex As Exception

                        End Try
                    Next
                    Dim nodes123 As XmlNodeList = doc.SelectNodes("/*[local-name()='TmiMatrix']/*[local-name()='VersionRanges']/*[local-name()='Range']")
                    For Each node123 As XmlNode In nodes123
                        Dim UPGMATRIX2 As New ListViewItem
                        Dim VRName As String = node123.Attributes(0).InnerText
                        Dim MinVersion As String = node123.Attributes(1).InnerText
                        Dim MaxVersion As String = node123.Attributes(2).InnerText
                        Dim MinSPVersion As String = node123.Attributes(3).InnerText
                        UPGMATRIX2.Text = VRName
                        UPGMATRIX2.SubItems.Add(MinVersion)
                        UPGMATRIX2.SubItems.Add(MaxVersion)
                        UPGMATRIX2.SubItems.Add(MinSPVersion)
                        ListView6.Items.Add(UPGMATRIX2)
                    Next
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs)
        If CheckBox4.Checked = True Then
            IgnoreReservePN = 1
        Else
            IgnoreReservePN = 0
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If lastloadedxrm.Text = "c:\" Then
            MetroFramework.MetroMessageBox.Show(Me, "You must load a pkeyconfig.xrm-ms file before opening the Product Key Checker", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Form2.Show()
            PictureBox2.Enabled = False
        End If
    End Sub

    Private Sub PictureBox2_MouseEnter(sender As Object, e As EventArgs) Handles PictureBox2.MouseEnter
        PictureBox2.BackColor = HoverColor
    End Sub

    Private Sub PictureBox2_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox2.MouseLeave
        PictureBox2.BackColor = BGColor
    End Sub

    Private Sub MetroCheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            Button9.Enabled = True
        Else
            Button9.Enabled = False
        End If
    End Sub

    Private Sub MetroCheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            IgnoreReservePN = 1
        Else
            IgnoreReservePN = 0
        End If
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        If MetroProgressSpinner1.Value = "100" Then
            MetroProgressSpinner1.Value = "0"
            MetroProgressSpinner1.Value = MetroProgressSpinner1.Value + 1
        Else
            MetroProgressSpinner1.Value = MetroProgressSpinner1.Value + 1
        End If
    End Sub

    Private Sub exporttoxmlworker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles exporttoxlsworker.DoWork
        Try
            Dim xls As New Excel.Application
            Dim sheet As Excel.Worksheet
            Dim i As Integer
            xls.Workbooks.Add()
            sheet = xls.ActiveWorkbook.Worksheets(WorkShetNumber)
            Dim col As Integer = 1
            For j As Integer = 0 To DirectCast(e.Argument, ListView).Columns.Count - 1
                sheet.Cells(1, col) = DirectCast(e.Argument, ListView).Columns(j).Text.ToString
                col = col + 1
            Next
            For i = 0 To DirectCast(e.Argument, ListView).Items.Count - 1
                Dim subitemscount As String = ""
                Dim columnscount As String = DirectCast(e.Argument, ListView).Columns.Count
                Dim currentccount As Object = 1
                Dim currentsubcount As Integer = 0
                While currentccount <= columnscount
                    sheet.Cells(i + 2, currentccount) = DirectCast(e.Argument, ListView).Items.Item(i).SubItems(currentsubcount).Text
                    currentccount = Val(currentccount) + 1
                    currentsubcount = Val(currentsubcount) + 1
                End While
            Next
            xls.ActiveWorkbook.SaveAs(FileNameXML)
            xls.Workbooks.Close()
            xls.Quit()
        Catch ex As Exception
            MetroFramework.MetroMessageBox.Show(Me, "Error saving the file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try
    End Sub

    Private Sub exporttoxlsworker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles exporttoxlsworker.RunWorkerCompleted
        Panel3.Visible = False
        Timer3.Enabled = False
        MetroTabControl1.Enabled = True
        Panel2.Enabled = True
    End Sub
End Class
