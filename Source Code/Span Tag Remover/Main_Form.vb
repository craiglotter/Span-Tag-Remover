Public Class Main_Form
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main_Form))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.CheckBox5 = New System.Windows.Forms.CheckBox
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.GroupBox1.Controls.Add(Me.CheckBox5)
        Me.GroupBox1.Controls.Add(Me.CheckBox4)
        Me.GroupBox1.Controls.Add(Me.CheckBox3)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ProgressBar1)
        Me.GroupBox1.Controls.Add(Me.Button1)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Controls.Add(Me.CheckBox2)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.Controls.Add(Me.Panel2)
        Me.GroupBox1.Controls.Add(Me.TextBox2)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox1.Location = New System.Drawing.Point(24, 24)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(518, 424)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'CheckBox5
        '
        Me.CheckBox5.Location = New System.Drawing.Point(347, 30)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(165, 24)
        Me.CheckBox5.TabIndex = 11
        Me.CheckBox5.Text = "Clear 'height=...' Style Info"
        Me.ToolTip1.SetToolTip(Me.CheckBox5, "Clears any located tag's entire contents should it contain a ""height="" identifier" & _
                ".")
        '
        'CheckBox4
        '
        Me.CheckBox4.Location = New System.Drawing.Point(347, 10)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(165, 24)
        Me.CheckBox4.TabIndex = 10
        Me.CheckBox4.Text = "Clear 'width=...' Style Info"
        Me.ToolTip1.SetToolTip(Me.CheckBox4, "Clears any located tag's entire contents should it contain a ""width="" identifier." & _
                "")
        '
        'CheckBox3
        '
        Me.CheckBox3.Location = New System.Drawing.Point(198, 52)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(176, 24)
        Me.CheckBox3.TabIndex = 9
        Me.CheckBox3.Text = "Clear 'style=...' Style Info"
        Me.ToolTip1.SetToolTip(Me.CheckBox3, "Clears any located tag's entire contents should it contain a ""style="" identifier." & _
                "")
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Firebrick
        Me.Label1.Location = New System.Drawing.Point(15, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(177, 32)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Span Tag Remover"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(201, 384)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(287, 16)
        Me.ProgressBar1.TabIndex = 4
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.Control
        Me.Button1.Enabled = False
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(32, 384)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 25)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Process Text"
        Me.ToolTip1.SetToolTip(Me.Button1, "Process the inputted text.")
        Me.Button1.UseVisualStyleBackColor = False
        '
        'TextBox1
        '
        Me.TextBox1.AllowDrop = True
        Me.TextBox1.BackColor = System.Drawing.Color.White
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.ForeColor = System.Drawing.Color.Black
        Me.TextBox1.Location = New System.Drawing.Point(32, 80)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(456, 296)
        Me.TextBox1.TabIndex = 2
        '
        'CheckBox2
        '
        Me.CheckBox2.Location = New System.Drawing.Point(198, 30)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(176, 24)
        Me.CheckBox2.TabIndex = 1
        Me.CheckBox2.Text = "Clear 'class=...' Style Info"
        Me.ToolTip1.SetToolTip(Me.CheckBox2, "Clears any located tag's entire contents should it contain a ""class="" identifier." & _
                "" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Note: this will destroy any hyperlink tags as the ""href="" will also be remove" & _
                "d.)")
        '
        'CheckBox1
        '
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(198, 10)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(176, 24)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "Remove '<span>' Tags"
        Me.ToolTip1.SetToolTip(Me.CheckBox1, "Removes all ""<span>"" and ""</span>"" tags from the text.")
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Location = New System.Drawing.Point(200, 383)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(289, 18)
        Me.Panel2.TabIndex = 5
        '
        'TextBox2
        '
        Me.TextBox2.AllowDrop = True
        Me.TextBox2.BackColor = System.Drawing.Color.White
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox2.ForeColor = System.Drawing.Color.Black
        Me.TextBox2.Location = New System.Drawing.Point(88, 96)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox2.Size = New System.Drawing.Size(10, 10)
        Me.TextBox2.TabIndex = 7
        Me.TextBox2.Visible = False
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem7})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem4, Me.MenuItem3, Me.MenuItem6, Me.MenuItem5})
        Me.MenuItem1.Text = "Actions"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Shortcut = System.Windows.Forms.Shortcut.CtrlO
        Me.MenuItem2.Text = "Load File"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.Shortcut = System.Windows.Forms.Shortcut.CtrlS
        Me.MenuItem4.Text = "Save File"
        '
        'MenuItem3
        '
        Me.MenuItem3.Enabled = False
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Shortcut = System.Windows.Forms.Shortcut.CtrlG
        Me.MenuItem3.Text = "Process Text"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 3
        Me.MenuItem6.Text = "-"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 4
        Me.MenuItem5.Shortcut = System.Windows.Forms.Shortcut.AltF4
        Me.MenuItem5.Text = "Exit"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 1
        Me.MenuItem7.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem8})
        Me.MenuItem7.Text = "Information"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 0
        Me.MenuItem8.Text = "Help"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "All files|*.*"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "htm"
        Me.SaveFileDialog1.Filter = "All files|*.*"
        '
        'Panel1
        '
        Me.Panel1.AllowDrop = True
        Me.Panel1.BackColor = System.Drawing.Color.WhiteSmoke
        Me.Panel1.Location = New System.Drawing.Point(18, 22)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(533, 432)
        Me.Panel1.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Firebrick
        Me.Label2.Location = New System.Drawing.Point(24, 456)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(488, 32)
        Me.Label2.TabIndex = 3
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Main_Form
        '
        Me.AllowDrop = True
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(567, 496)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(573, 549)
        Me.Menu = Me.MainMenu1
        Me.MinimumSize = New System.Drawing.Size(573, 549)
        Me.Name = "Main_Form"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Span Tag Remover"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub Error_Handler(ByVal message As String)
        Try
            Label2.Text = "Error Encountered"
            ProgressBar1.Value = 0
            MsgBox("Span Tag Remover 1.0 has trapped the following error: " & vbCrLf & message, MsgBoxStyle.Exclamation, "Span Tag Remover 1.0 Error Message")
            Label2.Text = "Error Encountered"
            ProgressBar1.Value = 0
        Catch ex As Exception
            MsgBox("Span Tag Remover 1.0 has trapped the following error: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Span Tag Remover 1.0 Error Message")
        End Try
    End Sub

    Private Sub Message_Handler(ByVal message As String)
        Try
            MsgBox(message, MsgBoxStyle.OkOnly, "Span Tag Remover 1.0")
        Catch ex As Exception
            MsgBox("Span Tag Remover 1.0 has trapped the following error: " & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Span Tag Remover 1.0 Error Message")
        End Try
    End Sub

    Private Function checkcount(ByVal startindex As Integer, ByVal endindex As Integer) As Boolean
        Dim result As Boolean
        result = True
        If endindex - startindex < 0 Then
            result = False
        End If
        Return result
    End Function

    Private Sub Process_Text()
        Try

            Dim counter, spancounter, classcounter, stylecounter, widthcounter, heightcounter As Integer
            Dim processline As String
            Dim startindex, endindex As Integer
            startindex = 0
            endindex = 0
            counter = 0
            Dim checkforclose As Boolean
            checkforclose = False
            Dim checkforcloseclass As Boolean
            checkforcloseclass = False
            Dim checkforclosestyle As Boolean
            checkforclosestyle = False

            Dim checkforclosewidth As Boolean
            checkforclosewidth = False
            Dim checkforcloseheight As Boolean
            checkforcloseheight = False


            spancounter = 0
            classcounter = 0
            stylecounter = 0
            widthcounter = 0
            heightcounter = 0

            ProgressBar1.Minimum = 0
            Dim testline As String
            testline = ""
            Dim checkedcount As Integer = 0

            If CheckBox1.Checked = True Then checkedcount = checkedcount + 1
            If CheckBox2.Checked = True Then checkedcount = checkedcount + 1
            If CheckBox3.Checked = True Then checkedcount = checkedcount + 1
            If CheckBox4.Checked = True Then checkedcount = checkedcount + 1
            If CheckBox5.Checked = True Then checkedcount = checkedcount + 1
            If checkedcount > 0 Then
                ProgressBar1.Maximum = (TextBox1.Lines.Length * checkedcount) + 1
            Else

                ProgressBar1.Maximum = 0
            End If

            ProgressBar1.Value = 0
            Label2.Text = "Processing Text..."
            TextBox2.Clear()
            If CheckBox1.Checked = True Then
                For counter = 0 To (TextBox1.Lines.Length - 1)
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    processline = (TextBox1.Lines(counter))
                    If checkforclose = True Then
                        startindex = 0
                        testline = processline.ToLower
                        endindex = testline.IndexOf(">", startindex)

                        If endindex > -1 Then
                            If checkcount(startindex, (endindex + 1)) = True Then
                                processline = processline.Remove(startindex, ((endindex - startindex) + 1))
                            Else
                                MsgBox("error raised " & processline & " " & startindex)
                            End If
                            checkforclose = False
                        Else
                            processline = processline.Remove(startindex, ((processline.Length - startindex)))
                            checkforclose = True
                        End If
                    End If


                    If checkforclose = False Then
                        testline = processline.ToLower
                        startindex = testline.IndexOf("<span")
                        If startindex > -1 Then
                            spancounter = spancounter + 1
                            endindex = testline.IndexOf(">", startindex)
                            If endindex > -1 Then
                                processline = processline.Remove(startindex, ((endindex - startindex) + 1))
                                checkforclose = False
                            Else
                                processline = processline.Remove(startindex, ((processline.Length - startindex)))
                                checkforclose = True
                            End If
                        End If
                    End If



                    If checkforclose = False Then
                        testline = processline.ToLower
                        startindex = testline.IndexOf("</span")
                        If startindex > -1 Then
                            endindex = testline.IndexOf(">", startindex)
                            If endindex > -1 Then
                                processline = processline.Remove(startindex, ((endindex - startindex) + 1))
                                checkforclose = False
                            Else
                                processline = processline.Remove(startindex, ((processline.Length - startindex)))
                                checkforclose = True
                            End If
                        End If
                    End If



                    TextBox2.Text = TextBox2.Text & processline.Replace(" >", ">")
                    If counter < (TextBox1.Lines.Length - 1) Then
                        TextBox2.Text = TextBox2.Text & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
                    End If

                Next
                TextBox1.Text = TextBox2.Text
            End If
            TextBox2.Clear()
            Label2.Text = "Processing Text..."
            If CheckBox2.Checked = True Then
                For counter = 0 To (TextBox1.Lines.Length - 1)
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    processline = (TextBox1.Lines(counter))

                    If checkforcloseclass = True Then
                        startindex = 0
                        testline = processline.ToLower
                        endindex = testline.IndexOf(">", startindex)
                        If endindex > -1 Then
                            If checkcount(startindex, endindex) = True Then
                                processline = processline.Remove(startindex, (endindex - startindex))
                            Else
                                MsgBox("error raised " & processline & " " & startindex & "here ")
                            End If

                            checkforcloseclass = False
                        Else
                            If checkcount(startindex, processline.Length) = True Then
                                processline = processline.Remove(startindex, ((processline.Length - startindex)))
                            Else
                                MsgBox("error raised " & processline & " " & startindex & "here2 ")
                            End If
                            'processline = processline.Remove(startindex, ((processline.Length - startindex)))
                            checkforcloseclass = True
                        End If
                    End If


                    If checkforcloseclass = False Then
                        testline = processline.ToLower
                        startindex = testline.IndexOf("class=")
                        If startindex > -1 Then
                            classcounter = classcounter + 1
                            endindex = testline.IndexOf(">", startindex)
                            If endindex > -1 Then
                                If checkcount(startindex, endindex) = True Then
                                    processline = processline.Remove(startindex, ((endindex - startindex)))
                                Else
                                    MsgBox("error raised " & processline & " " & startindex & "here3 " & endindex)
                                End If
                                'processline = processline.Remove(startindex, ((endindex - startindex)))
                                checkforcloseclass = False
                            Else
                                If checkcount(startindex, processline.Length) = True Then
                                    processline = processline.Remove(startindex, ((processline.Length - startindex)))
                                Else
                                    MsgBox("error raised " & processline & " " & startindex & "here4 ")
                                End If
                                'processline = processline.Remove(startindex, ((processline.Length - startindex)))
                                checkforcloseclass = True
                            End If
                        End If
                    End If

                    TextBox2.Text = TextBox2.Text & processline.Replace(" >", ">")
                    If counter < (TextBox1.Lines.Length - 1) Then
                        TextBox2.Text = TextBox2.Text & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
                    End If
                Next
                TextBox1.Text = TextBox2.Text
            End If



            TextBox2.Clear()
            Label2.Text = "Processing Text..."
            If CheckBox3.Checked = True Then
                For counter = 0 To (TextBox1.Lines.Length - 1)
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    processline = (TextBox1.Lines(counter))

                    If checkforclosestyle = True Then
                        startindex = 0
                        testline = processline.ToLower
                        endindex = testline.IndexOf(">", startindex)
                        If endindex > -1 Then
                            If checkcount(startindex, endindex) = True Then
                                processline = processline.Remove(startindex, (endindex - startindex))
                            Else
                                MsgBox("error raised " & processline & " " & startindex & "here ")
                            End If

                            checkforclosestyle = False
                        Else
                            If checkcount(startindex, processline.Length) = True Then
                                processline = processline.Remove(startindex, ((processline.Length - startindex)))
                            Else
                                MsgBox("error raised " & processline & " " & startindex & "here2 ")
                            End If
                            'processline = processline.Remove(startindex, ((processline.Length - startindex)))
                            checkforclosestyle = True
                        End If
                    End If


                    If checkforclosestyle = False Then
                        testline = processline.ToLower
                        startindex = testline.IndexOf("style=")
                        If startindex > -1 Then
                            stylecounter = stylecounter + 1
                            endindex = testline.IndexOf(">", startindex)
                            If endindex > -1 Then
                                If checkcount(startindex, endindex) = True Then
                                    processline = processline.Remove(startindex, ((endindex - startindex)))
                                Else
                                    MsgBox("error raised " & processline & " " & startindex & "here3 " & endindex)
                                End If
                                'processline = processline.Remove(startindex, ((endindex - startindex)))
                                checkforclosestyle = False
                            Else
                                If checkcount(startindex, processline.Length) = True Then
                                    processline = processline.Remove(startindex, ((processline.Length - startindex)))
                                Else
                                    MsgBox("error raised " & processline & " " & startindex & "here4 ")
                                End If
                                'processline = processline.Remove(startindex, ((processline.Length - startindex)))
                                checkforclosestyle = True
                            End If
                        End If
                    End If

                    TextBox2.Text = TextBox2.Text & processline.Replace(" >", ">")
                    If counter < (TextBox1.Lines.Length - 1) Then
                        TextBox2.Text = TextBox2.Text & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
                    End If
                Next
                TextBox1.Text = TextBox2.Text
            End If




            '***********************************************************
            TextBox2.Clear()
            Label2.Text = "Processing Text..."
            If CheckBox4.Checked = True Then
                For counter = 0 To (TextBox1.Lines.Length - 1)
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    processline = (TextBox1.Lines(counter))

                    If checkforclosewidth = True Then
                        startindex = 0
                        testline = processline.ToLower
                        endindex = testline.IndexOf(">", startindex)
                        If endindex > -1 Then
                            If checkcount(startindex, endindex) = True Then
                                processline = processline.Remove(startindex, (endindex - startindex))
                            Else
                                MsgBox("error raised " & processline & " " & startindex & "here ")
                            End If

                            checkforclosewidth = False
                        Else
                            If checkcount(startindex, processline.Length) = True Then
                                processline = processline.Remove(startindex, ((processline.Length - startindex)))
                            Else
                                MsgBox("error raised " & processline & " " & startindex & "here2 ")
                            End If
                            'processline = processline.Remove(startindex, ((processline.Length - startindex)))
                            checkforclosewidth = True
                        End If
                    End If


                    If checkforclosewidth = False Then
                        testline = processline.ToLower
                        startindex = testline.IndexOf("width=")
                        If startindex > -1 Then
                            widthcounter = widthcounter + 1
                            endindex = testline.IndexOf(">", startindex)
                            If endindex > -1 Then
                                If checkcount(startindex, endindex) = True Then
                                    processline = processline.Remove(startindex, ((endindex - startindex)))
                                Else
                                    MsgBox("error raised " & processline & " " & startindex & "here3 " & endindex)
                                End If
                                'processline = processline.Remove(startindex, ((endindex - startindex)))
                                checkforclosewidth = False
                            Else
                                If checkcount(startindex, processline.Length) = True Then
                                    processline = processline.Remove(startindex, ((processline.Length - startindex)))
                                Else
                                    MsgBox("error raised " & processline & " " & startindex & "here4 ")
                                End If
                                'processline = processline.Remove(startindex, ((processline.Length - startindex)))
                                checkforclosewidth = True
                            End If
                        End If
                    End If

                    TextBox2.Text = TextBox2.Text & processline.Replace(" >", ">")
                    If counter < (TextBox1.Lines.Length - 1) Then
                        TextBox2.Text = TextBox2.Text & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
                    End If
                Next
                TextBox1.Text = TextBox2.Text
            End If
            '***********************************************************


            TextBox2.Clear()
            Label2.Text = "Processing Text..."
            If CheckBox5.Checked = True Then
                For counter = 0 To (TextBox1.Lines.Length - 1)
                    ProgressBar1.Value = ProgressBar1.Value + 1
                    processline = (TextBox1.Lines(counter))

                    If checkforcloseheight = True Then
                        startindex = 0
                        testline = processline.ToLower
                        endindex = testline.IndexOf(">", startindex)
                        If endindex > -1 Then
                            If checkcount(startindex, endindex) = True Then
                                processline = processline.Remove(startindex, (endindex - startindex))
                            Else
                                MsgBox("error raised " & processline & " " & startindex & "here ")
                            End If

                            checkforcloseheight = False
                        Else
                            If checkcount(startindex, processline.Length) = True Then
                                processline = processline.Remove(startindex, ((processline.Length - startindex)))
                            Else
                                MsgBox("error raised " & processline & " " & startindex & "here2 ")
                            End If
                            'processline = processline.Remove(startindex, ((processline.Length - startindex)))
                            checkforcloseheight = True
                        End If
                    End If


                    If checkforcloseheight = False Then
                        testline = processline.ToLower
                        startindex = testline.IndexOf("height=")
                        If startindex > -1 Then
                            heightcounter = heightcounter + 1
                            endindex = testline.IndexOf(">", startindex)
                            If endindex > -1 Then
                                If checkcount(startindex, endindex) = True Then
                                    processline = processline.Remove(startindex, ((endindex - startindex)))
                                Else
                                    MsgBox("error raised " & processline & " " & startindex & "here3 " & endindex)
                                End If
                                'processline = processline.Remove(startindex, ((endindex - startindex)))
                                checkforcloseheight = False
                            Else
                                If checkcount(startindex, processline.Length) = True Then
                                    processline = processline.Remove(startindex, ((processline.Length - startindex)))
                                Else
                                    MsgBox("error raised " & processline & " " & startindex & "here4 ")
                                End If
                                'processline = processline.Remove(startindex, ((processline.Length - startindex)))
                                checkforcloseheight = True
                            End If
                        End If
                    End If

                    TextBox2.Text = TextBox2.Text & processline.Replace(" >", ">")
                    If counter < (TextBox1.Lines.Length - 1) Then
                        TextBox2.Text = TextBox2.Text & Microsoft.VisualBasic.ChrW(13) & Microsoft.VisualBasic.ChrW(10)
                    End If
                Next
                TextBox1.Text = TextBox2.Text
            End If




            ProgressBar1.Value = 0
            Label2.Text = "Text successfully processed. " & spancounter & " '<span>' tags and " & classcounter & " 'class=...' styles and " & stylecounter & " 'style=...' styles and " & vbCrLf & widthcounter & " 'width=...' styles and " & heightcounter & " 'height=...' styles were removed."
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Try
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Process_Text()
            TextBox1.Focus()
            TextBox1.Select(0, 0)
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim aboutscreen As About_Form = New About_Form
            aboutscreen.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Try
            If TextBox1.Text.Length > 0 Then
                Button1.Enabled = True
                MenuItem3.Enabled = True
                Label2.Text = ""
            Else
                Button1.Enabled = False
                MenuItem3.Enabled = False
                Label2.Text = ""
            End If

        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            Process_Text()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub


    Private Sub Form_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragEnter, TextBox1.DragEnter
        Try
            ' Check the format of the data being dropped.
            If (e.Data.GetDataPresent(DataFormats.Text)) Or (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
                ' Display the copy cursor.
                e.Effect = DragDropEffects.Copy
            Else
                ' Display the no-drop cursor.
                e.Effect = DragDropEffects.None
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Form_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles MyBase.DragDrop, TextBox1.DragDrop
        Try
            ' Paste the text.
            If (e.Data.GetDataPresent(DataFormats.Text)) Then
                TextBox1.Text = e.Data.GetData(DataFormats.Text)
            End If

            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim MyFiles() As String


                ' Assign the files to an array.
                MyFiles = e.Data.GetData(DataFormats.FileDrop)
                ' Loop through the array and add the files to the list.

                LoadFile(MyFiles(0))

            End If


        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub LoadFile(Optional ByVal filename As String = "")
        Try
            Dim processfile As String

            If filename = "" Then
                Dim result As DialogResult
                result = OpenFileDialog1.ShowDialog()
                If result = DialogResult.OK Then
                    filename = OpenFileDialog1.FileName
                    OpenFileDialog1.Dispose()
                End If
            End If

            If Not filename = "" Then
                Dim textreader As IO.StreamReader = New IO.StreamReader(filename, True)
                TextBox1.Text = textreader.ReadToEnd()
                Label2.Text = "File successfully loaded"
                textreader.Close()
            End If

        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub
    Private Sub savefile()
        Try
            Dim processfile As String
            Dim filename As String
            filename = ""
            Dim result As DialogResult
            result = SaveFileDialog1.ShowDialog()
            If result = DialogResult.OK Then
                filename = SaveFileDialog1.FileName
                SaveFileDialog1.Dispose()
            End If

            If Not filename = "" Then
                Dim textreader As IO.StreamWriter = New IO.StreamWriter(filename, False)
                textreader.Write(TextBox1.Text)
                Message_Handler("File successfully saved")
                Label2.Text = "File successfully saved"
                textreader.Close()
            End If
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub


    Private Sub Form_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress, TextBox1.KeyPress, Panel1.KeyPress
        Try
            If e.KeyChar.IsControl(e.KeyChar) And e.KeyChar.GetHashCode = 65537 Then
                TextBox1.SelectAll()
                e.Handled = True
            End If

        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub Main_Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            TextBox1.Focus()
            TextBox1.Select()
            Me.Text = "Span Tag Remover (" & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ")"
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Try
            LoadFile()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        Try
            savefile()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem8.Click
        Try
            Dim helpscreen As Help_Form = New Help_Form
            helpscreen.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub
End Class
