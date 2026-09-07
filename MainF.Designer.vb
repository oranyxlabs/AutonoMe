<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainF
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainF))
        ResponseRichTextBox = New TextBox()
        SendButton = New Button()
        ShellTimer = New Timer(components)
        RLTabControl = New TabControl()
        ThoughtsTabPage = New TabPage()
        QueryTextBox = New TextBox()
        ClearButton = New Button()
        KillButton = New Button()
        SettingsButton = New Button()
        QueryTabControl = New TabControl()
        CommandsTabPage = New TabPage()
        LogoPictureBox = New PictureBox()
        Label1 = New Label()
        Label2 = New Label()
        RecordButton = New Button()
        TabControl1 = New TabControl()
        TabPage1 = New TabPage()
        SceneListBox = New ListBox()
        AddButton = New Button()
        RemoveButton = New Button()
        RLTabControl.SuspendLayout()
        ThoughtsTabPage.SuspendLayout()
        QueryTabControl.SuspendLayout()
        CommandsTabPage.SuspendLayout()
        CType(LogoPictureBox, ComponentModel.ISupportInitialize).BeginInit()
        TabControl1.SuspendLayout()
        TabPage1.SuspendLayout()
        SuspendLayout()
        ' 
        ' ResponseRichTextBox
        ' 
        ResponseRichTextBox.BackColor = Color.Black
        ResponseRichTextBox.BorderStyle = BorderStyle.None
        ResponseRichTextBox.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        ResponseRichTextBox.ForeColor = Color.Gray
        ResponseRichTextBox.Location = New Point(6, 6)
        ResponseRichTextBox.Multiline = True
        ResponseRichTextBox.Name = "ResponseRichTextBox"
        ResponseRichTextBox.ReadOnly = True
        ResponseRichTextBox.Size = New Size(730, 325)
        ResponseRichTextBox.TabIndex = 0
        ' 
        ' SendButton
        ' 
        SendButton.BackColor = Color.Transparent
        SendButton.FlatAppearance.BorderColor = Color.FromArgb(CByte(64), CByte(64), CByte(64))
        SendButton.FlatAppearance.BorderSize = 0
        SendButton.FlatAppearance.CheckedBackColor = Color.Transparent
        SendButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        SendButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        SendButton.FlatStyle = FlatStyle.Flat
        SendButton.Font = New Font("Webdings", 12F)
        SendButton.ForeColor = Color.Lime
        SendButton.Location = New Point(675, 515)
        SendButton.Name = "SendButton"
        SendButton.Size = New Size(25, 25)
        SendButton.TabIndex = 2
        SendButton.Text = "4"
        SendButton.UseVisualStyleBackColor = False
        ' 
        ' ShellTimer
        ' 
        ShellTimer.Enabled = True
        ' 
        ' RLTabControl
        ' 
        RLTabControl.Controls.Add(ThoughtsTabPage)
        RLTabControl.Location = New Point(12, 12)
        RLTabControl.Name = "RLTabControl"
        RLTabControl.SelectedIndex = 0
        RLTabControl.Size = New Size(750, 365)
        RLTabControl.TabIndex = 3
        ' 
        ' ThoughtsTabPage
        ' 
        ThoughtsTabPage.BackColor = Color.Black
        ThoughtsTabPage.Controls.Add(ResponseRichTextBox)
        ThoughtsTabPage.ForeColor = Color.Gray
        ThoughtsTabPage.Location = New Point(4, 24)
        ThoughtsTabPage.Name = "ThoughtsTabPage"
        ThoughtsTabPage.Padding = New Padding(3)
        ThoughtsTabPage.Size = New Size(742, 337)
        ThoughtsTabPage.TabIndex = 0
        ThoughtsTabPage.Text = "Thoughts"
        ' 
        ' QueryTextBox
        ' 
        QueryTextBox.BackColor = Color.Black
        QueryTextBox.BorderStyle = BorderStyle.None
        QueryTextBox.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        QueryTextBox.ForeColor = Color.Gray
        QueryTextBox.Location = New Point(6, 6)
        QueryTextBox.Multiline = True
        QueryTextBox.Name = "QueryTextBox"
        QueryTextBox.Size = New Size(730, 80)
        QueryTextBox.TabIndex = 0
        ' 
        ' ClearButton
        ' 
        ClearButton.BackColor = Color.Transparent
        ClearButton.FlatAppearance.BorderColor = Color.FromArgb(CByte(64), CByte(64), CByte(64))
        ClearButton.FlatAppearance.BorderSize = 0
        ClearButton.FlatAppearance.CheckedBackColor = Color.Transparent
        ClearButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        ClearButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        ClearButton.FlatStyle = FlatStyle.Flat
        ClearButton.Font = New Font("Webdings", 12F)
        ClearButton.ForeColor = Color.RoyalBlue
        ClearButton.Location = New Point(737, 515)
        ClearButton.Name = "ClearButton"
        ClearButton.Size = New Size(25, 25)
        ClearButton.TabIndex = 5
        ClearButton.Text = "r"
        ClearButton.UseVisualStyleBackColor = False
        ' 
        ' KillButton
        ' 
        KillButton.BackColor = Color.Transparent
        KillButton.FlatAppearance.BorderColor = Color.FromArgb(CByte(64), CByte(64), CByte(64))
        KillButton.FlatAppearance.BorderSize = 0
        KillButton.FlatAppearance.CheckedBackColor = Color.Transparent
        KillButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        KillButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        KillButton.FlatStyle = FlatStyle.Flat
        KillButton.Font = New Font("Webdings", 12F)
        KillButton.ForeColor = Color.Red
        KillButton.Location = New Point(706, 515)
        KillButton.Name = "KillButton"
        KillButton.Size = New Size(25, 25)
        KillButton.TabIndex = 4
        KillButton.Text = "<"
        KillButton.UseVisualStyleBackColor = False
        ' 
        ' SettingsButton
        ' 
        SettingsButton.BackColor = Color.Transparent
        SettingsButton.FlatAppearance.BorderColor = Color.FromArgb(CByte(64), CByte(64), CByte(64))
        SettingsButton.FlatAppearance.BorderSize = 0
        SettingsButton.FlatAppearance.CheckedBackColor = Color.Transparent
        SettingsButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        SettingsButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        SettingsButton.FlatStyle = FlatStyle.Flat
        SettingsButton.Font = New Font("Wingdings 2", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(2))
        SettingsButton.ForeColor = Color.RoyalBlue
        SettingsButton.Location = New Point(12, 515)
        SettingsButton.Name = "SettingsButton"
        SettingsButton.Size = New Size(25, 25)
        SettingsButton.TabIndex = 0
        SettingsButton.Text = "ß"
        SettingsButton.UseVisualStyleBackColor = False
        ' 
        ' QueryTabControl
        ' 
        QueryTabControl.Controls.Add(CommandsTabPage)
        QueryTabControl.Location = New Point(12, 389)
        QueryTabControl.Name = "QueryTabControl"
        QueryTabControl.SelectedIndex = 0
        QueryTabControl.Size = New Size(750, 120)
        QueryTabControl.TabIndex = 1
        ' 
        ' CommandsTabPage
        ' 
        CommandsTabPage.BackColor = Color.Black
        CommandsTabPage.Controls.Add(QueryTextBox)
        CommandsTabPage.ForeColor = Color.Gray
        CommandsTabPage.Location = New Point(4, 24)
        CommandsTabPage.Name = "CommandsTabPage"
        CommandsTabPage.Padding = New Padding(3)
        CommandsTabPage.Size = New Size(742, 92)
        CommandsTabPage.TabIndex = 0
        CommandsTabPage.Text = "Commands"
        ' 
        ' LogoPictureBox
        ' 
        LogoPictureBox.Image = CType(resources.GetObject("LogoPictureBox.Image"), Image)
        LogoPictureBox.Location = New Point(814, 546)
        LogoPictureBox.Name = "LogoPictureBox"
        LogoPictureBox.Size = New Size(166, 94)
        LogoPictureBox.SizeMode = PictureBoxSizeMode.Zoom
        LogoPictureBox.TabIndex = 14
        LogoPictureBox.TabStop = False
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.ForeColor = Color.RoyalBlue
        Label1.Location = New Point(12, 610)
        Label1.Name = "Label1"
        Label1.Size = New Size(369, 30)
        Label1.TabIndex = 11
        Label1.Text = "© Copyright 2026 Oranyx Labs/Elliot Monteverde All Rights Reserved" & vbCrLf & "'GNU General Public License v3.0"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.ForeColor = Color.RoyalBlue
        Label2.Location = New Point(302, 520)
        Label2.Name = "Label2"
        Label2.Size = New Size(175, 15)
        Label2.TabIndex = 10
        Label2.Text = "Autono-Me can make mistakes."
        ' 
        ' RecordButton
        ' 
        RecordButton.BackColor = Color.Transparent
        RecordButton.FlatAppearance.BorderColor = Color.FromArgb(CByte(64), CByte(64), CByte(64))
        RecordButton.FlatAppearance.BorderSize = 0
        RecordButton.FlatAppearance.CheckedBackColor = Color.Transparent
        RecordButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        RecordButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        RecordButton.FlatStyle = FlatStyle.Flat
        RecordButton.Font = New Font("Webdings", 12F)
        RecordButton.ForeColor = Color.Lime
        RecordButton.Location = New Point(881, 515)
        RecordButton.Name = "RecordButton"
        RecordButton.Size = New Size(25, 25)
        RecordButton.TabIndex = 6
        RecordButton.Text = "="
        RecordButton.UseVisualStyleBackColor = False
        ' 
        ' TabControl1
        ' 
        TabControl1.Controls.Add(TabPage1)
        TabControl1.Location = New Point(768, 12)
        TabControl1.Name = "TabControl1"
        TabControl1.SelectedIndex = 0
        TabControl1.Size = New Size(200, 497)
        TabControl1.TabIndex = 7
        ' 
        ' TabPage1
        ' 
        TabPage1.BackColor = Color.Black
        TabPage1.Controls.Add(SceneListBox)
        TabPage1.ForeColor = Color.Gray
        TabPage1.Location = New Point(4, 24)
        TabPage1.Name = "TabPage1"
        TabPage1.Padding = New Padding(3)
        TabPage1.Size = New Size(192, 469)
        TabPage1.TabIndex = 0
        TabPage1.Text = "Scenes"
        ' 
        ' SceneListBox
        ' 
        SceneListBox.BackColor = Color.Black
        SceneListBox.BorderStyle = BorderStyle.None
        SceneListBox.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        SceneListBox.ForeColor = Color.Gray
        SceneListBox.FormattingEnabled = True
        SceneListBox.ItemHeight = 15
        SceneListBox.Location = New Point(6, 6)
        SceneListBox.Name = "SceneListBox"
        SceneListBox.Size = New Size(180, 450)
        SceneListBox.TabIndex = 0
        ' 
        ' AddButton
        ' 
        AddButton.BackColor = Color.Transparent
        AddButton.FlatAppearance.BorderColor = Color.FromArgb(CByte(64), CByte(64), CByte(64))
        AddButton.FlatAppearance.BorderSize = 0
        AddButton.FlatAppearance.CheckedBackColor = Color.Transparent
        AddButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        AddButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        AddButton.FlatStyle = FlatStyle.Flat
        AddButton.Font = New Font("Wingdings 2", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(2))
        AddButton.ForeColor = Color.RoyalBlue
        AddButton.Location = New Point(912, 517)
        AddButton.Name = "AddButton"
        AddButton.Size = New Size(25, 25)
        AddButton.TabIndex = 8
        AddButton.Text = "Ì"
        AddButton.UseVisualStyleBackColor = False
        ' 
        ' RemoveButton
        ' 
        RemoveButton.BackColor = Color.Transparent
        RemoveButton.FlatAppearance.BorderColor = Color.FromArgb(CByte(64), CByte(64), CByte(64))
        RemoveButton.FlatAppearance.BorderSize = 0
        RemoveButton.FlatAppearance.CheckedBackColor = Color.Transparent
        RemoveButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        RemoveButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        RemoveButton.FlatStyle = FlatStyle.Flat
        RemoveButton.Font = New Font("Webdings", 12F)
        RemoveButton.ForeColor = Color.RoyalBlue
        RemoveButton.Location = New Point(943, 515)
        RemoveButton.Name = "RemoveButton"
        RemoveButton.Size = New Size(25, 25)
        RemoveButton.TabIndex = 9
        RemoveButton.Text = "r"
        RemoveButton.UseVisualStyleBackColor = False
        ' 
        ' MainF
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.FromArgb(CByte(64), CByte(64), CByte(64))
        ClientSize = New Size(979, 656)
        Controls.Add(RemoveButton)
        Controls.Add(AddButton)
        Controls.Add(TabControl1)
        Controls.Add(RecordButton)
        Controls.Add(Label2)
        Controls.Add(SettingsButton)
        Controls.Add(ClearButton)
        Controls.Add(SendButton)
        Controls.Add(Label1)
        Controls.Add(KillButton)
        Controls.Add(LogoPictureBox)
        Controls.Add(QueryTabControl)
        Controls.Add(RLTabControl)
        FormBorderStyle = FormBorderStyle.FixedSingle
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        MaximizeBox = False
        Name = "MainF"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Autono-Me"
        RLTabControl.ResumeLayout(False)
        ThoughtsTabPage.ResumeLayout(False)
        ThoughtsTabPage.PerformLayout()
        QueryTabControl.ResumeLayout(False)
        CommandsTabPage.ResumeLayout(False)
        CommandsTabPage.PerformLayout()
        CType(LogoPictureBox, ComponentModel.ISupportInitialize).EndInit()
        TabControl1.ResumeLayout(False)
        TabPage1.ResumeLayout(False)
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents ResponseRichTextBox As TextBox
    Friend WithEvents SendButton As Button
    Friend WithEvents ShellTimer As Timer
    Friend WithEvents RLTabControl As TabControl
    Friend WithEvents ThoughtsTabPage As TabPage
    Friend WithEvents ClearButton As Button
    Friend WithEvents KillButton As Button
    Friend WithEvents QueryTextBox As TextBox
    Friend WithEvents SettingsButton As Button
    Friend WithEvents QueryTabControl As TabControl
    Friend WithEvents CommandsTabPage As TabPage
    Friend WithEvents LogoPictureBox As PictureBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents RecordButton As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents SceneListBox As ListBox
    Friend WithEvents AddButton As Button
    Friend WithEvents RemoveButton As Button

End Class
