'Autono-Me
'© 2026 Copyright Oranyx Labs/Elliot Monteverde All Rights Reserved
'GNU General Public License v3.0

Imports System.Net.Http
Imports System.Text.Json
Public Class SettingsF
    Private Sub SettingsF_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadBootFiles()
    End Sub
    Private checkboxEventsEnabled As Boolean = True
    Private Async Sub LoadBootFiles()
        Try
            Dim osPath = IO.Path.Combine(Application.StartupPath, "OS")
            If IO.File.Exists(osPath) Then
                OSRichTextBox.Text = IO.File.ReadAllText(osPath)
            End If
            Dim bootPath = IO.Path.Combine(Application.StartupPath, "BOOT")
            If IO.File.Exists(bootPath) Then
                BootRichTextBox.Text = IO.File.ReadAllText(bootPath)
            End If
            If String.IsNullOrWhiteSpace(My.Settings.CName) OrElse
           String.IsNullOrWhiteSpace(My.Settings.HAddress) OrElse
           String.IsNullOrWhiteSpace(My.Settings.MName) Then
                checkboxEventsEnabled = False
                OllamaCheckBox.Checked = False
                OpenAICheckBox.Checked = False
                AnthropicCheckBox.Checked = False
                GoogleCheckBox.Checked = False
                checkboxEventsEnabled = True
                HAddressTextBox.Clear()
                MNameListBox.Items.Clear()
                Return
            End If
            checkboxEventsEnabled = False
            Select Case My.Settings.CName
                Case "Ollama"
                    OllamaCheckBox.Checked = True
                Case "OpenAI"
                    OpenAICheckBox.Checked = True
                Case "Anthropic"
                    AnthropicCheckBox.Checked = True
                Case "Google"
                    GoogleCheckBox.Checked = True
            End Select
            HAddressTextBox.Text = My.Settings.HAddress
            checkboxEventsEnabled = True
            If OllamaCheckBox.Checked Then
                Await LoadOllamaModels(My.Settings.HAddress)
            ElseIf OpenAICheckBox.Checked Then
                Await LoadOpenAIModels(My.Settings.HAddress)
            ElseIf AnthropicCheckBox.Checked Then
                Await LoadAnthropicModels(My.Settings.HAddress)
            ElseIf GoogleCheckBox.Checked Then
                Await LoadGoogleModels(My.Settings.HAddress)
            End If
            If MNameListBox.Items.Contains(My.Settings.MName) Then
                MNameListBox.SelectedItem = My.Settings.MName
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading: " & ex.Message)
        End Try
    End Sub
    Private Sub ProviderCheckBoxChanged(sender As Object, e As EventArgs) _
    Handles OllamaCheckBox.CheckedChanged,
            OpenAICheckBox.CheckedChanged,
            AnthropicCheckBox.CheckedChanged,
            GoogleCheckBox.CheckedChanged
        If Not checkboxEventsEnabled Then Return
        Dim box = CType(sender, CheckBox)
        If box.Checked Then
            OllamaCheckBox.Checked = (box Is OllamaCheckBox)
            OpenAICheckBox.Checked = (box Is OpenAICheckBox)
            AnthropicCheckBox.Checked = (box Is AnthropicCheckBox)
            GoogleCheckBox.Checked = (box Is GoogleCheckBox)
            HAddressTextBox.Clear()
            MNameListBox.Items.Clear()
        End If
    End Sub
    Private Async Function LoadOllamaModels(serverUrl As String) As Task
        Try
            MNameListBox.Items.Clear()
            Dim apiUrl As String = BuildOllamaUrl(serverUrl)
            Using client As New HttpClient()
                Dim response = Await client.GetStringAsync(apiUrl)
                Dim json = JsonDocument.Parse(response)
                Dim models = json.RootElement.GetProperty("models")
                For Each model In models.EnumerateArray()
                    Dim name = model.GetProperty("name").GetString()
                    MNameListBox.Items.Add(name)
                Next
            End Using
        Catch ex As Exception
            MessageBox.Show("Ollama error: " & ex.Message)
        End Try
    End Function
    Private Async Function LoadOpenAIModels(apiKey As String) As Task
        Try
            MNameListBox.Items.Clear()
            Using client As New HttpClient()
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                Dim response = Await client.GetStringAsync("https://api.openai.com/v1/models")
                Dim json = JsonDocument.Parse(response)
                Dim models = json.RootElement.GetProperty("data")
                For Each model In models.EnumerateArray()
                    Dim name = model.GetProperty("id").GetString()
                    MNameListBox.Items.Add(name)
                Next
            End Using
        Catch ex As Exception
            MessageBox.Show("OpenAI error: " & ex.Message)
        End Try
    End Function
    Private Async Function LoadAnthropicModels(apiKey As String) As Task
        Try
            MNameListBox.Items.Clear()
            Using client As New HttpClient()
                client.DefaultRequestHeaders.Add("x-api-key", apiKey)
                client.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01")
                Dim response = Await client.GetStringAsync("https://api.anthropic.com/v1/models")
                Dim json = JsonDocument.Parse(response)
                Dim models = json.RootElement.GetProperty("data")
                For Each model In models.EnumerateArray()
                    Dim name = model.GetProperty("id").GetString()
                    MNameListBox.Items.Add(name)
                Next
            End Using
        Catch ex As Exception
            MessageBox.Show("Anthropic error: " & ex.Message)
        End Try
    End Function
    Private Async Function LoadGoogleModels(apiKey As String) As Task
        Try
            MNameListBox.Items.Clear()
            Using client As New HttpClient()
                Dim url = "https://generativelanguage.googleapis.com/v1/models?key=" & apiKey
                Dim response = Await client.GetStringAsync(url)
                Dim json = JsonDocument.Parse(response)
                Dim models = json.RootElement.GetProperty("models")
                For Each model In models.EnumerateArray()
                    Dim name = model.GetProperty("name").GetString()
                    MNameListBox.Items.Add(name)
                Next
            End Using
        Catch ex As Exception
            MessageBox.Show("Google Gemini error: " & ex.Message)
        End Try
    End Function
    Private Async Sub HAddressTextBox_Leave(sender As Object, e As EventArgs) Handles HAddressTextBox.Leave
        If String.IsNullOrWhiteSpace(HAddressTextBox.Text) Then Return
        If OllamaCheckBox.Checked Then
            Await LoadOllamaModels(HAddressTextBox.Text)
        ElseIf OpenAICheckBox.Checked Then
            Await LoadOpenAIModels(HAddressTextBox.Text)
        ElseIf AnthropicCheckBox.Checked Then
            Await LoadAnthropicModels(HAddressTextBox.Text)
        ElseIf GoogleCheckBox.Checked Then
            Await LoadGoogleModels(HAddressTextBox.Text)
        End If
    End Sub
    Private Sub ServerSButton_Click(sender As Object, e As EventArgs) Handles ServerSButton.Click
        Try
            If String.IsNullOrWhiteSpace(HAddressTextBox.Text) Then
                MessageBox.Show("Please enter a server address or API key.")
                Return
            End If
            If MNameListBox.SelectedItem Is Nothing Then
                MessageBox.Show("Please select a model from the list.")
                Return
            End If
            My.Settings.HAddress = HAddressTextBox.Text
            My.Settings.MName = MNameListBox.SelectedItem.ToString()
            If OllamaCheckBox.Checked Then
                My.Settings.CName = "Ollama"
            ElseIf OpenAICheckBox.Checked Then
                My.Settings.CName = "OpenAI"
            ElseIf AnthropicCheckBox.Checked Then
                My.Settings.CName = "Anthropic"
            ElseIf GoogleCheckBox.Checked Then
                My.Settings.CName = "Google"
            End If
            My.Settings.Save()
            MessageBox.Show("Server configuration saved successfully.")
        Catch ex As Exception
            MessageBox.Show("Error saving server configuration: " & ex.Message)
        End Try
    End Sub
    Private Sub SaveOSButton_Click(sender As Object, e As EventArgs) Handles SaveOSButton.Click
        Try
            Dim osPath = IO.Path.Combine(Application.StartupPath, "OS")
            IO.File.WriteAllText(osPath, OSRichTextBox.Text)
            MessageBox.Show("OS saved successfully.")
        Catch ex As Exception
            MessageBox.Show("Error saving OS: " & ex.Message)
        End Try
    End Sub
    Private Sub SaveBootButton_Click(sender As Object, e As EventArgs) Handles SaveBootButton.Click
        Try
            Dim bootPath = IO.Path.Combine(Application.StartupPath, "BOOT")
            IO.File.WriteAllText(bootPath, BootRichTextBox.Text)
            MessageBox.Show("BOOT saved successfully.")
        Catch ex As Exception
            MessageBox.Show("Error saving BOOT: " & ex.Message)
        End Try
    End Sub
    Private Sub ClearOSButton_Click(sender As Object, e As EventArgs) Handles ClearOSButton.Click
        OSRichTextBox.Clear()
    End Sub
    Private Sub ClearBootButton_Click(sender As Object, e As EventArgs) Handles ClearBootButton.Click
        BootRichTextBox.Clear()
    End Sub
    Private Sub ServerCButton_Click(sender As Object, e As EventArgs) Handles ServerCButton.Click
        HAddressTextBox.Clear()
        MNameListBox.Items.Clear()
    End Sub
    Private Sub LoadOSButton_Click(sender As Object, e As EventArgs) Handles LoadOSButton.Click
        Try
            Using ofd As New OpenFileDialog
                ofd.Title = "Select OS File"
                ofd.Filter = "All Files (*.*)|*.*"
                ofd.Multiselect = False
                If ofd.ShowDialog = DialogResult.OK Then
                    Dim selectedFile = ofd.FileName
                    Dim bytes = IO.File.ReadAllBytes(selectedFile)
                    If IsBinary(bytes) Then
                        MessageBox.Show("The selected OS file is not valid.")
                        Return
                    End If
                    Dim newContent = IO.File.ReadAllText(selectedFile)
                    Dim osPath = IO.Path.Combine(Application.StartupPath, "OS")
                    IO.File.WriteAllText(osPath, newContent)
                    OSRichTextBox.Text = newContent
                    MessageBox.Show("OS file loaded successfully.")
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("Error loading OS file: " & ex.Message)
        End Try
    End Sub
    Private Function IsBinary(bytes As Byte()) As Boolean
        For Each b In bytes
            If b = 0 Then Return True
        Next
        Try
            Dim text = System.Text.Encoding.UTF8.GetString(bytes)
            If text.Contains("") Then Return True
        Catch
            Return True
        End Try
        Return False
    End Function
    Private Function BuildOllamaUrl(input As String) As String
        Dim raw As String = input.Trim()
        If Not raw.StartsWith("http://") AndAlso Not raw.StartsWith("https://") Then
            raw = "http://" & raw
        End If
        Dim uri As New Uri(raw)
        Dim port As Integer = If(uri.IsDefaultPort, 11434, uri.Port)
        Return $"{uri.Scheme}://{uri.Host}:{port}/api/tags"
    End Function
End Class

'Autono-Me
'© 2026 Copyright Oranyx Labs/Elliot Monteverde All Rights Reserved
'GNU General Public License v3.0