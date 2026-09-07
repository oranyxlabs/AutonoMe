'Autono-Me
'© 2026 Copyright Oranyx Labs/Elliot Monteverde All Rights Reserved
'GNU General Public License v3.0

Imports System.Collections.Concurrent
Imports System.IO
Imports System.Net.Http
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Automation
Imports AutonoMe.MacroRecorder
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports Microsoft.Playwright
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Tesseract
Imports UglyToad.PdfPig
Imports UglyToad.PdfPig.DocumentLayoutAnalysis
Public Class MainF
    Private Const MaxRetries As Integer = 3
    Private Const MaxCorrectionAttempts As Integer = 3
    Private Const MaxConversationTurns As Integer = 20
    Private Const MaxDocumentChars As Integer = 25000
    Private Const MaxInputLength As Integer = 25000
    Private Const CommandTimeoutMs As Integer = 120000
    Private Shared ReadOnly httpClient As New HttpClient() With {
        .Timeout = TimeSpan.FromSeconds(300)
    }
    Private providerName As String = My.Settings.CName
    Private hostAddress As String = My.Settings.HAddress
    Private modelName As String = My.Settings.MName
    Private fullText As String = ""
    Private currentIndex As Integer = 0
    Private previousCommandPS As String = ""
    Private commandExecutedPS As Boolean = False
    Private gptResponse As String = ""
    Private isRunning As Integer = 0
    Private ReadOnly executionSemaphore As New SemaphoreSlim(1, 1)
    Private pendingDocContent As String = ""
    Private lastExtractedWebText As String = ""
    Private pendingTargetApp As String = ""
    Private pendingTaskQueue As New ConcurrentQueue(Of String)()
    Private pendingDocSavePath As String = ""
    Private allInstalledApps As New List(Of AppEntry)
    Private lastLogSize As Long = 0
    Private conversationHistory As New ConcurrentQueue(Of ConversationEntry)
    Private WithEvents NotifyIcon1 As New NotifyIcon()
    Private pendingNotification As String = ""
    Private selectedFiles As New List(Of String)
    Private selectedFolder As String = ""
    Private mode As String = "single"
    Public userFolders As New Dictionary(Of String, String)
    Private correctionAttempts As Integer = 0
    Private isMacroRecording As Boolean = False
    Private macroSavePath As String = ""
    Private WithEvents RenameTextBox As New TextBox()
    Private renameIndex As Integer = -1
    Private oldFileName As String = ""
    Private executionCts As CancellationTokenSource
    Private activeBrowser As IBrowser = Nothing
    Private activePlaywright As IPlaywright = Nothing
    Private activePage As IPage = Nothing
    Private activeContext As Microsoft.Playwright.IBrowserContext = Nothing
    Private lastUserInput As String = ""
    Private wantsChart As Boolean = False
    Private wantsFormulas As Boolean = False
    Private wantsConditionalFormat As Boolean = False
    Private wantsModifyExisting As Boolean = False
    Private wantsWordToC As Boolean = False
    Private wantsWordTable As Boolean = False
    Private wantsModifyExistingWord As Boolean = False
    Private wantsWordPdf As Boolean = False
    Private wantsPptNotes As Boolean = False
    Private wantsPptChart As Boolean = False
    Private wantsPptTable As Boolean = False
    Private wantsPptTransitions As Boolean = False
    Private wantsPptPdf As Boolean = False
    Private wantsPptTheme As Boolean = False
    Private wantsPptImages As Boolean = False
    Private wantsModifyExistingPpt As Boolean = False
    Private wantsOutlookSend As Boolean = False
    Private wantsOutlookReply As Boolean = False
    Private wantsOutlookForward As Boolean = False
    Private wantsOutlookRead As Boolean = False
    Private wantsOutlookCalendar As Boolean = False
    Private wantsOutlookTask As Boolean = False
    Private wantsOutlookContact As Boolean = False
    Private wantsOutlookExport As Boolean = False
    Private wantsOneNoteFormat As Boolean = False
    Private wantsOneNoteChecklist As Boolean = False
    Private wantsOneNoteTable As Boolean = False
    Private wantsOneNoteAppend As Boolean = False
    Private wantsOneNoteNewSection As Boolean = False
    Private wantsOneNoteSearch As Boolean = False
    Private wantsOneNoteExport As Boolean = False
    Private wantsAccessCreate As Boolean = False
    Private wantsAccessQuery As Boolean = False
    Private wantsAccessImport As Boolean = False
    Private wantsAccessExport As Boolean = False
    Private wantsAccessReport As Boolean = False
    Private wantsAccessModify As Boolean = False
    Private wantsLocalUserCreate As Boolean = False
    Private wantsLocalUserModify As Boolean = False
    Private wantsLocalUserDelete As Boolean = False
    Private wantsLocalUserList As Boolean = False
    Private wantsLocalGroupManage As Boolean = False
    Private wantsADUserCreate As Boolean = False
    Private wantsADUserModify As Boolean = False
    Private wantsADUserDelete As Boolean = False
    Private wantsADUserSearch As Boolean = False
    Private wantsADGroupManage As Boolean = False
    Private wantsADComputerManage As Boolean = False
    Private wantsADOUManage As Boolean = False
    Private wantsADPasswordReset As Boolean = False
    Private wantsADBulkOperation As Boolean = False
    Private Shared ReadOnly DangerousPatterns As String() = {
    "Invoke-Expression", "iex\s", "iex\(", "Invoke-WebRequest.*\|.*iex",
    "DownloadString", "DownloadFile", "Net\.WebClient",
    "Start-Process.*cmd", "Start-Process.*powershell",
    "Remove-Item.*-Recurse.*C:\\", "Remove-Item.*-Recurse.*\$env:",
    "del /s /q C:\\", "rmdir /s /q",
    "Format-Volume", "Clear-Disk", "Initialize-Disk",
    "net user.*\/add.*\/domain", "net user administrator.*\/active", "netsh",
    "reg add", "reg delete", "Remove-ItemProperty.*HKLM",
    "Set-ExecutionPolicy", "Disable-WindowsOptionalFeature",
    "schtasks /create", "sc create", "sc config",
    "bcdedit", "bcdboot", "bootrec",
    "cipher /w", "diskpart",
    "Stop-Service", "Stop-Computer", "Restart-Computer",
    "\$env:windir", "\$env:systemroot", "\$env:systemdrive\\windows",
    "\\windows\\system32", "\\windows\\syswow64",
    "New-Service", "Set-Service",
    "Add-MpPreference.*-ExclusionPath",
    "Set-MpPreference.*-DisableRealtimeMonitoring",
    "ConvertTo-SecureString",
    "Get-Credential", "Export-PfxCertificate",
    "Invoke-Mimikatz", "mimikatz",
    "certutil.*-decode", "certutil.*-urlcache"
}
    Private Async Sub MainF_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Me.KeyPreview = True
            If My.Settings.HAddress = String.Empty OrElse My.Settings.MName = String.Empty Then
                MessageBox.Show("Server configuration not found.")
            Else
                hostAddress = My.Settings.HAddress
                modelName = My.Settings.MName
                providerName = My.Settings.CName
            End If
            NotifyIcon1.Icon = SystemIcons.Information
            NotifyIcon1.Text = "Autono-Me"
            NotifyIcon1.Visible = True
            Await CheckPlaywrightBrowsersAsync()
            Dim registryApps = GetInstalledApplications()
            Dim storeApps = GetStoreApps()
            allInstalledApps = registryApps.Concat(storeApps).Distinct().ToList()
            ResponseRichTextBox.ReadOnly = True
            ResponseRichTextBox.ScrollBars = ScrollBars.Vertical
            ResponseRichTextBox.Multiline = True
            QueryTextBox.ScrollBars = ScrollBars.Vertical
            QueryTextBox.Multiline = True
            Dim userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            userFolders("UserProfile") = userProfile
            userFolders("Desktop") = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            userFolders("Documents") = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            userFolders("Downloads") = Path.Combine(userProfile, "Downloads")
            userFolders("Pictures") = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
            userFolders("Music") = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic)
            userFolders("Videos") = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos)
            userFolders("AppDataRoaming") = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            userFolders("AppDataLocal") = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            LoadExistingMacros()
            RenameTextBox.Visible = False
            SceneListBox.Parent.Controls.Add(RenameTextBox)
            QueryTextBox.Focus()
        Catch ex As Exception
            MessageBox.Show("Startup error: " & ex.Message, "Autono-Me", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Async Function CheckPlaywrightBrowsersAsync() As Task
        Try
            Dim browsersPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ms-playwright")
            Dim chromiumInstalled = Directory.Exists(browsersPath) AndAlso
                                    Directory.GetDirectories(browsersPath, "chromium-*").Length > 0
            If Not chromiumInstalled Then
                Dim result = MessageBox.Show(
                    "Web automation requires a one-time browser setup. Install now?" & vbCrLf & vbCrLf &
                    "This will download ~170MB and only needs to be done once.",
                    "Setup Required", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.Yes Then
                    Me.Text = "Installing browsers, please wait..."
                    Me.Enabled = False
                    Dim ok = Await PlaywrightInstallHelper.InstallBrowsersAsync()
                    Me.Text = "Autono-Me"
                    Me.Enabled = True
                    If ok Then
                        MessageBox.Show("Browser installation complete!", "Autono-Me",
                                       MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("Browser installation failed.", "Autono-Me",
                                       MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                End If
            End If
        Catch ex As Exception
            Debug.WriteLine("CheckPlaywrightBrowsers Error: " & ex.Message)
            LogToFile("BROWSER_CHECK_ERROR", ex.Message)
        End Try
    End Function
    Private Sub LogToFile(category As String, message As String)
        Try
            Dim logPath = Path.Combine(Application.StartupPath, "LOG")
            Dim entry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{category}] {message}{Environment.NewLine}"
            File.AppendAllText(logPath, entry)
        Catch
        End Try
    End Sub
    Private Sub AppendResponse(text As String)
        If ResponseRichTextBox.InvokeRequired Then
            ResponseRichTextBox.BeginInvoke(Sub()
                                                ResponseRichTextBox.AppendText(text & Environment.NewLine)
                                                ResponseRichTextBox.SelectionStart = ResponseRichTextBox.TextLength
                                                ResponseRichTextBox.ScrollToCaret()
                                            End Sub)
        Else
            ResponseRichTextBox.AppendText(text & Environment.NewLine)
            ResponseRichTextBox.SelectionStart = ResponseRichTextBox.TextLength
            ResponseRichTextBox.ScrollToCaret()
        End If
    End Sub
    Private Function IsCommandSafe(command As String) As (Safe As Boolean, Reason As String)
        If String.IsNullOrWhiteSpace(command) Then Return (False, "Empty command")
        Dim accountMgmtPatterns = {
        "Get-LocalUser", "New-LocalUser", "Set-LocalUser", "Remove-LocalUser",
        "Enable-LocalUser", "Disable-LocalUser", "Rename-LocalUser",
        "Get-LocalGroup", "New-LocalGroup", "Add-LocalGroupMember", "Remove-LocalGroupMember",
        "Get-ADUser", "New-ADUser", "Set-ADUser", "Remove-ADUser",
        "Enable-ADAccount", "Disable-ADAccount", "Unlock-ADAccount",
        "Get-ADGroup", "New-ADGroup", "Add-ADGroupMember", "Remove-ADGroupMember",
        "Get-ADComputer", "Get-ADOrganizationalUnit", "New-ADOrganizationalUnit",
        "Set-ADAccountPassword", "Search-ADAccount", "Move-ADObject",
        "Import-Module ActiveDirectory", "Get-ADDomain",
        "Get-LocalGroupMember"
    }
        Dim isAccountMgmt = accountMgmtPatterns.Any(
        Function(p) command.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
        For Each pattern In DangerousPatterns
            If Regex.IsMatch(command, pattern, RegexOptions.IgnoreCase) Then
                If isAccountMgmt Then
                    If pattern = "net user" OrElse pattern = "net localgroup" OrElse
                   pattern = "netsh" OrElse pattern = "ConvertTo-SecureString" OrElse
                   pattern = "Get-Credential" Then
                        Continue For
                    End If
                End If
                Return (False, $"Blocked dangerous pattern: {pattern}")
            End If
        Next
        Dim windowsDirPattern As String = "(?i)([a-z]:\\windows(?:\\|$))"
        If Regex.IsMatch(command, windowsDirPattern) Then
            Return (False, "Blocked: Windows system directory access")
        End If
        If command.ToLower().Contains("drivers\etc\hosts") Then
            Return (False, "Blocked: hosts file modification")
        End If
        If Regex.IsMatch(command, "-EncodedCommand\s+[A-Za-z0-9+/=]+", RegexOptions.IgnoreCase) Then
            Return (False, "Blocked: Nested encoded command execution")
        End If
        Return (True, "")
    End Function
    Private Function TranslateSceneToSteps(sceneText As String) As String
        Dim lines = sceneText.Split({Environment.NewLine, Chr(10).ToString()}, StringSplitOptions.RemoveEmptyEntries)
        Dim steps As New List(Of String)
        Dim stepNum As Integer = 1
        Dim currentApp As String = ""
        Dim typedBuffer As String = ""
        Dim searchQuery As String = ""
        Dim lastWasTaskbarSearch As Boolean = False
        For Each line In lines
            line = line.Trim()
            If line.StartsWith("===") Then Continue For
            If String.IsNullOrWhiteSpace(line) Then Continue For
            Dim appMatch = Regex.Match(line, "App:\s*'([^']+)'")
            Dim newApp = If(appMatch.Success, appMatch.Groups(1).Value, "")
            If typedBuffer.Length > 0 AndAlso (line.StartsWith("[CLICK") OrElse
            line.StartsWith("[SEARCH_LAUNCH]") OrElse
            (newApp <> currentApp AndAlso newApp <> "")) Then
                If currentApp.ToLower() = "searchhost" OrElse currentApp.ToLower() = "searchapp" OrElse
               currentApp.ToLower() = "explorer" Then
                    searchQuery = typedBuffer
                End If
                steps.Add($"STEP {stepNum}: TYPE ""{typedBuffer}"" in {currentApp}")
                stepNum += 1
                typedBuffer = ""
            End If
            If newApp <> "" Then currentApp = newApp
            If line.StartsWith("[TYPE]") Then
                Dim textMatch = Regex.Match(line, "Typed Text:\s*""([^""]*?)""")
                If textMatch.Success Then
                    typedBuffer &= textMatch.Groups(1).Value
                End If
            ElseIf line.StartsWith("[CLICK:TASKBAR_SEARCH]") Then
                steps.Add($"STEP {stepNum}: CLICK TASKBAR SEARCH BAR (click the search icon on taskbar)")
                stepNum += 1
                lastWasTaskbarSearch = True
            ElseIf line.StartsWith("[CLICK:SEARCH_RESULT_ADMIN]") Then
                Dim queryMatch = Regex.Match(line, "SearchQuery:\s*'([^']*)'")
                Dim searchQ = If(queryMatch.Success, queryMatch.Groups(1).Value, "unknown")
                steps.Add($"STEP {stepNum}: LAUNCH ""{searchQ}"" from taskbar search AS ADMINISTRATOR (Run as administrator)")
                stepNum += 1
                lastWasTaskbarSearch = False
                searchQuery = ""
            ElseIf line.StartsWith("[CLICK:SEARCH_RESULT]") Then
                Dim queryMatch = Regex.Match(line, "SearchQuery:\s*'([^']*)'")
                Dim nameMatch2 = Regex.Match(line, "Name:\s*'([^']*)'")
                Dim searchQ = If(queryMatch.Success, queryMatch.Groups(1).Value, "unknown")
                Dim resultName = If(nameMatch2.Success, nameMatch2.Groups(1).Value, "")
                steps.Add($"STEP {stepNum}: CLICK search result '{resultName}' for query ""{searchQ}"" (launch app from search)")
                stepNum += 1
                lastWasTaskbarSearch = False
                searchQuery = ""
            ElseIf line.StartsWith("[SEARCH_LAUNCH]") Then
                Dim launchMatch = Regex.Match(line, "Launched '([^']+)' from search query '([^']*)'")
                If launchMatch.Success Then
                    Dim appName = launchMatch.Groups(1).Value
                    Dim query = launchMatch.Groups(2).Value
                    steps.Add($"STEP {stepNum}: APP LAUNCHED '{appName}' from search ""{query}"" (search result was clicked but not captured)")
                    stepNum += 1
                End If
                lastWasTaskbarSearch = False
                searchQuery = ""
            ElseIf line.StartsWith("[CLICK]") Then
                Dim nameMatch = Regex.Match(line, "Name:\s*'([^']*)'")
                Dim idMatch = Regex.Match(line, "ID:\s*'([^']*)'")
                Dim windowMatch = Regex.Match(line, "Window:\s*'([^']*)'")
                Dim roleMatch = Regex.Match(line, "^\[CLICK\]\s*([^|]+)")
                Dim clickName = If(nameMatch.Success, nameMatch.Groups(1).Value.Trim(), "")
                Dim clickId = If(idMatch.Success, idMatch.Groups(1).Value.Trim(), "")
                Dim clickWindow = If(windowMatch.Success, windowMatch.Groups(1).Value.Trim(), "")
                Dim clickRole = If(roleMatch.Success, roleMatch.Groups(1).Value.Trim(), "")
                Dim previousWasSearch = lastWasTaskbarSearch OrElse
                currentApp.ToLower() = "searchhost" OrElse
                currentApp.ToLower() = "searchapp"
                Dim launchedNewApp = newApp <> "" AndAlso
                newApp.ToLower() <> "searchhost" AndAlso
                newApp.ToLower() <> "searchapp" AndAlso
                newApp.ToLower() <> "explorer"
                If previousWasSearch AndAlso launchedNewApp AndAlso Not String.IsNullOrWhiteSpace(searchQuery) Then
                    steps.Add($"STEP {stepNum}: LAUNCH ""{searchQuery}"" from taskbar search (user searched and launched app)")
                    stepNum += 1
                    searchQuery = ""
                End If
                lastWasTaskbarSearch = False
                Dim target = ""
                If Not String.IsNullOrWhiteSpace(clickName) AndAlso clickName <> "(unnamed)" Then
                    target = $"'{clickName}'"
                    If Not String.IsNullOrWhiteSpace(clickRole) Then target = $"{clickRole} '{clickName}'"
                ElseIf Not String.IsNullOrWhiteSpace(clickId) Then
                    target = $"element with ID '{clickId}'"
                Else
                    target = $"{clickRole} (unnamed element)"
                End If
                If Not String.IsNullOrWhiteSpace(clickWindow) Then
                    target &= $" in window '{clickWindow}'"
                End If
                steps.Add($"STEP {stepNum}: CLICK {target} in {currentApp}")
                stepNum += 1
            ElseIf line.StartsWith("[KEYPRESS]") Then
                Dim key = line.Replace("[KEYPRESS]", "").Trim()
                steps.Add($"STEP {stepNum}: PRESS KEY {key}")
                stepNum += 1
            ElseIf line.StartsWith("[HOTKEY]") Then
                Dim keys = line.Replace("[HOTKEY]", "").Trim()
                steps.Add($"STEP {stepNum}: HOTKEY {keys}")
                stepNum += 1
            End If
        Next
        If typedBuffer.Length > 0 Then
            steps.Add($"STEP {stepNum}: TYPE ""{typedBuffer}"" in {currentApp}")
            stepNum += 1
        End If
        Return String.Join(Environment.NewLine, steps)
    End Function
    Private Async Sub SendButton_Click(sender As Object, e As EventArgs) Handles SendButton.Click
        If Interlocked.CompareExchange(isRunning, 1, 0) <> 0 Then
            Return
        End If
        SetUIEnabled(False)
        Try
            executionCts?.Cancel()
        Catch : End Try
        executionCts = New CancellationTokenSource()
        Dim ct = executionCts.Token
        Try
            hostAddress = My.Settings.HAddress
            modelName = My.Settings.MName
            providerName = My.Settings.CName
            previousCommandPS = ""
            commandExecutedPS = False
            gptResponse = ""
            pendingNotification = ""
            correctionAttempts = 0
            Dim userInput = QueryTextBox.Text.Trim
            If String.IsNullOrWhiteSpace(userInput) Then
                MessageBox.Show("Error. Input required.")
                Return
            End If
            Const ScenePlaceholder As String = "|||SCENE_PLACEHOLDER|||"
            Dim extractedScene As String = ""
            Dim sceneExtractPattern As String = "(=== AUTONO-ME SCENE RECORDING STARTED.*?=== SCENE RECORDING STOPPED ===)"
            Dim sceneExtractMatch = Regex.Match(userInput, sceneExtractPattern, RegexOptions.Singleline)
            If sceneExtractMatch.Success Then
                extractedScene = sceneExtractMatch.Value
                userInput = userInput.Replace(extractedScene, ScenePlaceholder)
            End If
            Dim taskPattern = "(?:^|\n)\s*(\d+)\.\s+(.+?)(?=\n\s*\d+\.\s|$)"
            Dim taskMatches = Regex.Matches(userInput, taskPattern,
                            RegexOptions.Singleline Or RegexOptions.Multiline)
            If taskMatches.Count > 1 Then
                pendingTaskQueue.Clear()
                Dim allTasks As New List(Of String)
                For Each m As Match In taskMatches
                    Dim taskText = m.Groups(2).Value.Trim()
                    taskText = taskText.TrimEnd(Chr(10), Chr(13), " "c)
                    If Not String.IsNullOrWhiteSpace(taskText) Then
                        allTasks.Add(taskText)
                    End If
                Next
                If allTasks.Count < 2 Then
                    allTasks.Clear()
                    Dim lines = userInput.Split({Environment.NewLine, Chr(10).ToString()},
                                        StringSplitOptions.RemoveEmptyEntries)
                    For Each line In lines
                        Dim lineMatch = Regex.Match(line.Trim(), "^\d+\.\s+(.+)$")
                        If lineMatch.Success Then
                            Dim taskText = lineMatch.Groups(1).Value.Trim()
                            If Not String.IsNullOrWhiteSpace(taskText) Then
                                allTasks.Add(taskText)
                            End If
                        End If
                    Next
                End If
                If allTasks.Count > 1 Then
                    For i = 1 To allTasks.Count - 1
                        Dim taskText = allTasks(i)
                        If Not String.IsNullOrEmpty(extractedScene) Then
                            taskText = taskText.Replace(ScenePlaceholder, extractedScene)
                        End If
                        pendingTaskQueue.Enqueue(taskText)
                    Next
                    userInput = allTasks(0)
                    If Not String.IsNullOrEmpty(extractedScene) Then
                        userInput = userInput.Replace(ScenePlaceholder, extractedScene)
                    End If
                    ResponseRichTextBox.AppendText($"📋 {allTasks.Count} tasks detected. Running task 1 of {allTasks.Count}..." & Environment.NewLine)
                    ResponseRichTextBox.AppendText($"   Task 1: {userInput.Substring(0, Math.Min(80, userInput.Length))}..." & Environment.NewLine)
                    For i = 0 To Math.Min(pendingTaskQueue.Count - 1, 8)
                        Dim peekTasks = pendingTaskQueue.ToArray()
                        If i < peekTasks.Length Then
                            ResponseRichTextBox.AppendText($"   Task {i + 2}: {peekTasks(i).Substring(0, Math.Min(60, peekTasks(i).Length))}..." & Environment.NewLine)
                        End If
                    Next
                Else
                End If
            End If
            If Not String.IsNullOrEmpty(extractedScene) Then
                If userInput.Contains(ScenePlaceholder) Then
                    userInput = userInput.Replace(ScenePlaceholder, extractedScene)
                End If
            End If
            lastUserInput = userInput
            Dim lowerForExcel = userInput.ToLower()
            wantsChart = lowerForExcel.Contains("chart") OrElse lowerForExcel.Contains("graph") OrElse
         lowerForExcel.Contains("line chart") OrElse lowerForExcel.Contains("bar chart") OrElse
         lowerForExcel.Contains("pie chart") OrElse lowerForExcel.Contains("plot")
            wantsFormulas = lowerForExcel.Contains("formula") OrElse lowerForExcel.Contains("calculate") OrElse
            lowerForExcel.Contains("percentage change") OrElse lowerForExcel.Contains("percent change") OrElse
            lowerForExcel.Contains("average") OrElse lowerForExcel.Contains("sum") OrElse
            lowerForExcel.Contains("total") OrElse lowerForExcel.Contains("difference")
            wantsConditionalFormat = lowerForExcel.Contains("highlight") OrElse lowerForExcel.Contains("color") OrElse
                     lowerForExcel.Contains("colour") OrElse lowerForExcel.Contains("conditional") OrElse
                     lowerForExcel.Contains("heat") OrElse lowerForExcel.Contains("red") OrElse
                     lowerForExcel.Contains("green") OrElse lowerForExcel.Contains("dropped") OrElse
                     lowerForExcel.Contains("increased") OrElse lowerForExcel.Contains("flag")
            wantsModifyExisting = (lowerForExcel.Contains("open") OrElse lowerForExcel.Contains("modify") OrElse
                   lowerForExcel.Contains("edit") OrElse lowerForExcel.Contains("update") OrElse
                   lowerForExcel.Contains("add") OrElse lowerForExcel.Contains("insert")) AndAlso
                  (lowerForExcel.Contains(".xlsx") OrElse lowerForExcel.Contains("spreadsheet") OrElse
                   lowerForExcel.Contains("workbook") OrElse lowerForExcel.Contains("existing"))
            wantsWordToC = lowerForExcel.Contains("table of contents") OrElse lowerForExcel.Contains("toc") OrElse
           lowerForExcel.Contains("contents page") OrElse lowerForExcel.Contains("heading") OrElse
           lowerForExcel.Contains("chapters") OrElse lowerForExcel.Contains("sections")
            wantsWordTable = lowerForExcel.Contains("table") AndAlso
             (lowerForExcel.Contains("word") OrElse lowerForExcel.Contains("document") OrElse
              lowerForExcel.Contains("docx") OrElse lowerForExcel.Contains("report") OrElse
              lowerForExcel.Contains("write") OrElse lowerForExcel.Contains("create"))
            wantsModifyExistingWord = (lowerForExcel.Contains("open") OrElse lowerForExcel.Contains("modify") OrElse
                        lowerForExcel.Contains("edit") OrElse lowerForExcel.Contains("update") OrElse
                        lowerForExcel.Contains("add to") OrElse lowerForExcel.Contains("append") OrElse
                        lowerForExcel.Contains("insert")) AndAlso
                       (lowerForExcel.Contains(".docx") OrElse lowerForExcel.Contains(".doc") OrElse
                        (lowerForExcel.Contains("existing") AndAlso
                         (lowerForExcel.Contains("word") OrElse lowerForExcel.Contains("document"))))
            wantsWordPdf = lowerForExcel.Contains("pdf") OrElse lowerForExcel.Contains("portable") OrElse
           lowerForExcel.Contains("save as pdf") OrElse lowerForExcel.Contains("export to pdf") OrElse
           lowerForExcel.Contains("convert to pdf")
            wantsPptNotes = lowerForExcel.Contains("notes") OrElse lowerForExcel.Contains("speaker") OrElse
            lowerForExcel.Contains("presenter")
            wantsPptChart = lowerForExcel.Contains("chart") OrElse lowerForExcel.Contains("graph") OrElse
            lowerForExcel.Contains("plot")
            wantsPptTable = lowerForExcel.Contains("table") AndAlso
            (lowerForExcel.Contains("slide") OrElse lowerForExcel.Contains("powerpoint") OrElse
             lowerForExcel.Contains("ppt") OrElse lowerForExcel.Contains("presentation"))
            wantsPptTransitions = lowerForExcel.Contains("transition") OrElse lowerForExcel.Contains("animation") OrElse
                  lowerForExcel.Contains("animate") OrElse lowerForExcel.Contains("effect")
            wantsPptPdf = lowerForExcel.Contains("pdf") AndAlso
          (lowerForExcel.Contains("slide") OrElse lowerForExcel.Contains("powerpoint") OrElse
           lowerForExcel.Contains("ppt") OrElse lowerForExcel.Contains("presentation"))
            wantsPptTheme = lowerForExcel.Contains("theme") OrElse lowerForExcel.Contains("color scheme") OrElse
            lowerForExcel.Contains("colour scheme") OrElse lowerForExcel.Contains("professional") OrElse
            lowerForExcel.Contains("modern") OrElse lowerForExcel.Contains("dark") OrElse
            lowerForExcel.Contains("light") OrElse lowerForExcel.Contains("colorful")
            wantsPptImages = lowerForExcel.Contains("image") OrElse lowerForExcel.Contains("photo") OrElse
             lowerForExcel.Contains("picture") OrElse lowerForExcel.Contains("icon") OrElse
             lowerForExcel.Contains("logo")
            wantsModifyExistingPpt = (lowerForExcel.Contains("open") OrElse lowerForExcel.Contains("modify") OrElse
                       lowerForExcel.Contains("edit") OrElse lowerForExcel.Contains("update") OrElse
                       lowerForExcel.Contains("add") OrElse lowerForExcel.Contains("append") OrElse
                       lowerForExcel.Contains("insert")) AndAlso
                      (lowerForExcel.Contains(".pptx") OrElse lowerForExcel.Contains(".ppt") OrElse
                       (lowerForExcel.Contains("existing") AndAlso
                        (lowerForExcel.Contains("powerpoint") OrElse lowerForExcel.Contains("presentation") OrElse
                         lowerForExcel.Contains("ppt"))))
            wantsOutlookSend = (lowerForExcel.Contains("send") OrElse lowerForExcel.Contains("compose") OrElse
                lowerForExcel.Contains("write an email") OrElse lowerForExcel.Contains("draft")) AndAlso
               (lowerForExcel.Contains("email") OrElse lowerForExcel.Contains("mail") OrElse
                lowerForExcel.Contains("outlook"))
            wantsOutlookReply = (lowerForExcel.Contains("reply") OrElse lowerForExcel.Contains("respond")) AndAlso
                (lowerForExcel.Contains("email") OrElse lowerForExcel.Contains("mail") OrElse
                 lowerForExcel.Contains("outlook"))
            wantsOutlookForward = lowerForExcel.Contains("forward") AndAlso
                  (lowerForExcel.Contains("email") OrElse lowerForExcel.Contains("mail"))
            wantsOutlookRead = (lowerForExcel.Contains("read") OrElse lowerForExcel.Contains("check") OrElse
                lowerForExcel.Contains("search") OrElse lowerForExcel.Contains("find") OrElse
                lowerForExcel.Contains("show") OrElse lowerForExcel.Contains("list")) AndAlso
               (lowerForExcel.Contains("email") OrElse lowerForExcel.Contains("inbox") OrElse
                lowerForExcel.Contains("mail"))
            wantsOutlookCalendar = lowerForExcel.Contains("calendar") OrElse lowerForExcel.Contains("appointment") OrElse
                   lowerForExcel.Contains("meeting") OrElse lowerForExcel.Contains("schedule") OrElse
                   lowerForExcel.Contains("event") OrElse lowerForExcel.Contains("reminder")
            wantsOutlookTask = (lowerForExcel.Contains("task") OrElse lowerForExcel.Contains("to-do") OrElse
                lowerForExcel.Contains("todo") OrElse lowerForExcel.Contains("to do")) AndAlso
               (lowerForExcel.Contains("outlook") OrElse lowerForExcel.Contains("create") OrElse
                lowerForExcel.Contains("add") OrElse lowerForExcel.Contains("new"))
            wantsOutlookContact = lowerForExcel.Contains("contact") AndAlso
                  (lowerForExcel.Contains("outlook") OrElse lowerForExcel.Contains("add") OrElse
                   lowerForExcel.Contains("create") OrElse lowerForExcel.Contains("save") OrElse
                   lowerForExcel.Contains("new"))
            wantsOutlookExport = (lowerForExcel.Contains("export") OrElse lowerForExcel.Contains("save as") OrElse
                  lowerForExcel.Contains("convert")) AndAlso
                 (lowerForExcel.Contains("email") OrElse lowerForExcel.Contains("mail")) AndAlso
                 (lowerForExcel.Contains("word") OrElse lowerForExcel.Contains("pdf") OrElse
                  lowerForExcel.Contains("docx"))
            wantsOneNoteFormat = lowerForExcel.Contains("onenote") OrElse lowerForExcel.Contains("one note") OrElse
                 lowerForExcel.Contains("notebook")
            wantsOneNoteChecklist = (lowerForExcel.Contains("checklist") OrElse lowerForExcel.Contains("checkbox") OrElse
                      lowerForExcel.Contains("to-do") OrElse lowerForExcel.Contains("todo") OrElse
                      lowerForExcel.Contains("tick") OrElse lowerForExcel.Contains("check off")) AndAlso
                     (lowerForExcel.Contains("onenote") OrElse lowerForExcel.Contains("note"))
            wantsOneNoteTable = lowerForExcel.Contains("table") AndAlso
                (lowerForExcel.Contains("onenote") OrElse lowerForExcel.Contains("note") OrElse
                 lowerForExcel.Contains("notebook"))
            wantsOneNoteAppend = (lowerForExcel.Contains("append") OrElse lowerForExcel.Contains("add to") OrElse
                   lowerForExcel.Contains("insert into") OrElse lowerForExcel.Contains("update")) AndAlso
                  (lowerForExcel.Contains("onenote") OrElse lowerForExcel.Contains("notebook") OrElse
                   lowerForExcel.Contains("note"))
            wantsOneNoteNewSection = (lowerForExcel.Contains("new section") OrElse lowerForExcel.Contains("new page") OrElse
                       lowerForExcel.Contains("create section") OrElse lowerForExcel.Contains("add section") OrElse
                       lowerForExcel.Contains("new notebook")) AndAlso
                      (lowerForExcel.Contains("onenote") OrElse lowerForExcel.Contains("notebook"))
            wantsOneNoteSearch = (lowerForExcel.Contains("search") OrElse lowerForExcel.Contains("find") OrElse
                   lowerForExcel.Contains("look for")) AndAlso
                  (lowerForExcel.Contains("onenote") OrElse lowerForExcel.Contains("note") OrElse
                   lowerForExcel.Contains("notebook"))
            wantsOneNoteExport = (lowerForExcel.Contains("export") OrElse lowerForExcel.Contains("save as") OrElse
                   lowerForExcel.Contains("convert")) AndAlso
                  (lowerForExcel.Contains("onenote") OrElse lowerForExcel.Contains("note")) AndAlso
                  (lowerForExcel.Contains("word") OrElse lowerForExcel.Contains("pdf") OrElse
                   lowerForExcel.Contains("docx"))
            wantsAccessCreate = (lowerForExcel.Contains("create") OrElse lowerForExcel.Contains("new") OrElse
                 lowerForExcel.Contains("make")) AndAlso
                (lowerForExcel.Contains("database") OrElse lowerForExcel.Contains("access") OrElse
                 lowerForExcel.Contains(".accdb") OrElse lowerForExcel.Contains("table"))
            wantsAccessQuery = (lowerForExcel.Contains("query") OrElse lowerForExcel.Contains("select") OrElse
                lowerForExcel.Contains("search") OrElse lowerForExcel.Contains("find") OrElse
                lowerForExcel.Contains("show") OrElse lowerForExcel.Contains("list") OrElse
                lowerForExcel.Contains("get")) AndAlso
               (lowerForExcel.Contains("database") OrElse lowerForExcel.Contains("access") OrElse
                lowerForExcel.Contains(".accdb") OrElse lowerForExcel.Contains("table") OrElse
                lowerForExcel.Contains("record"))
            wantsAccessImport = (lowerForExcel.Contains("import") OrElse lowerForExcel.Contains("load") OrElse
                 lowerForExcel.Contains("insert from") OrElse lowerForExcel.Contains("add from")) AndAlso
                (lowerForExcel.Contains("excel") OrElse lowerForExcel.Contains("csv") OrElse
                 lowerForExcel.Contains(".xlsx") OrElse lowerForExcel.Contains(".csv"))
            wantsAccessExport = (lowerForExcel.Contains("export") OrElse lowerForExcel.Contains("save as") OrElse
                 lowerForExcel.Contains("convert")) AndAlso
                (lowerForExcel.Contains("database") OrElse lowerForExcel.Contains("access") OrElse
                 lowerForExcel.Contains(".accdb")) AndAlso
                (lowerForExcel.Contains("excel") OrElse lowerForExcel.Contains("pdf") OrElse
                 lowerForExcel.Contains("csv"))
            wantsAccessReport = lowerForExcel.Contains("report") AndAlso
                (lowerForExcel.Contains("database") OrElse lowerForExcel.Contains("access") OrElse
                 lowerForExcel.Contains(".accdb"))
            wantsAccessModify = (lowerForExcel.Contains("modify") OrElse lowerForExcel.Contains("update") OrElse
                 lowerForExcel.Contains("edit") OrElse lowerForExcel.Contains("delete") OrElse
                 lowerForExcel.Contains("add record") OrElse lowerForExcel.Contains("insert") OrElse
                 lowerForExcel.Contains("alter")) AndAlso
                (lowerForExcel.Contains("database") OrElse lowerForExcel.Contains("access") OrElse
                 lowerForExcel.Contains(".accdb") OrElse lowerForExcel.Contains("table") OrElse
                 lowerForExcel.Contains("record"))
            Dim lowerForAccounts = userInput.ToLower()
            Dim hasUserKeyword = lowerForAccounts.Contains("user") OrElse lowerForAccounts.Contains("account") OrElse
            lowerForAccounts.Contains("login") OrElse lowerForAccounts.Contains("credentials")
            Dim hasLocalKeyword = lowerForAccounts.Contains("local") OrElse
            (Not lowerForAccounts.Contains("active directory") AndAlso Not lowerForAccounts.Contains(" ad ") AndAlso
             Not lowerForAccounts.Contains("domain"))
            wantsLocalUserCreate = hasUserKeyword AndAlso hasLocalKeyword AndAlso
            (lowerForAccounts.Contains("create") OrElse lowerForAccounts.Contains("add") OrElse
             lowerForAccounts.Contains("new"))
            wantsLocalUserModify = hasUserKeyword AndAlso hasLocalKeyword AndAlso
            (lowerForAccounts.Contains("modify") OrElse lowerForAccounts.Contains("change") OrElse
             lowerForAccounts.Contains("update") OrElse lowerForAccounts.Contains("rename") OrElse
             lowerForAccounts.Contains("enable") OrElse lowerForAccounts.Contains("disable") OrElse
             lowerForAccounts.Contains("password"))
            wantsLocalUserDelete = hasUserKeyword AndAlso hasLocalKeyword AndAlso
            (lowerForAccounts.Contains("delete") OrElse lowerForAccounts.Contains("remove"))
            wantsLocalUserList = hasUserKeyword AndAlso hasLocalKeyword AndAlso
            (lowerForAccounts.Contains("list") OrElse lowerForAccounts.Contains("show") OrElse
             lowerForAccounts.Contains("get") OrElse lowerForAccounts.Contains("all") OrElse
             lowerForAccounts.Contains("who") OrElse lowerForAccounts.Contains("display"))
            wantsLocalGroupManage = (lowerForAccounts.Contains("group") OrElse lowerForAccounts.Contains("administrators") OrElse
            lowerForAccounts.Contains("remote desktop users")) AndAlso hasLocalKeyword AndAlso
            (lowerForAccounts.Contains("add") OrElse lowerForAccounts.Contains("remove") OrElse
             lowerForAccounts.Contains("list") OrElse lowerForAccounts.Contains("create") OrElse
             lowerForAccounts.Contains("member"))
            Dim hasADKeyword = lowerForAccounts.Contains("active directory") OrElse lowerForAccounts.Contains(" ad ") OrElse
            lowerForAccounts.Contains("domain") OrElse lowerForAccounts.Contains("ldap") OrElse
            lowerForAccounts.Contains("ad user") OrElse lowerForAccounts.Contains("ad group") OrElse
            lowerForAccounts.Contains("domain controller") OrElse lowerForAccounts.Contains("ou ") OrElse
            lowerForAccounts.Contains("organizational unit")
            wantsADUserCreate = hasADKeyword AndAlso
            (lowerForAccounts.Contains("create") OrElse lowerForAccounts.Contains("new") OrElse
             lowerForAccounts.Contains("add")) AndAlso hasUserKeyword
            wantsADUserModify = hasADKeyword AndAlso
            (lowerForAccounts.Contains("modify") OrElse lowerForAccounts.Contains("update") OrElse
             lowerForAccounts.Contains("change") OrElse lowerForAccounts.Contains("enable") OrElse
             lowerForAccounts.Contains("disable") OrElse lowerForAccounts.Contains("unlock")) AndAlso hasUserKeyword
            wantsADUserDelete = hasADKeyword AndAlso
            (lowerForAccounts.Contains("delete") OrElse lowerForAccounts.Contains("remove")) AndAlso hasUserKeyword
            wantsADUserSearch = hasADKeyword AndAlso
            (lowerForAccounts.Contains("search") OrElse lowerForAccounts.Contains("find") OrElse
             lowerForAccounts.Contains("list") OrElse lowerForAccounts.Contains("get") OrElse
             lowerForAccounts.Contains("show") OrElse lowerForAccounts.Contains("query") OrElse
             lowerForAccounts.Contains("who") OrElse lowerForAccounts.Contains("lookup"))
            wantsADGroupManage = hasADKeyword AndAlso
            (lowerForAccounts.Contains("group")) AndAlso
            (lowerForAccounts.Contains("add") OrElse lowerForAccounts.Contains("remove") OrElse
             lowerForAccounts.Contains("create") OrElse lowerForAccounts.Contains("member") OrElse
             lowerForAccounts.Contains("list"))
            wantsADComputerManage = hasADKeyword AndAlso
            (lowerForAccounts.Contains("computer") OrElse lowerForAccounts.Contains("machine") OrElse
             lowerForAccounts.Contains("workstation") OrElse lowerForAccounts.Contains("server"))
            wantsADOUManage = hasADKeyword AndAlso
            (lowerForAccounts.Contains("ou") OrElse lowerForAccounts.Contains("organizational unit") OrElse
             lowerForAccounts.Contains("container"))
            wantsADPasswordReset = hasADKeyword AndAlso
            (lowerForAccounts.Contains("password") OrElse lowerForAccounts.Contains("reset")) AndAlso
            (lowerForAccounts.Contains("reset") OrElse lowerForAccounts.Contains("change") OrElse
             lowerForAccounts.Contains("set") OrElse lowerForAccounts.Contains("expire"))
            wantsADBulkOperation = hasADKeyword AndAlso
            (lowerForAccounts.Contains("bulk") OrElse lowerForAccounts.Contains("csv") OrElse
             lowerForAccounts.Contains("import") OrElse lowerForAccounts.Contains("multiple") OrElse
             lowerForAccounts.Contains("batch") OrElse lowerForAccounts.Contains("mass"))
            Dim isSceneInput As Boolean = userInput.Contains("=== AUTONO-ME SCENE RECORDING STARTED")
            If isSceneInput Then
                ResponseRichTextBox.AppendText("You: [Scene Recording]" & Environment.NewLine)
                QueryTextBox.Clear()
                Await HandleSceneRecording(userInput)
                Await DequeueNextTaskAsync()
                QueryTextBox.Focus()
                Return
            End If
            Dim textForSearch As String = If(isSceneInput, "", Regex.Replace(userInput,
"=== AUTONO-ME SCENE RECORDING STARTED.*?=== SCENE RECORDING STOPPED ===|" &
"=== AUTONO-ME RECORDING STARTED.*?=== RECORDING STOPPED ===",
"", RegexOptions.Singleline))
            Dim lowerInput = textForSearch.ToLower
            Dim wantsToOpen = lowerInput.Contains("open") AndAlso
        Not lowerInput.Contains("read") AndAlso
        Not lowerInput.Contains("summarize") AndAlso
        Not lowerInput.Contains("analyze") AndAlso
        Not lowerInput.Contains("explain") AndAlso
        Not lowerInput.Contains("write") AndAlso
        Not lowerInput.Contains("essay")
            Dim wantsToWrite = lowerInput.Contains("write") OrElse
        lowerInput.Contains("essay") OrElse
        lowerInput.Contains("summarize") OrElse
        lowerInput.Contains("explain") OrElse
        lowerInput.Contains("analyze") OrElse
        lowerInput.Contains("create") OrElse
        lowerInput.Contains("presentation") OrElse
        lowerInput.Contains("slides") OrElse
        lowerInput.Contains("spreadsheet") OrElse
        lowerInput.Contains("workbook") OrElse
        lowerInput.Contains("note") OrElse
        lowerInput.Contains("notes")
            Dim targetApp = ""
            If lowerInput.Contains("notepad") Then
                targetApp = "notepad"
            ElseIf lowerInput.Contains("onenote") OrElse lowerInput.Contains("one note") Then
                targetApp = "onenote"
            ElseIf lowerInput.Contains("powerpoint") OrElse lowerInput.Contains("ppt") OrElse lowerInput.Contains("presentation") OrElse lowerInput.Contains("slides") Then
                targetApp = "powerpoint"
            ElseIf lowerInput.Contains("excel") OrElse lowerInput.Contains("spreadsheet") OrElse lowerInput.Contains("xlsx") OrElse lowerInput.Contains("workbook") Then
                targetApp = "excel"
            ElseIf lowerInput.Contains("outlook") OrElse lowerInput.Contains("email") OrElse lowerInput.Contains("mail") OrElse lowerInput.Contains("calendar") OrElse lowerInput.Contains("appointment") OrElse lowerInput.Contains("meeting request") OrElse lowerInput.Contains("send meeting") Then
                targetApp = "outlook"
            ElseIf lowerInput.Contains("word") OrElse lowerInput.Contains("docx") Then
                targetApp = "word"
            End If
            Dim outlookAction = ""
            If targetApp = "outlook" Then
                If lowerInput.Contains("send") OrElse lowerInput.Contains("email") OrElse lowerInput.Contains("mail") Then
                    outlookAction = "email"
                ElseIf lowerInput.Contains("meeting") OrElse lowerInput.Contains("invite") Then
                    outlookAction = "meeting"
                ElseIf lowerInput.Contains("appointment") OrElse lowerInput.Contains("calendar") OrElse lowerInput.Contains("schedule") Then
                    outlookAction = "appointment"
                ElseIf lowerInput.Contains("task") OrElse lowerInput.Contains("todo") OrElse lowerInput.Contains("to-do") Then
                    outlookAction = "task"
                ElseIf lowerInput.Contains("contact") Then
                    outlookAction = "contact"
                ElseIf lowerInput.Contains("reply") Then
                    outlookAction = "reply"
                ElseIf lowerInput.Contains("forward") Then
                    outlookAction = "forward"
                End If
            End If
            Dim onenoteAction = ""
            If targetApp = "onenote" Then
                If lowerInput.Contains("quick note") OrElse lowerInput.Contains("side note") Then
                    onenoteAction = "quicknote"
                ElseIf lowerInput.Contains("checklist") OrElse lowerInput.Contains("to-do") OrElse lowerInput.Contains("todo") OrElse lowerInput.Contains("checkbox") Then
                    onenoteAction = "checklist"
                ElseIf lowerInput.Contains("meeting") Then
                    onenoteAction = "meeting"
                ElseIf lowerInput.Contains("section") Then
                    onenoteAction = "section"
                Else
                    onenoteAction = "note"
                End If
            End If
            ResponseRichTextBox.AppendText("You: " & userInput & Environment.NewLine)
            ResponseRichTextBox.SelectionStart = ResponseRichTextBox.TextLength
            ResponseRichTextBox.ScrollToCaret()
            QueryTextBox.Clear()
            Dim documentContent = ""
            selectedFiles.Clear()
            selectedFolder = ""
            mode = "single"
            Dim allowedExtensions = {
        ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".tif",
        ".txt", ".csv",
        ".xlsx", ".pptx", ".docx",
        ".pdf"
    }
            Dim isAllowedFile = Function(filePath As String) As Boolean
                                    Dim ext = Path.GetExtension(filePath).ToLower
                                    Return allowedExtensions.Contains(ext)
                                End Function
            Dim userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Dim downloads = Path.Combine(userProfile, "Downloads")
            Dim documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            Dim possibleFolders As New List(Of String)
            Dim possibleNames As New List(Of String)
            Dim skipWords = {
        "read", "explain", "downloads", "download", "documents", "document",
        "desktop", "summarize", "analyze", "write", "notepad", "word",
        "folder", "folders", "files", "file", "search", "find",
        "check", "inside", "this", "that", "these", "those",
        "and", "the", "from", "with", "about", "into", "essay",
        "open", "create", "put", "make", "generate",
        "powerpoint", "pptx", "presentation", "slides", "slide",
        "excel", "xlsx", "spreadsheet", "workbook", "sheet",
        "outlook", "email", "mail", "send", "calendar", "appointment",
        "meeting", "task", "contact", "reply", "forward", "schedule",
        "onenote", "note", "notes", "quick", "section", "notebook",
        "checklist", "checkbox", "todo"
    }
            Dim subfolderPattern = "(downloads?|documents?|desktop)[\\\/\s]+([A-Za-z0-9_\-]+(?:[\\\/][A-Za-z0-9_\-]+)*)"
            Dim subfolderMatches = Regex.Matches(textForSearch, subfolderPattern, RegexOptions.IgnoreCase)
            For Each m As Match In subfolderMatches
                Dim folderName = m.Groups(1).Value.Trim.ToLower
                Dim subPath = m.Groups(2).Value.Trim.Replace("/", "\")
                Dim baseFolderPath = ""
                If folderName.StartsWith("download") Then baseFolderPath = downloads
                If folderName.StartsWith("document") Then baseFolderPath = documents
                If folderName = "desktop" Then baseFolderPath = desktop
                If Not String.IsNullOrEmpty(baseFolderPath) Then
                    Dim fullPath = Path.Combine(baseFolderPath, subPath)
                    If Directory.Exists(fullPath) Then
                        If Not possibleFolders.Contains(fullPath) Then
                            possibleFolders.Add(fullPath)
                            ResponseRichTextBox.AppendText($"📁 Found subfolder: {fullPath}" & Environment.NewLine)
                        End If
                    Else
                        Try
                            Dim searchName = subPath.Split("\"c)(0)
                            Dim matchingFolders = Directory.GetDirectories(baseFolderPath, "*" & searchName & "*", SearchOption.TopDirectoryOnly)
                            For Each folder In matchingFolders
                                If Not possibleFolders.Contains(folder) Then
                                    possibleFolders.Add(folder)
                                    ResponseRichTextBox.AppendText($"📁 Found matching folder: {folder}" & Environment.NewLine)
                                End If
                            Next
                        Catch
                        End Try
                    End If
                End If
            Next
            If possibleFolders.Count = 0 Then
                If lowerInput.Contains("downloads") OrElse lowerInput.Contains("download") Then possibleFolders.Add(downloads)
                If lowerInput.Contains("documents") OrElse lowerInput.Contains("document") Then possibleFolders.Add(documents)
                If lowerInput.Contains("desktop") Then possibleFolders.Add(desktop)
            End If
            Dim explicitPathPattern = "([a-zA-Z]:\\[^\""<>|\r\n,\s]+)"
            Dim explicitPathMatches = Regex.Matches(textForSearch, explicitPathPattern, RegexOptions.IgnoreCase)
            For Each m As Match In explicitPathMatches
                Dim ePath = m.Value.Trim.TrimEnd("."c, ","c, ";"c, ":"c)
                If Directory.Exists(ePath) Then
                    If Not possibleFolders.Contains(ePath) Then
                        possibleFolders.Add(ePath)
                        ResponseRichTextBox.AppendText($"📁 Found explicit folder: {ePath}" & Environment.NewLine)
                    End If
                ElseIf File.Exists(ePath) AndAlso isAllowedFile(ePath) Then
                    If Not selectedFiles.Contains(ePath) Then
                        selectedFiles.Add(ePath)
                        ResponseRichTextBox.AppendText($"📄 Found explicit file: {ePath}" & Environment.NewLine)
                    End If
                End If
            Next
            Dim nameMatches = Regex.Matches(textForSearch, "([A-Za-z0-9_\-]{4,})")
            For Each m As Match In nameMatches
                Dim cleaned = m.Value.Trim
                If Not skipWords.Contains(cleaned.ToLower) Then
                    Dim isFolderName = possibleFolders.Any(Function(f) f.ToLower.EndsWith("\" & cleaned.ToLower))
                    If Not isFolderName AndAlso Not possibleNames.Contains(cleaned) Then
                        possibleNames.Add(cleaned)
                    End If
                End If
            Next
            Dim underscoreNamePattern = "([A-Za-z][A-Za-z0-9_\-]{8,}(?:_[A-Za-z0-9]+)+)"
            Dim underscoreMatches = Regex.Matches(textForSearch, underscoreNamePattern)
            For Each m As Match In underscoreMatches
                Dim candidate = m.Value.Trim()
                If candidate.Length > 8 AndAlso Not possibleNames.Contains(candidate) Then
                    possibleNames.Insert(0, candidate)
                    ResponseRichTextBox.AppendText($"🔍 Detected filename pattern: {candidate}" & Environment.NewLine)
                End If
            Next
            Dim likelyFiles = ExtractLikelyFilenames(textForSearch)
            For Each lf In likelyFiles
                If Not possibleNames.Contains(lf) Then
                    possibleNames.Insert(0, lf)
                    ResponseRichTextBox.AppendText($"🔍 Likely filename: {lf}" & Environment.NewLine)
                End If
            Next
            If possibleFolders.Count > 0 Then
                ResponseRichTextBox.AppendText($"🔍 Searching in: {String.Join(", ", possibleFolders.Select(Function(f) Path.GetFileName(f)))}" & Environment.NewLine)
            End If
            If possibleNames.Count > 0 Then
                ResponseRichTextBox.AppendText($"🔍 Looking for: {String.Join(", ", possibleNames)}" & Environment.NewLine)
            End If
            For Each folder As String In possibleFolders
                If Directory.Exists(folder) Then
                    If possibleNames.Count > 0 Then
                        For Each searchName As String In possibleNames
                            Dim folderCopy = folder
                            Dim nameCopy = searchName
                            Dim found = Await Task.Run(
                        Function()
                            Try
                                Return Directory.GetFiles(folderCopy, "*" & nameCopy & "*.*", SearchOption.TopDirectoryOnly) _
                                    .Where(Function(f) isAllowedFile(f)).ToArray
                            Catch
                                Return New String() {}
                            End Try
                        End Function)
                            For Each f In found
                                If Not selectedFiles.Contains(f) Then
                                    selectedFiles.Add(f)
                                    ResponseRichTextBox.AppendText($"📄 Found: {Path.GetFileName(f)}" & Environment.NewLine)
                                End If
                            Next
                        Next
                    Else
                        Dim folderCopy = folder
                        Dim allFiles = Await Task.Run(
                    Function()
                        Try
                            Return Directory.GetFiles(folderCopy, "*.*", SearchOption.TopDirectoryOnly) _
                                .Where(Function(f) isAllowedFile(f)).ToArray
                        Catch
                            Return New String() {}
                        End Try
                    End Function)
                        For Each f In allFiles
                            If Not selectedFiles.Contains(f) Then
                                selectedFiles.Add(f)
                                ResponseRichTextBox.AppendText($"📄 Found: {Path.GetFileName(f)}" & Environment.NewLine)
                            End If
                        Next
                    End If
                End If
            Next
            If selectedFiles.Count = 0 AndAlso possibleFolders.Count > 0 Then
                For Each folder As String In possibleFolders
                    If Directory.Exists(folder) Then
                        Dim folderCopy = folder
                        Dim allFiles = Await Task.Run(
                        Function()
                            Try
                                Return Directory.GetFiles(folderCopy, "*.*", SearchOption.TopDirectoryOnly) _
                                    .Where(Function(f) isAllowedFile(f)).ToArray
                            Catch
                                Return New String() {}
                            End Try
                        End Function)
                        For Each f In allFiles
                            If Not selectedFiles.Contains(f) Then
                                selectedFiles.Add(f)
                                ResponseRichTextBox.AppendText($"📄 Found: {Path.GetFileName(f)}" & Environment.NewLine)
                            End If
                        Next
                    End If
                Next
            End If
            If selectedFiles.Count > 1 Then
                mode = "multi"
            ElseIf selectedFiles.Count = 1 Then
                mode = "single"
            End If
            If wantsToOpen AndAlso selectedFiles.Count >= 1 Then
                ResponseRichTextBox.AppendText($"📂 Will open file: {Path.GetFileName(selectedFiles(0))}" & Environment.NewLine)
                documentContent = ""
            ElseIf selectedFiles.Count = 1 Then
                ResponseRichTextBox.AppendText($"📖 Reading: {Path.GetFileName(selectedFiles(0))}" & Environment.NewLine)
                Try
                    Dim analysis = Await FileAnalysisHelper.AnalyzeFileAsync(selectedFiles(0))
                    documentContent = analysis.Text
                Catch ex As Exception
                    documentContent = $"(Error reading file: {ex.Message})"
                End Try
            ElseIf selectedFiles.Count > 1 Then
                ResponseRichTextBox.AppendText($"📖 Reading {selectedFiles.Count} files..." & Environment.NewLine)
                Dim sb As New StringBuilder
                For Each f In selectedFiles
                    Try
                        Dim analysis = Await FileAnalysisHelper.AnalyzeFileAsync(f)
                        sb.AppendLine("=== FILE: " & Path.GetFileName(f) & " ===")
                        sb.AppendLine(analysis.Text)
                        sb.AppendLine()
                    Catch ex As Exception
                        sb.AppendLine("=== FILE: " & Path.GetFileName(f) & " ===")
                        sb.AppendLine($"(Error reading: {ex.Message})")
                        sb.AppendLine()
                    End Try
                Next
                documentContent = sb.ToString
            End If
            If documentContent.Length > 25000 Then
                documentContent = documentContent.Substring(0, 25000) & vbCrLf & "... [TRUNCATED - Document too long]"
            End If
            If documentContent.Length > 0 Then
                ResponseRichTextBox.AppendText($"✅ Extracted {documentContent.Length} characters" & Environment.NewLine)
            End If
            Dim instructionHint = ""
            Dim needsWebData As Boolean = False
            If targetApp = "excel" Then
                needsWebData = lowerInput.Contains("stock") OrElse
                lowerInput.Contains("price") OrElse
                lowerInput.Contains("history") OrElse
                lowerInput.Contains("financial") OrElse
                lowerInput.Contains("market") OrElse
                lowerInput.Contains("one year") OrElse
                lowerInput.Contains("1 year") OrElse
                lowerInput.Contains("yearly") OrElse
                lowerInput.Contains("annual")
            End If
            If needsWebData Then
                pendingTargetApp = "excel"
                Dim saveHintXl = ""
                If lowerInput.Contains("downloads") Then saveHintXl = "Downloads"
                If lowerInput.Contains("desktop") Then saveHintXl = "Desktop"
                pendingDocSavePath = saveHintXl
                documentContent = ""
                instructionHint = Environment.NewLine &
                "--- INSTRUCTION ---" & Environment.NewLine &
                "User wants LIVE STOCK/FINANCIAL DATA from the web written into an Excel spreadsheet." & Environment.NewLine &
                "You MUST respond with a Playwright JSON workflow wrapped in ||| delimiters." & Environment.NewLine &
                "Go DIRECTLY to Yahoo Finance history pages for each ticker." & Environment.NewLine &
                "URL format: https://finance.yahoo.com/quote/TICKER/history/" & Environment.NewLine &
                "For EACH ticker:" & Environment.NewLine &
                "  1. goto the Yahoo Finance history page" & Environment.NewLine &
                "  2. wait 3000ms" & Environment.NewLine &
                "  3. use gettext to extract ALL visible data from the page" & Environment.NewLine &
                "Put ALL tickers in ONE single JSON workflow." & Environment.NewLine &
                "Do NOT use NetworkIdle — Yahoo Finance keeps live connections open forever." & Environment.NewLine &
                "Do NOT generate PowerShell. Do NOT try to create Excel yourself." & Environment.NewLine &
                "The app will AUTOMATICALLY create the Excel file from the extracted data." & Environment.NewLine &
                "STOCK TICKER MAPPING: Microsoft=MSFT, Google=GOOGL, Meta=META, Nvidia=NVDA, Intel=INTC, Apple=AAPL, Amazon=AMZN, Tesla=TSLA" & Environment.NewLine &
                Environment.NewLine &
                "EXAMPLE RESPONSE for Microsoft and Google:" & Environment.NewLine &
                "Fetching stock data from Yahoo Finance..." & Environment.NewLine &
                "|||" & Environment.NewLine &
                "{" & Environment.NewLine &
                "  ""steps"": [" & Environment.NewLine &
                "    { ""action"": ""goto"", ""url"": ""https://finance.yahoo.com/quote/MSFT/history/"", ""timeoutMs"": 60000 }," & Environment.NewLine &
                "    { ""action"": ""wait"", ""timeoutMs"": 3000 }," & Environment.NewLine &
                "    { ""action"": ""gettext"" }," & Environment.NewLine &
                "    { ""action"": ""goto"", ""url"": ""https://finance.yahoo.com/quote/GOOGL/history/"", ""timeoutMs"": 60000 }," & Environment.NewLine &
                "    { ""action"": ""wait"", ""timeoutMs"": 3000 }," & Environment.NewLine &
                "    { ""action"": ""gettext"" }" & Environment.NewLine &
                "  ]" & Environment.NewLine &
                "}" & Environment.NewLine &
                "|||" & Environment.NewLine &
                "--- END INSTRUCTION ---"
            ElseIf wantsToOpen AndAlso selectedFiles.Count >= 1 Then
                Dim filePath = selectedFiles(0)
                instructionHint = Environment.NewLine &
            "--- INSTRUCTION ---" & Environment.NewLine &
            $"User wants to OPEN this file: {filePath}" & Environment.NewLine &
            "You MUST respond with PowerShell code wrapped in ||| delimiters." & Environment.NewLine &
            "CORRECT FORMAT:" & Environment.NewLine &
            "Opening file..." & Environment.NewLine &
            "|||" & Environment.NewLine &
            $"Start-Process ""{filePath}""" & Environment.NewLine &
            "|||" & Environment.NewLine &
            "DO NOT read or analyze the file. Just open it." & Environment.NewLine &
            "--- END INSTRUCTION ---"
            ElseIf documentContent.Length > 0 AndAlso wantsToWrite AndAlso
               targetApp <> "outlook" AndAlso targetApp <> "onenote" Then
                pendingDocContent = documentContent
                pendingTargetApp = If(String.IsNullOrEmpty(targetApp), "word", targetApp)
                Dim saveHint = ""
                If lowerInput.Contains("downloads") Then saveHint = "Downloads"
                If lowerInput.Contains("desktop") Then saveHint = "Desktop"
                pendingDocSavePath = saveHint
                Dim appHint = ""
                Select Case targetApp
                    Case "word"
                        appHint = "Open Microsoft Word using: Start-Process ""winword.exe"" ""/w"" and use Word COM ($word = New-Object -ComObject Word.Application) to write content."
                    Case "notepad"
                        appHint = "Open Notepad using: Start-Process ""notepad.exe"" and use clipboard method to write."
                    Case "powerpoint"
                        appHint = "Create a PowerPoint presentation using the PowerPoint.Application COM object."
                    Case "excel"
                        appHint = "Open Excel using: Start-Process ""excel.exe"" then AppActivate(""Excel"") then send {ENTER} to select Blank Workbook from the Start Screen."
                    Case Else
                        appHint = "Open the appropriate application and write the content there."
                End Select
                instructionHint = Environment.NewLine &
            "--- INSTRUCTION ---" & Environment.NewLine &
            "Document content has been extracted above. User wants you to WRITE output to an application." & Environment.NewLine &
            $"Target application: {If(String.IsNullOrEmpty(targetApp), "Word (default)", targetApp.ToUpper)}" & Environment.NewLine &
            appHint & Environment.NewLine &
            "You MUST generate PowerShell code wrapped in ||| delimiters to open the target app and write your analysis/summary/essay there." & Environment.NewLine &
            "DO NOT just describe the content - you must OPEN THE APP AND WRITE TO IT!" & Environment.NewLine &
            "--- END INSTRUCTION ---"
            ElseIf targetApp = "outlook" Then
                Dim shouldSendImmediately = lowerInput.Contains("send") AndAlso
                                        Not lowerInput.Contains("draft") AndAlso
                                        Not lowerInput.Contains("compose") AndAlso
                                        Not lowerInput.Contains("show me")
                Dim outlookHint = ""
                Select Case outlookAction
                    Case "email"
                        If shouldSendImmediately Then
                            outlookHint = "Create and SEND a new email using Outlook COM. Use $mail.Send() — do NOT use $mail.Display()."
                        Else
                            outlookHint = "Create and display a new email using Outlook COM: $ol = New-Object -ComObject Outlook.Application; $mail = $ol.CreateItem(0); ... $mail.Display()"
                        End If
                    Case "meeting"
                        outlookHint = "Create a meeting request using: $ol = New-Object -ComObject Outlook.Application; $appt = $ol.CreateItem(1); $appt.MeetingStatus = 1"
                    Case "appointment"
                        outlookHint = "Create a calendar appointment using: $ol = New-Object -ComObject Outlook.Application; $appt = $ol.CreateItem(1)"
                    Case "task"
                        outlookHint = "Create a task using: $ol = New-Object -ComObject Outlook.Application; $task = $ol.CreateItem(3)"
                    Case "contact"
                        outlookHint = "Create a contact using: $ol = New-Object -ComObject Outlook.Application; $contact = $ol.CreateItem(2)"
                    Case "reply"
                        outlookHint = "Reply to the selected email using Ctrl+R"
                    Case "forward"
                        outlookHint = "Forward the selected email using Ctrl+F"
                    Case Else
                        outlookHint = "Perform the requested Outlook action"
                End Select
                Dim sendOrDisplayNote = If(shouldSendImmediately,
                "$mail.Send()  # Send immediately — user said 'send'",
                "$mail.Display()  # Show the email for review — user said 'compose/draft'")
                instructionHint = Environment.NewLine &
            "--- INSTRUCTION ---" & Environment.NewLine &
            $"User wants to perform an Outlook action: {If(String.IsNullOrEmpty(outlookAction), "email", outlookAction)}" & Environment.NewLine &
            outlookHint & Environment.NewLine &
            "You MUST generate PowerShell code wrapped in ||| delimiters." & Environment.NewLine &
            "ALWAYS use Outlook COM automation — NEVER use SendKeys or AppActivate for Outlook." & Environment.NewLine &
            "PATTERN FOR EMAIL:" & Environment.NewLine &
            "$ol = New-Object -ComObject Outlook.Application" & Environment.NewLine &
            "$mail = $ol.CreateItem(0)" & Environment.NewLine &
            "$mail.To = 'recipient@example.com'" & Environment.NewLine &
            "$mail.Subject = 'Subject'" & Environment.NewLine &
            "$mail.Body = 'Body text'" & Environment.NewLine &
            "# For attachments:" & Environment.NewLine &
            "# $mail.Attachments.Add('C:\path\to\file.ext')" & Environment.NewLine &
            sendOrDisplayNote & Environment.NewLine &
            Environment.NewLine &
            "PATTERN FOR SCREENSHOT ATTACHMENT:" & Environment.NewLine &
            "Add-Type -AssemblyName System.Windows.Forms" & Environment.NewLine &
            "$bmp = New-Object System.Drawing.Bitmap([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width, [System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height)" & Environment.NewLine &
            "$g = [System.Drawing.Graphics]::FromImage($bmp)" & Environment.NewLine &
            "$g.CopyFromScreen(0,0,0,0,$bmp.Size)" & Environment.NewLine &
            "$ssPath = [IO.Path]::Combine($env:TEMP, 'DesktopScreenshot.png')" & Environment.NewLine &
            "$bmp.Save($ssPath)" & Environment.NewLine &
            "$g.Dispose(); $bmp.Dispose()" & Environment.NewLine &
            "$mail.Attachments.Add($ssPath)" & Environment.NewLine &
            "--- END INSTRUCTION ---"
            ElseIf targetApp = "onenote" Then
                Dim onenoteHint = ""
                Select Case onenoteAction
                    Case "quicknote"
                        onenoteHint = "Create a quick note using: Start-Process ""onenote.exe"" ""/sidenote"""
                    Case "checklist"
                        onenoteHint = "Create a checklist in OneNote. Open with: Start-Process ""onenote.exe"", create new page with Ctrl+N, use Ctrl+1 for checkboxes."
                    Case "meeting"
                        onenoteHint = "Create meeting notes in OneNote. Open with: Start-Process ""onenote.exe"", create new page with Ctrl+N, format with headings and bullets."
                    Case "section"
                        onenoteHint = "Create a new section in OneNote using Ctrl+T after opening OneNote."
                    Case Else
                        onenoteHint = "Create a note in OneNote. Open with: Start-Process ""onenote.exe"", create new page with Ctrl+N."
                End Select
                instructionHint = Environment.NewLine &
            "--- INSTRUCTION ---" & Environment.NewLine &
            $"User wants to create content in OneNote: {If(String.IsNullOrEmpty(onenoteAction), "note", onenoteAction)}" & Environment.NewLine &
            onenoteHint & Environment.NewLine &
            "You MUST generate PowerShell code wrapped in ||| delimiters." & Environment.NewLine &
            "Use clipboard method for longer content. Use Ctrl+1 for checkboxes, Ctrl+. for bullets, Ctrl+/ for numbered list." & Environment.NewLine &
            "First line typed after Ctrl+N becomes the page title." & Environment.NewLine &
            "--- END INSTRUCTION ---"
            ElseIf targetApp = "powerpoint" AndAlso wantsToWrite Then
                instructionHint = Environment.NewLine &
            "--- INSTRUCTION ---" & Environment.NewLine &
            "User wants to create a PowerPoint presentation." & Environment.NewLine &
            "You MUST generate a PowerShell script wrapped in ||| that uses the PowerPoint.Application COM object to create the slides, add titles, and add body text." & Environment.NewLine &
            "Do NOT use SendKeys or UIAutomation for PowerPoint." & Environment.NewLine &
            "--- END INSTRUCTION ---"
            ElseIf targetApp = "excel" AndAlso wantsToWrite Then
                instructionHint = Environment.NewLine &
            "--- INSTRUCTION ---" & Environment.NewLine &
            "User wants to create an Excel spreadsheet." & Environment.NewLine &
            "Open Excel using: Start-Process ""excel.exe"" then AppActivate(""Excel"") then send {ENTER} to select Blank Workbook." & Environment.NewLine &
            "Use Tab to move between cells, Enter to move down." & Environment.NewLine &
            "You MUST generate PowerShell code wrapped in ||| delimiters." & Environment.NewLine &
            "--- END INSTRUCTION ---"
            ElseIf targetApp = "word" AndAlso wantsToWrite Then
                instructionHint = Environment.NewLine &
            "--- INSTRUCTION ---" & Environment.NewLine &
            "User wants to create a Word document." & Environment.NewLine &
            "You MUST use Word COM automation:" & Environment.NewLine &
            "$word = New-Object -ComObject Word.Application" & Environment.NewLine &
            "$word.Visible = $true" & Environment.NewLine &
            "$doc = $word.Documents.Add()" & Environment.NewLine &
            "$sel = $word.Selection" & Environment.NewLine &
            "NEVER use SendKeys or AppActivate for Word. NEVER use ^n (creates second window). NEVER use ^{ENTER} (page break)." & Environment.NewLine &
            "You MUST generate PowerShell code wrapped in ||| delimiters." & Environment.NewLine &
            "--- END INSTRUCTION ---"
            End If
            Dim appListText =
        String.Join(Environment.NewLine,
                    allInstalledApps.Select(Function(a) $"{a.Name} = {a.AUMID}"))
            Dim filePathsInfo = ""
            If wantsToOpen AndAlso selectedFiles.Count >= 1 Then
                filePathsInfo = Environment.NewLine &
            "--- FILE TO OPEN ---" & Environment.NewLine &
            selectedFiles(0) & Environment.NewLine &
            "--- END FILE ---" & Environment.NewLine
            End If
            Dim targetAppInfo = ""
            If Not String.IsNullOrEmpty(targetApp) Then
                targetAppInfo = Environment.NewLine &
                "--- TARGET APP ---" & Environment.NewLine &
                $"Application: {targetApp.ToUpper}" & Environment.NewLine &
                If(targetApp = "outlook" AndAlso Not String.IsNullOrEmpty(outlookAction), $"Action: {outlookAction}" & Environment.NewLine, "") &
                If(targetApp = "onenote" AndAlso Not String.IsNullOrEmpty(onenoteAction), $"Action: {onenoteAction}" & Environment.NewLine, "") &
                "--- END TARGET APP ---" & Environment.NewLine
            End If
            Dim translatedScene As String = ""
            If userInput.Contains("=== AUTONO-ME SCENE RECORDING STARTED") OrElse
           userInput.Contains("=== AUTONO-ME RECORDING STARTED") Then
                translatedScene = TranslateSceneToSteps(userInput)
            End If
            Dim sceneHint As String = ""
            If Not String.IsNullOrWhiteSpace(translatedScene) Then
                Dim lowerScene = translatedScene.ToLower()
                Dim lowerRaw = userInput.ToLower()
                Dim needsAdmin = lowerScene.Contains("administrator") OrElse
                lowerScene.Contains("as administrator") OrElse
                lowerScene.Contains("search_result_admin") OrElse
                lowerRaw.Contains("run as administrator") OrElse
                lowerRaw.Contains("[click:search_result_admin]")
                Dim searchQueryInScene = ""
                Dim searchMatch = Regex.Match(translatedScene, "TYPE ""([^""]+)"" in (?:SearchHost|searchhost|SearchApp|explorer)",
                RegexOptions.IgnoreCase)
                If searchMatch.Success Then searchQueryInScene = searchMatch.Groups(1).Value
                If String.IsNullOrWhiteSpace(searchQueryInScene) Then
                    Dim launchMatch = Regex.Match(translatedScene, "LAUNCH ""([^""]+)""", RegexOptions.IgnoreCase)
                    If launchMatch.Success Then searchQueryInScene = launchMatch.Groups(1).Value
                End If
                Dim targetAppInScene = ""
                Dim appMatchScene = Regex.Match(userInput,
                "App:\s*'(mmc|regedit|cmd|powershell|taskmgr|devmgmt|diskmgmt|compmgmt|services|eventvwr|gpedit|msconfig|perfmon)'",
                RegexOptions.IgnoreCase)
                If appMatchScene.Success Then targetAppInScene = appMatchScene.Groups(1).Value
                sceneHint = Environment.NewLine &
                "--- SCENE REPLAY INSTRUCTION ---" & Environment.NewLine &
                "The user recorded a sequence of actions. You must REPLAY these steps using PowerShell." & Environment.NewLine &
                "TRANSLATED STEPS:" & Environment.NewLine &
                translatedScene & Environment.NewLine &
                Environment.NewLine &
                "REPLAY RULES:" & Environment.NewLine &
                "1. Use Add-Type -AssemblyName System.Windows.Forms for SendKeys" & Environment.NewLine &
                "2. Use Add-Type -AssemblyName Microsoft.VisualBasic for AppActivate" & Environment.NewLine &
                "3. For TASKBAR SEARCH: Do NOT use search UI — launch the app directly with Start-Process" & Environment.NewLine &
                "4. For Word/Excel/PowerPoint: Prefer COM automation over SendKeys" & Environment.NewLine &
                "5. Add Start-Sleep between steps (500ms minimum)" & Environment.NewLine &
                "6. MAXIMIZE every window after opening" & Environment.NewLine &
                "7. Execute ALL steps — do not stop partway through" & Environment.NewLine &
                "8. Wait for each dialog/window to appear before interacting with it" & Environment.NewLine &
                "9. Use Start-Sleep -Seconds 3 after launching apps to wait for them to load" & Environment.NewLine &
                "10. For buttons in dialogs: use AppActivate to focus the window, then SendKeys {ENTER} or {TAB}{ENTER}" & Environment.NewLine
                If needsAdmin Then
                    Dim adminApp = If(Not String.IsNullOrWhiteSpace(targetAppInScene),
                    targetAppInScene & ".exe",
                    searchQueryInScene)
                    sceneHint &=
                    Environment.NewLine &
                    "CRITICAL — RUN AS ADMINISTRATOR DETECTED:" & Environment.NewLine &
                    "The user launched an app from taskbar search AS ADMINISTRATOR." & Environment.NewLine &
                    "You MUST use Start-Process with -Verb RunAs to elevate:" & Environment.NewLine &
                    $"  Start-Process ""{adminApp}"" -Verb RunAs" & Environment.NewLine &
                    "Do NOT just run the app normally — it MUST be elevated." & Environment.NewLine &
                    "After launching with RunAs, a UAC prompt will appear." & Environment.NewLine &
                    "Wait 5 seconds for the user to approve UAC before continuing." & Environment.NewLine &
                    Environment.NewLine
                End If
                sceneHint &= "--- END SCENE INSTRUCTION ---"
            End If
            Dim isAccountTask = wantsLocalUserCreate OrElse wantsLocalUserModify OrElse
            wantsLocalUserDelete OrElse wantsLocalUserList OrElse wantsLocalGroupManage OrElse
            wantsADUserCreate OrElse wantsADUserModify OrElse wantsADUserDelete OrElse
            wantsADUserSearch OrElse wantsADGroupManage OrElse wantsADComputerManage OrElse
            wantsADOUManage OrElse wantsADPasswordReset OrElse wantsADBulkOperation
            If isAccountTask AndAlso String.IsNullOrEmpty(instructionHint) Then
                instructionHint = BuildAccountManagementHint(lowerForExcel)
            End If
            Dim combinedInput As String
            If needsWebData Then
                combinedInput =
                userInput &
                Environment.NewLine &
                instructionHint
            ElseIf Not String.IsNullOrWhiteSpace(translatedScene) Then
                combinedInput =
                userInput &
                Environment.NewLine &
                sceneHint &
                If(Not String.IsNullOrEmpty(instructionHint), instructionHint, "")
            Else
                combinedInput =
                userInput &
                Environment.NewLine &
                filePathsInfo &
                targetAppInfo &
                If(Not String.IsNullOrEmpty(documentContent),
                   "--- DOCUMENT CONTENT ---" & Environment.NewLine &
                   documentContent & Environment.NewLine &
                   "--- END DOCUMENT ---" & Environment.NewLine,
                   "") &
                instructionHint &
                "--- INSTALLED APPS ---" & Environment.NewLine &
                appListText
            End If
            Dim systemMessage = GetBootPrompt()
            Dim osPath = Path.Combine(Application.StartupPath, "OS")
            If File.Exists(osPath) Then
                Dim osContent = File.ReadAllText(osPath)
                systemMessage = systemMessage & Environment.NewLine & Environment.NewLine &
            "====================================================" & Environment.NewLine &
            "COGNITIVE OPERATING SYSTEM" & Environment.NewLine &
            "====================================================" & Environment.NewLine &
            osContent
            End If
            AddToHistory("user", combinedInput)
            Dim rawResponse = Await GetChatResponseWithHistoryAsync(systemMessage)
            Dim parsed = Await ParseResponseAsync(rawResponse)
            If parsed.Contains("REQUEST_SCREENSHOT") Then
                Dim base64Image = CaptureScreenBase64()
                rawResponse = Await GetChatResponseWithVisionAsync(
            systemMessage,
            "I have attached the screenshot and the logs. Fix the issue.",
            base64Image)
                parsed = Await ParseResponseAsync(rawResponse)
            End If
            AddToHistory("assistant", parsed)
            TrimConversationHistory()
            Dim confirmationText = parsed
            If Not parsed.Contains("|||") Then
                ResponseRichTextBox.AppendText("AI: " & parsed & Environment.NewLine)
            Else
                confirmationText = parsed.Split({"|||"}, StringSplitOptions.None)(0).Trim
                confirmationText = confirmationText.Replace(vbCr, "").Replace(vbLf, "").Trim
                If String.IsNullOrWhiteSpace(confirmationText) Then
                    confirmationText = "Command received. Executing automation..."
                End If
                ResponseRichTextBox.AppendText("AI: " & confirmationText & Environment.NewLine)
            End If
            ResponseRichTextBox.SelectionStart = ResponseRichTextBox.TextLength
            ResponseRichTextBox.ScrollToCaret()
            If parsed.Contains("QUESTION:") Then
                Dim questionText = parsed
                Dim qIndex = parsed.IndexOf("QUESTION:")
                If qIndex >= 0 Then questionText = parsed.Substring(qIndex + 9).Trim
                Dim newlineIdx = questionText.IndexOf(vbLf)
                If newlineIdx > 0 Then questionText = questionText.Substring(0, newlineIdx).Trim
                If questionText.Length > 100 Then questionText = questionText.Substring(0, 100) & "..."
                ShowNotification("AI needs your input", questionText)
            End If
            If parsed.Contains("NOTIFY:") Then
                Dim notifyText = parsed
                Dim nIndex = parsed.IndexOf("NOTIFY:")
                If nIndex >= 0 Then notifyText = parsed.Substring(nIndex + 7).Trim
                Dim newlineIdx = notifyText.IndexOf(vbLf)
                If newlineIdx > 0 Then notifyText = notifyText.Substring(0, newlineIdx).Trim
                If notifyText.Length > 150 Then notifyText = notifyText.Substring(0, 150) & "..."
                pendingNotification = notifyText
            ElseIf lowerInput.Contains("notify") OrElse
           lowerInput.Contains("let me know") OrElse
           lowerInput.Contains("tell me when") Then
                pendingNotification = "Your requested task has been executed."
            End If
            Dim pattern = "\|\|\|(.*?)\|\|\|"
            Dim match = Regex.Match(parsed, pattern, RegexOptions.Singleline)
            If match.Success Then
                Dim cmd = match.Groups(1).Value.Trim
                If cmd.StartsWith("powershell", StringComparison.OrdinalIgnoreCase) Then
                    cmd = cmd.Substring(10).Trim
                End If
                cmd = cmd.Replace("```", "").Trim
                gptResponse = parsed
            Else
                gptResponse = ""
                If String.IsNullOrEmpty(documentContent) AndAlso Not parsed.Contains("QUESTION:") AndAlso Not wantsToOpen AndAlso String.IsNullOrEmpty(targetApp) Then
                    ResponseRichTextBox.AppendText("AI did not return a command." & Environment.NewLine)
                End If
                ResponseRichTextBox.SelectionStart = ResponseRichTextBox.TextLength
                ResponseRichTextBox.ScrollToCaret()
            End If
        Catch ex As OperationCanceledException
            AppendResponse("⚠️ Operation was cancelled.")
        Catch ex As Exception
            MessageBox.Show("Critical Error: " & ex.Message)
        Finally
            Interlocked.Exchange(isRunning, 0)
            SetUIEnabled(True)
        End Try
        QueryTextBox.Focus()
    End Sub
    Private Function BuildAccountManagementHint(lowerInput As String) As String
        Dim sb As New StringBuilder()
        sb.AppendLine()
        sb.AppendLine("--- INSTRUCTION ---")
        Dim isAD = wantsADUserCreate OrElse wantsADUserModify OrElse wantsADUserDelete OrElse
        wantsADUserSearch OrElse wantsADGroupManage OrElse wantsADComputerManage OrElse
        wantsADOUManage OrElse wantsADPasswordReset OrElse wantsADBulkOperation
        If isAD Then
            sb.AppendLine("User wants to perform ACTIVE DIRECTORY management.")
            sb.AppendLine("You MUST generate PowerShell code wrapped in ||| delimiters.")
            sb.AppendLine()
            sb.AppendLine("AD PREREQUISITES CHECK — always start scripts with:")
            sb.AppendLine("try {")
            sb.AppendLine("    Import-Module ActiveDirectory -ErrorAction Stop")
            sb.AppendLine("    Write-Output 'AD module loaded successfully.'")
            sb.AppendLine("} catch {")
            sb.AppendLine("    Write-Output 'ERROR: Active Directory module not available.'")
            sb.AppendLine("    Write-Output 'This machine may not have RSAT tools installed.'")
            sb.AppendLine("    Write-Output 'Install with: Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'")
            sb.AppendLine("    exit 1")
            sb.AppendLine("}")
            sb.AppendLine()
            sb.AppendLine("DOMAIN DETECTION — get current domain automatically:")
            sb.AppendLine("$domain = (Get-ADDomain).DNSRoot")
            sb.AppendLine("$domainDN = (Get-ADDomain).DistinguishedName")
            sb.AppendLine()
            If wantsADUserCreate Then
                sb.AppendLine("AD USER CREATION PATTERN:")
                sb.AppendLine("$password = ConvertTo-SecureString 'TempP@ss123!' -AsPlainText -Force")
                sb.AppendLine("New-ADUser -Name 'John Smith' -GivenName 'John' -Surname 'Smith' `")
                sb.AppendLine("    -SamAccountName 'jsmith' -UserPrincipalName 'jsmith@domain.com' `")
                sb.AppendLine("    -Path 'OU=Users,DC=domain,DC=com' `")
                sb.AppendLine("    -AccountPassword $password -Enabled $true `")
                sb.AppendLine("    -ChangePasswordAtLogon $true")
                sb.AppendLine("Write-Output 'AD user created: jsmith'")
                sb.AppendLine()
                sb.AppendLine("IMPORTANT: Extract the username, first name, last name from the user's request.")
                sb.AppendLine("Use Get-ADDomain to auto-detect the domain for UPN and Path.")
                sb.AppendLine("Always set -ChangePasswordAtLogon $true for new accounts.")
            End If
            If wantsADUserModify Then
                sb.AppendLine("AD USER MODIFICATION PATTERNS:")
                sb.AppendLine("# Enable/Disable: Enable-ADAccount -Identity 'jsmith' / Disable-ADAccount -Identity 'jsmith'")
                sb.AppendLine("# Unlock: Unlock-ADAccount -Identity 'jsmith'")
                sb.AppendLine("# Modify properties: Set-ADUser -Identity 'jsmith' -Title 'Manager' -Department 'IT'")
                sb.AppendLine("# Move to OU: Move-ADObject -Identity (Get-ADUser 'jsmith').DistinguishedName -TargetPath 'OU=Managers,DC=domain,DC=com'")
            End If
            If wantsADUserDelete Then
                sb.AppendLine("AD USER DELETION:")
                sb.AppendLine("# First confirm the user exists:")
                sb.AppendLine("$user = Get-ADUser -Identity 'jsmith' -ErrorAction Stop")
                sb.AppendLine("# Then disable first (safety):")
                sb.AppendLine("Disable-ADAccount -Identity 'jsmith'")
                sb.AppendLine("# Then remove:")
                sb.AppendLine("Remove-ADUser -Identity 'jsmith' -Confirm:$false")
                sb.AppendLine("SAFETY: Always disable before deleting. Confirm user exists first.")
            End If
            If wantsADUserSearch Then
                sb.AppendLine("AD USER SEARCH PATTERNS:")
                sb.AppendLine("# List all users: Get-ADUser -Filter * -Properties DisplayName,Department,Title,Enabled | Select-Object Name,SamAccountName,DisplayName,Department,Title,Enabled | Format-Table -AutoSize")
                sb.AppendLine("# Search by name: Get-ADUser -Filter ""Name -like '*smith*'"" -Properties DisplayName,Department,Enabled")
                sb.AppendLine("# Search by department: Get-ADUser -Filter ""Department -eq 'IT'"" -Properties DisplayName,Department,Title")
                sb.AppendLine("# Locked out users: Search-ADAccount -LockedOut | Select-Object Name,SamAccountName,LockedOut")
                sb.AppendLine("# Disabled users: Search-ADAccount -AccountDisabled | Select-Object Name,SamAccountName")
                sb.AppendLine("# Expiring passwords: Search-ADAccount -PasswordExpiring -TimeSpan 7.00:00:00")
                sb.AppendLine("# Users not logged in 90 days: Search-ADAccount -AccountInactive -TimeSpan 90.00:00:00 -UsersOnly")
                sb.AppendLine()
                sb.AppendLine("ALWAYS format output for readability using Format-Table or custom output.")
                sb.AppendLine("For large result sets, pipe to Out-String and then Write-Output.")
            End If
            If wantsADGroupManage Then
                sb.AppendLine("AD GROUP MANAGEMENT:")
                sb.AppendLine("# List groups: Get-ADGroup -Filter * | Select-Object Name,GroupScope,GroupCategory | Format-Table")
                sb.AppendLine("# Get group members: Get-ADGroupMember -Identity 'GroupName' | Select-Object Name,SamAccountName | Format-Table")
                sb.AppendLine("# Add to group: Add-ADGroupMember -Identity 'GroupName' -Members 'jsmith'")
                sb.AppendLine("# Remove from group: Remove-ADGroupMember -Identity 'GroupName' -Members 'jsmith' -Confirm:$false")
                sb.AppendLine("# Create group: New-ADGroup -Name 'NewGroup' -GroupScope Global -GroupCategory Security -Path 'OU=Groups,DC=domain,DC=com'")
            End If
            If wantsADComputerManage Then
                sb.AppendLine("AD COMPUTER MANAGEMENT:")
                sb.AppendLine("# List computers: Get-ADComputer -Filter * -Properties OperatingSystem,LastLogonDate | Select-Object Name,OperatingSystem,LastLogonDate,Enabled | Format-Table")
                sb.AppendLine("# Search by OS: Get-ADComputer -Filter ""OperatingSystem -like '*Windows 11*'"" -Properties OperatingSystem")
                sb.AppendLine("# Disable computer: Disable-ADAccount -Identity 'COMPUTERNAME$'")
                sb.AppendLine("# Stale computers: Get-ADComputer -Filter ""LastLogonDate -lt '$((Get-Date).AddDays(-90).ToFileTime())'"" -Properties LastLogonDate")
            End If
            If wantsADOUManage Then
                sb.AppendLine("AD ORGANIZATIONAL UNIT MANAGEMENT:")
                sb.AppendLine("# List OUs: Get-ADOrganizationalUnit -Filter * | Select-Object Name,DistinguishedName | Format-Table")
                sb.AppendLine("# Create OU: New-ADOrganizationalUnit -Name 'NewOU' -Path 'DC=domain,DC=com' -ProtectedFromAccidentalDeletion $true")
                sb.AppendLine("# Move object to OU: Move-ADObject -Identity $object.DistinguishedName -TargetPath 'OU=Target,DC=domain,DC=com'")
            End If
            If wantsADPasswordReset Then
                sb.AppendLine("AD PASSWORD MANAGEMENT:")
                sb.AppendLine("# Reset password: Set-ADAccountPassword -Identity 'jsmith' -Reset -NewPassword (ConvertTo-SecureString 'NewTempP@ss123!' -AsPlainText -Force)")
                sb.AppendLine("# Force change at next logon: Set-ADUser -Identity 'jsmith' -ChangePasswordAtLogon $true")
                sb.AppendLine("# Unlock account: Unlock-ADAccount -Identity 'jsmith'")
                sb.AppendLine("# Check password expiry: Get-ADUser 'jsmith' -Properties PasswordLastSet,PasswordExpired | Select-Object Name,PasswordLastSet,PasswordExpired")
            End If
            If wantsADBulkOperation Then
                sb.AppendLine("AD BULK OPERATIONS:")
                sb.AppendLine("# Import from CSV: Import-Csv 'C:\path\users.csv' | ForEach-Object {")
                sb.AppendLine("    New-ADUser -Name $_.Name -SamAccountName $_.SAM -UserPrincipalName ($_.SAM + '@domain.com') `")
                sb.AppendLine("        -AccountPassword (ConvertTo-SecureString $_.Password -AsPlainText -Force) -Enabled $true")
                sb.AppendLine("}")
                sb.AppendLine("# CSV format: Name,SAM,Password,Department,Title")
                sb.AppendLine("# Bulk disable: Get-ADUser -Filter ""Department -eq 'OldDept'"" | Disable-ADAccount")
                sb.AppendLine("# Bulk move: Get-ADUser -Filter ""Department -eq 'IT'"" | Move-ADObject -TargetPath 'OU=IT,DC=domain,DC=com'")
            End If
        Else
            sb.AppendLine("User wants to perform LOCAL USER ACCOUNT management.")
            sb.AppendLine("You MUST generate PowerShell code wrapped in ||| delimiters.")
            sb.AppendLine("IMPORTANT: Local user management requires ADMINISTRATOR privileges.")
            sb.AppendLine("If the script fails with access denied, the user needs to run Autono-Me as admin.")
            sb.AppendLine()
            If wantsLocalUserList Then
                sb.AppendLine("LIST LOCAL USERS:")
                sb.AppendLine("Get-LocalUser | Select-Object Name,Enabled,LastLogon,Description | Format-Table -AutoSize | Out-String | Write-Output")
                sb.AppendLine("# For detailed info: Get-LocalUser | ForEach-Object { $_ | Select-Object * } | Format-List")
                sb.AppendLine("# List group memberships: Get-LocalGroup | ForEach-Object { $g = $_.Name; Get-LocalGroupMember $g -ErrorAction SilentlyContinue | Select-Object @{N='Group';E={$g}},Name,ObjectClass } | Format-Table")
            End If
            If wantsLocalUserCreate Then
                sb.AppendLine("CREATE LOCAL USER:")
                sb.AppendLine("$password = ConvertTo-SecureString 'TempP@ss123!' -AsPlainText -Force")
                sb.AppendLine("New-LocalUser -Name 'Username' -Password $password -FullName 'Full Name' -Description 'Description' -AccountNeverExpires")
                sb.AppendLine("# Add to group: Add-LocalGroupMember -Group 'Users' -Member 'Username'")
                sb.AppendLine("# Add as admin: Add-LocalGroupMember -Group 'Administrators' -Member 'Username'")
                sb.AppendLine()
                sb.AppendLine("EXTRACT username and password from the user's request.")
                sb.AppendLine("If no password specified, generate a temporary one and display it.")
                sb.AppendLine("ALWAYS output the credentials at the end so the user knows them.")
            End If
            If wantsLocalUserModify Then
                sb.AppendLine("MODIFY LOCAL USER:")
                sb.AppendLine("# Change password: Set-LocalUser -Name 'Username' -Password (ConvertTo-SecureString 'NewP@ss!' -AsPlainText -Force)")
                sb.AppendLine("# Enable: Enable-LocalUser -Name 'Username'")
                sb.AppendLine("# Disable: Disable-LocalUser -Name 'Username'")
                sb.AppendLine("# Rename: Rename-LocalUser -Name 'OldName' -NewName 'NewName'")
                sb.AppendLine("# Change description: Set-LocalUser -Name 'Username' -Description 'New description'")
                sb.AppendLine("# Set password to never expire: Set-LocalUser -Name 'Username' -PasswordNeverExpires $true")
            End If
            If wantsLocalUserDelete Then
                sb.AppendLine("DELETE LOCAL USER:")
                sb.AppendLine("# Verify user exists first:")
                sb.AppendLine("$user = Get-LocalUser -Name 'Username' -ErrorAction Stop")
                sb.AppendLine("# Remove: Remove-LocalUser -Name 'Username'")
                sb.AppendLine("# IMPORTANT: This does NOT delete the user profile folder.")
                sb.AppendLine("# To also remove profile: Remove-Item ""C:\Users\Username"" -Recurse -Force")
                sb.AppendLine("SAFETY: Always confirm the user exists before attempting deletion.")
            End If
            If wantsLocalGroupManage Then
                sb.AppendLine("LOCAL GROUP MANAGEMENT:")
                sb.AppendLine("# List groups: Get-LocalGroup | Format-Table Name,Description")
                sb.AppendLine("# List members: Get-LocalGroupMember -Group 'Administrators' | Format-Table Name,ObjectClass")
                sb.AppendLine("# Add to group: Add-LocalGroupMember -Group 'Administrators' -Member 'Username'")
                sb.AppendLine("# Remove from group: Remove-LocalGroupMember -Group 'Administrators' -Member 'Username'")
                sb.AppendLine("# Create group: New-LocalGroup -Name 'GroupName' -Description 'Description'")
            End If
        End If
        sb.AppendLine()
        sb.AppendLine("GENERAL RULES FOR ACCOUNT MANAGEMENT:")
        sb.AppendLine("1. ALWAYS output results using Write-Output so the user can see what happened")
        sb.AppendLine("2. ALWAYS include error handling with try/catch")
        sb.AppendLine("3. For password creation, use ConvertTo-SecureString")
        sb.AppendLine("4. ALWAYS display the final state of the account after modification")
        sb.AppendLine("5. For destructive operations (delete), confirm the target exists first")
        sb.AppendLine("6. Format all output tables for readability")
        sb.AppendLine("--- END INSTRUCTION ---")
        Return sb.ToString()
    End Function
    Private Async Sub ShellTimer_Tick(sender As Object, e As EventArgs) Handles ShellTimer.Tick
        If Interlocked.CompareExchange(isRunning, 1, 0) <> 0 Then Return
        Try
            Dim content As String = gptResponse
            If String.IsNullOrWhiteSpace(content) Then
                Return
            End If
            Dim pattern As String = "\|\|\|(.*?)\|\|\|"
            Dim match As Match = Regex.Match(content, pattern, RegexOptions.Singleline)
            If Not match.Success Then
                gptResponse = ""
                Return
            End If
            Dim currentCommandPS As String = match.Groups(1).Value.Trim()
            If currentCommandPS.StartsWith("powershell", StringComparison.OrdinalIgnoreCase) Then
                currentCommandPS = currentCommandPS.Substring(10).Trim()
            End If
            currentCommandPS = currentCommandPS.Replace("```", "").Trim()
            Dim actualAppPath = Application.StartupPath.TrimEnd(Path.DirectorySeparatorChar)
            currentCommandPS = currentCommandPS.Replace("APP_PATH_PLACEHOLDER", actualAppPath)
            If String.IsNullOrWhiteSpace(currentCommandPS) Then
                gptResponse = ""
                Return
            End If
            Dim windowsDirPattern As String = "(?i)([a-z]:\\windows(?:\\|$))|(\$env:windir)|(\$env:systemroot)"
            If Regex.IsMatch(currentCommandPS, windowsDirPattern) Then
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("🛑 SECURITY OVERRIDE: Attempted interaction with the Windows system folder was blocked." & Environment.NewLine)
                                                    ResponseRichTextBox.SelectionStart = ResponseRichTextBox.TextLength
                                                    ResponseRichTextBox.ScrollToCaret()
                                                End Sub)
                gptResponse = ""
                commandExecutedPS = True
                Return
            End If
            Dim normalizedCurrent = NormalizeScript(currentCommandPS)
            Dim normalizedPrevious = NormalizeScript(previousCommandPS)
            If normalizedCurrent <> normalizedPrevious Then
                commandExecutedPS = False
                previousCommandPS = normalizedCurrent
            End If
            If Not commandExecutedPS Then
                gptResponse = ""
                commandExecutedPS = True
                Try
                    Dim trimmedCmd = currentCommandPS.Trim()
                    Dim looksLikeJson = trimmedCmd.StartsWith("{") AndAlso
                    (trimmedCmd.Contains("""steps""") OrElse trimmedCmd.Contains("""action"""))
                    If looksLikeJson Then
                        ResponseRichTextBox.BeginInvoke(Sub()
                                                            ResponseRichTextBox.AppendText("🌐 Running workflow..." & Environment.NewLine)
                                                            ResponseRichTextBox.SelectionStart = ResponseRichTextBox.TextLength
                                                            ResponseRichTextBox.ScrollToCaret()
                                                        End Sub)
                        Dim playwrightSuccess = Await RunAgentWorkflowAsync(trimmedCmd)
                        Dim autoWriteHandled As Boolean = False
                        If Not String.IsNullOrWhiteSpace(pendingDocContent) Then
                            Dim combinedContent = pendingDocContent
                            If Not String.IsNullOrWhiteSpace(lastExtractedWebText) Then
                                combinedContent = "=== LIVE WEB DATA ===" & Environment.NewLine &
                                    lastExtractedWebText & Environment.NewLine &
                                    "=== DOCUMENT CONTENT ===" & Environment.NewLine &
                                    pendingDocContent
                                lastExtractedWebText = ""
                            End If
                            Await AutoWriteDocumentAsync(combinedContent, pendingTargetApp, pendingDocSavePath)
                            pendingDocContent = ""
                            pendingTargetApp = ""
                            pendingDocSavePath = ""
                            autoWriteHandled = True
                        ElseIf Not String.IsNullOrWhiteSpace(lastExtractedWebText) Then
                            Dim webTarget As String = If(String.IsNullOrWhiteSpace(pendingTargetApp), "word", pendingTargetApp)
                            Await AutoWriteDocumentAsync(lastExtractedWebText, webTarget, pendingDocSavePath)
                            lastExtractedWebText = ""
                            pendingTargetApp = ""
                            pendingDocSavePath = ""
                            autoWriteHandled = True
                        Else
                            Dim allMatches = Regex.Matches(content, "\|\|\|(.*?)\|\|\|", RegexOptions.Singleline)
                            If allMatches.Count >= 2 Then
                                Dim secondBlock = allMatches(1).Groups(1).Value.Trim()
                                Dim looksLikePS = secondBlock.Contains("$") OrElse secondBlock.Contains("New-Object") OrElse
                                    secondBlock.Contains("Start-Process") OrElse secondBlock.Contains("Add-Type")
                                If Not String.IsNullOrWhiteSpace(secondBlock) AndAlso
                                   Not secondBlock.StartsWith("{") AndAlso looksLikePS Then
                                    Dim safeCheck = IsCommandSafe(secondBlock)
                                    If safeCheck.Safe Then
                                        Dim preview = secondBlock.Substring(0, Math.Min(60, secondBlock.Length)).Replace(Environment.NewLine, " ")
                                        ResponseRichTextBox.BeginInvoke(Sub()
                                                                            ResponseRichTextBox.AppendText("📝 Running follow-up PS: " & preview & "..." & Environment.NewLine)
                                                                        End Sub)
                                        Await Task.Run(Sub() RunPowerShellSync(secondBlock))
                                    Else
                                        ResponseRichTextBox.BeginInvoke(Sub()
                                                                            ResponseRichTextBox.AppendText($"🛑 Follow-up PS blocked: {safeCheck.Reason}" & Environment.NewLine)
                                                                        End Sub)
                                    End If
                                End If
                            End If
                        End If
                        If autoWriteHandled Then
                            Dim allMatches2 = Regex.Matches(content, "\|\|\|(.*?)\|\|\|", RegexOptions.Singleline)
                            If allMatches2.Count >= 2 Then
                                ResponseRichTextBox.BeginInvoke(Sub()
                                                                    ResponseRichTextBox.AppendText("⏭️ Document already created, skipping follow-up script." & Environment.NewLine)
                                                                End Sub)
                            End If
                        End If
                        If Not String.IsNullOrEmpty(pendingNotification) Then
                            ShowNotification("Task Complete", pendingNotification)
                            pendingNotification = ""
                        End If
                        Await DequeueNextTaskAsync()
                        Return
                    End If
                    Dim isCompleteOfficeScript = currentCommandPS.Contains("New-Object -ComObject") AndAlso
                    (currentCommandPS.Contains("PowerPoint.Application") OrElse
                     currentCommandPS.Contains("Word.Application") OrElse
                     currentCommandPS.Contains("Excel.Application") OrElse
                     currentCommandPS.Contains("Outlook.Application") OrElse
                     currentCommandPS.Contains("OneNote.Application"))
                    If isCompleteOfficeScript Then
                        pendingDocContent = ""
                        pendingTargetApp = ""
                        pendingDocSavePath = ""
                        ResponseRichTextBox.BeginInvoke(Sub()
                                                            ResponseRichTextBox.AppendText("📝 Executing Office automation script..." & Environment.NewLine)
                                                            ResponseRichTextBox.SelectionStart = ResponseRichTextBox.TextLength
                                                            ResponseRichTextBox.ScrollToCaret()
                                                        End Sub)
                    End If
                    Dim logPath As String = Path.Combine(Application.StartupPath, "LOG")
                    Dim commandLog As String = $"=== COMMAND EXECUTED: {DateTime.Now} ==={Environment.NewLine}" &
                    $"{currentCommandPS}{Environment.NewLine}" &
                    $"==========================================={Environment.NewLine}"
                    File.AppendAllText(logPath, commandLog)
                    Dim scriptFile = GetUniqueTempPath("ps_exec", ".ps1")
                    File.WriteAllText(scriptFile, currentCommandPS, Encoding.UTF8)
                    Dim psi As New ProcessStartInfo("powershell.exe") With {
    .WindowStyle = ProcessWindowStyle.Hidden,
    .CreateNoWindow = True,
    .UseShellExecute = False,
    .RedirectStandardOutput = True,
    .RedirectStandardError = True,
    .Arguments = $"-NoProfile -ExecutionPolicy Bypass -File ""{scriptFile}"""
}
                    Dim output As String = ""
                    Dim errors As String = ""
                    Using proc As Process = Process.Start(psi)
                        Dim outputTask As Task(Of String) = proc.StandardOutput.ReadToEndAsync()
                        Dim errorTask As Task(Of String) = proc.StandardError.ReadToEndAsync()
                        Await proc.WaitForExitAsync()
                        output = Await outputTask
                        errors = Await errorTask
                    End Using
                    Try : File.Delete(scriptFile) : Catch : End Try
                    Dim resultLog As String =
                    $"=== EXECUTION RESULT: {DateTime.Now} ==={Environment.NewLine}" &
                    $"OUTPUT:{Environment.NewLine}{output}{Environment.NewLine}" &
                    $"ERRORS:{Environment.NewLine}{errors}{Environment.NewLine}" &
                    $"==========================================={Environment.NewLine}"
                    File.AppendAllText(logPath, resultLog)
                    Dim hasOutput = Not String.IsNullOrWhiteSpace(output)
                    Dim hasErrors = Not String.IsNullOrWhiteSpace(errors)
                    Dim hasSuccessSignal As Boolean = False
                    If hasOutput Then
                        Dim successKeywords As String() = {
                        "Saved to", "Success", "Done", "Complete",
                        "Created", "Opened", "Saved", "Write-Output"
                    }
                        For Each kw In successKeywords
                            If output.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                hasSuccessSignal = True
                                Exit For
                            End If
                        Next
                    End If
                    Dim isFatalError As Boolean = False
                    If hasErrors AndAlso Not hasSuccessSignal Then
                        Dim fatalKeywords As String() = {
                        "Exception", "TerminatingError", "not recognized",
                        "cannot be loaded", "is not valid", "Access is denied",
                        "UnauthorizedAccess", "The term", "Cannot find",
                        "does not exist", "failed with", "At line:",
                        "CategoryInfo", "FullyQualifiedErrorId",
                        "is not a cmdlet", "cannot open", "No such file"
                    }
                        Dim errorLines = errors.Split(
                        New String() {Environment.NewLine, Chr(10)},
                        StringSplitOptions.RemoveEmptyEntries)
                        For Each line In errorLines
                            Dim trimmedLine = line.Trim()
                            If String.IsNullOrWhiteSpace(trimmedLine) Then Continue For
                            If trimmedLine.StartsWith("WARNING:", StringComparison.OrdinalIgnoreCase) Then Continue For
                            If trimmedLine.StartsWith("VERBOSE:", StringComparison.OrdinalIgnoreCase) Then Continue For
                            If trimmedLine.StartsWith("DEBUG:", StringComparison.OrdinalIgnoreCase) Then Continue For
                            If trimmedLine.StartsWith("INFORMATION:", StringComparison.OrdinalIgnoreCase) Then Continue For
                            For Each keyword In fatalKeywords
                                If trimmedLine.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                    isFatalError = True
                                    Exit For
                                End If
                            Next
                            If isFatalError Then Exit For
                        Next
                    End If
                    If hasSuccessSignal AndAlso Not String.IsNullOrWhiteSpace(pendingDocContent) Then
                        pendingDocContent = ""
                        pendingTargetApp = ""
                        pendingDocSavePath = ""
                    End If
                    If isFatalError AndAlso correctionAttempts < MaxCorrectionAttempts Then
                        Dim feedbackMsg As String =
                        "The PowerShell script produced errors:" &
                        Environment.NewLine &
                        "ERRORS:" & Environment.NewLine & errors.Trim() & Environment.NewLine &
                        If(hasOutput, "OUTPUT:" & Environment.NewLine & output.Trim() & Environment.NewLine, "") &
                        Environment.NewLine &
                        "Provide a corrected PowerShell script wrapped in ||| delimiters."
                        AddToHistory("user", feedbackMsg)
                        Dim systemMessage = GetBootPrompt()
                        Dim correctionResponse = Await GetChatResponseWithHistoryAsync(systemMessage)
                        Dim correctionParsed = Await ParseResponseAsync(correctionResponse)
                        AddToHistory("assistant", correctionParsed)
                        TrimConversationHistory()
                        Dim correctionMatch = Regex.Match(correctionParsed, "\|\|\|(.*?)\|\|\|", RegexOptions.Singleline)
                        Dim isDone = correctionParsed.Trim().Equals("DONE", StringComparison.OrdinalIgnoreCase)
                        If correctionMatch.Success AndAlso Not isDone Then
                            correctionAttempts += 1
                            gptResponse = correctionParsed
                            commandExecutedPS = False
                            ResponseRichTextBox.BeginInvoke(Sub()
                                                                ResponseRichTextBox.AppendText($"🔄 AI self-correcting (attempt {correctionAttempts}/{MaxCorrectionAttempts})..." & Environment.NewLine)
                                                                ResponseRichTextBox.SelectionStart = ResponseRichTextBox.TextLength
                                                                ResponseRichTextBox.ScrollToCaret()
                                                            End Sub)
                        Else
                            correctionAttempts = 0
                            If Not String.IsNullOrEmpty(pendingNotification) Then
                                ShowNotification("Task Complete", pendingNotification)
                                pendingNotification = ""
                            End If
                            Await DequeueNextTaskAsync()
                        End If
                    ElseIf correctionAttempts >= MaxCorrectionAttempts Then
                        correctionAttempts = 0
                        ResponseRichTextBox.BeginInvoke(Sub()
                                                            ResponseRichTextBox.AppendText($"⚠️ Max correction attempts reached. Please check the result manually." & Environment.NewLine)
                                                            ResponseRichTextBox.SelectionStart = ResponseRichTextBox.TextLength
                                                            ResponseRichTextBox.ScrollToCaret()
                                                        End Sub)
                        If Not String.IsNullOrEmpty(pendingNotification) Then
                            ShowNotification("Task Complete", pendingNotification)
                            pendingNotification = ""
                        End If
                        Await DequeueNextTaskAsync()
                    Else
                        If Not String.IsNullOrEmpty(pendingNotification) Then
                            ShowNotification("Task Complete", pendingNotification)
                            pendingNotification = ""
                        End If
                        Await DequeueNextTaskAsync()
                    End If
                Catch ex As Exception
                    Debug.WriteLine("PS Execution Error: " & ex.Message)
                    LogToFile("PS_EXEC_ERROR", ex.ToString())
                End Try
            End If
        Catch ex As Exception
            Debug.WriteLine("Execution Error: " & ex.Message)
            LogToFile("TIMER_ERROR", ex.ToString())
        Finally
            Interlocked.Exchange(isRunning, 0)
        End Try
    End Sub
    Private Async Function ExecutePowerShellCommand(command As String, fullContent As String,
                                                 Optional ct As CancellationToken = Nothing) As Task
        Dim logPath As String = Path.Combine(Application.StartupPath, "LOG")
        Dim separator As String = "==========================================="
        Dim commandLog As String = $"=== COMMAND EXECUTED: {DateTime.Now} ==={Environment.NewLine}{command}{Environment.NewLine}{separator}{Environment.NewLine}"
        Try
            File.AppendAllText(logPath, commandLog)
        Catch ex As Exception
            LogToFile("LOG_WRITE_ERROR", ex.Message)
        End Try
        If ct.IsCancellationRequested Then
            AppendResponse("⚠️ PowerShell command was cancelled before starting.")
            Return
        End If
        Dim safeCheck = IsCommandSafe(command)
        If Not safeCheck.Safe Then
            AppendResponse($"🛑 SECURITY: {safeCheck.Reason}")
            SendCompletionNotification()
            Await DequeueNextTaskAsync()
            Return
        End If
        Dim bytes As Byte() = Encoding.Unicode.GetBytes(command)
        Dim base64Command As String = Convert.ToBase64String(bytes)
        Dim scriptFile = GetUniqueTempPath("ps_cmd", ".ps1")
        File.WriteAllText(scriptFile, command, Encoding.UTF8)
        Dim psi As New ProcessStartInfo("powershell.exe") With {
    .WindowStyle = ProcessWindowStyle.Hidden,
    .CreateNoWindow = True,
    .UseShellExecute = False,
    .RedirectStandardOutput = True,
    .RedirectStandardError = True,
    .Arguments = $"-NoProfile -ExecutionPolicy Bypass -File ""{scriptFile}"""
}
        Dim output As String = ""
        Dim errors As String = ""
        Dim wasCancelled As Boolean = False
        Dim wasStartError As Boolean = False
        Dim wasGeneralError As Boolean = False
        Dim wasTimeout As Boolean = False
        Dim errorMessage As String = ""
        Try
            Using proc = StartProcessSafe(psi)
                Dim registration As CancellationTokenRegistration = Nothing
                If Not ct.Equals(CancellationToken.None) Then
                    registration = ct.Register(Sub()
                                                   Try
                                                       If Not proc.HasExited Then
                                                           proc.Kill()
                                                       End If
                                                   Catch
                                                   End Try
                                               End Sub)
                End If
                Try
                    Dim outputTask = proc.StandardOutput.ReadToEndAsync()
                    Dim errorTask = proc.StandardError.ReadToEndAsync()
                    Dim exited = Await Task.Run(Function() proc.WaitForExit(CommandTimeoutMs))
                    If Not exited Then
                        Try
                            If Not proc.HasExited Then proc.Kill()
                        Catch
                        End Try
                        wasTimeout = True
                    Else
                        If ct.IsCancellationRequested Then
                            wasCancelled = True
                        Else
                            output = Await outputTask
                            errors = Await errorTask
                        End If
                    End If
                Finally
                    If Not registration.Equals(Nothing) Then
                        registration.Dispose()
                    End If
                End Try
            End Using
            Try : File.Delete(scriptFile) : Catch : End Try
        Catch ex As OperationCanceledException
            wasCancelled = True
        Catch ex As InvalidOperationException
            wasStartError = True
            errorMessage = ex.Message
            LogToFile("PS_START_ERROR", ex.ToString())
        Catch ex As Exception
            wasGeneralError = True
            errorMessage = ex.Message
            LogToFile("PS_EXEC_ERROR", ex.ToString())
        End Try
        If wasCancelled Then
            AppendResponse("⚠️ PowerShell command was cancelled.")
            Return
        End If
        If wasTimeout Then
            AppendResponse("⚠️ PowerShell command timed out after " &
            (CommandTimeoutMs / 1000) & " seconds.")
            SendCompletionNotification()
            Await DequeueNextTaskAsync()
            Return
        End If
        If wasStartError Then
            AppendResponse("⚠️ Failed to start PowerShell: " & errorMessage)
            SendCompletionNotification()
            Await DequeueNextTaskAsync()
            Return
        End If
        If wasGeneralError Then
            AppendResponse("⚠️ PowerShell execution error: " & errorMessage)
            SendCompletionNotification()
            Await DequeueNextTaskAsync()
            Return
        End If
        Try
            Dim resultLog As String = $"=== EXECUTION RESULT: {DateTime.Now} ==={Environment.NewLine}" &
            $"OUTPUT:{Environment.NewLine}{output}{Environment.NewLine}" &
            $"ERRORS:{Environment.NewLine}{errors}{Environment.NewLine}{separator}{Environment.NewLine}"
            File.AppendAllText(logPath, resultLog)
        Catch ex As Exception
            LogToFile("LOG_WRITE_ERROR", ex.Message)
        End Try
        Dim hasOutput = Not String.IsNullOrWhiteSpace(output)
        Dim hasErrors = Not String.IsNullOrWhiteSpace(errors)
        Dim hasSuccessSignal As Boolean = False
        If hasOutput Then
            Dim successKeywords As String() = {
            "Saved to", "Success", "Done", "Complete",
            "Created", "Opened", "Saved", "Write-Output"
        }
            hasSuccessSignal = successKeywords.Any(
            Function(kw) output.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
        End If
        Dim isFatalError = AnalyzePowerShellErrors(errors, hasSuccessSignal)
        If hasOutput AndAlso Not hasSuccessSignal Then
            Dim trimmedOutput = output.Trim()
            If trimmedOutput.Length > 500 Then
                trimmedOutput = trimmedOutput.Substring(0, 500) & "... [TRUNCATED]"
            End If
            AppendResponse("📋 Output: " & trimmedOutput)
        End If
        If isFatalError AndAlso correctionAttempts < MaxCorrectionAttempts Then
            If ct.IsCancellationRequested Then
                AppendResponse("⚠️ Correction cancelled.")
                correctionAttempts = 0
                Return
            End If
            Await HandlePowerShellCorrection(errors, output)
        ElseIf correctionAttempts >= MaxCorrectionAttempts Then
            correctionAttempts = 0
            AppendResponse("⚠️ Max correction attempts reached. Please check the result manually.")
            SendCompletionNotification()
            Await DequeueNextTaskAsync()
        Else
            correctionAttempts = 0
            SendCompletionNotification()
            Await DequeueNextTaskAsync()
        End If
    End Function
    Private Function AnalyzePowerShellErrors(errors As String, hasSuccessSignal As Boolean) As Boolean
        If String.IsNullOrWhiteSpace(errors) OrElse hasSuccessSignal Then Return False
        Dim fatalKeywords As String() = {
            "Exception", "TerminatingError", "not recognized",
            "cannot be loaded", "is not valid", "Access is denied",
            "UnauthorizedAccess", "The term", "Cannot find",
            "does not exist", "failed with", "At line:",
            "CategoryInfo", "FullyQualifiedErrorId", "is not a cmdlet",
            "cannot open", "No such file"
        }
        For Each line In errors.Split({Environment.NewLine, Chr(10).ToString()}, StringSplitOptions.RemoveEmptyEntries)
            Dim trimmedLine = line.Trim()
            If String.IsNullOrWhiteSpace(trimmedLine) Then Continue For
            If trimmedLine.StartsWith("WARNING:", StringComparison.OrdinalIgnoreCase) Then Continue For
            If trimmedLine.StartsWith("VERBOSE:", StringComparison.OrdinalIgnoreCase) Then Continue For
            If trimmedLine.StartsWith("DEBUG:", StringComparison.OrdinalIgnoreCase) Then Continue For
            If trimmedLine.StartsWith("INFORMATION:", StringComparison.OrdinalIgnoreCase) Then Continue For
            For Each keyword In fatalKeywords
                If trimmedLine.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return True
                End If
            Next
        Next
        Return False
    End Function
    Private Async Function HandlePowerShellCorrection(errors As String, output As String) As Task
        Dim feedbackMsg = "The PowerShell script produced errors:" & Environment.NewLine &
            "ERRORS:" & Environment.NewLine & errors.Trim() & Environment.NewLine &
            If(Not String.IsNullOrWhiteSpace(output),
               "OUTPUT:" & Environment.NewLine & output.Trim() & Environment.NewLine, "") &
            Environment.NewLine &
            "Provide a corrected PowerShell script wrapped in ||| delimiters."
        AddToHistory("user", feedbackMsg)
        Dim systemMessage = GetBootPrompt()
        Dim correctionResponse = Await GetChatResponseWithHistoryAsync(systemMessage)
        Dim correctionParsed = Await ParseResponseAsync(correctionResponse)
        AddToHistory("assistant", correctionParsed)
        TrimConversationHistory()
        Dim correctionMatch = Regex.Match(correctionParsed, "\|\|\|(.*?)\|\|\|", RegexOptions.Singleline)
        Dim isDone = correctionParsed.Trim().Equals("DONE", StringComparison.OrdinalIgnoreCase)
        If correctionMatch.Success AndAlso Not isDone Then
            correctionAttempts += 1
            gptResponse = correctionParsed
            commandExecutedPS = False
            AppendResponse($"🔄 AI self-correcting (attempt {correctionAttempts}/{MaxCorrectionAttempts})...")
        Else
            correctionAttempts = 0
            SendCompletionNotification()
            Await DequeueNextTaskAsync()
        End If
    End Function
    Private Sub SendCompletionNotification()
        If Not String.IsNullOrEmpty(pendingNotification) Then
            ShowNotification("Task Complete", pendingNotification)
            pendingNotification = ""
        End If
    End Sub
    Private Async Function ExecuteWebWorkflow(jsonCmd As String, fullContent As String) As Task
        AppendResponse("🌐 Running workflow...")
        Dim playwrightSuccess = Await RunAgentWorkflowAsync(jsonCmd)
        Dim autoWriteHandled As Boolean = False
        If Not String.IsNullOrWhiteSpace(pendingDocContent) Then
            Dim combinedContent = pendingDocContent
            If Not String.IsNullOrWhiteSpace(lastExtractedWebText) Then
                combinedContent = "=== LIVE WEB DATA ===" & Environment.NewLine &
                lastExtractedWebText & Environment.NewLine &
                "=== DOCUMENT CONTENT ===" & Environment.NewLine &
                pendingDocContent
                lastExtractedWebText = ""
            End If
            Await AutoWriteDocumentAsync(combinedContent, pendingTargetApp, pendingDocSavePath)
            pendingDocContent = ""
            pendingTargetApp = ""
            pendingDocSavePath = ""
            autoWriteHandled = True
        ElseIf Not String.IsNullOrWhiteSpace(lastExtractedWebText) Then
            Dim webTarget = If(String.IsNullOrWhiteSpace(pendingTargetApp), "word", pendingTargetApp)
            Await AutoWriteDocumentAsync(lastExtractedWebText, webTarget, pendingDocSavePath)
            lastExtractedWebText = ""
            pendingTargetApp = ""
            pendingDocSavePath = ""
            autoWriteHandled = True
        Else
            Dim allMatches = Regex.Matches(fullContent, "\|\|\|(.*?)\|\|\|", RegexOptions.Singleline)
            If allMatches.Count >= 2 Then
                Dim secondBlock = allMatches(1).Groups(1).Value.Trim()
                Dim looksLikePS = secondBlock.Contains("$") OrElse secondBlock.Contains("New-Object") OrElse
                secondBlock.Contains("Start-Process") OrElse secondBlock.Contains("Add-Type")
                If Not String.IsNullOrWhiteSpace(secondBlock) AndAlso
               Not secondBlock.StartsWith("{") AndAlso looksLikePS Then
                    Dim isDuplicate = autoWriteHandled AndAlso
                    (secondBlock.Contains("Excel.Application") OrElse
                     secondBlock.Contains("Word.Application") OrElse
                     secondBlock.Contains("PowerPoint.Application"))
                    If Not isDuplicate Then
                        Dim safeCheck = IsCommandSafe(secondBlock)
                        If safeCheck.Safe Then
                            Dim preview = secondBlock.Substring(0, Math.Min(60, secondBlock.Length)).Replace(Environment.NewLine, " ")
                            AppendResponse("📝 Running follow-up PS: " & preview & "...")
                            Await Task.Run(Sub() RunPowerShellSync(secondBlock))
                        Else
                            AppendResponse($"🛑 Follow-up PS blocked: {safeCheck.Reason}")
                        End If
                    Else
                        AppendResponse("⏭️ Skipping duplicate document creation script.")
                    End If
                End If
            End If
        End If
        SendCompletionNotification()
        Await DequeueNextTaskAsync()
    End Function
    Private Async Function DequeueNextTaskAsync() As Task
        Dim nextTask As String = Nothing
        If Not pendingTaskQueue.TryDequeue(nextTask) Then Return
        pendingDocContent = ""
        pendingTargetApp = ""
        pendingDocSavePath = ""
        lastExtractedWebText = ""
        pendingNotification = ""
        correctionAttempts = 0
        gptResponse = ""
        commandExecutedPS = False
        previousCommandPS = ""
        wantsChart = False
        wantsFormulas = False
        wantsConditionalFormat = False
        wantsModifyExisting = False
        wantsWordToC = False
        wantsWordTable = False
        wantsModifyExistingWord = False
        wantsWordPdf = False
        wantsPptNotes = False
        wantsPptChart = False
        wantsPptTable = False
        wantsPptTransitions = False
        wantsPptPdf = False
        wantsPptTheme = False
        wantsPptImages = False
        wantsModifyExistingPpt = False
        wantsOutlookSend = False
        wantsOutlookReply = False
        wantsOutlookForward = False
        wantsOutlookRead = False
        wantsOutlookCalendar = False
        wantsOutlookTask = False
        wantsOutlookContact = False
        wantsOutlookExport = False
        wantsOneNoteFormat = False
        wantsOneNoteChecklist = False
        wantsOneNoteTable = False
        wantsOneNoteAppend = False
        wantsOneNoteNewSection = False
        wantsOneNoteSearch = False
        wantsOneNoteExport = False
        wantsAccessCreate = False
        wantsAccessQuery = False
        wantsAccessImport = False
        wantsAccessExport = False
        wantsAccessReport = False
        wantsAccessModify = False
        wantsLocalUserCreate = False
        wantsLocalUserModify = False
        wantsLocalUserDelete = False
        wantsLocalUserList = False
        wantsLocalGroupManage = False
        wantsADUserCreate = False
        wantsADUserModify = False
        wantsADUserDelete = False
        wantsADUserSearch = False
        wantsADGroupManage = False
        wantsADComputerManage = False
        wantsADOUManage = False
        wantsADPasswordReset = False
        wantsADBulkOperation = False
        selectedFiles.Clear()
        selectedFolder = ""
        mode = "single"
        Interlocked.Exchange(isRunning, 0)
        Await Task.Delay(2000)
        Dim remaining = pendingTaskQueue.Count
        AppendResponse($"▶ Starting next task ({remaining} remaining after this):" &
               Environment.NewLine & $"   {nextTask}")
        Try
            Me.Invoke(Sub()
                          Try
                              QueryTextBox.Text = nextTask
                              SendButton_Click(SendButton, EventArgs.Empty)
                          Catch ex As Exception
                              LogToFile("DEQUEUE_INVOKE_ERROR", ex.ToString())
                              AppendResponse("❌ Failed to start next task: " & ex.Message)
                          End Try
                      End Sub)
        Catch ex As Exception
            LogToFile("DEQUEUE_ERROR", ex.ToString())
            AppendResponse("❌ Task queue error: " & ex.Message)
        End Try
    End Function
    Private Sub AddToHistory(role As String, text As String)
        conversationHistory.Enqueue(New ConversationEntry(role, text))
    End Sub
    Private Sub TrimConversationHistory()
        While conversationHistory.Count > MaxConversationTurns
            Dim discard As ConversationEntry = Nothing
            conversationHistory.TryDequeue(discard)
        End While
        Dim totalChars = conversationHistory.Sum(Function(e) If(e.Content, "").Length)
        While totalChars > 100000 AndAlso conversationHistory.Count > 2
            Dim removed As ConversationEntry = Nothing
            If conversationHistory.TryDequeue(removed) Then
                totalChars -= If(removed.Content, "").Length
            Else
                Exit While
            End If
        End While
    End Sub
    Private Function DetectTargetApp(lowerInput As String) As String
        If lowerInput.Contains("notepad") Then Return "notepad"
        If lowerInput.Contains("onenote") OrElse lowerInput.Contains("one note") Then Return "onenote"
        If lowerInput.Contains("powerpoint") OrElse lowerInput.Contains("ppt") OrElse
           lowerInput.Contains("presentation") OrElse lowerInput.Contains("slides") Then Return "powerpoint"
        If lowerInput.Contains("excel") OrElse lowerInput.Contains("spreadsheet") OrElse
           lowerInput.Contains("xlsx") OrElse lowerInput.Contains("workbook") Then Return "excel"
        If lowerInput.Contains("outlook") OrElse lowerInput.Contains("email") OrElse
           lowerInput.Contains("mail") OrElse lowerInput.Contains("calendar") OrElse
           lowerInput.Contains("appointment") OrElse lowerInput.Contains("meeting request") OrElse
           lowerInput.Contains("send meeting") Then Return "outlook"
        If lowerInput.Contains("word") OrElse lowerInput.Contains("docx") Then Return "word"
        Return ""
    End Function
    Private Function DetectOutlookAction(lowerInput As String, targetApp As String) As String
        If targetApp <> "outlook" Then Return ""
        If lowerInput.Contains("send") OrElse lowerInput.Contains("email") OrElse lowerInput.Contains("mail") Then Return "email"
        If lowerInput.Contains("meeting") OrElse lowerInput.Contains("invite") Then Return "meeting"
        If lowerInput.Contains("appointment") OrElse lowerInput.Contains("calendar") OrElse lowerInput.Contains("schedule") Then Return "appointment"
        If lowerInput.Contains("task") OrElse lowerInput.Contains("todo") OrElse lowerInput.Contains("to-do") Then Return "task"
        If lowerInput.Contains("contact") Then Return "contact"
        If lowerInput.Contains("reply") Then Return "reply"
        If lowerInput.Contains("forward") Then Return "forward"
        Return ""
    End Function
    Private Function DetectOneNoteAction(lowerInput As String, targetApp As String) As String
        If targetApp <> "onenote" Then Return ""
        If lowerInput.Contains("quick note") OrElse lowerInput.Contains("side note") Then Return "quicknote"
        If lowerInput.Contains("checklist") OrElse lowerInput.Contains("to-do") OrElse
           lowerInput.Contains("todo") OrElse lowerInput.Contains("checkbox") Then Return "checklist"
        If lowerInput.Contains("meeting") Then Return "meeting"
        If lowerInput.Contains("section") Then Return "section"
        Return "note"
    End Function
    Private Function ExtractLikelyFilenames(input As String) As List(Of String)
        Dim results As New List(Of String)
        Dim withExt = Regex.Matches(input,
        "(?<!\w)([A-Za-z0-9_\-]+\.(?:docx|xlsx|pptx|pdf|txt|csv|png|jpg|jpeg|bmp))(?!\w)",
        RegexOptions.IgnoreCase)
        For Each m As Match In withExt
            If Not results.Contains(m.Groups(1).Value) Then
                results.Add(m.Groups(1).Value)
            End If
        Next
        Dim underscored = Regex.Matches(input,
        "(?<!\w)([A-Za-z][A-Za-z0-9]*(?:_[A-Za-z0-9]+){2,})(?!\w)")
        For Each m As Match In underscored
            If m.Value.Length > 10 AndAlso Not results.Contains(m.Value) Then
                results.Add(m.Value)
            End If
        Next
        Dim readPattern = Regex.Match(input,
        "(?:read|open|find|get)\s+(.+?)\s+(?:in|from|and|then)",
        RegexOptions.IgnoreCase)
        If readPattern.Success Then
            Dim candidate = readPattern.Groups(1).Value.Trim()
            candidate = Regex.Replace(candidate, "^(?:the|a|an)\s+", "",
            RegexOptions.IgnoreCase).Trim()
            If candidate.Length > 4 AndAlso Not results.Contains(candidate) Then
                results.Add(candidate)
            End If
        End If
        Return results
    End Function
    Private Async Function DetectAndReadFiles(textForSearch As String, lowerInput As String,
                                           wantsToOpen As Boolean) As Task
        Dim allowedExtensions = {
        ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".tif",
        ".txt", ".csv", ".xlsx", ".pptx", ".docx", ".pdf"
    }
        Dim isAllowedFile = Function(fp As String) allowedExtensions.Contains(Path.GetExtension(fp).ToLower())
        Dim userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        Dim downloads = Path.Combine(userProfile, "Downloads")
        Dim documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim possibleFolders As New List(Of String)
        Dim subfolderPattern = "(downloads?|documents?|desktop)[\\\/\s]+([A-Za-z0-9_\-]+(?:[\\\/][A-Za-z0-9_\-]+)*)"
        For Each m As Match In Regex.Matches(textForSearch, subfolderPattern, RegexOptions.IgnoreCase)
            Dim folderName = m.Groups(1).Value.Trim().ToLower()
            Dim subPath = m.Groups(2).Value.Trim().Replace("/", "\")
            Dim baseFolderPath = ""
            If folderName.StartsWith("download") Then baseFolderPath = downloads
            If folderName.StartsWith("document") Then baseFolderPath = documents
            If folderName = "desktop" Then baseFolderPath = desktop
            If Not String.IsNullOrEmpty(baseFolderPath) Then
                Dim fullPath = Path.Combine(baseFolderPath, subPath)
                If Directory.Exists(fullPath) Then
                    If Not possibleFolders.Contains(fullPath) Then
                        possibleFolders.Add(fullPath)
                        AppendResponse($"📁 Found subfolder: {fullPath}")
                    End If
                Else
                    Try
                        Dim searchName = subPath.Split("\"c)(0)
                        Dim matchingFolders = Directory.GetDirectories(baseFolderPath, "*" & searchName & "*", SearchOption.TopDirectoryOnly)
                        For Each folder In matchingFolders
                            If Not possibleFolders.Contains(folder) Then
                                possibleFolders.Add(folder)
                                AppendResponse($"📁 Found matching folder: {folder}")
                            End If
                        Next
                    Catch
                    End Try
                End If
            End If
        Next
        If possibleFolders.Count = 0 Then
            If lowerInput.Contains("downloads") OrElse lowerInput.Contains("download") Then possibleFolders.Add(downloads)
            If lowerInput.Contains("documents") OrElse lowerInput.Contains("document") Then possibleFolders.Add(documents)
            If lowerInput.Contains("desktop") Then possibleFolders.Add(desktop)
        End If
        Dim explicitPathPattern = "([a-zA-Z]:\\[^\""<>|\r\n,\s]+)"
        For Each m As Match In Regex.Matches(textForSearch, explicitPathPattern, RegexOptions.IgnoreCase)
            Dim filePath = m.Value.Trim().TrimEnd("."c, ","c, ";"c, ":"c)
            If Directory.Exists(filePath) Then
                If Not possibleFolders.Contains(filePath) Then possibleFolders.Add(filePath)
            ElseIf File.Exists(filePath) AndAlso isAllowedFile(filePath) Then
                If Not selectedFiles.Contains(filePath) Then
                    selectedFiles.Add(filePath)
                    AppendResponse($"📄 Found explicit file: {filePath}")
                End If
            End If
        Next
        Dim quotedPattern = """([^""]+\.(?:docx|xlsx|pptx|pdf|txt|csv|png|jpg|jpeg|bmp))"""
        For Each m As Match In Regex.Matches(textForSearch, quotedPattern, RegexOptions.IgnoreCase)
            Dim fileName = m.Groups(1).Value.Trim()
            For Each folder In If(possibleFolders.Count > 0, possibleFolders, New List(Of String) From {downloads, documents, desktop})
                Dim fullPath = Path.Combine(folder, fileName)
                If File.Exists(fullPath) AndAlso Not selectedFiles.Contains(fullPath) Then
                    selectedFiles.Add(fullPath)
                    AppendResponse($"📄 Found: {fileName}")
                End If
            Next
        Next
        Dim fileNamePattern = "(?<!\w)([A-Za-z0-9_\-]+\.(?:docx|xlsx|pptx|pdf|txt|csv|png|jpg|jpeg|bmp|gif|tiff?))(?!\w)"
        For Each m As Match In Regex.Matches(textForSearch, fileNamePattern, RegexOptions.IgnoreCase)
            Dim fileName = m.Groups(1).Value.Trim()
            For Each folder In If(possibleFolders.Count > 0, possibleFolders, New List(Of String) From {downloads, documents, desktop})
                Try
                    Dim found = Directory.GetFiles(folder, "*" & Path.GetFileNameWithoutExtension(fileName) & "*" & Path.GetExtension(fileName),
                                               SearchOption.TopDirectoryOnly).Where(Function(f) isAllowedFile(f)).ToArray()
                    For Each f In found
                        If Not selectedFiles.Contains(f) Then
                            selectedFiles.Add(f)
                            AppendResponse($"📄 Found: {Path.GetFileName(f)}")
                        End If
                    Next
                Catch
                End Try
            Next
        Next
        If selectedFiles.Count = 0 AndAlso possibleFolders.Count > 0 Then
            Dim skipWords = {
            "read", "explain", "downloads", "download", "documents", "document",
            "desktop", "summarize", "analyze", "analysis", "write", "notepad", "word",
            "folder", "folders", "files", "file", "search", "find",
            "check", "inside", "this", "that", "these", "those",
            "and", "the", "from", "with", "about", "into", "essay",
            "open", "create", "put", "make", "generate", "all",
            "powerpoint", "pptx", "presentation", "slides", "slide",
            "excel", "xlsx", "spreadsheet", "workbook", "sheet",
            "outlook", "email", "mail", "send", "calendar", "appointment",
            "meeting", "task", "contact", "reply", "forward", "schedule",
            "onenote", "note", "notes", "quick", "section", "notebook",
            "checklist", "checkbox", "todo", "summary", "test",
            "comprehensive", "detailed", "complete", "full", "entire",
            "using", "based", "then", "also", "each", "every", "its"
        }
            Dim namePattern = "([A-Za-z0-9_\-]{6,})"
            Dim possibleNames As New List(Of String)
            For Each m As Match In Regex.Matches(textForSearch, namePattern)
                Dim cleaned = m.Value.Trim()
                If Not skipWords.Contains(cleaned.ToLower()) Then
                    Dim isFolderName = possibleFolders.Any(Function(f) f.ToLower().EndsWith("\" & cleaned.ToLower()))
                    If Not isFolderName AndAlso Not possibleNames.Contains(cleaned) Then
                        possibleNames.Add(cleaned)
                    End If
                End If
            Next
            If possibleNames.Count > 0 Then
                AppendResponse($"🔍 Looking for: {String.Join(", ", possibleNames)}")
            End If
            For Each folder In possibleFolders
                If Not Directory.Exists(folder) Then Continue For
                For Each searchName In possibleNames
                    Try
                        Dim folderCopy = folder
                        Dim nameCopy = searchName
                        Dim found = Await Task.Run(Function()
                                                       Try
                                                           Return Directory.GetFiles(folderCopy, "*" & nameCopy & "*.*", SearchOption.TopDirectoryOnly).
                                                           Where(Function(f) isAllowedFile(f)).ToArray()
                                                       Catch
                                                           Return New String() {}
                                                       End Try
                                                   End Function)
                        For Each f In found
                            If Not selectedFiles.Contains(f) Then
                                selectedFiles.Add(f)
                                AppendResponse($"📄 Found: {Path.GetFileName(f)}")
                            End If
                        Next
                    Catch
                    End Try
                Next
            Next
        End If
        If selectedFiles.Count = 0 AndAlso possibleFolders.Count > 0 Then
            For Each folder In possibleFolders
                Try
                    Dim folderCopy = folder
                    Dim allFiles = Await Task.Run(Function() Directory.GetFiles(folderCopy, "*.*", SearchOption.TopDirectoryOnly).
                                                Where(Function(f) isAllowedFile(f)).ToArray())
                    For Each f In allFiles
                        If Not selectedFiles.Contains(f) Then
                            selectedFiles.Add(f)
                            AppendResponse($"📄 Found: {Path.GetFileName(f)}")
                        End If
                    Next
                Catch
                End Try
            Next
        End If
        If selectedFiles.Count > 1 Then mode = "multi"
    End Function
    Private Async Function ReadSelectedFilesContent(wantsToOpen As Boolean) As Task(Of String)
        If wantsToOpen AndAlso selectedFiles.Count >= 1 Then
            AppendResponse($"📂 Will open file: {Path.GetFileName(selectedFiles(0))}")
            Return ""
        End If
        If selectedFiles.Count = 1 Then
            AppendResponse($"📖 Reading: {Path.GetFileName(selectedFiles(0))}")
            Try
                Dim analysis = Await FileAnalysisHelper.AnalyzeFileAsync(selectedFiles(0))
                If analysis.Text.StartsWith("(OCR not available") OrElse
                   analysis.Text.StartsWith("(Unsupported file type") Then
                    AppendResponse($"⚠️ {analysis.Text}")
                    Return ""
                End If
                Return analysis.Text
            Catch ex As Exception
                Return $"(Error reading file: {ex.Message})"
            End Try
        End If
        If selectedFiles.Count > 1 Then
            AppendResponse($"📖 Reading {selectedFiles.Count} files...")
            Dim sb As New StringBuilder()
            For Each f In selectedFiles
                Try
                    Dim analysis = Await FileAnalysisHelper.AnalyzeFileAsync(f)
                    If Not analysis.Text.StartsWith("(OCR not available") Then
                        sb.AppendLine("=== FILE: " & Path.GetFileName(f) & " ===")
                        sb.AppendLine(analysis.Text)
                        sb.AppendLine()
                    End If
                Catch ex As Exception
                    sb.AppendLine("=== FILE: " & Path.GetFileName(f) & " ===")
                    sb.AppendLine($"(Error reading: {ex.Message})")
                    sb.AppendLine()
                End Try
            Next
            Return sb.ToString()
        End If
        Return ""
    End Function
    Private Function SmartTruncate(text As String, maxLength As Integer) As String
        If text.Length <= maxLength Then Return text
        Dim keepStart = CInt(maxLength * 0.7)
        Dim keepEnd = CInt(maxLength * 0.25)
        Dim marker = Environment.NewLine & Environment.NewLine &
            $"... [TRUNCATED: {text.Length - keepStart - keepEnd} characters omitted from middle] ..." &
            Environment.NewLine & Environment.NewLine
        Return text.Substring(0, keepStart) & marker & text.Substring(text.Length - keepEnd)
    End Function
    Private Function BuildInstructionHint(wantsToOpen As Boolean, wantsToWrite As Boolean,
                                           targetApp As String, outlookAction As String,
                                           onenoteAction As String, lowerInput As String,
                                           documentContent As String) As String
        If wantsToOpen AndAlso selectedFiles.Count >= 1 Then
            Dim filePath = selectedFiles(0)
            Return Environment.NewLine &
                "--- INSTRUCTION ---" & Environment.NewLine &
                $"User wants to OPEN this file: {filePath}" & Environment.NewLine &
                "You MUST respond with PowerShell code wrapped in ||| delimiters." & Environment.NewLine &
                "CORRECT FORMAT:" & Environment.NewLine &
                "Opening file..." & Environment.NewLine &
                "|||" & Environment.NewLine &
                $"Start-Process ""{filePath}""" & Environment.NewLine &
                "|||" & Environment.NewLine &
                "DO NOT read or analyze the file. Just open it." & Environment.NewLine &
                "--- END INSTRUCTION ---"
        End If
        If documentContent.Length > 0 AndAlso wantsToWrite Then
            pendingDocContent = documentContent
            pendingTargetApp = If(String.IsNullOrEmpty(targetApp), "word", targetApp)
            If lowerInput.Contains("downloads") Then pendingDocSavePath = "Downloads"
            If lowerInput.Contains("desktop") Then pendingDocSavePath = "Desktop"
            Dim appHint = GetAppHint(targetApp)
            Return Environment.NewLine &
                "--- INSTRUCTION ---" & Environment.NewLine &
                "Document content has been extracted above. User wants you to WRITE output to an application." & Environment.NewLine &
                $"Target application: {If(String.IsNullOrEmpty(targetApp), "Word (default)", targetApp.ToUpper)}" & Environment.NewLine &
                appHint & Environment.NewLine &
                "You MUST generate PowerShell code wrapped in ||| delimiters to open the target app and write your analysis/summary/essay there." & Environment.NewLine &
                "DO NOT just describe the content - you must OPEN THE APP AND WRITE TO IT!" & Environment.NewLine &
                "--- END INSTRUCTION ---"
        End If
        If targetApp = "outlook" Then
            Return BuildOutlookHint(outlookAction)
        End If
        If targetApp = "onenote" Then
            Return BuildOneNoteHint(onenoteAction)
        End If
        If targetApp = "powerpoint" AndAlso wantsToWrite Then
            Return Environment.NewLine &
                "--- INSTRUCTION ---" & Environment.NewLine &
                "User wants to create a PowerPoint presentation." & Environment.NewLine &
                "You MUST generate a PowerShell script wrapped in ||| that uses the PowerPoint.Application COM object to create the slides, add titles, and add body text." & Environment.NewLine &
                "Do NOT use SendKeys or UIAutomation for PowerPoint." & Environment.NewLine &
                "--- END INSTRUCTION ---"
        End If
        If targetApp = "excel" AndAlso wantsToWrite Then
            Return BuildExcelHint(lowerInput)
        End If
        If targetApp = "word" AndAlso wantsToWrite Then
            Return Environment.NewLine &
                "--- INSTRUCTION ---" & Environment.NewLine &
                "User wants to create a Word document." & Environment.NewLine &
                "Open Word using: Start-Process ""winword.exe"" ""/w""" & Environment.NewLine &
                "Use clipboard method for longer text content." & Environment.NewLine &
                "NEVER use ^n (creates second window), NEVER use ^{ENTER} (page break)." & Environment.NewLine &
                "You MUST generate PowerShell code wrapped in ||| delimiters." & Environment.NewLine &
                "--- END INSTRUCTION ---"
        End If
        Return ""
    End Function
    Private Function GetAppHint(targetApp As String) As String
        Select Case targetApp
            Case "word" : Return "Open Microsoft Word using: Start-Process ""winword.exe"" ""/w"" and use clipboard method to write."
            Case "notepad" : Return "Open Notepad using: Start-Process ""notepad.exe"" and use clipboard method to write."
            Case "powerpoint" : Return "Create a PowerPoint presentation using the PowerPoint.Application COM object."
            Case "excel" : Return "Open Excel using: Start-Process ""excel.exe"" then AppActivate(""Excel"") then send {ENTER} to select Blank Workbook."
            Case "onenote" : Return "Open OneNote using: Start-Process ""onenote.exe"", create new page with Ctrl+N."
            Case Else : Return "Open the appropriate application and write the content there."
        End Select
    End Function
    Private Function BuildOutlookHint(outlookAction As String) As String
        Dim hint = ""
        Select Case outlookAction
            Case "email" : hint = "Create and send a new email using: Start-Process ""outlook.exe"" ""/c ipm.note"""
            Case "meeting" : hint = "Create a meeting request using: Start-Process ""outlook.exe"" ""/c ipm.appointment"" then add attendees"
            Case "appointment" : hint = "Create a calendar appointment using: Start-Process ""outlook.exe"" ""/c ipm.appointment"""
            Case "task" : hint = "Create a task using: Start-Process ""outlook.exe"" ""/c ipm.task"""
            Case "contact" : hint = "Create a contact using: Start-Process ""outlook.exe"" ""/c ipm.contact"""
            Case "reply" : hint = "Reply to the selected email using Ctrl+R"
            Case "forward" : hint = "Forward the selected email using Ctrl+F"
            Case Else : hint = "Perform the requested Outlook action"
        End Select
        Return Environment.NewLine &
            "--- INSTRUCTION ---" & Environment.NewLine &
            $"User wants to perform an Outlook action: {If(String.IsNullOrEmpty(outlookAction), "email", outlookAction)}" & Environment.NewLine &
            hint & Environment.NewLine &
            "You MUST generate PowerShell code wrapped in ||| delimiters." & Environment.NewLine &
            "After typing the To address, press {TAB}{TAB} to skip Cc and reach Subject. Then {TAB} once to reach Body. Send with ^{ENTER}." & Environment.NewLine &
            "--- END INSTRUCTION ---"
    End Function
    Private Function BuildOneNoteHint(onenoteAction As String) As String
        Dim hint = ""
        Select Case onenoteAction
            Case "quicknote" : hint = "Create a quick note using: Start-Process ""onenote.exe"" ""/sidenote"""
            Case "checklist" : hint = "Create a checklist in OneNote. Open with: Start-Process ""onenote.exe"", create new page with Ctrl+N, use Ctrl+1 for checkboxes."
            Case "meeting" : hint = "Create meeting notes in OneNote. Open with: Start-Process ""onenote.exe"", create new page with Ctrl+N."
            Case "section" : hint = "Create a new section in OneNote using Ctrl+T after opening OneNote."
            Case Else : hint = "Create a note in OneNote. Open with: Start-Process ""onenote.exe"", create new page with Ctrl+N."
        End Select
        Return Environment.NewLine &
            "--- INSTRUCTION ---" & Environment.NewLine &
            $"User wants to create content in OneNote: {If(String.IsNullOrEmpty(onenoteAction), "note", onenoteAction)}" & Environment.NewLine &
            hint & Environment.NewLine &
            "You MUST generate PowerShell code wrapped in ||| delimiters." & Environment.NewLine &
            "Use clipboard method for longer content. Use Ctrl+1 for checkboxes." & Environment.NewLine &
            "--- END INSTRUCTION ---"
    End Function
    Private Function BuildExcelHint(lowerInput As String) As String
        Dim needsWebData = lowerInput.Contains("stock") OrElse lowerInput.Contains("price") OrElse
            lowerInput.Contains("history") OrElse lowerInput.Contains("financial") OrElse
            lowerInput.Contains("market") OrElse lowerInput.Contains("analysis")
        If needsWebData Then
            pendingTargetApp = "excel"
            If lowerInput.Contains("downloads") Then pendingDocSavePath = "Downloads"
            If lowerInput.Contains("desktop") Then pendingDocSavePath = "Desktop"
            Return Environment.NewLine &
                "--- INSTRUCTION ---" & Environment.NewLine &
                "User wants LIVE DATA from the web written into an Excel spreadsheet." & Environment.NewLine &
                "You MUST respond with a Playwright JSON workflow wrapped in ||| delimiters." & Environment.NewLine &
                "--- END INSTRUCTION ---"
        Else
            Return Environment.NewLine &
                "--- INSTRUCTION ---" & Environment.NewLine &
                "User wants to create an Excel spreadsheet." & Environment.NewLine &
                "Open Excel using: Start-Process ""excel.exe"" then AppActivate(""Excel"") then send {ENTER}." & Environment.NewLine &
                "You MUST generate PowerShell code wrapped in ||| delimiters." & Environment.NewLine &
                "--- END INSTRUCTION ---"
        End If
    End Function
    Private Sub HandleNotifications(parsed As String, lowerInput As String)
        If parsed.Contains("QUESTION:") Then
            Dim qIndex = parsed.IndexOf("QUESTION:")
            Dim questionText = parsed.Substring(qIndex + 9).Trim()
            Dim nlIdx = questionText.IndexOf(vbLf)
            If nlIdx > 0 Then questionText = questionText.Substring(0, nlIdx).Trim()
            If questionText.Length > 100 Then questionText = questionText.Substring(0, 100) & "..."
            ShowNotification("AI needs your input", questionText)
        End If
        If parsed.Contains("NOTIFY:") Then
            Dim nIndex = parsed.IndexOf("NOTIFY:")
            Dim notifyText = parsed.Substring(nIndex + 7).Trim()
            Dim nlIdx = notifyText.IndexOf(vbLf)
            If nlIdx > 0 Then notifyText = notifyText.Substring(0, nlIdx).Trim()
            If notifyText.Length > 150 Then notifyText = notifyText.Substring(0, 150) & "..."
            pendingNotification = notifyText
        ElseIf lowerInput.Contains("notify") OrElse lowerInput.Contains("let me know") OrElse
               lowerInput.Contains("tell me when") Then
            pendingNotification = "Your requested task has been executed."
        End If
    End Sub
    Private Async Function HandleSceneRecording(userInput As String) As Task
        Dim systemMessage = GetBootPrompt()
        Dim osPath = Path.Combine(Application.StartupPath, "OS")
        If File.Exists(osPath) Then
            systemMessage &= Environment.NewLine & Environment.NewLine &
        "====================================================" & Environment.NewLine &
        "COGNITIVE OPERATING SYSTEM" & Environment.NewLine &
        "====================================================" & Environment.NewLine &
        File.ReadAllText(osPath)
        End If
        Dim cleanedScene = ScenePreprocessor.CleanScene(userInput)
        Dim scenesFolder = Path.Combine(Application.StartupPath, "Scenes")
        Dim allFrames As New List(Of String)
        If Directory.Exists(scenesFolder) Then
            Dim frameDirs = Directory.GetDirectories(scenesFolder, "*_frames").
                OrderByDescending(Function(d) Directory.GetCreationTime(d)).ToList()
            For Each framesDir In frameDirs
                Dim frames = Directory.GetFiles(framesDir, "frame_*.jpg").OrderBy(Function(f) f).ToList()
                If frames.Count > 0 Then
                    allFrames = frames
                    Exit For
                End If
            Next
        End If
        Dim sceneLines = cleanedScene.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
        Dim contentBlocks As New List(Of Object)
        Dim sceneIntro = "=== MACRO TRANSLATION PROTOCOL ===" & Environment.NewLine &
            "Translate this recorded scene into automation." & Environment.NewLine &
            "Each action is paired with a frame from a screen recording." & Environment.NewLine &
            "CRITICAL RULES:" & Environment.NewLine &
            "- If ANY clicks are on msedge/chrome/brave, use Playwright JSON wrapped in |||, NOT PowerShell UIA." & Environment.NewLine &
            "- NEVER use UIAutomationClient.CUIAutomation or System.Windows.Automation in PowerShell." & Environment.NewLine &
            "- Use each frame image to understand exactly what was on screen during that action." & Environment.NewLine &
            "- Reconstruct INTENT from visual context, not just raw coordinates." & Environment.NewLine &
            "- [TYPE] lines immediately before a [CLICK:TASKBAR_SEARCH] are the search query." & Environment.NewLine &
            "=== SCENE LOG WITH FRAMES ==="
        contentBlocks.Add(New With {.type = "text", .text = sceneIntro})
        Dim stepNumber = 0
        Dim frameCount = 0
        Dim maxFrames = 10
        Dim actionLines = sceneLines.Where(Function(l)
                                               Dim t = l.Trim()
                                               Return t.StartsWith("[CLICK") OrElse t.StartsWith("[TYPE") OrElse t.StartsWith("[KEYPRESS")
                                           End Function).Count()
        Dim frameStep = If(allFrames.Count > 0 AndAlso actionLines > 0,
                           Math.Max(1, allFrames.Count \ Math.Min(actionLines, maxFrames)), 1)
        For Each line In sceneLines
            Dim t = line.Trim()
            If String.IsNullOrWhiteSpace(t) Then Continue For
            If t.StartsWith("[CLICK") OrElse t.StartsWith("[TYPE") OrElse t.StartsWith("[KEYPRESS") Then
                stepNumber += 1
                contentBlocks.Add(New With {.type = "text", .text = $"Step {stepNumber}: {t}"})
                If allFrames.Count > 0 AndAlso frameCount < maxFrames Then
                    Dim frameIndex = Math.Min((stepNumber - 1) * frameStep, allFrames.Count - 1)
                    Dim framePath = allFrames(frameIndex)
                    If File.Exists(framePath) Then
                        Try
                            Dim imgBytes = File.ReadAllBytes(framePath)
                            Dim b64 = Convert.ToBase64String(imgBytes)
                            contentBlocks.Add(New With {.type = "text", .text = $"[Screen at Step {stepNumber}:]"})
                            contentBlocks.Add(New With {
                                .type = "image",
                                .source = New With {.type = "base64", .media_type = "image/jpeg", .data = b64}
                            })
                            frameCount += 1
                        Catch
                        End Try
                    End If
                End If
            Else
                contentBlocks.Add(New With {.type = "text", .text = t})
            End If
        Next
        Dim textOnly = String.Join(Environment.NewLine,
            contentBlocks.Where(Function(b) b.GetType().GetProperty("type")?.GetValue(b)?.ToString() = "text").
            Select(Function(b) b.GetType().GetProperty("text")?.GetValue(b)?.ToString()))
        AddToHistory("user", textOnly)
        Dim rawResponse = Await GetChatResponseWithHistoryAsync(systemMessage)
        Dim parsed = Await ParseResponseAsync(rawResponse)
        AddToHistory("assistant", parsed)
        If Not parsed.Contains("|||") Then
            AppendResponse("AI: " & parsed)
        Else
            Dim confirmText = parsed.Split({"|||"}, StringSplitOptions.None)(0).Trim()
            If String.IsNullOrWhiteSpace(confirmText) Then confirmText = "Replaying recorded scene..."
            AppendResponse("AI: " & confirmText)
        End If
        gptResponse = parsed
        commandExecutedPS = False
    End Function
    Private Async Function ParseResponseAsync(raw As String) As Task(Of String)
        Await Task.Yield()
        If String.IsNullOrWhiteSpace(raw) Then Return ""
        Try
            Dim json = Newtonsoft.Json.Linq.JObject.Parse(raw)
            Dim ollamaMsg = json("message")
            If ollamaMsg IsNot Nothing Then
                Dim ollamaContent = ollamaMsg("content")
                If ollamaContent IsNot Nothing AndAlso
               Not String.IsNullOrWhiteSpace(ollamaContent.ToString()) Then
                    Return CleanAIResponse(ollamaContent.ToString())
                End If
            End If
            Dim choices = json("choices")
            If choices IsNot Nothing Then
                Dim firstChoice = choices(0)
                If firstChoice IsNot Nothing Then
                    Dim msgToken = firstChoice("message")
                    If msgToken IsNot Nothing Then
                        Dim oaiContent = msgToken("content")
                        If oaiContent IsNot Nothing AndAlso
                       Not String.IsNullOrWhiteSpace(oaiContent.ToString()) Then
                            Return CleanAIResponse(oaiContent.ToString())
                        End If
                    End If
                End If
            End If
            Dim anthropicContent = json("content")
            If anthropicContent IsNot Nothing AndAlso anthropicContent.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                For Each block In anthropicContent
                    If block("type")?.ToString() = "text" Then
                        Dim txt = block("text")?.ToString()
                        If Not String.IsNullOrWhiteSpace(txt) Then
                            Return CleanAIResponse(txt)
                        End If
                    End If
                Next
            End If
            Dim candidates = json("candidates")
            If candidates IsNot Nothing Then
                Dim firstCandidate = candidates(0)
                If firstCandidate IsNot Nothing Then
                    Dim candidateContent = firstCandidate("content")
                    If candidateContent IsNot Nothing Then
                        Dim parts = candidateContent("parts")
                        If parts IsNot Nothing Then
                            Dim firstPart = parts(0)
                            If firstPart IsNot Nothing Then
                                Dim geminiText = firstPart("text")?.ToString()
                                If Not String.IsNullOrWhiteSpace(geminiText) Then
                                    Return CleanAIResponse(geminiText)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            Dim errorToken = json("error")
            If errorToken IsNot Nothing Then
                Dim errMsg = errorToken("message")?.ToString()
                If Not String.IsNullOrWhiteSpace(errMsg) Then
                    Return "API Error: " & errMsg
                End If
            End If
        Catch
            Return CleanAIResponse(raw.Trim())
        End Try
        Return CleanAIResponse(raw.Trim())
    End Function
    Private Function CleanAIResponse(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return text
        text = StripThinkingTags(text)
        Return text.Replace("```powershell", "").Replace("```PowerShell", "").
            Replace("```shell", "").Replace("```json", "").Replace("```", "").Trim()
    End Function
    Private Function StripThinkingTags(input As String) As String
        If String.IsNullOrWhiteSpace(input) Then Return input
        Dim result = Regex.Replace(input, "<think>.*?</think>", "", RegexOptions.Singleline Or RegexOptions.IgnoreCase)
        result = Regex.Replace(result, "</?think>", "", RegexOptions.IgnoreCase)
        Return result.Trim()
    End Function
    Private Async Function GetChatResponseWithHistoryAsync(systemMsg As String) As Task(Of String)
        Try
            Dim messages As New List(Of Object)
            messages.Add(New With {.role = "system", .content = systemMsg})
            For Each entry In conversationHistory.ToArray()
                If Not String.IsNullOrWhiteSpace(entry.Content) Then
                    messages.Add(New With {.role = entry.Role, .content = entry.Content})
                End If
            Next
            Return Await SendToProviderAsync(messages.ToArray(), Nothing, Nothing)
        Catch ex As Exception
            LogToFile("AI_ERROR", ex.ToString())
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Async Function GetChatResponseWithVisionAsync(systemMsg As String,
                                                       userMsg As String,
                                                       base64Image As String) As Task(Of String)
        Try
            Dim messages As New List(Of Object)
            messages.Add(New With {.role = "system", .content = systemMsg})
            messages.Add(New With {
            .role = "user",
            .content = userMsg,
            .images = If(Not String.IsNullOrEmpty(base64Image),
                         New String() {base64Image},
                         New String() {})
        })
            Return Await SendToProviderAsync(messages.ToArray(), Nothing, Nothing)
        Catch ex As Exception
            LogToFile("VISION_ERROR", ex.ToString())
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Async Function GetChatResponseAsync(systemMsg As String,
                                             userMsg As String) As Task(Of String)
        Try
            Dim messages As New List(Of Object)
            messages.Add(New With {.role = "system", .content = systemMsg})
            messages.Add(New With {.role = "user", .content = userMsg})
            Return Await SendToProviderAsync(messages.ToArray(), Nothing, Nothing)
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Async Function GetSimpleAIResponseAsync(prompt As String,
                                                 Optional ct As CancellationToken = Nothing) As Task(Of String)
        Try
            Dim messages As New List(Of Object)
            messages.Add(New With {.role = "user", .content = prompt})
            Dim raw = Await SendToProviderAsync(messages.ToArray(), Nothing, ct)
            Return StripThinkingTags(raw)
        Catch ex As OperationCanceledException
            LogToFile("SIMPLE_AI_CANCELLED", "Request was cancelled")
            Return ""
        Catch ex As Exception
            LogToFile("SIMPLE_AI_ERROR", ex.ToString())
            Return ""
        End Try
    End Function
    Private Async Function SendToProviderAsync(messages As Object(),
                                            Optional systemOverride As String = Nothing,
                                            Optional ct As CancellationToken = Nothing) As Task(Of String)
        Dim provider = If(String.IsNullOrWhiteSpace(providerName), "Ollama", providerName)
        Dim apiKey = hostAddress
        Dim model = modelName
        Select Case provider
            Case "Ollama"
                Dim payload = New With {
                .model = model,
                .messages = messages,
                .stream = False
            }
                Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(payload)
                Dim content = New System.Net.Http.StringContent(json,
                               System.Text.Encoding.UTF8, "application/json")
                Dim response As System.Net.Http.HttpResponseMessage
                If ct <> Nothing AndAlso ct <> CancellationToken.None Then
                    response = Await httpClient.PostAsync(hostAddress, content, ct)
                Else
                    response = Await httpClient.PostAsync(hostAddress, content)
                End If
                Return Await response.Content.ReadAsStringAsync()
            Case "OpenAI"
                Dim payload = New With {
                .model = model,
                .messages = messages
            }
                Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(payload)
                Dim content = New System.Net.Http.StringContent(json,
                              System.Text.Encoding.UTF8, "application/json")
                Using client As New System.Net.Http.HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(300)
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                    Dim response As System.Net.Http.HttpResponseMessage
                    If ct <> Nothing AndAlso ct <> CancellationToken.None Then
                        response = Await client.PostAsync(
                        "https://api.openai.com/v1/chat/completions", content, ct)
                    Else
                        response = Await client.PostAsync(
                        "https://api.openai.com/v1/chat/completions", content)
                    End If
                    Return Await response.Content.ReadAsStringAsync()
                End Using
            Case "Anthropic"
                Dim systemText As String = ""
                Dim anthropicMessages As New List(Of Object)
                For Each m In messages
                    Dim mType = m.GetType()
                    Dim roleProp = mType.GetProperty("role")
                    Dim contentProp = mType.GetProperty("content")
                    If roleProp Is Nothing OrElse contentProp Is Nothing Then Continue For
                    Dim role = roleProp.GetValue(m)?.ToString()
                    Dim msgContent = contentProp.GetValue(m)?.ToString()
                    If role = "system" Then
                        systemText = If(msgContent, "")
                    Else
                        anthropicMessages.Add(New With {.role = role, .content = msgContent})
                    End If
                Next
                Dim payload = New With {
                .model = model,
                .max_tokens = 8096,
                .system = systemText,
                .messages = anthropicMessages.ToArray()
            }
                Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(payload)
                Dim content = New System.Net.Http.StringContent(json,
                              System.Text.Encoding.UTF8, "application/json")
                Using client As New System.Net.Http.HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(300)
                    client.DefaultRequestHeaders.Add("x-api-key", apiKey)
                    client.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01")
                    Dim response As System.Net.Http.HttpResponseMessage
                    If ct <> Nothing AndAlso ct <> CancellationToken.None Then
                        response = Await client.PostAsync(
                        "https://api.anthropic.com/v1/messages", content, ct)
                    Else
                        response = Await client.PostAsync(
                        "https://api.anthropic.com/v1/messages", content)
                    End If
                    Return Await response.Content.ReadAsStringAsync()
                End Using
            Case "Google"
                Dim geminiContents As New List(Of Object)
                Dim systemParts As New List(Of String)
                For Each m In messages
                    Dim mType = m.GetType()
                    Dim roleProp = mType.GetProperty("role")
                    Dim contentProp = mType.GetProperty("content")
                    If roleProp Is Nothing OrElse contentProp Is Nothing Then Continue For
                    Dim role = roleProp.GetValue(m)?.ToString()
                    Dim msgContent = contentProp.GetValue(m)?.ToString()
                    If role = "system" Then
                        systemParts.Add(If(msgContent, ""))
                    ElseIf role = "assistant" Then
                        geminiContents.Add(New With {
                        .role = "model",
                        .parts = New Object() {New With {.text = msgContent}}
                    })
                    Else
                        geminiContents.Add(New With {
                        .role = "user",
                        .parts = New Object() {New With {.text = msgContent}}
                    })
                    End If
                Next
                If systemParts.Count > 0 Then
                    geminiContents.Insert(0, New With {
                    .role = "user",
                    .parts = New Object() {New With {
                        .text = String.Join(Environment.NewLine, systemParts)
                    }}
                })
                End If
                Dim geminiPayload = New With {
                .contents = geminiContents.ToArray()
            }
                Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(geminiPayload)
                Dim content = New System.Net.Http.StringContent(json,
                              System.Text.Encoding.UTF8, "application/json")
                Dim url = $"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={apiKey}"
                Using client As New System.Net.Http.HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(300)
                    Dim response As System.Net.Http.HttpResponseMessage
                    If ct <> Nothing AndAlso ct <> CancellationToken.None Then
                        response = Await client.PostAsync(url, content, ct)
                    Else
                        response = Await client.PostAsync(url, content)
                    End If
                    Return Await response.Content.ReadAsStringAsync()
                End Using

            Case Else
                Return "Error: Unknown provider '" & provider & "'"
        End Select
    End Function
    Private Function RunPowerShellSync(script As String) As String
        Try
            Dim safeCheck = IsCommandSafe(script)
            If Not safeCheck.Safe Then Return $"(Blocked: {safeCheck.Reason})"
            Dim scriptFile = GetUniqueTempPath("ps_script", ".ps1")
            File.WriteAllText(scriptFile, script, Encoding.UTF8)
            Dim psi As New ProcessStartInfo("powershell.exe") With {
            .Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -File ""{scriptFile}""",
            .RedirectStandardOutput = True,
            .RedirectStandardError = True,
            .UseShellExecute = False,
            .CreateNoWindow = True
        }
            Using proc = StartProcessSafe(psi)
                Dim output = proc.StandardOutput.ReadToEnd()
                Dim errors = proc.StandardError.ReadToEnd()
                If Not proc.WaitForExit(CommandTimeoutMs) Then
                    Try : proc.Kill() : Catch : End Try
                    Try : File.Delete(scriptFile) : Catch : End Try
                    Return "(Timed out after " & (CommandTimeoutMs / 1000) & "s)"
                End If
                Try : File.Delete(scriptFile) : Catch : End Try

                If Not String.IsNullOrWhiteSpace(errors) Then
                    Return "(Error: " & errors.Trim() & ") " & output.Trim()
                End If
                Return output.Trim()
            End Using
        Catch ex As InvalidOperationException
            LogToFile("PS_START_ERROR", ex.Message)
            Return "(Process start failed: " & ex.Message & ")"
        Catch ex As Exception
            Return "(PowerShell failed: " & ex.Message & ")"
        End Try
    End Function
    Private Function CaptureScreenBase64() As String
        Using bmp As New Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
            Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bmp)
                g.CopyFromScreen(Point.Empty, Point.Empty, Screen.PrimaryScreen.Bounds.Size)
            End Using
            Using ms As New MemoryStream()
                bmp.Save(ms, Imaging.ImageFormat.Jpeg)
                Return Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
    End Function
    Private Function CaptureScreenWithInfo() As (Base64 As String, Metadata As String)
        Dim base64 = CaptureScreenBase64()
        Dim metadata = $"Resolution: {Screen.PrimaryScreen.Bounds.Width}x{Screen.PrimaryScreen.Bounds.Height}" & Environment.NewLine &
            $"Scaling: {GetDpiScaling()}%" & Environment.NewLine &
            $"Monitor: {Screen.PrimaryScreen.DeviceName}"
        Return (base64, metadata)
    End Function
    Private Function GetDpiScaling() As Integer
        Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromHwnd(IntPtr.Zero)
            Dim dpiX = g.DpiX
            Return CInt((dpiX / 96) * 100)
        End Using
    End Function
    Private Sub ShowNotification(title As String, message As String)
        Try
            NotifyIcon1.ShowBalloonTip(5000, title, message, ToolTipIcon.Info)
        Catch ex As Exception
            LogToFile("NOTIFICATION_ERROR", ex.Message)
        End Try
    End Sub
    Private Sub NotifyIcon1_BalloonTipClicked(sender As Object, e As EventArgs) Handles NotifyIcon1.BalloonTipClicked
        Me.WindowState = FormWindowState.Normal
        Me.Activate()
        Me.BringToFront()
        QueryTextBox.Focus()
    End Sub
    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        QueryTextBox.Clear()
        gptResponse = ""
        While conversationHistory.Count > 0
            Dim discard As ConversationEntry = Nothing
            conversationHistory.TryDequeue(discard)
        End While
        correctionAttempts = 0
        QueryTextBox.Focus()
    End Sub
    Private Sub KillButton_Click(sender As Object, e As EventArgs) Handles KillButton.Click
        Try
            If pendingTaskQueue.Count > 0 AndAlso sender IsNot KillButton Then Return
            ShellTimer.Stop()
            Try
                executionCts?.Cancel()
            Catch : End Try
            Interlocked.Exchange(isRunning, 0)
            commandExecutedPS = True
            gptResponse = ""
            correctionAttempts = 0
            previousCommandPS = ""
            pendingTargetApp = ""
            pendingDocContent = ""
            pendingDocSavePath = ""
            pendingNotification = ""
            lastExtractedWebText = ""
            Dim discardTask As String = Nothing
            While pendingTaskQueue.TryDequeue(discardTask)
            End While
            For Each proc In Process.GetProcessesByName("powershell")
                Try : proc.Kill() : Catch : End Try
            Next
            For Each proc In Process.GetProcessesByName("pwsh")
                Try : proc.Kill() : Catch : End Try
            Next
            CleanupBrowserAsync().Wait(15000)
            For Each bName In {"msedge", "chrome", "chromium"}
                Try
                    For Each proc In Process.GetProcessesByName(bName)
                        If proc.MainWindowTitle = "" Then
                            Try : proc.Kill() : Catch : End Try
                        End If
                    Next
                Catch : End Try
            Next
            Try : Clipboard.Clear() : Catch : End Try
            AppendResponse("🛑 Automation stopped. Browser closed and session data wiped.")
            SetUIEnabled(True)
            ShellTimer.Start()
        Catch ex As Exception
            LogToFile("KILL_ERROR", ex.Message)
            Try : SetUIEnabled(True) : Catch : End Try
            Try : ShellTimer.Start() : Catch : End Try
        End Try
    End Sub
    Private Sub SettingsButton_Click(sender As Object, e As EventArgs) Handles SettingsButton.Click
        SettingsF.Show()
    End Sub
    Private Sub SetUIEnabled(enabled As Boolean)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() SetUIEnabled(enabled))
            Return
        End If
        SendButton.Enabled = enabled
        QueryTextBox.Enabled = enabled
    End Sub
    Private Function StartProcessSafe(psi As ProcessStartInfo) As Process
        Dim proc = Process.Start(psi)
        If proc Is Nothing Then
            Throw New InvalidOperationException(
            $"Process.Start returned Nothing for '{psi.FileName} {psi.Arguments}'")
        End If
        Return proc
    End Function
    Private Function GetBootPrompt() As String
        Try
            Dim bootPath = Path.Combine(Application.StartupPath, "BOOT")
            If File.Exists(bootPath) Then
                Return File.ReadAllText(bootPath, Encoding.UTF8)
            Else
                LogToFile("BOOT_ERROR", "BOOT file not found at: " & bootPath)
                Return "You are a Windows Automation Engine. Generate PowerShell scripts wrapped in ||| delimiters."
            End If
        Catch ex As Exception
            LogToFile("BOOT_ERROR", "Failed to read BOOT file: " & ex.Message)
            Return "You are a Windows Automation Engine. Generate PowerShell scripts wrapped in ||| delimiters."
        End Try
    End Function
    Private Async Function RunAgentWorkflowAsync(gptPlanJson As String) As Task(Of Boolean)
        Dim retryCount As Integer = 0
        Dim currentPlanJson As String = gptPlanJson
        While retryCount < MaxRetries
            retryCount += 1
            Try
                Dim workflow = JsonConvert.DeserializeObject(Of WebWorkflow)(currentPlanJson)
                If workflow Is Nothing OrElse workflow.Steps Is Nothing OrElse workflow.Steps.Count = 0 Then
                    AppendResponse("   ❌ Empty or invalid workflow JSON.")
                    Return False
                End If
                AppendResponse($"   Attempt {retryCount}/{MaxRetries} ({workflow.Steps.Count} steps)...")
                Dim webResult = Await RunWebWorkflowAsync(workflow)
                Dim meaningfulStepCount As Integer = workflow.Steps.Where(
    Function(s) s.Action IsNot Nothing AndAlso
        {"goto", "fill", "click", "selectoption"}.Contains(s.Action.ToLower())).Count()
                Dim seemsReal As Boolean = webResult.Success AndAlso
                (webResult.StepsCompleted >= meaningfulStepCount OrElse
                 meaningfulStepCount = 0)
                If seemsReal Then
                    AppendResponse($"   ✅ Succeeded on attempt {retryCount}! ({webResult.StepsCompleted}/{workflow.Steps.Count} steps completed)")
                    Return True
                End If
                If webResult.Success AndAlso webResult.StepsCompleted < meaningfulStepCount Then
                    webResult.Success = False
                    webResult.Message = If(String.IsNullOrEmpty(webResult.Message),
                    $"Only {webResult.StepsCompleted}/{meaningfulStepCount} meaningful steps completed — workflow likely failed silently.",
                    webResult.Message)
                    AppendResponse($"   ⚠️ Reported success but only {webResult.StepsCompleted}/{meaningfulStepCount} steps actually worked.")
                End If
                If retryCount >= MaxRetries Then
                    AppendResponse($"   ❌ Failed after {MaxRetries} attempts.")
                    AppendResponse($"   Last error: {webResult.Message}")
                    Return False
                End If
                If IsBrowserClosedException(webResult.Message) Then
                    AppendResponse("   🛑 Browser closed — workflow stopped.")
                    Return False
                End If
                AppendResponse($"   ❌ Attempt {retryCount} failed: {webResult.Message}")
                Dim errorParts As New List(Of String)
                errorParts.Add($"Web automation failed on attempt {retryCount} of {MaxRetries}.")
                errorParts.Add($"Error: {webResult.Message}")
                errorParts.Add($"Steps completed: {webResult.StepsCompleted}/{workflow.Steps.Count}")
                errorParts.Add($"Current URL: {If(webResult.CurrentUrl, "unknown")}")
                errorParts.Add($"Page title: {If(webResult.PageTitle, "unknown")}")
                If Not String.IsNullOrEmpty(webResult.FailedStepDescription) Then
                    errorParts.Add($"Failed step: {webResult.FailedStepDescription}")
                End If
                If Not String.IsNullOrEmpty(webResult.ElementMap) Then
                    Dim truncMap = webResult.ElementMap
                    If truncMap.Length > 4000 Then truncMap = truncMap.Substring(0, 4000) & "... [TRUNCATED]"
                    errorParts.Add($"Available elements on page:{Environment.NewLine}{truncMap}")
                End If
                If Not String.IsNullOrEmpty(webResult.DomPath) AndAlso File.Exists(webResult.DomPath) Then
                    Dim domText = File.ReadAllText(webResult.DomPath)
                    If domText.Length > 3000 Then domText = domText.Substring(0, 3000) & "... [TRUNCATED]"
                    errorParts.Add($"DOM:{Environment.NewLine}{domText}")
                End If
                If Not String.IsNullOrEmpty(webResult.ConsoleLogPath) AndAlso File.Exists(webResult.ConsoleLogPath) Then
                    Dim consoleText = File.ReadAllText(webResult.ConsoleLogPath)
                    If consoleText.Length > 2000 Then consoleText = consoleText.Substring(0, 2000) & "... [TRUNCATED]"
                    errorParts.Add($"Console:{Environment.NewLine}{consoleText}")
                End If
                If Not String.IsNullOrEmpty(webResult.OcrText) Then
                    errorParts.Add($"OCR text from screenshot:{Environment.NewLine}{webResult.OcrText}")
                End If
                Dim errorSummary = String.Join(Environment.NewLine, errorParts) & Environment.NewLine &
                "Analyze what went wrong and return ONLY the corrected JSON workflow wrapped in ||| delimiters." & Environment.NewLine &
                "Make sure every fill and click step has observeBefore: true and a specific description."
                Dim screenData = CaptureScreenWithInfo()
                Dim correctionResponse = Await GetChatResponseWithVisionAsync(
                "You are a web automation expert. Analyze the failure and fix the JSON workflow. " &
                "Common fixes: use more specific descriptions, add wait steps between actions, " &
                "use observeBefore: true on all click/fill steps, add scroll steps if elements are below the fold. " &
                "Return ONLY the corrected JSON wrapped in ||| delimiters.",
                errorSummary, screenData.Base64)
                Dim correctionMatch = Regex.Match(correctionResponse, "\|\|\|(.*?)\|\|\|", RegexOptions.Singleline)
                If correctionMatch.Success Then
                    Dim correctedJson = correctionMatch.Groups(1).Value.Trim()
                    Try
                        Dim testParse = JsonConvert.DeserializeObject(Of WebWorkflow)(correctedJson)
                        If testParse IsNot Nothing AndAlso testParse.Steps IsNot Nothing AndAlso testParse.Steps.Count > 0 Then
                            currentPlanJson = correctedJson
                            AppendResponse($"   🔄 Got corrected workflow ({testParse.Steps.Count} steps) for attempt {retryCount + 1}...")
                        Else
                            AppendResponse("   ⚠️ Corrected workflow was empty. Retrying original...")
                        End If
                    Catch
                        AppendResponse("   ⚠️ Corrected workflow was invalid JSON. Retrying original...")
                    End Try
                Else
                    AppendResponse("   ⚠️ AI didn't provide a corrected workflow. Retrying original...")
                End If
                Await Task.Delay(2000)
            Catch ex As Exception When IsBrowserClosedException(ex.Message)
                AppendResponse("   🛑 Browser closed — workflow stopped.")
                Return False
            Catch ex As Exception
                AppendResponse($"   ❌ Workflow error: {ex.Message}")
                LogToFile("WORKFLOW_ERROR", ex.ToString())
                If retryCount >= MaxRetries Then Return False
            End Try
        End While
        Return False
    End Function
    Private Async Function RunWebWorkflowAsync(workflow As WebWorkflow) As Task(Of WebAutomationResult)
        Dim result As New WebAutomationResult With {.Success = False, .StepsCompleted = 0}
        Dim consoleLogs As New List(Of String)
        Dim page As IPage = Nothing
        Dim browser As IBrowser = Nothing
        Dim pw As IPlaywright = Nothing
        Dim ctx As IBrowserContext = Nothing
        Dim captureScreenshot As Boolean = False
        Dim errorMessage As String = ""
        Dim stepsCompleted As Integer = 0
        Try
            ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🔧 Starting browser..." & Environment.NewLine))
            pw = Await Playwright.CreateAsync()
            activePlaywright = pw
            browser = Await pw.Chromium.LaunchAsync(New BrowserTypeLaunchOptions With {
            .Headless = False,
            .Args = New String() {
                "--start-maximized", "--disable-infobars",
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox", "--disable-dev-shm-usage", "--lang=en-US"
            }
        })
            activeBrowser = browser
            ctx = Await browser.NewContextAsync(New BrowserNewContextOptions With {
            .ViewportSize = New ViewportSize() With {.Width = 1920, .Height = 1080},
            .UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
        })
            activeContext = ctx
            page = Await ctx.NewPageAsync()
            activePage = page
            Await page.AddInitScriptAsync(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined});" &
            "Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});" &
            "Object.defineProperty(navigator, 'languages', {get: () => ['en-US','en']});" &
            "window.chrome = {runtime: {}};" &
            "Object.defineProperty(navigator, 'permissions', {get: () => ({query: () => Promise.resolve({state: 'granted'})})});" &
            "const origQuery = window.navigator.permissions.query;" &
            "window.navigator.permissions.query = (p) => p.name==='notifications' ? Promise.resolve({state: Notification.permission}) : origQuery(p);"
        )
            Await BringBrowserToFrontAsync()
            result.WindowInFocus = True
            Try
                AddHandler page.Console, Sub(sender, msg) consoleLogs.Add("[" & msg.Type & "] " & msg.Text)
            Catch
            End Try
            Dim workflowFailed As Boolean = False
            For Each stepItem In workflow.Steps
                If workflowFailed Then Exit For
                If stepItem.ObserveBefore Then
                    ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🔍 Observing before: " & stepItem.Action & "..." & Environment.NewLine))
                    Try
                        Await page.WaitForLoadStateAsync(LoadState.NetworkIdle, New PageWaitForLoadStateOptions With {.Timeout = 15000})
                    Catch
                    End Try
                    Try
                        Await page.WaitForLoadStateAsync(LoadState.DOMContentLoaded)
                    Catch
                    End Try
                    Await Task.Delay(3000)
                    Dim preObs = Await GetPageObservationAsync(page)
                    Dim obsForCount = preObs.Text
                    Dim axSplit = obsForCount.IndexOf("=== ACCESSIBILITY")
                    If axSplit > 0 Then obsForCount = obsForCount.Substring(0, axSplit)
                    Dim elemCount = obsForCount.Split({Environment.NewLine, Chr(10)}, StringSplitOptions.RemoveEmptyEntries) _
                    .Where(Function(l) l.Contains(" | ")).Count()
                    If elemCount < 5 Then
                        ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   ⏳ Few elements (" & elemCount & "), waiting..." & Environment.NewLine))
                        Await Task.Delay(5000)
                        Try
                            Await page.EvaluateAsync("window.scrollBy(0, 100)")
                            Await Task.Delay(1000)
                            Await page.EvaluateAsync("window.scrollBy(0, -100)")
                        Catch
                        End Try
                        preObs = Await GetPageObservationAsync(page)
                        obsForCount = preObs.Text
                        axSplit = obsForCount.IndexOf("=== ACCESSIBILITY")
                        If axSplit > 0 Then obsForCount = obsForCount.Substring(0, axSplit)
                        elemCount = obsForCount.Split({Environment.NewLine, Chr(10)}, StringSplitOptions.RemoveEmptyEntries) _
                        .Where(Function(l) l.Contains(" | ")).Count()
                    End If
                    result.AccessibilityTree = preObs.Text
                    ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   📋 Elements found: " & elemCount & Environment.NewLine))
                    If Not String.IsNullOrWhiteSpace(stepItem.Description) Then
                        Dim jsMapOnly As String = preObs.Text
                        Dim axDivider = jsMapOnly.IndexOf("=== ACCESSIBILITY ROLES")
                        If axDivider > 0 Then jsMapOnly = jsMapOnly.Substring(0, axDivider)
                        Dim knownSelector = GetKnownSelector(page.Url, stepItem.Description, jsMapOnly)
                        If Not String.IsNullOrWhiteSpace(knownSelector) Then
                            ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   📌 Known selector: " & knownSelector & Environment.NewLine))
                            stepItem.Selector = knownSelector
                        End If
                        Dim aiSelector = Await AskAIForBestSelectorAsync(stepItem.Description, preObs)
                        Dim bareSelectorTags = {
                        "a", "p", "div", "span", "input", "button", "form", "plaintext",
                        "textarea", "select", "#thumbnail", "thumbnail",
                        "#button", "#logo", "#guide", "#back", "#menu", "#close",
                        "a[href=""#main""]", "a[href='#main']"
                    }
                        Dim aiIsGood = Not String.IsNullOrWhiteSpace(aiSelector) AndAlso
                        Not bareSelectorTags.Contains(aiSelector.ToLower().Trim()) AndAlso
                        Not aiSelector.ToLower().StartsWith("nodeid:") AndAlso
                        (aiSelector.Contains("#") OrElse aiSelector.Contains("[") OrElse
                         aiSelector.Contains(".") OrElse aiSelector.Contains("=") OrElse
                         aiSelector.StartsWith("text=") OrElse aiSelector.StartsWith("role="))
                        If aiIsGood Then
                            ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🎯 AI selector: " & aiSelector & Environment.NewLine))
                            stepItem.Selector = aiSelector
                        End If
                        If String.IsNullOrWhiteSpace(stepItem.Selector) OrElse
                       stepItem.Selector.ToLower().StartsWith("nodeid:") Then
                            Dim textMatched = FindBestElementByText(jsMapOnly, stepItem.Description)
                            If Not String.IsNullOrWhiteSpace(textMatched) Then
                                ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🎯 Text-matched: " & textMatched & Environment.NewLine))
                                stepItem.Selector = textMatched
                            End If
                        End If
                        If String.IsNullOrWhiteSpace(stepItem.Selector) Then
                            stepItem.Selector = "text=" & stepItem.Description
                            ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🔤 Text fallback: " & stepItem.Selector & Environment.NewLine))
                        End If
                    End If
                End If
                Dim stepFailed As Boolean = False
                Dim stepError As String = ""
                ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText(
                "   ▶ Step: " & stepItem.Action &
                If(Not String.IsNullOrEmpty(stepItem.Url), " → " & stepItem.Url.Substring(0, Math.Min(50, stepItem.Url.Length)), "") &
                Environment.NewLine))
                Try
                    Select Case stepItem.Action.ToLower()
                        Case "goto"
                            Dim gotoFailed As Boolean = False
                            Try
                                Await page.GotoAsync(stepItem.Url, New PageGotoOptions With {
                                .Timeout = If(stepItem.TimeoutMs > 0, stepItem.TimeoutMs, 60000),
                                .WaitUntil = WaitUntilState.DOMContentLoaded
                            })
                            Catch ex As Exception
                                Debug.WriteLine("Goto DOMContentLoaded failed: " & ex.Message)
                                gotoFailed = True
                            End Try
                            If gotoFailed Then
                                Await page.GotoAsync(stepItem.Url, New PageGotoOptions With {
                                .Timeout = If(stepItem.TimeoutMs > 0, stepItem.TimeoutMs, 60000),
                                .WaitUntil = WaitUntilState.Commit
                            })
                            End If
                            Try
                                Dim chromProcs = Process.GetProcessesByName("chrome")
                                If chromProcs.Length = 0 Then chromProcs = Process.GetProcessesByName("chromium")
                                For Each p In chromProcs.OrderByDescending(Function(x) x.StartTime)
                                    If p.MainWindowHandle <> IntPtr.Zero Then
                                        NativeMethods.ShowWindow(p.MainWindowHandle, 3)
                                        Exit For
                                    End If
                                Next
                            Catch
                            End Try
                            result.CurrentUrl = page.Url
                            result.PageTitle = Await page.TitleAsync()
                            ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🌍 Loaded: " & result.CurrentUrl & Environment.NewLine))
                            stepsCompleted += 1
                        Case "click"
                            If String.IsNullOrWhiteSpace(stepItem.Selector) Then
                                Throw New Exception("click: no selector for '" & stepItem.Description & "'")
                            End If
                            Dim clickSelectorBad = stepItem.Selector.ToLower().StartsWith("nodeid:") OrElse
                            (stepItem.Selector.StartsWith("text=") AndAlso stepItem.Selector.Length > 60)
                            If clickSelectorBad AndAlso Not String.IsNullOrWhiteSpace(stepItem.Description) Then
                                ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🔍 Bad selector, re-scanning..." & Environment.NewLine))
                                Dim liveScan = Await GetCDPElementMapFallbackAsync(page)
                                Dim liveObs = (Base64:="", Text:="URL: " & page.Url & Environment.NewLine &
                                "=== INTERACTIVE ELEMENTS ===" & Environment.NewLine & liveScan)
                                Dim freshSelector = Await AskAIForBestSelectorAsync(stepItem.Description, liveObs)
                                If Not String.IsNullOrWhiteSpace(freshSelector) AndAlso
                               Not freshSelector.ToLower().StartsWith("nodeid:") Then
                                    stepItem.Selector = freshSelector
                                Else
                                    Dim bestMatch = FindBestElementByText(liveScan, stepItem.Description)
                                    If Not String.IsNullOrWhiteSpace(bestMatch) Then stepItem.Selector = bestMatch
                                End If
                            End If
                            Await ShowClickIndicatorAsync(page, stepItem.Selector)
                            Dim clickDone As Boolean = False
                            Try
                                Await page.ClickAsync(stepItem.Selector, New PageClickOptions With {
                                .Timeout = If(stepItem.TimeoutMs > 0, stepItem.TimeoutMs, 10000)})
                                clickDone = True
                            Catch
                            End Try
                            If Not clickDone Then
                                Try
                                    Await page.Locator(stepItem.Selector).First.ClickAsync(New LocatorClickOptions With {.Timeout = 10000})
                                    clickDone = True
                                Catch
                                End Try
                            End If
                            If Not clickDone Then
                                Try
                                    Await page.EvaluateAsync("(sel) => { var el = document.querySelector(sel); if(el) el.click(); }", stepItem.Selector)
                                    clickDone = True
                                Catch
                                End Try
                            End If
                            If Not clickDone Then
                                Throw New Exception("All click attempts failed for: " & stepItem.Selector)
                            End If
                            stepsCompleted += 1
                        Case "fill"
                            If String.IsNullOrWhiteSpace(stepItem.Selector) Then
                                Throw New Exception("fill: no selector for '" & stepItem.Description & "'")
                            End If
                            Dim fillSelectorBad = stepItem.Selector.ToLower().StartsWith("nodeid:") OrElse
                            stepItem.Selector.ToLower() = "#button" OrElse
                            stepItem.Selector.ToLower() = "#logo"
                            If fillSelectorBad AndAlso Not String.IsNullOrWhiteSpace(stepItem.Description) Then
                                Dim liveScanFill = Await GetCDPElementMapFallbackAsync(page)
                                Dim liveObsFill = (Base64:="", Text:="URL: " & page.Url & Environment.NewLine &
                                "=== INTERACTIVE ELEMENTS ===" & Environment.NewLine & liveScanFill)
                                Dim freshFillSel = Await AskAIForBestSelectorAsync(stepItem.Description, liveObsFill)
                                If Not String.IsNullOrWhiteSpace(freshFillSel) AndAlso
                               Not freshFillSel.ToLower().StartsWith("nodeid:") Then
                                    stepItem.Selector = freshFillSel
                                Else
                                    Dim textMatch = FindBestElementByText(liveScanFill, stepItem.Description)
                                    If Not String.IsNullOrWhiteSpace(textMatch) Then stepItem.Selector = textMatch
                                End If
                            End If
                            Dim fillDone As Boolean = False
                            Try
                                Await page.FillAsync(stepItem.Selector, stepItem.Value, New PageFillOptions With {.Timeout = 3000})
                                fillDone = True
                            Catch
                            End Try
                            If Not fillDone Then
                                Try
                                    Await page.ClickAsync(stepItem.Selector, New PageClickOptions With {.Timeout = 5000})
                                    Await page.Keyboard.PressAsync("Control+a")
                                    Await page.Keyboard.TypeAsync(stepItem.Value, New KeyboardTypeOptions With {.Delay = 40})
                                    fillDone = True
                                Catch
                                End Try
                            End If
                            If Not fillDone Then
                                Try
                                    Await page.Locator(stepItem.Selector).First.FillAsync(stepItem.Value, New LocatorFillOptions With {.Timeout = 3000})
                                    fillDone = True
                                Catch
                                End Try
                            End If
                            If Not fillDone Then
                                Try
                                    Await page.Locator("[contenteditable]").First.FillAsync(stepItem.Value, New LocatorFillOptions With {.Timeout = 3000})
                                    fillDone = True
                                Catch
                                End Try
                            End If
                            If Not fillDone Then
                                Try
                                    Await page.ClickAsync("input,textarea,[contenteditable],[role=searchbox],[role=textbox],[role=combobox]", New PageClickOptions With {.Timeout = 5000})
                                    Await page.Keyboard.TypeAsync(stepItem.Value, New KeyboardTypeOptions With {.Delay = 40})
                                    fillDone = True
                                Catch
                                End Try
                            End If
                            If Not fillDone Then
                                ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   ⏭️ No input found — skipping fill" & Environment.NewLine))
                                GoTo ContinueLoop
                            End If
                            stepsCompleted += 1
                        Case "press"
                            Await page.Keyboard.PressAsync(stepItem.Value)
                            stepsCompleted += 1
                        Case "wait"
                            Await page.WaitForTimeoutAsync(If(stepItem.TimeoutMs > 0, stepItem.TimeoutMs, 1000))
                            stepsCompleted += 1
                        Case "waitforselector"
                            Try
                                Await page.WaitForSelectorAsync(stepItem.Selector, New PageWaitForSelectorOptions With {
                                .Timeout = If(stepItem.TimeoutMs > 0, stepItem.TimeoutMs, 30000)})
                            Catch
                            End Try
                            stepsCompleted += 1
                        Case "scroll"
                            Await ShowScrollIndicatorAsync(page)
                            Dim scrollVal = If(stepItem.Value IsNot Nothing, stepItem.Value.ToLower(), "")
                            If scrollVal = "bottom" Then
                                Await page.EvaluateAsync("window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'})")
                            ElseIf scrollVal = "top" Then
                                Await page.EvaluateAsync("window.scrollTo({top: 0, behavior: 'smooth'})")
                            ElseIf scrollVal.StartsWith("by:") Then
                                Dim amount = Integer.Parse(stepItem.Value.Substring(3))
                                Await page.EvaluateAsync("window.scrollBy({top: " & amount & ", behavior: 'smooth'})")
                            ElseIf Not String.IsNullOrEmpty(stepItem.Selector) Then
                                Await page.EvaluateAsync("document.querySelector('" & stepItem.Selector & "').scrollIntoView({behavior: 'smooth'})")
                            End If
                            stepsCompleted += 1
                        Case "slowscroll"
                            Await ShowScrollIndicatorAsync(page)
                            Dim scrollBudgetMs As Integer = If(stepItem.TimeoutMs > 0, Math.Min(stepItem.TimeoutMs, 30000), 8000)
                            Dim stepMs As Integer = 300
                            Dim maxSteps As Integer = scrollBudgetMs \ stepMs
                            Dim previousScrolled As Long = -1
                            Dim stuckCount As Integer = 0
                            For scrollStep = 1 To maxSteps
                                Await page.EvaluateAsync("window.scrollBy(0, window.innerHeight * 0.6); document.documentElement.scrollBy(0, document.documentElement.clientHeight * 0.6);")
                                Await page.WaitForTimeoutAsync(stepMs)
                                Dim scrolledVal As Long = CLng(Await page.EvaluateAsync(Of Long)("Math.max(window.scrollY + window.innerHeight, document.documentElement.scrollTop + document.documentElement.clientHeight)"))
                                Dim totalVal As Long = CLng(Await page.EvaluateAsync(Of Long)("Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)"))
                                If scrolledVal >= totalVal - 20 Then Exit For
                                If scrolledVal = previousScrolled Then
                                    stuckCount += 1
                                    If stuckCount >= 3 Then Exit For
                                Else
                                    stuckCount = 0
                                End If
                                previousScrolled = scrolledVal
                            Next
                            stepsCompleted += 1
                        Case "extracttext"
                            Dim extractedTxt As String = ""
                            Dim extractFailed As Boolean = False
                            If Not String.IsNullOrWhiteSpace(stepItem.Selector) Then
                                Try
                                    extractedTxt = Await page.InnerTextAsync(stepItem.Selector,
                New PageInnerTextOptions With {.Timeout = 5000})
                                Catch
                                    extractFailed = True
                                End Try
                            Else
                                extractFailed = True
                            End If
                            If extractFailed Then
                                extractedTxt = Await page.EvaluateAsync(Of String)("document.body.innerText")
                            End If
                            If extractedTxt.Length > 8000 Then
                                extractedTxt = extractedTxt.Substring(0, 8000) & "... [TRUNCATED]"
                            End If
                            Dim currentExt = If(result.ExtractedText, "")
                            If currentExt.Length + extractedTxt.Length > 50000 Then
                                extractedTxt = extractedTxt.Substring(0,
            Math.Max(0, 50000 - currentExt.Length)) & "... [CUMULATIVE LIMIT]"
                            End If
                            result.ExtractedText = currentExt & extractedTxt
                            result.Message = If(result.Message, "") & "[extractText] " &
        extractedTxt.Substring(0, Math.Min(200, extractedTxt.Length)) & Environment.NewLine
                            stepsCompleted += 1
                        Case "gettext"
                            Try
                                Await page.WaitForLoadStateAsync(LoadState.DOMContentLoaded,
            New PageWaitForLoadStateOptions With {.Timeout = 10000})
                            Catch
                            End Try
                            Await Task.Delay(2000)
                            Dim bodyTxt = Await page.EvaluateAsync(Of String)("document.body.innerText")
                            If bodyTxt.Length > 8000 Then bodyTxt = bodyTxt.Substring(0, 8000) & "... [TRUNCATED]"
                            Dim currentExtracted = If(result.ExtractedText, "")
                            If currentExtracted.Length + bodyTxt.Length > 50000 Then
                                bodyTxt = bodyTxt.Substring(0, Math.Max(0, 50000 - currentExtracted.Length)) & "... [CUMULATIVE LIMIT]"
                            End If
                            result.ExtractedText = currentExtracted & bodyTxt & Environment.NewLine
                            result.Message = If(result.Message, "") & "[gettext] " &
        bodyTxt.Substring(0, Math.Min(200, bodyTxt.Length)) & Environment.NewLine
                            ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText(
        "   📖 Got " & bodyTxt.Length & " chars" & Environment.NewLine))
                            stepsCompleted += 1
                        Case "waitfornavigation"
                            Await page.WaitForURLAsync("**", New PageWaitForURLOptions With {
                            .Timeout = If(stepItem.TimeoutMs > 0, stepItem.TimeoutMs, 30000),
                            .WaitUntil = WaitUntilState.Load
                        })
                            result.CurrentUrl = page.Url
                            result.PageTitle = Await page.TitleAsync()
                            stepsCompleted += 1
                        Case "presssequentially"
                            For Each c In stepItem.Value.ToCharArray()
                                Await page.Keyboard.PressAsync(c.ToString())
                                Await Task.Delay(50 + New Random().Next(0, 50))
                            Next
                            stepsCompleted += 1
                        Case "hover"
                            If Not String.IsNullOrWhiteSpace(stepItem.Selector) Then
                                Await page.HoverAsync(stepItem.Selector)
                            End If
                            stepsCompleted += 1
                        Case "selectoption"
                            Await page.SelectOptionAsync(stepItem.Selector, stepItem.Value)
                            stepsCompleted += 1
                        Case "screenshot"
                            Dim screenshotPath = Path.Combine(Path.GetTempPath(), "screenshot_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".png")
                            Await page.ScreenshotAsync(New PageScreenshotOptions With {.Path = screenshotPath})
                            result.Message = If(result.Message, "") & "[screenshot] Saved to " & screenshotPath & Environment.NewLine
                            stepsCompleted += 1
                        Case "geturl"
                            result.CurrentUrl = page.Url
                            result.Message = If(result.Message, "") & "[url] " & result.CurrentUrl & Environment.NewLine
                            stepsCompleted += 1
                        Case "gettitle"
                            result.PageTitle = Await page.TitleAsync()
                            result.Message = If(result.Message, "") & "[title] " & result.PageTitle & Environment.NewLine
                            stepsCompleted += 1
                        Case "download"
                            Dim download = Await page.RunAndWaitForDownloadAsync(
                            Async Function()
                                Await page.ClickAsync(stepItem.Selector)
                            End Function)
                            Dim dlSavePath = Path.Combine(Application.StartupPath, download.SuggestedFilename)
                            Await download.SaveAsAsync(dlSavePath)
                            result.Message = If(result.Message, "") & "[download] Saved to " & dlSavePath & Environment.NewLine
                            stepsCompleted += 1
                        Case "uiascan"
                            ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🧠 Scanning foreground UIA..." & Environment.NewLine))
                            Dim uiaResult = Await Task.Run(Function() GetForegroundWindowUIATree())
                            result.AccessibilityTree = uiaResult
                            result.Message = If(result.Message, "") & "[uiascan] " & uiaResult.Substring(0, Math.Min(200, uiaResult.Length)) & Environment.NewLine
                            stepsCompleted += 1
                        Case "uiascanall"
                            ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🧠 Scanning all UIA..." & Environment.NewLine))
                            Dim uiaAllResult = Await Task.Run(Function() GetAllWindowsUIATree())
                            result.AccessibilityTree = uiaAllResult
                            result.Message = If(result.Message, "") & "[uiascanall] Done." & Environment.NewLine
                            stepsCompleted += 1
                        Case "snapshot"
                            ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🔍 Observing page..." & Environment.NewLine))
                            Dim snapObs = Await GetPageObservationAsync(page)
                            result.AccessibilityTree = snapObs.Text
                            result.Message = If(result.Message, "") & "[snapshot] Done." & Environment.NewLine
                            stepsCompleted += 1
                        Case "cdplocate"
                            Dim cdpMap = Await GetCDPElementMapAsync(page)
                            result.ElementMap = cdpMap
                            result.Message = If(result.Message, "") & "[cdplocate] " & cdpMap.Substring(0, Math.Min(200, cdpMap.Length)) & Environment.NewLine
                            stepsCompleted += 1
                        Case Else
                            Throw New Exception("Unknown action: " & stepItem.Action)
                    End Select
                Catch stepEx As Exception
                    stepFailed = True
                    stepError = stepEx.Message
                    ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText(
                    "   ❌ Step '" & stepItem.Action & "' error: " & stepEx.Message & Environment.NewLine))
                    result.Message = If(result.Message, "") & "[STEP FAILED: " & stepItem.Action & "] " & stepEx.Message & Environment.NewLine
                    result.FailedStepDescription = stepItem.Action & ": " & If(stepItem.Description, stepItem.Selector)
                    If IsBrowserClosedException(stepEx.Message) Then
                        result.Success = False
                        result.Message = stepEx.Message
                        GoTo CleanupPhase
                    End If
                    If stepItem.Action.ToLower() = "goto" OrElse
                   stepItem.Action.ToLower() = "click" OrElse
                   stepItem.Action.ToLower() = "fill" Then
                        workflowFailed = True
                    End If
                End Try
                If workflowFailed AndAlso page IsNot Nothing Then
                    Try
                        result.ElementMap = Await GetCDPElementMapFallbackAsync(page)
                    Catch
                    End Try
                    Exit For
                End If
ContinueLoop:
                If stepFailed Then
                    Dim captchaDetected As Boolean = False
                    Try
                        captchaDetected = page.Url.Contains("sorry/index") OrElse page.Url.Contains("google.com/sorry")
                        If Not captchaDetected Then
                            Dim bodyText = Await page.InnerTextAsync("body")
                            captchaDetected = bodyText.ToLower().Contains("unusual traffic") OrElse
                            bodyText.ToLower().Contains("i'm not a robot")
                        End If
                    Catch
                    End Try
                    If captchaDetected Then
                        result.KeepOpen = True
                        ResponseRichTextBox.BeginInvoke(Sub() ResponseRichTextBox.AppendText("   🚧 CAPTCHA detected — browser left open." & Environment.NewLine))
                        result.Message = If(result.Message, "") & "[CAPTCHA] Manual intervention required." & Environment.NewLine
                        GoTo CleanupPhase
                    End If
                End If
            Next
            If Not workflowFailed Then
                result.Success = True
            End If
            result.KeepOpen = workflow.KeepOpen
            result.Message = If(String.IsNullOrWhiteSpace(result.Message), "Success", result.Message)
            If Not String.IsNullOrWhiteSpace(result.ExtractedText) Then
                lastExtractedWebText = result.ExtractedText
            End If
        Catch ex As Exception
            If IsBrowserClosedException(ex.Message) Then
                captureScreenshot = False
            Else
                captureScreenshot = True
                errorMessage = ex.Message
            End If
            result.Success = False
            result.Message = "Error: " & ex.Message
        End Try
CleanupPhase:
        result.StepsCompleted = stepsCompleted
        If captureScreenshot AndAlso page IsNot Nothing Then
            Try
                Dim errorDir = Path.Combine(Application.StartupPath, "PlaywrightErrors")
                Directory.CreateDirectory(errorDir)
                Dim timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                result.ScreenshotPath = Path.Combine(errorDir, "error_" & timestamp & ".png")
                result.DomPath = Path.Combine(errorDir, "error_" & timestamp & ".html")
                result.ConsoleLogPath = Path.Combine(errorDir, "error_" & timestamp & "_console.txt")
                Await page.ScreenshotAsync(New PageScreenshotOptions With {.Path = result.ScreenshotPath})
                Dim domContent = Await page.ContentAsync()
                File.WriteAllText(result.DomPath, domContent)
                File.WriteAllLines(result.ConsoleLogPath, consoleLogs)
                result.CurrentUrl = page.Url
                result.PageTitle = Await page.TitleAsync()
            Catch ex As Exception
                LogToFile("CLEANUP_SCREENSHOT_ERROR", ex.Message)
            End Try
        End If
        If consoleLogs.Count > 0 Then
            Try
                Dim logPath = Path.Combine(Path.GetTempPath(),
                $"WebConsole_{DateTime.Now:yyyyMMdd_HHmmss}.txt")
                File.WriteAllLines(logPath, consoleLogs)
                If String.IsNullOrEmpty(result.ConsoleLogPath) Then result.ConsoleLogPath = logPath
            Catch
            End Try
        End If
        If Not result.KeepOpen Then
            Await CleanupBrowserAsync()
        End If
        Return result
    End Function
    Private Shared Function GetUniqueTempPath(prefix As String, extension As String) As String
        Return Path.Combine(Path.GetTempPath(),
        prefix & "_" & Guid.NewGuid().ToString("N").Substring(0, 8) & extension)
    End Function
    Private Shared Function IsBrowserClosedException(message As String) As Boolean
        If String.IsNullOrWhiteSpace(message) Then Return False
        Dim keywords = {"Target page", "context or browser has been closed",
        "Browser has been closed", "Target closed", "Connection refused",
        "disconnected", "Execution context was destroyed", "Session closed"}
        Return keywords.Any(Function(kw) message.IndexOf(kw, StringComparison.OrdinalIgnoreCase) >= 0)
    End Function
    Private Async Function CleanupBrowserAsync() As Task
        Dim localPage = activePage
        activePage = Nothing
        If localPage IsNot Nothing Then
            Try : Await localPage.CloseAsync() : Catch : End Try
        End If
        Dim localCtx = activeContext
        activeContext = Nothing
        If localCtx IsNot Nothing Then
            Try : Await localCtx.CloseAsync() : Catch : End Try
        End If
        Dim localBrowser = activeBrowser
        activeBrowser = Nothing
        If localBrowser IsNot Nothing Then
            Try : Await localBrowser.CloseAsync() : Catch : End Try
        End If
        Dim localPw = activePlaywright
        activePlaywright = Nothing
        If localPw IsNot Nothing Then
            Try : localPw.Dispose() : Catch : End Try
        End If
        Try
            Dim userDataDir = Path.Combine(Path.GetTempPath(), "autono-me-browser-profile")
            If Directory.Exists(userDataDir) Then
                Directory.Delete(userDataDir, True)
            End If
        Catch : End Try
    End Function
    Private Async Function AutoWriteDocumentAsync(docContent As String, targetApp As String, savePath As String) As Task
        Try
            If targetApp = "outlook" Then
                Dim outlookResult = Await Task.Run(Function() HandleOutlookViaPS(docContent, lastUserInput.ToLower()))
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("✅ Outlook: " & outlookResult & Environment.NewLine)
                                                End Sub)
                Return
            End If
            If targetApp = "onenote" Then
                Dim oneNotePrompt As String = "Write a clear, well-structured note about the following content. " &
                "Use headings marked with ## for sections, bullet points starting with - for lists, " &
                "and plain paragraphs for body text. Start with the main title on the first line. Content:" &
                Environment.NewLine & docContent.Substring(0, Math.Min(12000, docContent.Length))
                Dim oneNoteText = Await GetSimpleAIResponseAsync(oneNotePrompt)
                If String.IsNullOrWhiteSpace(oneNoteText) Then oneNoteText = docContent.Substring(0, Math.Min(5000, docContent.Length))
                Dim oneNoteResult = Await Task.Run(Function() HandleOneNoteViaPS(oneNoteText, lastUserInput.ToLower()))
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("✅ OneNote: " & oneNoteResult & Environment.NewLine)
                                                End Sub)
                Return
            End If
            If targetApp = "excel" Then
                Try
                    Dim existingExcelProcs = Process.GetProcessesByName("EXCEL")
                    If existingExcelProcs.Length > 0 Then
                        Dim searchFolders = {
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    }
                        For Each folder In searchFolders
                            If Not Directory.Exists(folder) Then Continue For
                            Dim recentFiles = Directory.GetFiles(folder, "*.xlsx", SearchOption.TopDirectoryOnly).
                            Where(Function(f) File.GetCreationTime(f) > DateTime.Now.AddSeconds(-60)).
                            ToArray()
                            If recentFiles.Length > 0 Then
                                Dim recentFile = recentFiles.OrderByDescending(Function(f) File.GetCreationTime(f)).First()
                                ResponseRichTextBox.BeginInvoke(Sub()
                                                                    ResponseRichTextBox.AppendText("⏭️ Excel file already created recently: " & Path.GetFileName(recentFile) & Environment.NewLine)
                                                                End Sub)
                                Return
                            End If
                        Next
                    End If
                Catch
                End Try
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("📊 Creating Excel spreadsheet from extracted web data..." & Environment.NewLine)
                                                End Sub)
                Dim identifyPrompt As String =
                "Look at this raw web page text and identify what tabular data is present." & Environment.NewLine &
                "Tell me ONLY: what tickers/items are there, and what columns (Date, Open, High, Low, Close, Volume, etc.)." & Environment.NewLine &
                "Reply in 2-3 sentences max. No markdown." & Environment.NewLine &
                Environment.NewLine &
                docContent.Substring(0, Math.Min(4000, docContent.Length))
                Dim dataDescription = Await GetSimpleAIResponseAsync(identifyPrompt)
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("   📋 Data identified: " & dataDescription.Substring(0, Math.Min(150, dataDescription.Length)) & Environment.NewLine)
                                                End Sub)
                Dim tablePrompt As String =
                "You are a data extraction engine. Convert the following raw web page text into clean tab-separated tables." & Environment.NewLine &
                Environment.NewLine &
                "STRICT RULES:" & Environment.NewLine &
                "1. Extract ONLY actual data rows (dates with numbers). Ignore ALL navigation text, menus, headers, footers, ads." & Environment.NewLine &
                "2. A valid stock data row looks like: Feb 27, 2026  390.88  396.82  389.88  392.74  392.74  51,276,300" & Environment.NewLine &
                "3. If there are multiple tickers/items, create a SEPARATE section for each starting with SHEET:NAME" & Environment.NewLine &
                "4. Column headers go on the first line after SHEET:NAME" & Environment.NewLine &
                "5. Use TAB character (not spaces, not commas) between columns" & Environment.NewLine &
                "6. Remove commas from numbers (51,276,300 becomes 51276300)" & Environment.NewLine &
                "7. Format dates consistently (e.g., Feb 27, 2026)" & Environment.NewLine &
                "8. Output ALL rows — never truncate, never write '...' or 'continues'" & Environment.NewLine &
                "9. NO markdown, NO bold, NO code fences, NO explanation text" & Environment.NewLine &
                "10. Start output IMMEDIATELY with SHEET: or column headers — no intro text" & Environment.NewLine &
                Environment.NewLine &
                "CORRECT OUTPUT EXAMPLE:" & Environment.NewLine &
                "SHEET:MSFT" & Environment.NewLine &
                "Date" & Chr(9) & "Open" & Chr(9) & "High" & Chr(9) & "Low" & Chr(9) & "Close" & Chr(9) & "Adj Close" & Chr(9) & "Volume" & Environment.NewLine &
                "Feb 27, 2026" & Chr(9) & "390.88" & Chr(9) & "396.82" & Chr(9) & "389.88" & Chr(9) & "392.74" & Chr(9) & "392.74" & Chr(9) & "51276300" & Environment.NewLine &
                "Feb 26, 2026" & Chr(9) & "395.11" & Chr(9) & "400.12" & Chr(9) & "394.16" & Chr(9) & "397.23" & Chr(9) & "397.23" & Chr(9) & "33960000" & Environment.NewLine &
                Environment.NewLine &
                "WRONG — never do any of these:" & Environment.NewLine &
                "- Starting with: 'Here is the data:' or 'I found the following'" & Environment.NewLine &
                "- Including navigation text like 'Yahoo Finance' 'Search' 'News' 'Markets'" & Environment.NewLine &
                "- Using spaces instead of tabs between columns" & Environment.NewLine &
                "- Using **bold** or ```code fences```" & Environment.NewLine &
                "- Writing '// more rows...' instead of actual data" & Environment.NewLine &
                Environment.NewLine &
                "DATA DESCRIPTION: " & dataDescription & Environment.NewLine &
                Environment.NewLine &
                "RAW WEB PAGE TEXT TO EXTRACT FROM:" & Environment.NewLine &
                docContent.Substring(0, Math.Min(12000, docContent.Length))
                Dim tableText = Await GetSimpleAIResponseAsync(tablePrompt)
                tableText = StripThinkingTags(tableText)
                tableText = tableText.Replace("```", "").Trim()
                If tableText.StartsWith("Here") OrElse tableText.StartsWith("I ") OrElse tableText.StartsWith("Below") Then
                    Dim firstNewline = tableText.IndexOf(Environment.NewLine)
                    If firstNewline > 0 AndAlso firstNewline < 100 Then
                        tableText = tableText.Substring(firstNewline).Trim()
                    End If
                End If
                If String.IsNullOrWhiteSpace(tableText) Then
                    ResponseRichTextBox.BeginInvoke(Sub()
                                                        ResponseRichTextBox.AppendText("   ⚠️ AI returned empty table, using raw data..." & Environment.NewLine)
                                                    End Sub)
                    tableText = docContent.Substring(0, Math.Min(3000, docContent.Length))
                End If
                Dim xlDocsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                Dim xlFileName = "Data_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"
                Dim xlFullPath = Path.Combine(xlDocsFolder, xlFileName)
                If savePath = "Downloads" Then
                    xlFullPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    "Downloads", xlFileName)
                ElseIf savePath = "Desktop" Then
                    xlFullPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop), xlFileName)
                End If
                Dim lowerLast = lastUserInput.ToLower()
                Dim wantsPivot As Boolean = lowerLast.Contains("pivot")
                If wantsModifyExisting AndAlso selectedFiles.Count > 0 Then
                    Await Task.Run(Function() ModifyExistingExcelViaPS(tableText, selectedFiles(0), lowerLast))
                ElseIf wantsPivot AndAlso wantsChart Then
                    Await Task.Run(Function() CreateExcelWithPivotAndChartViaPS(tableText, xlFullPath, lowerLast))
                ElseIf wantsPivot Then
                    Await Task.Run(Function() CreateExcelWithPivotViaPS(tableText, xlFullPath))
                ElseIf wantsChart Then
                    Await Task.Run(Function() CreateExcelWithChartViaPS(tableText, xlFullPath, lowerLast))
                ElseIf wantsFormulas OrElse wantsConditionalFormat Then
                    Await Task.Run(Function() CreateExcelWithExtrasViaPS(tableText, xlFullPath, lowerLast))
                Else
                    Await Task.Run(Function() CreateExcelDocumentViaPS(tableText, xlFullPath))
                End If
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("✅ Excel file saved: " & xlFullPath & Environment.NewLine)
                                                End Sub)
                Return
            End If
            If targetApp = "powerpoint" Then
                Try
                    Dim existingPptProcs = Process.GetProcessesByName("POWERPNT")
                    If existingPptProcs.Length > 0 Then
                        Dim searchFolders = {
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    }
                        For Each folder In searchFolders
                            If Not Directory.Exists(folder) Then Continue For
                            Dim recentFiles = Directory.GetFiles(folder, "*.pptx", SearchOption.TopDirectoryOnly).
                            Where(Function(f) File.GetCreationTime(f) > DateTime.Now.AddSeconds(-60)).
                            ToArray()
                            If recentFiles.Length > 0 Then
                                Dim recentFile = recentFiles.OrderByDescending(Function(f) File.GetCreationTime(f)).First()
                                ResponseRichTextBox.BeginInvoke(Sub()
                                                                    ResponseRichTextBox.AppendText("⏭️ Presentation already created recently: " & Path.GetFileName(recentFile) & Environment.NewLine)
                                                                End Sub)
                                Return
                            End If
                        Next
                    End If
                Catch
                End Try
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("📊 Creating PowerPoint presentation..." & Environment.NewLine)
                                                End Sub)
                Dim pptPrompt As String = "You are creating content for a PowerPoint presentation. " &
                "Structure your response as a series of slides using this EXACT format for each slide:" & Environment.NewLine &
                "SLIDE|<layout>|<title>" & Environment.NewLine &
                "CONTENT|<bullet point or paragraph>" & Environment.NewLine &
                "CONTENT|<bullet point or paragraph>" & Environment.NewLine &
                "NOTES|<speaker notes for this slide>" & Environment.NewLine &
                "---" & Environment.NewLine &
                "Layout options: TITLE, CONTENT, TWOCOL, BLANK" & Environment.NewLine &
                "First slide should always be TITLE layout." & Environment.NewLine &
                "Generate slides based on this content:" & Environment.NewLine &
                docContent.Substring(0, Math.Min(12000, docContent.Length))
                Dim pptStructured = Await GetSimpleAIResponseAsync(pptPrompt)
                If String.IsNullOrWhiteSpace(pptStructured) Then
                    pptStructured = "SLIDE|TITLE|Presentation" & Environment.NewLine &
                    "CONTENT|Content based on provided material" & Environment.NewLine &
                    "NOTES|Speaker notes here" & Environment.NewLine & "---"
                End If
                Dim pptDocsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                Dim pptFileName = "Presentation_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".pptx"
                Dim pptFullPath = Path.Combine(pptDocsFolder, pptFileName)
                If savePath = "Downloads" Then
                    pptFullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", pptFileName)
                ElseIf savePath = "Desktop" Then
                    pptFullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), pptFileName)
                End If
                If wantsModifyExistingPpt AndAlso selectedFiles.Count > 0 Then
                    Dim existingPptPath As String = ""
                    For Each f In selectedFiles
                        Dim ext = Path.GetExtension(f).ToLower()
                        If ext = ".pptx" OrElse ext = ".ppt" Then
                            existingPptPath = f
                            Exit For
                        End If
                    Next
                    If Not String.IsNullOrEmpty(existingPptPath) Then
                        Await Task.Run(Function() ModifyExistingPptViaPS(pptStructured, existingPptPath, lastUserInput.ToLower()))
                        ResponseRichTextBox.BeginInvoke(Sub()
                                                            ResponseRichTextBox.AppendText("✅ Presentation modified: " & existingPptPath & Environment.NewLine)
                                                        End Sub)
                        Return
                    End If
                End If
                Await Task.Run(Function() CreatePptEnhancedViaPS(pptStructured, pptFullPath, lastUserInput.ToLower()))
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("✅ Presentation saved: " & pptFullPath & Environment.NewLine)
                                                    If wantsPptPdf Then
                                                        Dim pdfSaved = Path.ChangeExtension(pptFullPath, ".pdf")
                                                        If File.Exists(pdfSaved) Then
                                                            ResponseRichTextBox.AppendText("✅ PDF exported: " & pdfSaved & Environment.NewLine)
                                                        End If
                                                    End If
                                                End Sub)
                Return
            End If
            If targetApp = "outlook" Then
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("📧 Processing Outlook request..." & Environment.NewLine)
                                                End Sub)
                Dim outlookResult = Await Task.Run(Function() HandleOutlookViaPS(docContent, lastUserInput.ToLower()))
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("✅ Outlook: " & outlookResult & Environment.NewLine)
                                                End Sub)
                Return
            End If
            If targetApp = "onenote" Then
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("📓 Creating OneNote note..." & Environment.NewLine)
                                                End Sub)
                Dim oneNotePrompt As String = "Write a clear, well-structured note about the following content. " &
                "Use headings marked with ## for sections, bullet points starting with - for lists, " &
                "and plain paragraphs for body text. Start with the main title on the first line. Content:" &
                Environment.NewLine & docContent.Substring(0, Math.Min(12000, docContent.Length))
                Dim oneNoteText = Await GetSimpleAIResponseAsync(oneNotePrompt)
                If String.IsNullOrWhiteSpace(oneNoteText) Then
                    oneNoteText = docContent.Substring(0, Math.Min(5000, docContent.Length))
                End If
                Dim oneNoteResult = Await Task.Run(Function() HandleOneNoteViaPS(oneNoteText, lastUserInput.ToLower()))
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("✅ OneNote: " & oneNoteResult & Environment.NewLine)
                                                End Sub)
                Return
            End If
            If targetApp = "access" Then
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("🗄️ Processing Access database request..." & Environment.NewLine)
                                                End Sub)
                Dim accessResult = Await Task.Run(Function() HandleAccessViaPS(docContent, lastUserInput.ToLower()))
                ResponseRichTextBox.BeginInvoke(Sub()
                                                    ResponseRichTextBox.AppendText("✅ Access: " & accessResult & Environment.NewLine)
                                                End Sub)
                Return
            End If
            Try
                Dim existingWordProcs = Process.GetProcessesByName("WINWORD")
                If existingWordProcs.Length > 0 Then
                    Dim searchFolders = {
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads"),
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                }
                    For Each folder In searchFolders
                        If Not Directory.Exists(folder) Then Continue For
                        Dim recentFiles = Directory.GetFiles(folder, "*.docx", SearchOption.TopDirectoryOnly).
                        Where(Function(f) File.GetCreationTime(f) > DateTime.Now.AddSeconds(-60)).
                        ToArray()
                        If recentFiles.Length > 0 Then
                            Dim recentFile = recentFiles.OrderByDescending(Function(f) File.GetCreationTime(f)).First()
                            ResponseRichTextBox.BeginInvoke(Sub()
                                                                ResponseRichTextBox.AppendText("⏭️ Word document already created recently: " & Path.GetFileName(recentFile) & Environment.NewLine)
                                                            End Sub)
                            Return
                        End If
                    Next
                End If
            Catch
            End Try
            ResponseRichTextBox.BeginInvoke(Sub()
                                                ResponseRichTextBox.AppendText("📝 Creating Word document from extracted content..." & Environment.NewLine)
                                            End Sub)
            Dim summaryPrompt As String = "Write a clear, well-structured summary of this document. " &
            "Use plain text paragraphs. Start with a title line. Content:" &
            Environment.NewLine & docContent.Substring(0, Math.Min(12000, docContent.Length))
            Dim summaryText = Await GetSimpleAIResponseAsync(summaryPrompt)
            If String.IsNullOrWhiteSpace(summaryText) Then
                summaryText = docContent.Substring(0, Math.Min(3000, docContent.Length))
            End If
            Dim docsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim fileName = "Summary_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".docx"
            Dim fullPath = Path.Combine(docsFolder, fileName)
            If savePath = "Downloads" Then
                fullPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Downloads", fileName)
            ElseIf savePath = "Desktop" Then
                fullPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName)
            End If
            Dim wordResult As String
            If wantsModifyExistingWord AndAlso selectedFiles.Count > 0 Then
                Dim existingDocPath As String = ""
                For Each f In selectedFiles
                    Dim ext = Path.GetExtension(f).ToLower()
                    If ext = ".docx" OrElse ext = ".doc" Then
                        existingDocPath = f
                        Exit For
                    End If
                Next
                If Not String.IsNullOrEmpty(existingDocPath) Then
                    wordResult = Await Task.Run(Function() ModifyExistingWordViaPS(summaryText, existingDocPath, lastUserInput.ToLower()))
                Else
                    wordResult = Await Task.Run(Function() CreateWordDocumentEnhancedViaPS(summaryText, fullPath, lastUserInput.ToLower()))
                End If
            Else
                wordResult = Await Task.Run(Function() CreateWordDocumentEnhancedViaPS(summaryText, fullPath, lastUserInput.ToLower()))
            End If
            Dim finalWordPath As String = fullPath
            If wantsModifyExistingWord AndAlso selectedFiles.Count > 0 Then
                For Each f In selectedFiles
                    Dim ext = Path.GetExtension(f).ToLower()
                    If ext = ".docx" OrElse ext = ".doc" Then
                        finalWordPath = f
                        Exit For
                    End If
                Next
            End If
            ResponseRichTextBox.BeginInvoke(Sub()
                                                ResponseRichTextBox.AppendText("✅ Word document saved: " & finalWordPath & Environment.NewLine)
                                                If wantsWordPdf Then
                                                    Dim pdfSaved = Path.ChangeExtension(finalWordPath, ".pdf")
                                                    If File.Exists(pdfSaved) Then
                                                        ResponseRichTextBox.AppendText("✅ PDF exported: " & pdfSaved & Environment.NewLine)
                                                    End If
                                                End If
                                            End Sub)
        Catch ex As Exception
            ResponseRichTextBox.BeginInvoke(Sub()
                                                ResponseRichTextBox.AppendText("❌ Document creation error: " & ex.Message & Environment.NewLine)
                                            End Sub)
        End Try
    End Function
    Private Function NormalizeScript(s As String) As String
        If s Is Nothing Then Return ""
        Return Regex.Replace(s, "\s+", " ").Trim()
    End Function
    Private Function GetInstalledApplications() As List(Of AppEntry)
        Dim apps As New List(Of AppEntry)
        Dim registryPaths = {"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                             "SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"}
        For Each regPath In registryPaths
            Using key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(regPath)
                If key IsNot Nothing Then
                    For Each subKeyName In key.GetSubKeyNames()
                        Using subKey = key.OpenSubKey(subKeyName)
                            Dim displayName = CStr(subKey.GetValue("DisplayName"))
                            If Not String.IsNullOrWhiteSpace(displayName) Then
                                apps.Add(New AppEntry With {.Name = displayName, .AUMID = ""})
                            End If
                        End Using
                    Next
                End If
            End Using
        Next
        Return apps
    End Function
    Private Function GetStoreApps() As List(Of AppEntry)
        Dim result As New List(Of AppEntry)
        Try
            Dim psi As New ProcessStartInfo("powershell.exe") With {
                .Arguments = "-NoProfile -ExecutionPolicy Bypass -Command ""Get-StartApps | ConvertTo-Json | Out-String""",
                .RedirectStandardOutput = True, .UseShellExecute = False, .CreateNoWindow = True
            }
            Dim proc = Process.Start(psi)
            Dim output = proc.StandardOutput.ReadToEnd().Trim()
            proc.WaitForExit(10000)
            If String.IsNullOrWhiteSpace(output) OrElse Not output.StartsWith("[") Then Return result
            Dim json = JArray.Parse(output)
            For Each item In json
                Dim name = item("Name")?.ToString()
                Dim aumid = item("AppID")?.ToString()
                If Not String.IsNullOrWhiteSpace(name) AndAlso Not String.IsNullOrWhiteSpace(aumid) Then
                    result.Add(New AppEntry With {.Name = name, .AUMID = aumid})
                End If
            Next
        Catch ex As Exception
            LogToFile("STORE_APPS_ERROR", ex.Message)
        End Try
        Return result
    End Function
    Private Async Function GetOnlineDateAsync() As Task(Of Date?)
        Try
            Using tempClient As New HttpClient() With {.Timeout = TimeSpan.FromSeconds(3)}
                Dim response = Await tempClient.SendAsync(New HttpRequestMessage(HttpMethod.Head, "https://www.oranyxlabs.com"))
                If response.Headers.Date.HasValue Then Return response.Headers.Date.Value.LocalDateTime
            End Using
        Catch
            Return Nothing
        End Try
        Return Nothing
    End Function
    Private Sub LoadExistingMacros()
        SceneListBox.Items.Clear()
        Dim scenesFolder = Path.Combine(Application.StartupPath, "Scenes")
        If Not Directory.Exists(scenesFolder) Then Directory.CreateDirectory(scenesFolder)
        For Each file In Directory.GetFiles(scenesFolder, "*")
            SceneListBox.Items.Add(Path.GetFileName(file))
        Next
    End Sub
    Private Sub RecordButton_Click(sender As Object, e As EventArgs) Handles RecordButton.Click
        If Not isMacroRecording Then
            isMacroRecording = True
            RecordButton.ForeColor = System.Drawing.Color.Red
            Dim scenesFolder = Path.Combine(Application.StartupPath, "Scenes")
            If Not Directory.Exists(scenesFolder) Then Directory.CreateDirectory(scenesFolder)
            macroSavePath = Path.Combine(scenesFolder, $"Scene_{Date.Now:yyyyMMdd_HHmmss}")
            AppendResponse($"🔴 Recording started. Logging to {macroSavePath}")
            MacroRecorder.StartRecording(macroSavePath)
            ScreenRecorder.StartRecording(macroSavePath)
        Else
            isMacroRecording = False
            RecordButton.ForeColor = System.Drawing.Color.Lime
            MacroRecorder.StopRecording(macroSavePath)
            ScreenRecorder.StopRecording()
            SceneListBox.Items.Add(Path.GetFileName(macroSavePath))
            AppendResponse($"✅ Recording saved to: {macroSavePath}")
            QueryTextBox.Text = File.ReadAllText(macroSavePath)
        End If
    End Sub
    Private Sub AddButton_Click(sender As Object, e As EventArgs) Handles AddButton.Click
        If SceneListBox.SelectedItem IsNot Nothing Then
            Dim selectedFileName = SceneListBox.SelectedItem.ToString()
            Dim scenesFolder = Path.Combine(Application.StartupPath, "Scenes")
            Dim fullPath = Path.Combine(scenesFolder, selectedFileName)
            If File.Exists(fullPath) Then
                QueryTextBox.SelectedText = File.ReadAllText(fullPath)
                QueryTextBox.Focus()
            Else
                MessageBox.Show("Scene file not found.", "Autono-Me", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("Please select a scene from the list first.", "Autono-Me", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
    Private Sub SceneListBox_DoubleClick(sender As Object, e As EventArgs) Handles SceneListBox.DoubleClick
        If SceneListBox.SelectedIndex >= 0 Then
            renameIndex = SceneListBox.SelectedIndex
            oldFileName = SceneListBox.SelectedItem.ToString()
            Dim rect = SceneListBox.GetItemRectangle(renameIndex)
            Dim screenPt = SceneListBox.PointToScreen(New Point(rect.Left, rect.Top))
            Dim parentPt = SceneListBox.Parent.PointToClient(screenPt)
            RenameTextBox.Location = parentPt
            RenameTextBox.Size = New Size(SceneListBox.ClientRectangle.Width, rect.Height)
            RenameTextBox.Text = oldFileName
            RenameTextBox.Visible = True
            RenameTextBox.BringToFront()
            RenameTextBox.Focus()
            RenameTextBox.SelectAll()
        End If
    End Sub
    Private Sub RenameTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles RenameTextBox.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            CommitRename()
        ElseIf e.KeyCode = Keys.Escape Then
            e.SuppressKeyPress = True
            RenameTextBox.Visible = False
        End If
    End Sub
    Private Sub RenameTextBox_Leave(sender As Object, e As EventArgs) Handles RenameTextBox.Leave
        CommitRename()
    End Sub
    Private Sub CommitRename()
        If Not RenameTextBox.Visible Then Return
        RenameTextBox.Visible = False
        Dim newName = RenameTextBox.Text.Trim()
        If String.IsNullOrWhiteSpace(newName) OrElse newName = oldFileName Then Return
        For Each c In Path.GetInvalidFileNameChars()
            newName = newName.Replace(c.ToString(), "")
        Next
        If String.IsNullOrWhiteSpace(newName) OrElse newName = oldFileName Then Return
        Dim scenesFolder = Path.Combine(Application.StartupPath, "Scenes")
        Dim oldFullPath = Path.Combine(scenesFolder, oldFileName)
        Dim newFullPath = Path.Combine(scenesFolder, newName)
        Try
            If File.Exists(newFullPath) Then
                MessageBox.Show("A scene with that name already exists.", "Autono-Me", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            File.Move(oldFullPath, newFullPath)
            Dim oldFramesDir = oldFullPath & "_frames"
            If Directory.Exists(oldFramesDir) Then
                Directory.Move(oldFramesDir, newFullPath & "_frames")
            End If
            If File.Exists(oldFullPath & ".mp4") Then
                File.Move(oldFullPath & ".mp4", newFullPath & ".mp4")
            End If
            For Each domFile In Directory.GetFiles(scenesFolder, oldFileName & "_DOM_*")
                Dim domNewName = domFile.Replace(oldFileName, newName)
                File.Move(domFile, domNewName)
            Next
            SceneListBox.Items(renameIndex) = newName
            AppendResponse($"✏️ Renamed scene '{oldFileName}' to '{newName}'")
        Catch ex As Exception
            MessageBox.Show("Error renaming scene: " & ex.Message, "Autono-Me", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Sub RemoveButton_Click(sender As Object, e As EventArgs) Handles RemoveButton.Click
        If SceneListBox.SelectedItem IsNot Nothing Then
            Dim selectedFileName = SceneListBox.SelectedItem.ToString()
            Dim result = MessageBox.Show($"Are you sure you want to permanently delete '{selectedFileName}'?",
                                         "Autono-Me", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            If result = DialogResult.Yes Then
                Dim scenesFolder = Path.Combine(Application.StartupPath, "Scenes")
                Dim fullPath = Path.Combine(scenesFolder, selectedFileName)
                Try
                    If File.Exists(fullPath) Then File.Delete(fullPath)
                    Dim framesDir = fullPath & "_frames"
                    If Directory.Exists(framesDir) Then Directory.Delete(framesDir, True)
                    If File.Exists(fullPath & ".mp4") Then File.Delete(fullPath & ".mp4")
                    For Each domFile In Directory.GetFiles(scenesFolder, selectedFileName & "_DOM_*")
                        File.Delete(domFile)
                    Next
                    SceneListBox.Items.Remove(SceneListBox.SelectedItem)
                    AppendResponse($"🗑️ Deleted scene: '{selectedFileName}'")
                Catch ex As Exception
                    MessageBox.Show("Error deleting scene: " & ex.Message, "Autono-Me", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            MessageBox.Show("Please select a scene from the list to delete.", "Autono-Me", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
    Private ReadOnly UIAScript As String =
        "Add-Type -AssemblyName UIAutomationClient" & Environment.NewLine &
        "Add-Type -AssemblyName UIAutomationTypes" & Environment.NewLine &
        "$target = $null" & Environment.NewLine &
        "if ($args[0] -eq 'foreground') {" & Environment.NewLine &
        "  Add-Type @'" & Environment.NewLine &
        "using System; using System.Runtime.InteropServices;" & Environment.NewLine &
        "public class WinHelper { [DllImport(""user32.dll"")] public static extern IntPtr GetForegroundWindow(); }" & Environment.NewLine &
        "'@" & Environment.NewLine &
        "  $hwnd = [WinHelper]::GetForegroundWindow()" & Environment.NewLine &
        "  $target = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)" & Environment.NewLine &
        "} else {" & Environment.NewLine &
        "  $target = [System.Windows.Automation.AutomationElement]::RootElement" & Environment.NewLine &
        "}" & Environment.NewLine &
        "function Walk($el, $depth) {" & Environment.NewLine &
        "  if ($depth -gt 7) { return }" & Environment.NewLine &
        "  try {" & Environment.NewLine &
        "    $n = $el.Current.Name; $ct = $el.Current.LocalizedControlType" & Environment.NewLine &
        "    $aid = $el.Current.AutomationId" & Environment.NewLine &
        "    $r = $el.Current.BoundingRectangle" & Environment.NewLine &
        "    $off = $el.Current.IsOffscreen" & Environment.NewLine &
        "    if (-not $off -and $n) {" & Environment.NewLine &
        "      $pad = ' ' * ($depth * 2)" & Environment.NewLine &
        "      $pos = if ($r.IsEmpty) { '' } else { ' | x:' + [int]$r.X + ' y:' + [int]$r.Y + ' w:' + [int]$r.Width + ' h:' + [int]$r.Height }" & Environment.NewLine &
        "      $idStr = if ($aid) { ' | id:' + $aid } else { '' }" & Environment.NewLine &
        "      Write-Output ($pad + $ct + ' | ' + $n + $idStr + $pos)" & Environment.NewLine &
        "    }" & Environment.NewLine &
        "  } catch {}" & Environment.NewLine &
        "  try {" & Environment.NewLine &
        "    $kids = $el.FindAll([System.Windows.Automation.TreeScope]::Children, [System.Windows.Automation.Condition]::TrueCondition)" & Environment.NewLine &
        "    foreach ($k in $kids) { Walk $k ($depth+1) }" & Environment.NewLine &
        "  } catch {}" & Environment.NewLine &
        "}" & Environment.NewLine &
        "Walk $target 0"
    Private Function GetForegroundWindowUIATree() As String
        Try
            Dim scriptFile = Path.Combine(Path.GetTempPath(), "uia_scan.ps1")
            File.WriteAllText(scriptFile, UIAScript)
            Dim psi As New ProcessStartInfo("powershell.exe",
                $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -File ""{scriptFile}"" foreground") With {
                .RedirectStandardOutput = True, .RedirectStandardError = True,
                .UseShellExecute = False, .CreateNoWindow = True
            }
            Using proc = Process.Start(psi)
                Dim output = proc.StandardOutput.ReadToEnd().Trim()
                proc.WaitForExit(15000)
                If output.Length > 10000 Then output = output.Substring(0, 10000) & "... [TRUNCATED]"
                Return If(String.IsNullOrWhiteSpace(output), "(no UIA elements found)", output)
            End Using
        Catch ex As Exception
            Return "(foreground UIA scan failed: " & ex.Message & ")"
        End Try
    End Function
    Private Async Function BringBrowserToFrontAsync() As Task
        Await Task.Run(Sub()
                           Try
                               Dim browserNames = {"chrome", "msedge", "brave", "chromium"}
                               For Each bName As String In browserNames
                                   Dim procs = Process.GetProcessesByName(bName)
                                   Dim target = procs.Where(Function(p) p.MainWindowHandle <> IntPtr.Zero).
                                   OrderByDescending(Function(p) p.StartTime).FirstOrDefault()
                                   If target IsNot Nothing Then
                                       NativeMethods.ShowWindow(target.MainWindowHandle, 9)
                                       System.Threading.Thread.Sleep(300)
                                       NativeMethods.SetForegroundWindow(target.MainWindowHandle)
                                       Exit For
                                   End If
                               Next
                           Catch ex As Exception
                               Debug.WriteLine("BringBrowserToFront failed: " & ex.Message)
                           End Try
                       End Sub)
        Await Task.Delay(500)
    End Function
    Private Async Function ShowClickIndicatorAsync(page As IPage, selector As String) As Task
        Try
            Dim script As String =
"(selector) => {
    const element = document.querySelector(selector);
    if (element) {
        const originalOutline = element.style.outline;
        element.style.outline = '3px solid red';
        const rect = element.getBoundingClientRect();
        const indicator = document.createElement('div');
        indicator.textContent = '🤖 CLICK';
        indicator.style.position = 'absolute';
        indicator.style.left = rect.left + 'px';
        indicator.style.top = (rect.top - 30) + 'px';
        indicator.style.backgroundColor = 'yellow';
        indicator.style.border = '2px solid red';
        indicator.style.padding = '4px';
        indicator.style.borderRadius = '4px';
        indicator.style.zIndex = '9999';
        indicator.style.fontWeight = 'bold';
        document.body.appendChild(indicator);
        setTimeout(() => {
            element.style.outline = originalOutline;
            indicator.remove();
        }, 1000);
    }
}"
            Await page.EvaluateAsync(script, selector)
            Await Task.Delay(500)
        Catch
        End Try
    End Function
    Private Async Function ShowScrollIndicatorAsync(page As IPage) As Task
        Try
            Dim script As String =
"const indicator = document.createElement('div');
indicator.textContent = '⬇️ SCROLLING ⬇️';
indicator.style.position = 'fixed';
indicator.style.bottom = '20px';
indicator.style.right = '20px';
indicator.style.backgroundColor = 'blue';
indicator.style.color = 'white';
indicator.style.padding = '10px';
indicator.style.borderRadius = '5px';
indicator.style.zIndex = '9999';
indicator.style.fontWeight = 'bold';
document.body.appendChild(indicator);
setTimeout(() => indicator.remove(), 3000);"
            Await page.EvaluateAsync(script)
        Catch
        End Try
    End Function
    Private Async Function GetCDPElementMapAsync(page As IPage) As Task(Of String)
        Dim fallbackNeeded As Boolean = False
        Dim axOutput As String = ""
        Try
            Dim cdp = Await page.Context.NewCDPSessionAsync(page)
            Dim axParams = New Dictionary(Of String, Object)
            Dim axResult = Await cdp.SendAsync("Accessibility.getFullAXTree", axParams)
            Try : Await cdp.DetachAsync() : Catch : End Try
            If Not axResult.HasValue Then
                fallbackNeeded = True
            Else
                Dim root = JObject.Parse(axResult.Value.ToString())
                Dim nodes = root("nodes")
                If nodes Is Nothing OrElse Not nodes.HasValues Then
                    fallbackNeeded = True
                Else
                    Dim skipRoles = {"none", "generic", "group", "region", "list",
                                     "listitem", "paragraph", "LineBreak", "StaticText", "InlineTextBox"}
                    Dim results As New List(Of String)
                    For Each node In nodes
                        Try
                            Dim role = node("role")?("value")?.ToString()
                            If String.IsNullOrEmpty(role) OrElse skipRoles.Contains(role) Then Continue For
                            Dim nameVal = If(node("name")?("value")?.ToString(), "(unnamed)")
                            Dim nodeId = node("nodeId")?.ToString()
                            results.Add(role & " | " & nameVal & " | nodeId:" & nodeId)
                            If results.Count >= 200 Then Exit For
                        Catch
                        End Try
                    Next
                    If results.Count = 0 Then
                        fallbackNeeded = True
                    Else
                        axOutput = String.Join(Environment.NewLine, results)
                        If axOutput.Length > 10000 Then axOutput = axOutput.Substring(0, 10000) & "... [TRUNCATED]"
                    End If
                End If
            End If
        Catch
            fallbackNeeded = True
        End Try
        If fallbackNeeded Then
            Return Await GetCDPElementMapFallbackAsync(page)
        End If
        Return axOutput
    End Function
    Private Async Function GetCDPElementMapFallbackAsync(page As IPage) As Task(Of String)
        Try
            Try
                Await page.WaitForLoadStateAsync(LoadState.NetworkIdle,
                    New PageWaitForLoadStateOptions With {.Timeout = 10000})
            Catch
            End Try
            Dim pollJs As String =
                "(function() {" &
                "  var sel = 'input,button,a[href],select,textarea," &
                "[contenteditable=true],[contenteditable=\""\""],[role=textbox],[role=button]," &
                "[role=link],[role=searchbox],[role=combobox],[role=menuitem],[role=option],[role=tab]';" &
                "  var found = Array.from(document.querySelectorAll(sel)).filter(function(el) {" &
                "    var r = el.getBoundingClientRect();" &
                "    return r.width > 0 && r.height > 0 && r.x >= 0 && r.y >= 0;" &
                "  });" &
                "  return found.length;" &
                "})()"
            For pollAttempt = 1 To 12
                Dim visibleCount = 0
                Try
                    visibleCount = Await page.EvaluateAsync(Of Integer)(pollJs)
                Catch
                End Try
                If visibleCount >= 3 Then Exit For
                If pollAttempt = 4 Then
                    Try
                        Await page.EvaluateAsync("window.scrollBy(0, 200)")
                        Await Task.Delay(300)
                        Await page.EvaluateAsync("window.scrollBy(0, -200)")
                    Catch
                    End Try
                End If
                Await Task.Delay(800)
            Next
            Await Task.Delay(500)
            Dim q As String = Chr(34)
            Dim scanJs As String =
                "(function() {" &
                "  var results = [];" &
                "  var seen = new WeakSet();" &
                "  function getSelector(el) {" &
                "    if (el.id && el.id.trim() && !el.id.match(/^[0-9]/)) return '#' + el.id.trim();" &
                "    var dt = el.getAttribute('data-testid');" &
                "    if (dt) return '[data-testid=' + JSON.stringify(dt) + ']';" &
                "    var al = el.getAttribute('aria-label');" &
                "    if (al && al.trim()) return '[aria-label=' + JSON.stringify(al.trim()) + ']';" &
                "    var ph = el.getAttribute('placeholder');" &
                "    if (ph && ph.trim()) return '[placeholder=' + JSON.stringify(ph.trim()) + ']';" &
                "    var nm = el.getAttribute('name');" &
                "    if (nm) return el.tagName.toLowerCase() + '[name=' + JSON.stringify(nm) + ']';" &
                "    if (el.getAttribute('contenteditable') === 'true' || el.getAttribute('contenteditable') === '') return '[contenteditable]';" &
                "    var href = el.getAttribute('href');" &
                "    if (href && el.tagName === 'A') {" &
                "      var clean = href.split('?')[0].substring(0,80);" &
                "      return 'a[href*=' + JSON.stringify(clean) + ']';" &
                "    }" &
                "    var cls = (typeof el.className === 'string') ? el.className.trim().split(/\s+/)[0] : '';" &
                "    return el.tagName.toLowerCase() + (cls ? '.' + cls : '');" &
                "  }" &
                "  function getText(el) {" &
                "    var t = el.getAttribute('aria-label') ||" &
                "            el.getAttribute('title') ||" &
                "            el.getAttribute('placeholder') ||" &
                "            el.getAttribute('alt') ||" &
                "            el.getAttribute('data-testid') || '';" &
                "    if (!t.trim()) {" &
                "      t = (el.innerText || el.value || el.textContent || '').trim();" &
                "    }" &
                "    return t.replace(/\s+/g,' ').substring(0,120);" &
                "  }" &
                "  function isVisible(el) {" &
                "    var r = el.getBoundingClientRect();" &
                "    return r.width > 0 && r.height > 0 && r.right >= 0 && r.bottom >= 0;" &
                "  }" &
                "  function addEl(el) {" &
                "    if (!el || seen.has(el)) return;" &
                "    seen.add(el);" &
                "    if (el.getAttribute('aria-hidden') === 'true') return;" &
                "    if (!isVisible(el)) return;" &
                "    var role = el.getAttribute('role') || el.tagName.toLowerCase();" &
                "    var text = getText(el);" &
                "    var sel = getSelector(el);" &
                "    var r = el.getBoundingClientRect();" &
                "    results.push(role + ' | ' + text + ' | ' + sel +" &
                "      ' | x:' + Math.round(r.x) + ' y:' + Math.round(r.y) +" &
                "      ' w:' + Math.round(r.width) + ' h:' + Math.round(r.height));" &
                "  }" &
                "  function scanNode(root) {" &
                "    if (!root) return;" &
                "    var iSel = 'input,button,a[href],select,textarea," &
                "[contenteditable=true],[contenteditable=" & q & q & "]," &
                "[role=textbox],[role=button],[role=link],[role=searchbox]," &
                "[role=combobox],[role=menuitem],[role=option],[role=tab]," &
                "[role=treeitem],[role=gridcell],[role=listitem]';" &
                "    try {" &
                "      root.querySelectorAll(iSel).forEach(function(el) {" &
                "        addEl(el);" &
                "        if (el.shadowRoot) scanNode(el.shadowRoot);" &
                "      });" &
                "    } catch(e) {}" &
                "    try {" &
                "      Array.from(root.querySelectorAll('*')).slice(0,500).forEach(function(el) {" &
                "        if (el.shadowRoot) scanNode(el.shadowRoot);" &
                "      });" &
                "    } catch(e) {}" &
                "  }" &
                "  scanNode(document);" &
                "  try {" &
                "    document.querySelectorAll('iframe').forEach(function(iframe, idx) {" &
                "      try {" &
                "        var iDoc = iframe.contentDocument || iframe.contentWindow.document;" &
                "        if (iDoc) {" &
                "          results.push('=== IFRAME ' + idx + ' ===');" &
                "          scanNode(iDoc);" &
                "        }" &
                "      } catch(e) {}" &
                "    });" &
                "  } catch(e) {}" &
                "  return results.slice(0,200).join('\n');" &
                "})()"
            Dim raw As String = ""
            For scanAttempt = 1 To 5
                Try
                    raw = Await page.EvaluateAsync(Of String)(scanJs)
                Catch
                End Try
                If Not String.IsNullOrWhiteSpace(raw) Then
                    Dim count = raw.Split({Environment.NewLine, Chr(10).ToString()}, StringSplitOptions.RemoveEmptyEntries).
                        Where(Function(l) l.Contains(" | ") AndAlso Not l.TrimStart().StartsWith("===")).Count()
                    If count >= 3 Then Exit For
                End If
                Await Task.Delay(1000)
            Next
            Return If(String.IsNullOrWhiteSpace(raw), "(no elements found)", raw)
        Catch ex As Exception
            Return "(element scan failed: " & ex.Message & ")"
        End Try
    End Function
    Private Async Function GetAccessibilitySnapshotAsync(page As IPage) As Task(Of String)
        Try
            Dim js As String =
                "(function() {" &
                "  var results = [];" &
                "  var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT, null, false);" &
                "  var count = 0;" &
                "  while (walker.nextNode() && count < 300) {" &
                "    var el = walker.currentNode;" &
                "    var r = el.getBoundingClientRect();" &
                "    if (r.width === 0 && r.height === 0) continue;" &
                "    var role = el.getAttribute('role') || el.tagName.toLowerCase();" &
                "    var name = el.getAttribute('aria-label') || el.getAttribute('title') || (el.innerText || '').trim().substring(0, 60);" &
                "    if (!name) continue;" &
                "    var id = el.id ? '#' + el.id : '';" &
                "    results.push(role + ' | ' + name + ' | ' + id + ' | x:' + Math.round(r.x) + ' y:' + Math.round(r.y));" &
                "    count++;" &
                "  }" &
                "  return results.join('\n');" &
                "})()"
            Dim rawObj = Await page.EvaluateAsync(js)
            Dim raw = If(rawObj IsNot Nothing, rawObj.ToString(), "")
            If String.IsNullOrWhiteSpace(raw) Then Return "(no accessibility data found)"
            If raw.Length > 8000 Then raw = raw.Substring(0, 8000) & "... [TRUNCATED]"
            Return raw
        Catch ex As Exception
            Return "(accessibility snapshot failed: " & ex.Message & ")"
        End Try
    End Function
    Private Async Function GetPageScreenshotBase64Async(page As IPage) As Task(Of String)
        Try
            Dim bytes = Await page.ScreenshotAsync(New PageScreenshotOptions With {
                .Type = ScreenshotType.Jpeg, .Quality = 75
            })
            Return Convert.ToBase64String(bytes)
        Catch
            Return String.Empty
        End Try
    End Function
    Private Async Function GetPageObservationAsync(page As IPage) As Task(Of (Base64 As String, Text As String))
        Dim screenshotB64 = Await GetPageScreenshotBase64Async(page)
        Dim jsMap = Await GetCDPElementMapFallbackAsync(page)
        Dim axTree = Await GetCDPElementMapAsync(page)
        Dim currentUrl = page.Url
        Dim title = Await page.TitleAsync()
        jsMap = jsMap.Replace(Chr(10), Environment.NewLine).Replace(Chr(13) & Chr(13), Chr(13))
        axTree = axTree.Replace(Chr(10), Environment.NewLine).Replace(Chr(13) & Chr(13), Chr(13))
        Dim textObs As String =
            "URL: " & currentUrl & Environment.NewLine &
            "Title: " & title & Environment.NewLine &
            Environment.NewLine &
            "=== INTERACTIVE ELEMENTS (role | text | CSS_SELECTOR | x y w h) ===" & Environment.NewLine &
            jsMap & Environment.NewLine &
            Environment.NewLine &
            "=== ACCESSIBILITY ROLES (role | name) ===" & Environment.NewLine &
            axTree
        Return (screenshotB64, textObs)
    End Function
    Private Function GetDesktopObservation() As (Base64 As String, Text As String)
        Dim screenData = CaptureScreenWithInfo()
        Dim uiaTree = GetForegroundWindowUIATree()
        Dim text As String =
            "=== DESKTOP SCREENSHOT TAKEN ===" & Environment.NewLine &
            screenData.Metadata & Environment.NewLine &
            Environment.NewLine &
            "=== FOREGROUND WINDOW UIA TREE ===" & Environment.NewLine &
            uiaTree
        Return (screenData.Base64, text)
    End Function
    Private Function GetKnownSelector(url As String, description As String,
                              Optional elementMap As String = "") As String
        Dim urlLower = If(url, "").ToLower()
        Dim descLower = If(description, "").ToLower()
        If urlLower.Contains("chatgpt.com") OrElse urlLower.Contains("chat.openai.com") Then
            If descLower.Contains("input") OrElse descLower.Contains("chat") OrElse
               descLower.Contains("message") OrElse descLower.Contains("textarea") OrElse
               descLower.Contains("text") OrElse descLower.Contains("prompt") OrElse
               descLower.Contains("type") OrElse descLower.Contains("box") Then
                Return "#prompt-textarea"
            End If
            If descLower.Contains("send") OrElse descLower.Contains("submit") Then
                Return "[data-testid=""send-button""]"
            End If
        End If
        If urlLower.Contains("youtube.com") Then
            If descLower.Contains("search") OrElse descLower.Contains("input") OrElse descLower.Contains("query") Then
                Return "input[name=""search_query""]"
            End If
            If descLower.Contains("video") OrElse descLower.Contains("result") OrElse
           descLower.Contains("title") OrElse descLower.Contains("link") OrElse
           descLower.Contains("play") Then
                If Not String.IsNullOrWhiteSpace(elementMap) Then
                    Dim ytSkipWords = {"the", "video", "result", "link", "first", "title",
                                 "whose", "contains", "that", "with", "about",
                                 "click", "play", "watch", "open", "from", "search"}
                    Dim ytKeywords = descLower.Split({" "c, ","c, "'"c, "."c},
                    StringSplitOptions.RemoveEmptyEntries) _
                    .Where(Function(w) w.Length > 2 AndAlso Not ytSkipWords.Contains(w)) _
                    .ToArray()
                    Dim ytBestSelector As String = ""
                    Dim ytBestScore As Integer = 0
                    For Each line In elementMap.Split({Environment.NewLine, Chr(10)},
                    StringSplitOptions.RemoveEmptyEntries)
                        If Not line.Contains(" | ") OrElse line.TrimStart().StartsWith("===") Then Continue For
                        Dim parts = line.Split({" | "}, StringSplitOptions.None)
                        If parts.Length < 3 Then Continue For
                        Dim role = parts(0).Trim().ToLower()
                        Dim text = If(parts.Length > 1, parts(1).Trim().ToLower(), "")
                        Dim selector = If(parts.Length > 2, parts(2).Trim(), "")
                        If Not (role = "a" OrElse role = "link" OrElse role = "yt-formatted-string") Then Continue For
                        If String.IsNullOrWhiteSpace(text) OrElse text = "(unnamed)" Then Continue For
                        If String.IsNullOrWhiteSpace(selector) Then Continue For
                        If text.Length < 10 Then Continue For
                        If selector.ToLower().Contains("#logo") OrElse
                       selector.ToLower().Contains("#guide") OrElse
                       selector.ToLower().Contains("header") Then Continue For
                        Dim ytScore As Integer = 0
                        For Each kw In ytKeywords
                            If text.Contains(kw) Then ytScore += 5
                        Next
                        If selector.Contains("/watch") OrElse selector.Contains("/shorts/") Then ytScore += 3
                        If selector.Contains("video-title") OrElse selector.Contains("#video-title") Then ytScore += 4
                        If ytScore > ytBestScore Then
                            ytBestScore = ytScore
                            ytBestSelector = selector
                        End If
                    Next
                    If ytBestScore >= 5 AndAlso Not String.IsNullOrWhiteSpace(ytBestSelector) Then
                        Return ytBestSelector
                    End If
                End If
            End If
        End If
        If urlLower.Contains("google.com") Then
            If descLower.Contains("search") OrElse descLower.Contains("input") OrElse descLower.Contains("query") Then
                Return "[name=""q""]"
            End If
        End If
        If urlLower.Contains("finance.yahoo.com") Then
            If descLower.Contains("search") Then
                Return "#yfin-usr-qry"
            End If
        End If
        If String.IsNullOrWhiteSpace(elementMap) Then Return String.Empty
        Dim d = description.ToLower()
        Dim keywords = d.Split({" "c, ","c, "/"c}, StringSplitOptions.RemoveEmptyEntries)
        Dim inputWords = {"input", "search", "query", "box", "type", "enter", "fill",
                          "write", "chat", "message", "textarea", "text", "username",
                          "password", "email", "address", "field", "prompt"}
        Dim buttonWords = {"button", "click", "submit", "send", "press", "sign", "login", "log"}
        Dim linkWords = {"link", "video", "result", "title", "article", "post",
                         "play", "open", "navigate", "visit"}
        Dim wantsInput = keywords.Any(Function(k) inputWords.Contains(k))
        Dim wantsButton = keywords.Any(Function(k) buttonWords.Contains(k))
        Dim wantsLink = keywords.Any(Function(k) linkWords.Contains(k))
        Dim blacklistedSelectors = {"#thumbnail", "#logo", "#guide", "#header", "#footer",
                                    "#avatar", "#button", "#back", "#menu", "#close", "#dismiss",
                                    "a[href=""#main""]", "a[href='#main']"}
        Dim nonInputRoles = {"button", "a", "link"}
        Dim bestSelector As String = String.Empty
        Dim bestScore As Integer = -1
        For Each line In elementMap.Split({Environment.NewLine, Chr(10).ToString()}, StringSplitOptions.RemoveEmptyEntries)
            If Not line.Contains(" | ") OrElse line.TrimStart().StartsWith("===") Then Continue For
            Dim parts = line.Split({" | "}, StringSplitOptions.None)
            If parts.Length < 3 Then Continue For
            Dim role = parts(0).Trim().ToLower()
            Dim text = If(parts.Length > 1, parts(1).Trim().ToLower(), "")
            Dim selector = If(parts.Length > 2, parts(2).Trim(), "")
            If String.IsNullOrWhiteSpace(selector) Then Continue For
            Dim sLow = selector.ToLower()
            If blacklistedSelectors.Any(Function(b) sLow = b.ToLower() OrElse sLow.StartsWith(b.ToLower())) Then Continue For
            Dim score As Integer = 0
            For Each kw In keywords
                If text.Contains(kw) Then score += 3
                If sLow.Contains(kw) Then score += 2
            Next
            If wantsInput Then
                If role = "input" OrElse role = "searchbox" OrElse role = "textbox" OrElse
                   role = "combobox" OrElse role = "textarea" OrElse
                   selector.StartsWith("input") OrElse selector.Contains("[contenteditable") OrElse
                   selector.Contains("placeholder") OrElse selector.Contains("search") Then
                    score += 8
                End If
                If nonInputRoles.Contains(role) Then score -= 10
            End If
            If wantsButton Then
                If role = "button" OrElse selector.StartsWith("button") Then score += 5
                If role = "link" OrElse role = "a" Then score -= 5
            End If
            If wantsLink Then
                If role = "link" OrElse role = "a" OrElse selector.StartsWith("a[href") Then score += 4
            End If
            If selector.StartsWith("#") Then score += 3
            If selector.Contains("[") Then score += 2
            If score > bestScore Then
                bestScore = score
                bestSelector = selector
            End If
        Next
        Return If(bestScore > 2, bestSelector, String.Empty)
    End Function
    Private Async Function AskAIForBestSelectorAsync(description As String,
                                                      obs As (Base64 As String, Text As String)) As Task(Of String)
        Try
            Dim elementMapLines = obs.Text.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
            Dim truncatedMap = String.Join(Environment.NewLine, elementMapLines.Take(80))
            Dim systemMsg =
                "You are a Playwright selector expert. Return ONE selector to locate the described element." & Environment.NewLine &
                "Pick a CSS selector from the INTERACTIVE ELEMENTS list that matches." & Environment.NewLine &
                "FALLBACK: Use Playwright role locator syntax: role=textbox[name=""Message""]" & Environment.NewLine &
                "NEVER return: #button #logo #guide #back #menu a[href=""#main""] or any navigation element." & Environment.NewLine &
                "For search/input requests ALWAYS prefer: input, [placeholder=...], [aria-label=...], [role=searchbox], [role=textbox], [role=combobox], [contenteditable]." & Environment.NewLine &
                "Return ONE line only. No explanation. No thinking."
            Dim userMsg =
                "Find the CSS selector for: " & description & Environment.NewLine &
                Environment.NewLine &
                "ELEMENT MAP:" & Environment.NewLine &
                truncatedMap
            Dim response As String
            If Not String.IsNullOrEmpty(obs.Base64) Then
                response = Await GetChatResponseWithVisionAsync(systemMsg, userMsg, obs.Base64)
            Else
                response = Await GetChatResponseAsync(systemMsg, userMsg)
            End If
            Dim parsed = StripThinkingTags(Await ParseResponseAsync(response))
            Dim bareTags = {
                "a", "p", "div", "span", "li", "ul", "ol", "section", "article",
                "main", "nav", "header", "footer", "css", "selector", "element",
                "button", "input", "textarea", "form", "body", "html",
                "#thumbnail", "thumbnail", "#button", "#logo", "#guide",
                "#back", "#menu", "#close", "a[href=""#main""]", "a[href='#main']"
            }
            For Each line In parsed.Split({Environment.NewLine, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                Dim trimmed = line.Trim().Trim("'"c).Trim("`"c)
                If String.IsNullOrWhiteSpace(trimmed) Then Continue For
                If trimmed.StartsWith("//") Then Continue For
                If trimmed.Length > 200 Then Continue For
                If bareTags.Contains(trimmed.ToLower().Trim()) Then Continue For
                Dim outsideBrackets = Regex.Replace(trimmed, "\[[^\]]*\]", "")
                Dim wordCount = outsideBrackets.Trim().Split({" "c}, StringSplitOptions.RemoveEmptyEntries).Length
                Dim looksLikeSentence = outsideBrackets.Contains(". ") OrElse
                    outsideBrackets.Contains("  ") OrElse
                    (wordCount > 4 AndAlso Not trimmed.StartsWith("role=")) OrElse
                    (wordCount > 1 AndAlso Char.IsUpper(trimmed(0)) AndAlso
                     Not trimmed.StartsWith("role=") AndAlso Not trimmed.StartsWith("text="))
                If looksLikeSentence Then Continue For
                Dim firstChar = trimmed(0)
                If firstChar = "#"c OrElse firstChar = "."c OrElse firstChar = "["c OrElse
                   firstChar = "*"c OrElse Char.IsLetter(firstChar) Then
                    If Not outsideBrackets.Contains(" ") OrElse
                       trimmed.StartsWith("text=") OrElse trimmed.StartsWith("role=") Then
                        Return trimmed
                    End If
                End If
            Next
            Return FallbackSelectorFromMap(description, truncatedMap)
        Catch
            Return String.Empty
        End Try
    End Function
    Private Function FallbackSelectorFromMap(description As String, elementMap As String) As String
        Try
            Dim keywords = description.ToLower().Split({" "c, ","c}, StringSplitOptions.RemoveEmptyEntries)
            Dim bestLine As String = ""
            Dim bestScore As Integer = 0
            For Each line In elementMap.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                Dim lower = line.ToLower()
                Dim score = keywords.Count(Function(k) lower.Contains(k))
                If score > bestScore Then
                    bestScore = score
                    bestLine = line
                End If
            Next
            If String.IsNullOrEmpty(bestLine) Then Return String.Empty
            If bestLine.ToLower().Contains("aria-hidden") Then Return String.Empty
            Dim parts = bestLine.Split({" | "}, StringSplitOptions.RemoveEmptyEntries)
            If parts.Length >= 3 Then
                Dim sel = parts(2).Trim()
                Dim badSelectors = {"#thumbnail", "#header", "#footer", "#logo"}
                If badSelectors.Contains(sel.ToLower()) Then Return String.Empty
                If sel.StartsWith("#") OrElse sel.StartsWith(".") OrElse
                   sel.StartsWith("[") OrElse Char.IsLetter(sel(0)) Then
                    Return sel
                End If
            End If
            Return String.Empty
        Catch
            Return String.Empty
        End Try
    End Function
    Private Function FindBestElementByText(elementMap As String, description As String) As String
        If String.IsNullOrWhiteSpace(elementMap) OrElse String.IsNullOrWhiteSpace(description) Then
            Return String.Empty
        End If
        Dim genericWords = {"the", "a", "an", "this", "that", "first", "second",
                            "click", "open", "find", "select", "choose",
                            "button", "link", "result", "whose", "contains",
                            "title", "with", "from", "about", "video"}
        Dim descWords = description.ToLower().
            Split({" "c, "'"c, ","c, "."c}, StringSplitOptions.RemoveEmptyEntries).
            Where(Function(w) w.Length > 2).ToArray()
        Dim contentWords = descWords.Where(Function(w) Not genericWords.Contains(w)).ToArray()
        Dim bestSelector As String = String.Empty
        Dim bestScore As Integer = 0
        For Each line In elementMap.Split({Environment.NewLine, Chr(10).ToString()}, StringSplitOptions.RemoveEmptyEntries)
            If Not line.Contains(" | ") OrElse line.TrimStart().StartsWith("===") Then Continue For
            Dim parts = line.Split({" | "}, StringSplitOptions.None)
            If parts.Length < 3 Then Continue For
            Dim role = parts(0).Trim().ToLower()
            Dim text = If(parts.Length > 1, parts(1).Trim().ToLower(), "")
            Dim selector = If(parts.Length > 2, parts(2).Trim(), "")
            If String.IsNullOrWhiteSpace(selector) Then Continue For
            If selector.ToLower().StartsWith("nodeid:") Then Continue For
            Dim score As Integer = 0
            For Each word In contentWords
                If text.Contains(word) Then score += 6
                If selector.ToLower().Contains(word) Then score += 2
            Next
            For Each word In descWords
                If genericWords.Contains(word) Then Continue For
            Next
            If {"a", "link", "button", "listitem", "gridcell", "option", "treeitem", "row", "tab"}.Contains(role) Then
                score += 2
            End If
            If selector.StartsWith("#") Then score += 1
            If selector.Contains("[") Then score += 1
            If selector.Contains("/watch") OrElse selector.Contains("video-title") Then
                score += 5
            End If
            If text.Length < 10 Then score -= 5
            If selector.ToLower().Contains("#logo") OrElse
               selector.ToLower().Contains("#guide") OrElse
               selector.ToLower().Contains("#menu") OrElse
               selector.ToLower().Contains("header") OrElse
               selector.ToLower().Contains("footer") Then
                score -= 8
            End If
            If score > bestScore Then
                bestScore = score
                bestSelector = selector
            End If
        Next
        Return If(bestScore >= 4, bestSelector, String.Empty)
    End Function
    Private Async Function AskAIForUIATargetAsync(description As String,
                                                    obs As (Base64 As String, Text As String)) As Task(Of String)
        Try
            Dim systemMsg = "You are a Windows UI automation expert. Given a screenshot and UIA tree, " &
                "return ONLY a PowerShell snippet using UIAutomation to find and click the described element. " &
                "Use AutomationId or Name from the tree. Return nothing else."
            Dim userMsg = "I need to interact with: " & description & Environment.NewLine & Environment.NewLine & obs.Text
            Dim response As String
            If Not String.IsNullOrEmpty(obs.Base64) Then
                response = Await GetChatResponseWithVisionAsync(systemMsg, userMsg, obs.Base64)
            Else
                response = Await GetChatResponseAsync(systemMsg, userMsg)
            End If
            Return Await ParseResponseAsync(response)
        Catch
            Return String.Empty
        End Try
    End Function
    Private Async Function HandleAutocompleteAsync(page As IPage, selector As String, value As String) As Task
        Dim focusDone = False
        Try
            Await page.ClickAsync(selector, New PageClickOptions With {.Timeout = 8000})
            focusDone = True
        Catch
        End Try
        If Not focusDone Then
            Try
                Await page.Locator(selector).First.ClickAsync(New LocatorClickOptions With {.Timeout = 5000})
            Catch ex As Exception
                Throw New Exception("autocomplete: could not focus input '" & selector & "': " & ex.Message)
            End Try
        End If
        Await page.Keyboard.PressAsync("Control+a")
        Await page.Keyboard.PressAsync("Delete")
        Await Task.Delay(150)
        Await page.Keyboard.TypeAsync(value, New KeyboardTypeOptions With {.Delay = 60})
        Dim suggestionSelectors = {
        ".ui-autocomplete li", ".ui-menu-item",
        "[role='listbox'] [role='option']", "[role='option']",
        ".autocomplete-suggestion", ".tt-suggestion",
        ".dropdown-menu li", "[class*='autocomplete'] li",
        "[class*='suggestion']", "[class*='dropdown'] li[data-value]"
    }
        Dim foundSelector As String = ""
        Dim waited = 0
        While waited < 8000 AndAlso String.IsNullOrEmpty(foundSelector)
            Await Task.Delay(300)
            waited += 300
            For Each sel In suggestionSelectors
                Try
                    If Await page.Locator(sel).CountAsync() > 0 Then
                        foundSelector = sel
                        Exit For
                    End If
                Catch
                End Try
            Next
        End While
        If String.IsNullOrEmpty(foundSelector) Then
            AppendResponse("   ⚠️ No suggestion list appeared — pressing Enter")
            Await page.Keyboard.PressAsync("Enter")
            Return
        End If
        Dim allItems = Await page.Locator(foundSelector).AllAsync()
        Dim matchVal = value.ToLower()
        Dim bestItem As ILocator = Nothing
        Dim bestScore = -1
        AppendResponse($"   📋 {allItems.Count} suggestions via: {foundSelector}")
        For Each item In allItems
            Try
                Dim txt = (Await item.InnerTextAsync()).Trim().ToLower()
                Dim score As Integer = 0
                If txt = matchVal Then
                    score = 100
                ElseIf txt.StartsWith(matchVal) Then
                    score = 50
                ElseIf txt.Contains(matchVal) Then
                    score = 25
                ElseIf matchVal.Contains(txt) AndAlso txt.Length > 3 Then
                    score = 10
                End If
                If score > bestScore Then
                    bestScore = score
                    bestItem = item
                End If
            Catch
            End Try
        Next
        If bestItem Is Nothing Then
            AppendResponse("   ⚠️ No matching suggestion — clicking first item")
            If allItems.Count > 0 Then
                Await allItems(0).ClickAsync(New LocatorClickOptions With {.Timeout = 5000})
            Else
                Await page.Keyboard.PressAsync("Enter")
            End If
            Return
        End If
        Dim matchedText = (Await bestItem.InnerTextAsync()).Trim()
        AppendResponse($"   ✅ Clicking suggestion: ""{matchedText}""")
        Await bestItem.ClickAsync(New LocatorClickOptions With {.Timeout = 5000})
        Await Task.Delay(500)
    End Function
    Private Async Function AutoWriteDocumentAsyncFull(docContent As String, targetApp As String, savePath As String) As Task
        Try
            If targetApp = "excel" Then
                AppendResponse("📊 Creating Excel spreadsheet from extracted content...")
                Dim tablePrompt = "Convert the following data into tab-separated tables for Excel." & Environment.NewLine &
                    "STRICT OUTPUT RULES:" & Environment.NewLine &
                    "1. If multiple items, output each as SEPARATE section starting with SHEET:Name" & Environment.NewLine &
                    "2. Use TAB CHARACTER between columns — not spaces, not commas." & Environment.NewLine &
                    "3. NO markdown, NO blank lines between rows, NO commentary." & Environment.NewLine &
                    "4. Output ALL sections completely." & Environment.NewLine &
                    Environment.NewLine &
                    "DATA TO CONVERT:" & Environment.NewLine &
                    docContent.Substring(0, Math.Min(12000, docContent.Length))
                Dim tableText = Await GetSimpleAIResponseAsync(tablePrompt)
                If String.IsNullOrWhiteSpace(tableText) Then
                    tableText = docContent.Substring(0, Math.Min(3000, docContent.Length))
                End If
                Dim xlDocsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                Dim xlFileName = "Data_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx"
                Dim xlFullPath = Path.Combine(xlDocsFolder, xlFileName)
                If savePath = "Downloads" Then
                    xlFullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", xlFileName)
                End If
                Await Task.Run(Function() CreateExcelDocumentViaPS(tableText, xlFullPath))
                AppendResponse("✅ Excel file saved: " & xlFullPath)
                Return
            End If
            AppendResponse("📝 Creating Word document from extracted content...")
            Dim summaryPrompt = "Write a clear, well-structured summary of this document. " &
                "Use plain text paragraphs. Start with a title line. Content:" &
                Environment.NewLine & docContent.Substring(0, Math.Min(12000, docContent.Length))
            Dim summaryText = Await GetSimpleAIResponseAsync(summaryPrompt)
            If String.IsNullOrWhiteSpace(summaryText) Then
                summaryText = docContent.Substring(0, Math.Min(3000, docContent.Length))
            End If
            Dim docsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim fileName = "Summary_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".docx"
            Dim fullPath = Path.Combine(docsFolder, fileName)
            If savePath = "Downloads" Then
                fullPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", fileName)
            End If
            Await Task.Run(Function() CreateWordDocumentViaPS(summaryText, fullPath))
            AppendResponse("✅ Word document saved: " & fullPath)
        Catch ex As Exception
            AppendResponse("❌ Document creation error: " & ex.Message)
        End Try
    End Function
    Private Async Function ResolveElementSelectorAsync(page As IPage, description As String,
                                                    Optional existingSelector As String = Nothing) As Task(Of String)
        Try
            Dim obs = Await GetPageObservationAsync(page)
            Dim selector = GetKnownSelector(page.Url, description, obs.Text)
            If String.IsNullOrEmpty(selector) AndAlso Not String.IsNullOrEmpty(existingSelector) Then
                Try
                    Dim count = Await page.Locator(existingSelector).CountAsync()
                    If count > 0 Then Return existingSelector
                Catch
                End Try
            End If
            If String.IsNullOrEmpty(selector) Then
                selector = Await AskAIForBestSelectorAsync(description, obs)
            End If
            If String.IsNullOrEmpty(selector) Then
                selector = FindBestElementByText(obs.Text, description)
            End If
            If Not String.IsNullOrEmpty(selector) Then
                Try
                    Dim count = Await page.Locator(selector).CountAsync()
                    If count = 0 Then
                        AppendResponse($"   ⚠️ Selector '{selector}' found 0 elements, discarding...")
                        selector = String.Empty
                    End If
                Catch
                    selector = String.Empty
                End Try
            End If
            If String.IsNullOrEmpty(selector) AndAlso Not String.IsNullOrEmpty(description) Then
                selector = "text=" & description
                AppendResponse($"   🔤 Using text fallback: {selector}")
            End If
            Return If(selector, String.Empty)
        Catch
            If Not String.IsNullOrEmpty(description) Then
                Return "text=" & description
            End If
            Return If(existingSelector, String.Empty)
        End Try
    End Function
    Private Function CreateExcelDocumentViaPS(content As String, savePath As String) As String
        Try
            If IsRecentlyCreatedFile(savePath) Then
                Return "File already exists recently: " & savePath
            End If
            If File.Exists(savePath) Then
                Dim fileAge = DateTime.Now - File.GetCreationTime(savePath)
                If fileAge.TotalSeconds < 120 Then
                    Try
                        Process.Start(savePath)
                    Catch
                    End Try
                    Return "File already exists (created " & CInt(fileAge.TotalSeconds) & "s ago): " & savePath
                End If
            End If
            Dim tempTxt = GetUniqueTempPath("excel_content", ".txt")
            File.WriteAllText(tempTxt, content, Encoding.UTF8)
            Dim script =
            "Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | ForEach-Object {" & vbLf &
            "    try { $_.CloseMainWindow() | Out-Null } catch {}" & vbLf &
            "}" & vbLf &
            "Start-Sleep -Milliseconds 800" & vbLf &
            "Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | ForEach-Object {" & vbLf &
            "    try { $_.Kill() } catch {}" & vbLf &
            "}" & vbLf &
            "Start-Sleep -Milliseconds 800" & vbLf &
            "$excel = New-Object -ComObject Excel.Application" & vbLf &
            "$excel.Visible = $true" & vbLf &
            "$wb = $excel.Workbooks.Add()" & vbLf &
            "$allLines = [System.IO.File]::ReadAllLines('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)" & vbLf &
            "$sections = [System.Collections.Generic.List[hashtable]]::new()" & vbLf &
            "$currentSection = $null" & vbLf &
            "foreach ($line in $allLines) {" & vbLf &
            "    if ($line -match '^SHEET:(.+)$') {" & vbLf &
            "        $sheetName = $Matches[1].Trim()" & vbLf &
            "        $currentSection = @{ Name = $sheetName; Lines = [System.Collections.Generic.List[string]]::new() }" & vbLf &
            "        $sections.Add($currentSection)" & vbLf &
            "    } elseif ($currentSection -ne $null) {" & vbLf &
            "        $currentSection.Lines.Add($line)" & vbLf &
            "    } else {" & vbLf &
            "        if ($sections.Count -eq 0) {" & vbLf &
            "            $currentSection = @{ Name = 'Sheet1'; Lines = [System.Collections.Generic.List[string]]::new() }" & vbLf &
            "            $sections.Add($currentSection)" & vbLf &
            "        }" & vbLf &
            "        $sections[0].Lines.Add($line)" & vbLf &
            "    }" & vbLf &
            "}" & vbLf &
            "$sheetIndex = 1" & vbLf &
            "foreach ($section in $sections) {" & vbLf &
            "    if ($sheetIndex -le $wb.Worksheets.Count) {" & vbLf &
            "        $ws = $wb.Worksheets.Item($sheetIndex)" & vbLf &
            "    } else {" & vbLf &
            "        $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))" & vbLf &
            "    }" & vbLf &
            "    try { $ws.Name = $section.Name } catch {}" & vbLf &
            "    $row = 1" & vbLf &
            "    foreach ($line in $section.Lines) {" & vbLf &
            "        if ([string]::IsNullOrWhiteSpace($line)) { continue }" & vbLf &
            "        $cols = $line.Split([char]9)" & vbLf &
            "        $col = 1" & vbLf &
            "        foreach ($cell in $cols) {" & vbLf &
            "            $val = $cell.Trim()" & vbLf &
            "            $ws.Cells.Item($row, $col).Value2 = $val" & vbLf &
            "            $col++" & vbLf &
            "        }" & vbLf &
            "        $row++" & vbLf &
            "    }" & vbLf &
            "    $ws.Rows.Item(1).Font.Bold = $true" & vbLf &
            "    $ws.Columns.AutoFit() | Out-Null" & vbLf &
            "    try {" & vbLf &
            "        $range = $ws.Range('A1').CurrentRegion" & vbLf &
            "        if ($range.Rows.Count -gt 1) {" & vbLf &
            "            $ws.ListObjects.Add(1, $range, $null, 1) | Out-Null" & vbLf &
            "        }" & vbLf &
            "    } catch {}" & vbLf &
            "    $sheetIndex++" & vbLf &
            "}" & vbLf &
            "$excel.DisplayAlerts = $false" & vbLf &
            "while ($wb.Worksheets.Count -gt $sections.Count) {" & vbLf &
            "    $wb.Worksheets.Item($wb.Worksheets.Count).Delete()" & vbLf &
            "}" & vbLf &
            "$excel.DisplayAlerts = $true" & vbLf &
            "$wb.Worksheets.Item(1).Activate()" & vbLf &
            "$wb.SaveAs('" & savePath.Replace("'", "''") & "', 51)" & vbLf &
            "Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue" & vbLf &
            "$excel.WindowState = -4137" & vbLf &
            "try { $excel.ActiveWindow.WindowState = -4137 } catch {}" & vbLf &
            "Write-Output ('Saved to: ' + '" & savePath & "')"
            Return RunPowerShellSync(script)
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function CreateExcelWithPivotViaPS(content As String, savePath As String) As String
        Try
            If IsRecentlyCreatedFile(savePath) Then
                Return "File already exists recently: " & savePath
            End If
            Dim tempTxt = GetUniqueTempPath("excel_pivot_", ".txt")
            File.WriteAllText(tempTxt, content, Encoding.UTF8)
            Dim script =
            "Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | ForEach-Object {" & vbLf &
            "    try { $_.CloseMainWindow() | Out-Null } catch {}" & vbLf &
            "}" & vbLf &
            "Start-Sleep -Milliseconds 800" & vbLf &
            "Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | ForEach-Object {" & vbLf &
            "    try { $_.Kill() } catch {}" & vbLf &
            "}" & vbLf &
            "Start-Sleep -Milliseconds 800" & vbLf &
            "$excel = New-Object -ComObject Excel.Application" & vbLf &
            "$excel.Visible = $true" & vbLf &
            "$wb = $excel.Workbooks.Add()" & vbLf &
            "$allLines = [System.IO.File]::ReadAllLines('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)" & vbLf &
            "$sections = [System.Collections.Generic.List[hashtable]]::new()" & vbLf &
            "$currentSection = $null" & vbLf &
            "foreach ($line in $allLines) {" & vbLf &
            "    if ($line -match '^SHEET:(.+)$') {" & vbLf &
            "        $sheetName = $Matches[1].Trim()" & vbLf &
            "        $currentSection = @{ Name = $sheetName; Lines = [System.Collections.Generic.List[string]]::new() }" & vbLf &
            "        $sections.Add($currentSection)" & vbLf &
            "    } elseif ($currentSection -ne $null) {" & vbLf &
            "        $currentSection.Lines.Add($line)" & vbLf &
            "    } else {" & vbLf &
            "        if ($sections.Count -eq 0) {" & vbLf &
            "            $currentSection = @{ Name = 'Data'; Lines = [System.Collections.Generic.List[string]]::new() }" & vbLf &
            "            $sections.Add($currentSection)" & vbLf &
            "        }" & vbLf &
            "        $sections[0].Lines.Add($line)" & vbLf &
            "    }" & vbLf &
            "}" & vbLf &
            "$sheetIndex = 1" & vbLf &
            "$pivotSheets = [System.Collections.Generic.List[object]]::new()" & vbLf &
            "foreach ($section in $sections) {" & vbLf &
            "    if ($sheetIndex -le $wb.Worksheets.Count) {" & vbLf &
            "        $ws = $wb.Worksheets.Item($sheetIndex)" & vbLf &
            "    } else {" & vbLf &
            "        $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))" & vbLf &
            "    }" & vbLf &
            "    try { $ws.Name = $section.Name + '_Data' } catch {}" & vbLf &
            "    $row = 1" & vbLf &
            "    $colCount = 0" & vbLf &
            "    foreach ($line in $section.Lines) {" & vbLf &
            "        if ([string]::IsNullOrWhiteSpace($line)) { continue }" & vbLf &
            "        $cols = $line.Split([char]9)" & vbLf &
            "        if ($cols.Count -gt $colCount) { $colCount = $cols.Count }" & vbLf &
            "        $col = 1" & vbLf &
            "        foreach ($cell in $cols) {" & vbLf &
            "            $ws.Cells.Item($row, $col).Value2 = $cell.Trim()" & vbLf &
            "            $col++" & vbLf &
            "        }" & vbLf &
            "        $row++" & vbLf &
            "    }" & vbLf &
            "    $ws.Rows.Item(1).Font.Bold = $true" & vbLf &
            "    $ws.Columns.AutoFit() | Out-Null" & vbLf &
            "    $dataRange = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($row-1, $colCount))" & vbLf &
            "    $pivotSheets.Add(@{ DataWs = $ws; DataRange = $dataRange; Name = $section.Name; RowCount = $row-1 })" & vbLf &
            "    $sheetIndex++" & vbLf &
            "}" & vbLf &
            "foreach ($ps in $pivotSheets) {" & vbLf &
            "    try {" & vbLf &
            "        $pivotWs = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))" & vbLf &
            "        $pivotWs.Name = $ps.Name + '_Pivot'" & vbLf &
            "        $pivotCache = $wb.PivotCaches().Create(1, $ps.DataRange)" & vbLf &
            "        $pivotTable = $pivotCache.CreatePivotTable($pivotWs.Range('A3'), $ps.Name + '_PT')" & vbLf &
            "        $pivotTable.ManualUpdate = $true" & vbLf &
            "        $headers = @()" & vbLf &
            "        $headerRow = $ps.DataWs.Rows.Item(1)" & vbLf &
            "        for ($c = 1; $c -le $ps.DataRange.Columns.Count; $c++) {" & vbLf &
            "            $h = $ps.DataWs.Cells.Item(1, $c).Value2" & vbLf &
            "            if ($h) { $headers += $h }" & vbLf &
            "        }" & vbLf &
            "        if ($headers -contains 'Date') {" & vbLf &
            "            $dateField = $pivotTable.PivotFields('Date')" & vbLf &
            "            $dateField.Orientation = 1" & vbLf &
            "            $dateField.Position = 1" & vbLf &
            "        }" & vbLf &
            "        $valueFields = @('Close', 'Adj Close', 'Open', 'High', 'Low', 'Volume')" & vbLf &
            "        foreach ($vf in $valueFields) {" & vbLf &
            "            if ($headers -contains $vf) {" & vbLf &
            "                $field = $pivotTable.PivotFields($vf)" & vbLf &
            "                $field.Orientation = 4" & vbLf &
            "                $field.Function = -4106" & vbLf &
            "                $field.NumberFormat = '#,##0.00'" & vbLf &
            "            }" & vbLf &
            "        }" & vbLf &
            "        if ($headers -contains 'Volume') {" & vbLf &
            "            try {" & vbLf &
            "                $volField = $pivotTable.PivotFields('Volume')" & vbLf &
            "                $volField.NumberFormat = '#,##0'" & vbLf &
            "            } catch {}" & vbLf &
            "        }" & vbLf &
            "        $pivotTable.ManualUpdate = $false" & vbLf &
            "        $pivotTable.RefreshTable() | Out-Null" & vbLf &
            "        $pivotWs.Columns.AutoFit() | Out-Null" & vbLf &
            "    } catch {" & vbLf &
            "        Write-Output ('Pivot error for ' + $ps.Name + ': ' + $_.Exception.Message)" & vbLf &
            "    }" & vbLf &
            "}" & vbLf &
            "$excel.DisplayAlerts = $false" & vbLf &
            "while ($wb.Worksheets.Count -gt ($sections.Count * 2)) {" & vbLf &
            "    $wb.Worksheets.Item($wb.Worksheets.Count).Delete()" & vbLf &
            "}" & vbLf &
            "$excel.DisplayAlerts = $true" & vbLf &
            "$wb.Worksheets.Item(1).Activate()" & vbLf &
            "$wb.SaveAs('" & savePath.Replace("'", "''") & "', 51)" & vbLf &
            "Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue" & vbLf &
            "$excel.WindowState = -4137" & vbLf &
            "Write-Output ('Saved to: ' + '" & savePath & "')"
            Return RunPowerShellSync(script)
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function CreateExcelWithChartViaPS(content As String, savePath As String, userIntent As String) As String
        Try
            If IsRecentlyCreatedFile(savePath) Then
                Return "File already exists recently: " & savePath
            End If
            Dim tempTxt = GetUniqueTempPath("excel_chart_", ".txt")
            File.WriteAllText(tempTxt, content, Encoding.UTF8)
            Dim chartType = "4"
            If userIntent.Contains("bar") OrElse userIntent.Contains("column") Then chartType = "51"
            If userIntent.Contains("pie") Then chartType = "5"
            If userIntent.Contains("area") Then chartType = "1"
            If userIntent.Contains("scatter") Then chartType = "74"
            Dim script =
            "Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Kill() } catch {} }" & vbLf &
            "Start-Sleep -Milliseconds 800" & vbLf &
            "$excel = New-Object -ComObject Excel.Application" & vbLf &
            "$excel.Visible = $true" & vbLf &
            "$wb = $excel.Workbooks.Add()" & vbLf &
            "$allLines = [System.IO.File]::ReadAllLines('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)" & vbLf &
            "$sections = [System.Collections.Generic.List[hashtable]]::new()" & vbLf &
            "$currentSection = $null" & vbLf &
            "foreach ($line in $allLines) {" & vbLf &
            "    if ($line -match '^SHEET:(.+)$') {" & vbLf &
            "        $currentSection = @{ Name = $Matches[1].Trim(); Lines = [System.Collections.Generic.List[string]]::new() }" & vbLf &
            "        $sections.Add($currentSection)" & vbLf &
            "    } elseif ($currentSection -ne $null) { $currentSection.Lines.Add($line) }" & vbLf &
            "    else {" & vbLf &
            "        if ($sections.Count -eq 0) {" & vbLf &
            "            $currentSection = @{ Name = 'Data'; Lines = [System.Collections.Generic.List[string]]::new() }" & vbLf &
            "            $sections.Add($currentSection) }" & vbLf &
            "        $sections[0].Lines.Add($line) }" & vbLf &
            "}" & vbLf &
            "$sheetIndex = 1" & vbLf &
            "foreach ($section in $sections) {" & vbLf &
            "    if ($sheetIndex -le $wb.Worksheets.Count) { $ws = $wb.Worksheets.Item($sheetIndex) }" & vbLf &
            "    else { $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count)) }" & vbLf &
            "    try { $ws.Name = $section.Name } catch {}" & vbLf &
            "    $row = 1; $colCount = 0" & vbLf &
            "    foreach ($line in $section.Lines) {" & vbLf &
            "        if ([string]::IsNullOrWhiteSpace($line)) { continue }" & vbLf &
            "        $cols = $line.Split([char]9)" & vbLf &
            "        if ($cols.Count -gt $colCount) { $colCount = $cols.Count }" & vbLf &
            "        $col = 1" & vbLf &
            "        foreach ($cell in $cols) { $ws.Cells.Item($row, $col).Value2 = $cell.Trim(); $col++ }" & vbLf &
            "        $row++" & vbLf &
            "    }" & vbLf &
            "    $ws.Rows.Item(1).Font.Bold = $true" & vbLf &
            "    $ws.Columns.AutoFit() | Out-Null" & vbLf &
            "    $dataRange = $ws.Range($ws.Cells.Item(1,1), $ws.Cells.Item($row-1, $colCount))" & vbLf &
            "    try {" & vbLf &
            "        $chartObj = $ws.ChartObjects().Add(50, ($row * 15) + 20, 600, 300)" & vbLf &
            "        $chart = $chartObj.Chart" & vbLf &
            "        $chart.SetSourceData($dataRange)" & vbLf &
            "        $chart.ChartType = " & chartType & vbLf &
            "        $chart.HasTitle = $true" & vbLf &
            "        $chart.ChartTitle.Text = $section.Name + ' Chart'" & vbLf &
            "        $chart.HasLegend = $true" & vbLf &
            "        try { $chart.Axes(1).HasTitle = $true; $chart.Axes(1).AxisTitle.Text = 'Date' } catch {}" & vbLf &
            "        try { $chart.Axes(2).HasTitle = $true; $chart.Axes(2).AxisTitle.Text = 'Price' } catch {}" & vbLf &
            "        $chart.PlotBy = 2" & vbLf &
            "    } catch { Write-Output ('Chart error: ' + $_.Exception.Message) }" & vbLf &
            "    $sheetIndex++" & vbLf &
            "}" & vbLf &
            "$excel.DisplayAlerts = $false" & vbLf &
            "while ($wb.Worksheets.Count -gt $sections.Count) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }" & vbLf &
            "$excel.DisplayAlerts = $true" & vbLf &
            "$wb.SaveAs('" & savePath.Replace("'", "''") & "', 51)" & vbLf &
            "Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue" & vbLf &
            "$excel.WindowState = -4137" & vbLf &
            "Write-Output ('Saved to: ' + '" & savePath & "')"
            Return RunPowerShellSync(script)
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function CreateExcelWithPivotAndChartViaPS(content As String, savePath As String, userIntent As String) As String
        Try
            If IsRecentlyCreatedFile(savePath) Then
                Return "File already exists recently: " & savePath
            End If
            Dim pivotResult = CreateExcelWithPivotViaPS(content, savePath)
            If Not pivotResult.Contains("Saved to") Then Return pivotResult
            Dim script =
            "Start-Sleep -Milliseconds 500" & vbLf &
            "$excel = New-Object -ComObject Excel.Application" & vbLf &
            "$excel.Visible = $true" & vbLf &
            "$wb = $excel.Workbooks.Open('" & savePath.Replace("'", "''") & "')" & vbLf &
            "foreach ($ws in $wb.Worksheets) {" & vbLf &
            "    if ($ws.Name -like '*_Pivot') {" & vbLf &
            "        try {" & vbLf &
            "            $usedRange = $ws.UsedRange" & vbLf &
            "            $lastRow = $usedRange.Rows.Count + $usedRange.Row" & vbLf &
            "            $chartObj = $ws.ChartObjects().Add(50, ($lastRow * 15) + 30, 600, 300)" & vbLf &
            "            $chart = $chartObj.Chart" & vbLf &
            "            $chart.SetSourceData($usedRange)" & vbLf &
            "            $chart.ChartType = 4" & vbLf &
            "            $chart.HasTitle = $true" & vbLf &
            "            $chart.ChartTitle.Text = $ws.Name + ' Chart'" & vbLf &
            "        } catch {}" & vbLf &
            "    }" & vbLf &
            "}" & vbLf &
            "$wb.Save()" & vbLf &
            "$excel.WindowState = -4137" & vbLf &
            "Write-Output ('Saved to: ' + '" & savePath & "')"
            Return RunPowerShellSync(script)
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function CreateExcelWithExtrasViaPS(content As String, savePath As String, userIntent As String) As String
        Try
            If IsRecentlyCreatedFile(savePath) Then
                Return "File already exists recently: " & savePath
            End If
            Dim tempTxt = GetUniqueTempPath("excel_extras_", ".txt")
            File.WriteAllText(tempTxt, content, Encoding.UTF8)
            Dim addPctChange = wantsFormulas AndAlso (userIntent.Contains("percent") OrElse userIntent.Contains("change") OrElse userIntent.Contains("difference"))
            Dim addAverage = wantsFormulas AndAlso (userIntent.Contains("average") OrElse userIntent.Contains("mean"))
            Dim addSum = wantsFormulas AndAlso (userIntent.Contains("sum") OrElse userIntent.Contains("total"))
            Dim addHighlight = wantsConditionalFormat
            Dim highlightDrops = userIntent.Contains("drop") OrElse userIntent.Contains("decrease") OrElse userIntent.Contains("fell") OrElse userIntent.Contains("red")
            Dim highlightGains = userIntent.Contains("gain") OrElse userIntent.Contains("increase") OrElse userIntent.Contains("rose") OrElse userIntent.Contains("green")
            Dim script =
            "Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Kill() } catch {} }" & vbLf &
            "Start-Sleep -Milliseconds 800" & vbLf &
            "$excel = New-Object -ComObject Excel.Application" & vbLf &
            "$excel.Visible = $true" & vbLf &
            "$wb = $excel.Workbooks.Add()" & vbLf &
            "$allLines = [System.IO.File]::ReadAllLines('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)" & vbLf &
            "$sections = [System.Collections.Generic.List[hashtable]]::new()" & vbLf &
            "$currentSection = $null" & vbLf &
            "foreach ($line in $allLines) {" & vbLf &
            "    if ($line -match '^SHEET:(.+)$') {" & vbLf &
            "        $currentSection = @{ Name = $Matches[1].Trim(); Lines = [System.Collections.Generic.List[string]]::new() }" & vbLf &
            "        $sections.Add($currentSection)" & vbLf &
            "    } elseif ($currentSection -ne $null) { $currentSection.Lines.Add($line) }" & vbLf &
            "    else {" & vbLf &
            "        if ($sections.Count -eq 0) {" & vbLf &
            "            $currentSection = @{ Name = 'Data'; Lines = [System.Collections.Generic.List[string]]::new() }" & vbLf &
            "            $sections.Add($currentSection) }" & vbLf &
            "        $sections[0].Lines.Add($line) }" & vbLf &
            "}" & vbLf &
            "$sheetIndex = 1" & vbLf &
            "foreach ($section in $sections) {" & vbLf &
            "    if ($sheetIndex -le $wb.Worksheets.Count) { $ws = $wb.Worksheets.Item($sheetIndex) }" & vbLf &
            "    else { $ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count)) }" & vbLf &
            "    try { $ws.Name = $section.Name } catch {}" & vbLf &
            "    $row = 1; $colCount = 0; $headers = @()" & vbLf &
            "    foreach ($line in $section.Lines) {" & vbLf &
            "        if ([string]::IsNullOrWhiteSpace($line)) { continue }" & vbLf &
            "        $cols = $line.Split([char]9)" & vbLf &
            "        if ($cols.Count -gt $colCount) { $colCount = $cols.Count }" & vbLf &
            "        if ($row -eq 1) { $headers = $cols | ForEach-Object { $_.Trim() } }" & vbLf &
            "        $col = 1" & vbLf &
            "        foreach ($cell in $cols) { $ws.Cells.Item($row, $col).Value2 = $cell.Trim(); $col++ }" & vbLf &
            "        $row++" & vbLf &
            "    }" & vbLf &
            "    $ws.Rows.Item(1).Font.Bold = $true" & vbLf &
            "    $lastDataRow = $row - 1" & vbLf &
            "    $closeColIdx = -1" & vbLf &
            "    for ($c = 0; $c -lt $headers.Count; $c++) {" & vbLf &
            "        if ($headers[$c] -eq 'Close' -or $headers[$c] -eq 'Adj Close') { $closeColIdx = $c + 1; break }" & vbLf &
            "    }" & vbLf &
            If(addPctChange,
            "    if ($closeColIdx -gt 0) {" & vbLf &
            "        $pctCol = $colCount + 1" & vbLf &
            "        $ws.Cells.Item(1, $pctCol).Value2 = '% Change'" & vbLf &
            "        $ws.Cells.Item(1, $pctCol).Font.Bold = $true" & vbLf &
            "        for ($r = 3; $r -le $lastDataRow; $r++) {" & vbLf &
            "            $closeRef = [char](64 + $closeColIdx)" & vbLf &
            "            $formula = '=(' + $closeRef + $r + '-' + $closeRef + ($r-1) + ')/' + $closeRef + ($r-1) + '*100'" & vbLf &
            "            $ws.Cells.Item($r, $pctCol).Formula = $formula" & vbLf &
            "            $ws.Cells.Item($r, $pctCol).NumberFormat = '0.00""%%""'" & vbLf &
            "        }" & vbLf &
            "        $colCount = $pctCol" & vbLf &
            "    }" & vbLf, "") &
            If(addAverage,
            "    if ($closeColIdx -gt 0) {" & vbLf &
            "        $avgRow = $lastDataRow + 2" & vbLf &
            "        $closeRef = [char](64 + $closeColIdx)" & vbLf &
            "        $ws.Cells.Item($avgRow, $closeColIdx - 1).Value2 = 'Average:'" & vbLf &
            "        $ws.Cells.Item($avgRow, $closeColIdx - 1).Font.Bold = $true" & vbLf &
            "        $ws.Cells.Item($avgRow, $closeColIdx).Formula = '=AVERAGE(' + $closeRef + '2:' + $closeRef + $lastDataRow + ')'" & vbLf &
            "        $ws.Cells.Item($avgRow, $closeColIdx).NumberFormat = '#,##0.00'" & vbLf &
            "    }" & vbLf, "") &
            If(addSum,
            "    if ($closeColIdx -gt 0) {" & vbLf &
            "        $sumRow = $lastDataRow + 3" & vbLf &
            "        $closeRef = [char](64 + $closeColIdx)" & vbLf &
            "        $ws.Cells.Item($sumRow, $closeColIdx - 1).Value2 = 'Total:'" & vbLf &
            "        $ws.Cells.Item($sumRow, $closeColIdx - 1).Font.Bold = $true" & vbLf &
            "        $ws.Cells.Item($sumRow, $closeColIdx).Formula = '=SUM(' + $closeRef + '2:' + $closeRef + $lastDataRow + ')'" & vbLf &
            "        $ws.Cells.Item($sumRow, $closeColIdx).NumberFormat = '#,##0.00'" & vbLf &
            "    }" & vbLf, "") &
            If(addHighlight,
            "    if ($closeColIdx -gt 0) {" & vbLf &
            "        $dataRangeAddr = [char](64 + $closeColIdx) + '2:' + [char](64 + $closeColIdx) + $lastDataRow" & vbLf &
            "        $cfRange = $ws.Range($dataRangeAddr)" & vbLf &
            If(highlightDrops OrElse (Not highlightGains),
            "        $fc1 = $cfRange.FormatConditions.Add(2, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value, [System.Reflection.Missing]::Value)" & vbLf &
            "        try {" & vbLf &
            "            $pctColAddr = [char](64 + $colCount) + '2:' + [char](64 + $colCount) + $lastDataRow" & vbLf &
            "            $fcDrop = $cfRange.FormatConditions.Add(1, 6, 0)" & vbLf &
            "            $fcDrop.Interior.Color = 16711680" & vbLf &
            "        } catch {}" & vbLf, "") &
            If(highlightGains OrElse (Not highlightDrops),
            "        try {" & vbLf &
            "            $fcGain = $cfRange.FormatConditions.Add(1, 5, 0)" & vbLf &
            "            $fcGain.Interior.Color = 5287936" & vbLf &
            "        } catch {}" & vbLf, "") &
            "    }" & vbLf, "") &
            "    $ws.Columns.AutoFit() | Out-Null" & vbLf &
            "    $sheetIndex++" & vbLf &
            "}" & vbLf &
            "$excel.DisplayAlerts = $false" & vbLf &
            "while ($wb.Worksheets.Count -gt $sections.Count) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }" & vbLf &
            "$excel.DisplayAlerts = $true" & vbLf &
            "$wb.SaveAs('" & savePath.Replace("'", "''") & "', 51)" & vbLf &
            "Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue" & vbLf &
            "$excel.WindowState = -4137" & vbLf &
            "Write-Output ('Saved to: ' + '" & savePath & "')"
            Return RunPowerShellSync(script)
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function ModifyExistingExcelViaPS(newContent As String, existingPath As String, userIntent As String) As String
        Try
            Dim addSheet = userIntent.Contains("add") AndAlso (userIntent.Contains("sheet") OrElse userIntent.Contains("tab"))
            Dim addChart = userIntent.Contains("chart") OrElse userIntent.Contains("graph")
            Dim addPivot = userIntent.Contains("pivot")
            Dim tempTxt = GetUniqueTempPath("excel_modify_", ".txt")
            If Not String.IsNullOrWhiteSpace(newContent) Then
                File.WriteAllText(tempTxt, newContent, Encoding.UTF8)
            End If
            Dim script =
            "Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Kill() } catch {} }" & vbLf &
            "Start-Sleep -Milliseconds 800" & vbLf &
            "$excel = New-Object -ComObject Excel.Application" & vbLf &
            "$excel.Visible = $true" & vbLf &
            "$wb = $excel.Workbooks.Open('" & existingPath.Replace("'", "''") & "')" & vbLf &
            If(addSheet AndAlso Not String.IsNullOrWhiteSpace(newContent),
            "$newWs = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))" & vbLf &
            "$newWs.Name = 'New Data'" & vbLf &
            "$allLines = [System.IO.File]::ReadAllLines('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)" & vbLf &
            "$row = 1" & vbLf &
            "foreach ($line in $allLines) {" & vbLf &
            "    if ([string]::IsNullOrWhiteSpace($line)) { continue }" & vbLf &
            "    $cols = $line.Split([char]9); $col = 1" & vbLf &
            "    foreach ($cell in $cols) { $newWs.Cells.Item($row, $col).Value2 = $cell.Trim(); $col++ }" & vbLf &
            "    $row++" & vbLf &
            "}" & vbLf &
            "$newWs.Rows.Item(1).Font.Bold = $true" & vbLf &
            "$newWs.Columns.AutoFit() | Out-Null" & vbLf, "") &
            If(addChart,
            "foreach ($ws in $wb.Worksheets) {" & vbLf &
            "    $used = $ws.UsedRange" & vbLf &
            "    if ($used.Rows.Count -lt 2) { continue }" & vbLf &
            "    try {" & vbLf &
            "        $co = $ws.ChartObjects().Add(50, ($used.Rows.Count * 15) + 30, 600, 300)" & vbLf &
            "        $co.Chart.SetSourceData($used)" & vbLf &
            "        $co.Chart.ChartType = 4" & vbLf &
            "        $co.Chart.HasTitle = $true" & vbLf &
            "        $co.Chart.ChartTitle.Text = $ws.Name + ' Chart'" & vbLf &
            "    } catch {}" & vbLf &
            "}" & vbLf, "") &
            If(addPivot,
            "$firstWs = $wb.Worksheets.Item(1)" & vbLf &
            "$pivotWs = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))" & vbLf &
            "$pivotWs.Name = 'Pivot'" & vbLf &
            "try {" & vbLf &
            "    $cache = $wb.PivotCaches().Create(1, $firstWs.UsedRange)" & vbLf &
            "    $pt = $cache.CreatePivotTable($pivotWs.Range('A3'), 'PivotTable1')" & vbLf &
            "    $pt.ManualUpdate = $false" & vbLf &
            "} catch { Write-Output ('Pivot error: ' + $_.Exception.Message) }" & vbLf, "") &
            "$wb.Save()" & vbLf &
            "if (Test-Path '" & tempTxt.Replace("'", "''") & "') { Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue }" & vbLf &
            "$excel.WindowState = -4137" & vbLf &
            "Write-Output ('Modified: ' + '" & existingPath & "')"
            Return RunPowerShellSync(script)
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function CreateWordDocumentViaPS(content As String, savePath As String) As String
        Try
            Dim tempTxt = GetUniqueTempPath("doc_content_", ".txt")
            File.WriteAllText(tempTxt, content, Encoding.UTF8)
            Dim script =
                "$word = New-Object -ComObject Word.Application" & vbLf &
                "$word.Visible = $true" & vbLf &
                "$doc = $word.Documents.Add()" & vbLf &
                "$sel = $word.Selection" & vbLf &
                "try { $sel.Style = $doc.Styles.Item('Heading 1') } catch {}" & vbLf &
                "$sel.TypeText('Document Summary')" & vbLf &
                "$sel.TypeParagraph()" & vbLf &
                "try { $sel.Style = $doc.Styles.Item('Normal') } catch {}" & vbLf &
                "$body = [System.IO.File]::ReadAllText('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)" & vbLf &
                "$sel.TypeText($body)" & vbLf &
                "$doc.SaveAs2('" & savePath.Replace("'", "''") & "')" & vbLf &
                "Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue" & vbLf &
                "try { $word.ActiveWindow.WindowState = 1 } catch {}" & vbLf &
                "Write-Output ('Saved to: ' + '" & savePath & "')"
            Return RunPowerShellSync(script)
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function CreateWordDocumentEnhancedViaPS(content As String, savePath As String, userIntent As String) As String
        Try
            Dim tempTxt = GetUniqueTempPath("doc_enhanced_", ".txt")
            File.WriteAllText(tempTxt, content, Encoding.UTF8)
            Dim addToC = wantsWordToC
            Dim addTable = wantsWordTable
            Dim saveAsPdf = wantsWordPdf
            Dim pdfPath = Path.ChangeExtension(savePath, ".pdf")
            Dim titleText = "Document"
            Dim lines = content.Split({Environment.NewLine, Chr(10)}, StringSplitOptions.RemoveEmptyEntries)
            If lines.Length > 0 Then titleText = lines(0).Trim().TrimEnd(":"c).Trim()
            If titleText.Length > 100 Then titleText = titleText.Substring(0, 100)
            titleText = titleText.Replace("'", "''")
            Dim sb As New StringBuilder()
            sb.AppendLine("Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue | ForEach-Object { try { $_.CloseMainWindow() | Out-Null } catch {} }")
            sb.AppendLine("Start-Sleep -Milliseconds 800")
            sb.AppendLine("$word = New-Object -ComObject Word.Application")
            sb.AppendLine("$word.Visible = $true")
            sb.AppendLine("$doc = $word.Documents.Add()")
            sb.AppendLine("$sel = $word.Selection")
            sb.AppendLine("$doc.PageSetup.TopMargin = $word.InchesToPoints(1)")
            sb.AppendLine("$doc.PageSetup.BottomMargin = $word.InchesToPoints(1)")
            sb.AppendLine("$doc.PageSetup.LeftMargin = $word.InchesToPoints(1.25)")
            sb.AppendLine("$doc.PageSetup.RightMargin = $word.InchesToPoints(1.25)")
            sb.AppendLine("$sec = $doc.Sections.Item(1)")
            sb.AppendLine("$hdr = $sec.Headers.Item(1)")
            sb.AppendLine("$hdr.Range.Text = '" & titleText & "'")
            sb.AppendLine("$hdr.Range.Font.Size = 9")
            sb.AppendLine("$hdr.Range.Font.Color = 8421504")
            sb.AppendLine("$ftr = $sec.Footers.Item(1)")
            sb.AppendLine("$ftr.PageNumbers.Add(1) | Out-Null")
            sb.AppendLine("$ftr.Range.ParagraphFormat.Alignment = 1")
            sb.AppendLine("try { $sel.Style = $doc.Styles.Item('Title') } catch { try { $sel.Style = $doc.Styles.Item('Heading 1') } catch {} }")
            sb.AppendLine("$sel.TypeText('" & titleText & "')")
            sb.AppendLine("$sel.TypeParagraph()")
            sb.AppendLine("$sel.TypeParagraph()")
            If addToC Then
                sb.AppendLine("try {")
                sb.AppendLine("    $sel.Style = $doc.Styles.Item('Normal')")
                sb.AppendLine("    $tocRange = $sel.Range")
                sb.AppendLine("    $toc = $doc.TablesOfContents.Add($tocRange, $true, 1, 3)")
                sb.AppendLine("    $toc.TabLeader = 0")
                sb.AppendLine("    $sel.TypeParagraph()")
                sb.AppendLine("    $sel.InsertBreak(7)")
                sb.AppendLine("} catch { $sel.TypeText('Table of Contents'); $sel.TypeParagraph() }")
            End If
            sb.AppendLine("$body = [System.IO.File]::ReadAllText('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)")
            sb.AppendLine("$bodyLines = $body -split '\r?\n'")
            sb.AppendLine("foreach ($line in $bodyLines) {")
            sb.AppendLine("    $trimmed = $line.Trim()")
            sb.AppendLine("    if ($trimmed -eq '') { $sel.TypeParagraph(); continue }")
            sb.AppendLine("    $isH1 = $trimmed -match '^#{1}\s+(.+)$' -or ($trimmed -match '^[A-Z][A-Z\s]{4,}$' -and $trimmed.Length -lt 60)")
            sb.AppendLine("    $isH2 = $trimmed -match '^#{2}\s+(.+)$' -or ($trimmed -match '^\d+\.\s+[A-Z]' -and $trimmed.Length -lt 80)")
            sb.AppendLine("    $isH3 = $trimmed -match '^#{3}\s+(.+)$' -or ($trimmed -match '^\d+\.\d+\s+' -and $trimmed.Length -lt 80)")
            sb.AppendLine("    $isBullet = $trimmed -match '^[-\*\•]\s+'")
            sb.AppendLine("    $isNumbered = $trimmed -match '^\d+[\.\)]\s+'")
            sb.AppendLine("    if ($isH1) {")
            sb.AppendLine("        $clean = $trimmed -replace '^#+\s*',''")
            sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('Heading 1') } catch {}")
            sb.AppendLine("        $sel.TypeText($clean)")
            sb.AppendLine("        $sel.TypeParagraph()")
            sb.AppendLine("    } elseif ($isH2) {")
            sb.AppendLine("        $clean = $trimmed -replace '^#+\s*','' -replace '^\d+\.\s*',''")
            sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('Heading 2') } catch {}")
            sb.AppendLine("        $sel.TypeText($clean)")
            sb.AppendLine("        $sel.TypeParagraph()")
            sb.AppendLine("    } elseif ($isH3) {")
            sb.AppendLine("        $clean = $trimmed -replace '^#+\s*','' -replace '^\d+\.\d+\s*',''")
            sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('Heading 3') } catch {}")
            sb.AppendLine("        $sel.TypeText($clean)")
            sb.AppendLine("        $sel.TypeParagraph()")
            sb.AppendLine("    } elseif ($isBullet) {")
            sb.AppendLine("        $clean = $trimmed -replace '^[-\*\•]\s*',''")
            sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('List Bullet') } catch {}")
            sb.AppendLine("        $sel.TypeText($clean)")
            sb.AppendLine("        $sel.TypeParagraph()")
            sb.AppendLine("    } elseif ($isNumbered) {")
            sb.AppendLine("        $clean = $trimmed -replace '^\d+[\.\)]\s*',''")
            sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('List Number') } catch {}")
            sb.AppendLine("        $sel.TypeText($clean)")
            sb.AppendLine("        $sel.TypeParagraph()")
            sb.AppendLine("    } else {")
            sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('Normal') } catch {}")
            sb.AppendLine("        $sel.TypeText($trimmed)")
            sb.AppendLine("        $sel.TypeParagraph()")
            sb.AppendLine("    }")
            sb.AppendLine("}")
            If addToC Then
                sb.AppendLine("try { $doc.TablesOfContents.Item(1).Update() } catch {}")
            End If
            If addTable Then
                sb.AppendLine("try {")
                sb.AppendLine("    $sel.TypeParagraph()")
                sb.AppendLine("    try { $sel.Style = $doc.Styles.Item('Heading 2') } catch {}")
                sb.AppendLine("    $sel.TypeText('Summary Table')")
                sb.AppendLine("    $sel.TypeParagraph()")
                sb.AppendLine("    try { $sel.Style = $doc.Styles.Item('Normal') } catch {}")
                sb.AppendLine("    $tableRange = $sel.Range")
                sb.AppendLine("    $tbl = $doc.Tables.Add($tableRange, 4, 3)")
                sb.AppendLine("    $tbl.Style = 'Table Grid'")
                sb.AppendLine("    $tbl.Rows.Item(1).Shading.BackgroundPatternColor = 4626167")
                sb.AppendLine("    $tbl.Rows.Item(1).Range.Font.Bold = $true")
                sb.AppendLine("    $tbl.Rows.Item(1).Range.Font.Color = 16777215")
                sb.AppendLine("    $tbl.Cell(1,1).Range.Text = 'Item'")
                sb.AppendLine("    $tbl.Cell(1,2).Range.Text = 'Description'")
                sb.AppendLine("    $tbl.Cell(1,3).Range.Text = 'Notes'")
                sb.AppendLine("    $tbl.Columns.AutoFit()")
                sb.AppendLine("} catch { Write-Output ('Table error: ' + $_.Exception.Message) }")
            End If
            sb.AppendLine("$doc.SaveAs2('" & savePath.Replace("'", "''") & "')")
            If saveAsPdf Then
                sb.AppendLine("try { $doc.ExportAsFixedFormat('" & pdfPath.Replace("'", "''") & "', 17) } catch { Write-Output ('PDF error: ' + $_.Exception.Message) }")
            End If
            sb.AppendLine("Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue")
            sb.AppendLine("try { $word.ActiveWindow.WindowState = 1 } catch {}")
            sb.AppendLine("Write-Output ('Saved to: ' + '" & savePath & "')")
            Return RunPowerShellSync(sb.ToString())
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function ModifyExistingWordViaPS(newContent As String, existingPath As String, userIntent As String) As String
        Try
            Dim tempTxt = GetUniqueTempPath("doc_modify_", ".txt")
            If Not String.IsNullOrWhiteSpace(newContent) Then
                File.WriteAllText(tempTxt, newContent, Encoding.UTF8)
            End If
            Dim appendContent = userIntent.Contains("append") OrElse userIntent.Contains("add") OrElse userIntent.Contains("insert")
            Dim addTable = wantsWordTable
            Dim saveAsPdf = wantsWordPdf
            Dim pdfPath = Path.ChangeExtension(existingPath, "_modified.pdf")
            Dim sb As New StringBuilder()
            sb.AppendLine("Get-Process -Name 'WINWORD' -ErrorAction SilentlyContinue | ForEach-Object { try { $_.CloseMainWindow() | Out-Null } catch {} }")
            sb.AppendLine("Start-Sleep -Milliseconds 800")
            sb.AppendLine("$word = New-Object -ComObject Word.Application")
            sb.AppendLine("$word.Visible = $true")
            sb.AppendLine("$doc = $word.Documents.Open('" & existingPath.Replace("'", "''") & "')")
            sb.AppendLine("$sel = $word.Selection")
            If appendContent AndAlso Not String.IsNullOrWhiteSpace(newContent) Then
                sb.AppendLine("$sel.EndKey(6) | Out-Null")
                sb.AppendLine("$sel.TypeParagraph()")
                sb.AppendLine("try { $sel.Style = $doc.Styles.Item('Heading 1') } catch {}")
                sb.AppendLine("$sel.TypeText('Additional Content')")
                sb.AppendLine("$sel.TypeParagraph()")
                sb.AppendLine("try { $sel.Style = $doc.Styles.Item('Normal') } catch {}")
                sb.AppendLine("$body = [System.IO.File]::ReadAllText('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)")
                sb.AppendLine("$bodyLines = $body -split '\r?\n'")
                sb.AppendLine("foreach ($line in $bodyLines) {")
                sb.AppendLine("    $trimmed = $line.Trim()")
                sb.AppendLine("    if ($trimmed -eq '') { $sel.TypeParagraph(); continue }")
                sb.AppendLine("    $isBullet = $trimmed -match '^[-\*\•]\s+'")
                sb.AppendLine("    $isH1 = $trimmed -match '^#{1}\s+' -or ($trimmed -match '^[A-Z][A-Z\s]{4,}$' -and $trimmed.Length -lt 60)")
                sb.AppendLine("    if ($isH1) {")
                sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('Heading 2') } catch {}")
                sb.AppendLine("        $sel.TypeText($trimmed -replace '^#+\s*','')")
                sb.AppendLine("    } elseif ($isBullet) {")
                sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('List Bullet') } catch {}")
                sb.AppendLine("        $sel.TypeText($trimmed -replace '^[-\*\•]\s*','')")
                sb.AppendLine("    } else {")
                sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('Normal') } catch {}")
                sb.AppendLine("        $sel.TypeText($trimmed)")
                sb.AppendLine("    }")
                sb.AppendLine("    $sel.TypeParagraph()")
                sb.AppendLine("}")
            End If
            If addTable Then
                sb.AppendLine("$sel.EndKey(6) | Out-Null")
                sb.AppendLine("$sel.TypeParagraph()")
                sb.AppendLine("try { $sel.Style = $doc.Styles.Item('Heading 2') } catch {}")
                sb.AppendLine("$sel.TypeText('Data Table')")
                sb.AppendLine("$sel.TypeParagraph()")
                sb.AppendLine("try { $sel.Style = $doc.Styles.Item('Normal') } catch {}")
                sb.AppendLine("try {")
                sb.AppendLine("    $tbl = $doc.Tables.Add($sel.Range, 3, 3)")
                sb.AppendLine("    $tbl.Style = 'Table Grid'")
                sb.AppendLine("    $tbl.Rows.Item(1).Range.Font.Bold = $true")
                sb.AppendLine("    $tbl.Rows.Item(1).Shading.BackgroundPatternColor = 4626167")
                sb.AppendLine("    $tbl.Cell(1,1).Range.Text = 'Column 1'")
                sb.AppendLine("    $tbl.Cell(1,2).Range.Text = 'Column 2'")
                sb.AppendLine("    $tbl.Cell(1,3).Range.Text = 'Column 3'")
                sb.AppendLine("    $tbl.Columns.AutoFit()")
                sb.AppendLine("} catch { Write-Output ('Table error: ' + $_.Exception.Message) }")
            End If
            sb.AppendLine("try {")
            sb.AppendLine("    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }")
            sb.AppendLine("} catch {}")
            sb.AppendLine("$doc.Save()")
            If saveAsPdf Then
                sb.AppendLine("try { $doc.ExportAsFixedFormat('" & pdfPath.Replace("'", "''") & "', 17) } catch {}")
            End If
            sb.AppendLine("if (Test-Path '" & tempTxt.Replace("'", "''") & "') { Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue }")
            sb.AppendLine("try { $word.ActiveWindow.WindowState = 1 } catch {}")
            sb.AppendLine("Write-Output ('Modified: ' + '" & existingPath & "')")
            Return RunPowerShellSync(sb.ToString())
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function CreatePptEnhancedViaPS(structuredContent As String, savePath As String, userIntent As String) As String
        Try
            Dim tempTxt = GetUniqueTempPath("ppt_content_", ".txt")
            File.WriteAllText(tempTxt, structuredContent, Encoding.UTF8)
            Dim addTransitions = wantsPptTransitions
            Dim addNotes = wantsPptNotes OrElse structuredContent.Contains("NOTES|")
            Dim addPdf = wantsPptPdf
            Dim pdfPath = Path.ChangeExtension(savePath, ".pdf")
            Dim themeAccent = "4472C4"
            If userIntent.Contains("red") Then themeAccent = "C00000"
            If userIntent.Contains("green") Then themeAccent = "375623"
            If userIntent.Contains("dark") Then themeAccent = "1F3864"
            If userIntent.Contains("orange") Then themeAccent = "E26B0A"
            If userIntent.Contains("purple") Then themeAccent = "7030A0"
            If userIntent.Contains("teal") Then themeAccent = "1F7391"
            Dim sb As New StringBuilder()
            sb.AppendLine("Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue | ForEach-Object { try { $_.CloseMainWindow() | Out-Null } catch {} }")
            sb.AppendLine("Start-Sleep -Milliseconds 800")
            sb.AppendLine("$ppt = New-Object -ComObject PowerPoint.Application")
            sb.AppendLine("$ppt.Visible = $true")
            sb.AppendLine("$pres = $ppt.Presentations.Add()")
            sb.AppendLine("$pres.PageSetup.SlideWidth = 720")
            sb.AppendLine("$pres.PageSetup.SlideHeight = 405")
            sb.AppendLine("try {")
            sb.AppendLine("    $scheme = $pres.SlideMaster.Theme.ThemeColorScheme")
            sb.AppendLine("    $scheme.Colors(5).RGB = [System.Convert]::ToInt32('FF" & themeAccent & "', 16)")
            sb.AppendLine("} catch {}")
            sb.AppendLine("$allLines = [System.IO.File]::ReadAllLines('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)")
            sb.AppendLine("$slideNum = 0")
            sb.AppendLine("$currentSlide = $null")
            sb.AppendLine("$currentLayout = 'CONTENT'")
            sb.AppendLine("$currentTitle = ''")
            sb.AppendLine("$currentContents = [System.Collections.Generic.List[string]]::new()")
            sb.AppendLine("$currentNotes = ''")
            sb.AppendLine("$slideData = [System.Collections.Generic.List[hashtable]]::new()")
            sb.AppendLine("foreach ($line in $allLines) {")
            sb.AppendLine("    if ($line -match '^SLIDE\|(.+?)\|(.+)$') {")
            sb.AppendLine("        if ($currentTitle -ne '' -or $slideNum -gt 0) {")
            sb.AppendLine("            $slideData.Add(@{ Layout=$currentLayout; Title=$currentTitle; Contents=[string[]]$currentContents.ToArray(); Notes=$currentNotes })")
            sb.AppendLine("        }")
            sb.AppendLine("        $currentLayout = $Matches[1].Trim().ToUpper()")
            sb.AppendLine("        $currentTitle = $Matches[2].Trim()")
            sb.AppendLine("        $currentContents = [System.Collections.Generic.List[string]]::new()")
            sb.AppendLine("        $currentNotes = ''")
            sb.AppendLine("        $slideNum++")
            sb.AppendLine("    } elseif ($line -match '^CONTENT\|(.+)$') {")
            sb.AppendLine("        $currentContents.Add($Matches[1].Trim())")
            sb.AppendLine("    } elseif ($line -match '^NOTES\|(.+)$') {")
            sb.AppendLine("        $currentNotes = $Matches[1].Trim()")
            sb.AppendLine("    }")
            sb.AppendLine("}")
            sb.AppendLine("if ($currentTitle -ne '') {")
            sb.AppendLine("    $slideData.Add(@{ Layout=$currentLayout; Title=$currentTitle; Contents=[string[]]$currentContents.ToArray(); Notes=$currentNotes })")
            sb.AppendLine("}")
            sb.AppendLine("$slideIndex = 1")
            sb.AppendLine("foreach ($sd in $slideData) {")
            sb.AppendLine("    $slide = $pres.Slides.Add($slideIndex, 1)")
            sb.AppendLine("    $slide.Layout = 1")
            sb.AppendLine("    try {")
            sb.AppendLine("        $titleShape = $slide.Shapes.AddTextbox(1, 30, 20, 660, 60)")
            sb.AppendLine("        $titleShape.TextFrame.WordWrap = $true")
            sb.AppendLine("        $titleRange = $titleShape.TextFrame.TextRange")
            sb.AppendLine("        $titleRange.Text = $sd.Title")
            sb.AppendLine("        $titleRange.Font.Size = if ($sd.Layout -eq 'TITLE') { 36 } else { 28 }")
            sb.AppendLine("        $titleRange.Font.Bold = $true")
            sb.AppendLine("        $titleRange.Font.Color.RGB = [System.Convert]::ToInt32('FF" & themeAccent & "', 16)")
            sb.AppendLine("        $titleRange.ParagraphFormat.Alignment = 2")
            sb.AppendLine("    } catch {}")
            sb.AppendLine("    try {")
            sb.AppendLine("        $line = $slide.Shapes.AddLine(30, 85, 690, 85)")
            sb.AppendLine("        $line.Line.ForeColor.RGB = [System.Convert]::ToInt32('FF" & themeAccent & "', 16)")
            sb.AppendLine("        $line.Line.Weight = 2")
            sb.AppendLine("    } catch {}")
            sb.AppendLine("    if ($sd.Layout -eq 'TITLE') {")
            sb.AppendLine("        try {")
            sb.AppendLine("            $subShape = $slide.Shapes.AddTextbox(1, 100, 150, 520, 180)")
            sb.AppendLine("            $subRange = $subShape.TextFrame.TextRange")
            sb.AppendLine("            $subRange.Text = if ($sd.Contents.Count -gt 0) { $sd.Contents[0] } else { '' }")
            sb.AppendLine("            $subRange.Font.Size = 20")
            sb.AppendLine("            $subRange.Font.Color.RGB = 6710886")
            sb.AppendLine("            $subRange.ParagraphFormat.Alignment = 2")
            sb.AppendLine("        } catch {}")
            sb.AppendLine("        try {")
            sb.AppendLine("            $bg = $slide.Shapes.AddShape(1, 0, 320, 720, 85)")
            sb.AppendLine("            $bg.Fill.ForeColor.RGB = [System.Convert]::ToInt32('FF" & themeAccent & "', 16)")
            sb.AppendLine("            $bg.Line.Visible = $false")
            sb.AppendLine("            $bg.ZOrder(1)")
            sb.AppendLine("        } catch {}")
            sb.AppendLine("    } elseif ($sd.Layout -eq 'TWOCOL') {")
            sb.AppendLine("        try {")
            sb.AppendLine("            $col1 = $slide.Shapes.AddTextbox(1, 30, 100, 320, 260)")
            sb.AppendLine("            $col2 = $slide.Shapes.AddTextbox(1, 370, 100, 320, 260)")
            sb.AppendLine("            $half = [Math]::Ceiling($sd.Contents.Count / 2)")
            sb.AppendLine("            $col1Text = ''")
            sb.AppendLine("            $col2Text = ''")
            sb.AppendLine("            for ($i = 0; $i -lt $sd.Contents.Count; $i++) {")
            sb.AppendLine("                if ($i -lt $half) { $col1Text += '• ' + $sd.Contents[$i] + [char]13 }")
            sb.AppendLine("                else { $col2Text += '• ' + $sd.Contents[$i] + [char]13 }")
            sb.AppendLine("            }")
            sb.AppendLine("            $col1.TextFrame.TextRange.Text = $col1Text.Trim()")
            sb.AppendLine("            $col2.TextFrame.TextRange.Text = $col2Text.Trim()")
            sb.AppendLine("            $col1.TextFrame.TextRange.Font.Size = 16")
            sb.AppendLine("            $col2.TextFrame.TextRange.Font.Size = 16")
            sb.AppendLine("        } catch {}")
            sb.AppendLine("    } elseif ($sd.Layout -ne 'BLANK') {")
            sb.AppendLine("        try {")
            sb.AppendLine("            $contentShape = $slide.Shapes.AddTextbox(1, 30, 100, 660, 260)")
            sb.AppendLine("            $contentShape.TextFrame.WordWrap = $true")
            sb.AppendLine("            $contentShape.TextFrame.AutoSize = 0")
            sb.AppendLine("            $tf = $contentShape.TextFrame.TextRange")
            sb.AppendLine("            $fullText = ''")
            sb.AppendLine("            foreach ($c in $sd.Contents) { $fullText += '• ' + $c + [char]13 }")
            sb.AppendLine("            $tf.Text = $fullText.Trim()")
            sb.AppendLine("            $tf.Font.Size = 18")
            sb.AppendLine("            $tf.Font.Color.RGB = 2302755")
            sb.AppendLine("            $tf.ParagraphFormat.SpaceAfter = 8")
            sb.AppendLine("        } catch {}")
            sb.AppendLine("    }")
            sb.AppendLine("    }")
            sb.AppendLine("    try {")
            sb.AppendLine("        if ($sd.Layout -ne 'TITLE') {")
            sb.AppendLine("            $slide.Background.Fill.ForeColor.RGB = 16777215")
            sb.AppendLine("            $slide.Background.Fill.Solid()")
            sb.AppendLine("        }")
            sb.AppendLine("    } catch {}")
            If addNotes Then
                sb.AppendLine("    try {")
                sb.AppendLine("        if ($sd.Notes -ne '') {")
                sb.AppendLine("            $notesPage = $slide.NotesPage")
                sb.AppendLine("            $notesShape = $notesPage.Shapes | Where-Object { $_.PlaceholderFormat.Type -eq 2 } | Select-Object -First 1")
                sb.AppendLine("            if ($notesShape) { $notesShape.TextFrame.TextRange.Text = $sd.Notes }")
                sb.AppendLine("        }")
                sb.AppendLine("    } catch {}")
            End If
            If addTransitions Then
                sb.AppendLine("    try {")
                sb.AppendLine("        $slide.SlideShowTransition.Type = 1")
                sb.AppendLine("        $slide.SlideShowTransition.Speed = 2")
                sb.AppendLine("        $slide.SlideShowTransition.AdvanceOnTime = $true")
                sb.AppendLine("        $slide.SlideShowTransition.AdvanceTime = 5")
                sb.AppendLine("    } catch {}")
            End If
            sb.AppendLine("    try {")
            sb.AppendLine("        $numShape = $slide.Shapes.AddTextbox(1, 650, 370, 60, 25)")
            sb.AppendLine("        $numShape.TextFrame.TextRange.Text = $slideIndex.ToString()")
            sb.AppendLine("        $numShape.TextFrame.TextRange.Font.Size = 10")
            sb.AppendLine("        $numShape.TextFrame.TextRange.Font.Color.RGB = 8421504")
            sb.AppendLine("        $numShape.TextFrame.TextRange.ParagraphFormat.Alignment = 3")
            sb.AppendLine("    } catch {}")
            sb.AppendLine("    $slideIndex++")
            sb.AppendLine("}")
            sb.AppendLine("$pres.SaveAs('" & savePath.Replace("'", "''") & "')")
            If addPdf Then
                sb.AppendLine("try { $pres.ExportAsFixedFormat('" & pdfPath.Replace("'", "''") & "', 2) } catch { Write-Output ('PDF error: ' + $_.Exception.Message) }")
            End If
            sb.AppendLine("Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue")
            sb.AppendLine("$ppt.WindowState = 3")
            sb.AppendLine("Write-Output ('Saved to: ' + '" & savePath & "')")
            Return RunPowerShellSync(sb.ToString())
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function ModifyExistingPptViaPS(newContent As String, existingPath As String, userIntent As String) As String
        Try
            Dim tempTxt = GetUniqueTempPath("ppt_modify_", ".txt")
            If Not String.IsNullOrWhiteSpace(newContent) Then
                File.WriteAllText(tempTxt, newContent, Encoding.UTF8)
            End If
            Dim addSlides = userIntent.Contains("add") OrElse userIntent.Contains("append") OrElse userIntent.Contains("insert")
            Dim addTransitions = wantsPptTransitions
            Dim saveAsPdf = wantsPptPdf
            Dim pdfPath = Path.ChangeExtension(existingPath, "_modified.pdf")
            Dim sb As New StringBuilder()
            sb.AppendLine("Get-Process -Name 'POWERPNT' -ErrorAction SilentlyContinue | ForEach-Object { try { $_.CloseMainWindow() | Out-Null } catch {} }")
            sb.AppendLine("Start-Sleep -Milliseconds 800")
            sb.AppendLine("$ppt = New-Object -ComObject PowerPoint.Application")
            sb.AppendLine("$ppt.Visible = $true")
            sb.AppendLine("$pres = $ppt.Presentations.Open('" & existingPath.Replace("'", "''") & "')")
            sb.AppendLine("$existingCount = $pres.Slides.Count")
            If addSlides AndAlso Not String.IsNullOrWhiteSpace(newContent) Then
                sb.AppendLine("$allLines = [System.IO.File]::ReadAllLines('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)")
                sb.AppendLine("$slideData = [System.Collections.Generic.List[hashtable]]::new()")
                sb.AppendLine("$currentLayout = 'CONTENT'")
                sb.AppendLine("$currentTitle = ''")
                sb.AppendLine("$currentContents = [System.Collections.Generic.List[string]]::new()")
                sb.AppendLine("$currentNotes = ''")
                sb.AppendLine("$started = $false")
                sb.AppendLine("foreach ($line in $allLines) {")
                sb.AppendLine("    if ($line -match '^SLIDE\|(.+?)\|(.+)$') {")
                sb.AppendLine("        if ($started) { $slideData.Add(@{ Layout=$currentLayout; Title=$currentTitle; Contents=[string[]]$currentContents.ToArray(); Notes=$currentNotes }) }")
                sb.AppendLine("        $currentLayout = $Matches[1].Trim().ToUpper()")
                sb.AppendLine("        $currentTitle = $Matches[2].Trim()")
                sb.AppendLine("        $currentContents = [System.Collections.Generic.List[string]]::new()")
                sb.AppendLine("        $currentNotes = ''")
                sb.AppendLine("        $started = $true")
                sb.AppendLine("    } elseif ($line -match '^CONTENT\|(.+)$') { $currentContents.Add($Matches[1].Trim()) }")
                sb.AppendLine("    elseif ($line -match '^NOTES\|(.+)$') { $currentNotes = $Matches[1].Trim() }")
                sb.AppendLine("}")
                sb.AppendLine("if ($started) { $slideData.Add(@{ Layout=$currentLayout; Title=$currentTitle; Contents=[string[]]$currentContents.ToArray(); Notes=$currentNotes }) }")
                sb.AppendLine("$slideIndex = $existingCount + 1")
                sb.AppendLine("foreach ($sd in $slideData) {")
                sb.AppendLine("    $slide = $pres.Slides.Add($slideIndex, 1)")
                sb.AppendLine("    try {")
                sb.AppendLine("        $ts = $slide.Shapes.AddTextbox(1, 30, 20, 660, 60)")
                sb.AppendLine("        $ts.TextFrame.TextRange.Text = $sd.Title")
                sb.AppendLine("        $ts.TextFrame.TextRange.Font.Size = 28")
                sb.AppendLine("        $ts.TextFrame.TextRange.Font.Bold = $true")
                sb.AppendLine("    } catch {}")
                sb.AppendLine("    try {")
                sb.AppendLine("        $cs = $slide.Shapes.AddTextbox(1, 30, 100, 660, 260)")
                sb.AppendLine("        $fullText = ''")
                sb.AppendLine("        foreach ($c in $sd.Contents) { $fullText += '• ' + $c + [char]13 }")
                sb.AppendLine("        $cs.TextFrame.TextRange.Text = $fullText.Trim()")
                sb.AppendLine("        $cs.TextFrame.TextRange.Font.Size = 18")
                sb.AppendLine("    } catch {}")
                If wantsPptNotes Then
                    sb.AppendLine("    try {")
                    sb.AppendLine("        if ($sd.Notes -ne '') {")
                    sb.AppendLine("            $np = $slide.NotesPage")
                    sb.AppendLine("            $ns = $np.Shapes | Where-Object { $_.PlaceholderFormat.Type -eq 2 } | Select-Object -First 1")
                    sb.AppendLine("            if ($ns) { $ns.TextFrame.TextRange.Text = $sd.Notes }")
                    sb.AppendLine("        }")
                    sb.AppendLine("    } catch {}")
                End If
                sb.AppendLine("    $slideIndex++")
                sb.AppendLine("}")
            End If
            If addTransitions Then
                sb.AppendLine("foreach ($slide in $pres.Slides) {")
                sb.AppendLine("    try {")
                sb.AppendLine("        $slide.SlideShowTransition.Type = 1")
                sb.AppendLine("        $slide.SlideShowTransition.Speed = 2")
                sb.AppendLine("    } catch {}")
                sb.AppendLine("}")
            End If
            sb.AppendLine("$pres.Save()")
            If saveAsPdf Then
                sb.AppendLine("try { $pres.ExportAsFixedFormat('" & pdfPath.Replace("'", "''") & "', 2) } catch {}")
            End If
            sb.AppendLine("if (Test-Path '" & tempTxt.Replace("'", "''") & "') { Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue }")
            sb.AppendLine("$ppt.WindowState = 3")
            sb.AppendLine("Write-Output ('Modified: ' + '" & existingPath & "')")
            Return RunPowerShellSync(sb.ToString())
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function HandleOutlookViaPS(content As String, userIntent As String) As String
        Try
            Dim sb As New StringBuilder()
            sb.AppendLine("Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Out-Null")
            sb.AppendLine("$outlook = New-Object -ComObject Outlook.Application")
            sb.AppendLine("$ns = $outlook.GetNamespace('MAPI')")
            If wantsOutlookRead OrElse (Not wantsOutlookSend AndAlso Not wantsOutlookCalendar AndAlso
                                    Not wantsOutlookTask AndAlso Not wantsOutlookContact AndAlso
                                    Not wantsOutlookReply AndAlso Not wantsOutlookForward) Then
                Dim searchTerm = ""
                Dim intentWords = userIntent.Split(" "c)
                Dim skipRead = {"read", "check", "search", "find", "show", "list", "email", "emails",
                            "inbox", "mail", "outlook", "me", "my", "the", "in", "from", "for"}
                For Each w In intentWords
                    If w.Length > 2 AndAlso Not skipRead.Contains(w) Then
                        searchTerm = w
                        Exit For
                    End If
                Next
                sb.AppendLine("try {")
                sb.AppendLine("    $inbox = $ns.GetDefaultFolder(6)")
                sb.AppendLine("    $items = $inbox.Items")
                sb.AppendLine("    $items.Sort('[ReceivedTime]', $true)")
                If Not String.IsNullOrWhiteSpace(searchTerm) Then
                    sb.AppendLine("    $filtered = $items | Where-Object { $_.Subject -like '*" & searchTerm & "*' -or $_.SenderName -like '*" & searchTerm & "*' } | Select-Object -First 10")
                Else
                    sb.AppendLine("    $filtered = $items | Select-Object -First 10")
                End If
                sb.AppendLine("    $output = ''")
                sb.AppendLine("    foreach ($mail in $filtered) {")
                sb.AppendLine("        $output += '📧 ' + $mail.ReceivedTime.ToString('MM/dd/yyyy HH:mm') + ' | From: ' + $mail.SenderName + ' | ' + $mail.Subject + [char]13 + [char]10")
                sb.AppendLine("    }")
                sb.AppendLine("    if ($output -eq '') { $output = 'No emails found.' }")
                sb.AppendLine("    Write-Output $output")
                sb.AppendLine("} catch { Write-Output ('Read error: ' + $_.Exception.Message) }")
            End If
            If wantsOutlookSend AndAlso Not wantsOutlookReply AndAlso Not wantsOutlookForward Then
                Dim tempTxt = GetUniqueTempPath("outlook_body_", ".txt")
                File.WriteAllText(tempTxt, content, Encoding.UTF8)
                sb.AppendLine("try {")
                sb.AppendLine("    $mail = $outlook.CreateItem(0)")
                sb.AppendLine("    $bodyText = [System.IO.File]::ReadAllText('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)")
                Dim toAddr = ""
                Dim toMatch = Regex.Match(userIntent, "to\s+([\w\.\-]+@[\w\.\-]+)")
                If toMatch.Success Then toAddr = toMatch.Groups(1).Value
                Dim subjectText = "Message from Autono-Me"
                Dim subjMatch = Regex.Match(userIntent, "(?:subject|re:|about)\s+['""]?(.+?)['""]?(?:\s|$)", RegexOptions.IgnoreCase)
                If subjMatch.Success Then subjectText = subjMatch.Groups(1).Value.Trim()
                sb.AppendLine("    $mail.To = '" & toAddr.Replace("'", "''") & "'")
                sb.AppendLine("    $mail.Subject = '" & subjectText.Replace("'", "''") & "'")
                sb.AppendLine("    $mail.Body = $bodyText")
                Dim shouldSend = userIntent.Contains("send") AndAlso
                 Not userIntent.Contains("draft") AndAlso
                 Not userIntent.Contains("compose") AndAlso
                 Not userIntent.Contains("show")
                If shouldSend Then
                    sb.AppendLine("    $mail.Send()")
                    sb.AppendLine("    Write-Output ('Email sent to: " & toAddr & "')")
                Else
                    sb.AppendLine("    $mail.Display()")
                    sb.AppendLine("    Write-Output ('Email composed to: " & toAddr & "')")
                End If
                sb.AppendLine("} catch { Write-Output ('Send error: ' + $_.Exception.Message) }")
            End If
            If wantsOutlookReply OrElse wantsOutlookForward Then
                sb.AppendLine("try {")
                sb.AppendLine("    $inbox = $ns.GetDefaultFolder(6)")
                sb.AppendLine("    $items = $inbox.Items")
                sb.AppendLine("    $items.Sort('[ReceivedTime]', $true)")
                sb.AppendLine("    $latest = $items | Select-Object -First 1")
                If wantsOutlookForward Then
                    sb.AppendLine("    $fwd = $latest.Forward()")
                    Dim fwdTo = ""
                    Dim fwdMatch = Regex.Match(userIntent, "to\s+([\w\.\-]+@[\w\.\-]+)")
                    If fwdMatch.Success Then fwdTo = fwdMatch.Groups(1).Value
                    sb.AppendLine("    $fwd.To = '" & fwdTo.Replace("'", "''") & "'")
                    sb.AppendLine("    $fwd.Display()")
                    sb.AppendLine("    Write-Output ('Forwarded to: " & fwdTo & "')")
                Else
                    sb.AppendLine("    $reply = $latest.Reply()")
                    sb.AppendLine("    $reply.Body = '" & content.Replace("'", "''").Replace(Environment.NewLine, "' + [char]13 + [char]10 + '") & "' + [char]13 + [char]10 + $reply.Body")
                    sb.AppendLine("    $reply.Display()")
                    sb.AppendLine("    Write-Output ('Reply opened for: ' + $latest.Subject)")
                End If
                sb.AppendLine("} catch { Write-Output ('Reply/Forward error: ' + $_.Exception.Message) }")
            End If
            If wantsOutlookCalendar Then
                Dim subjectText = "New Appointment"
                Dim subjMatch = Regex.Match(userIntent, "(?:meeting|appointment|event|schedule|called|titled|named|about)\s+['""]?(.+?)['""]?(?:\s+on|\s+at|\s+for|$)", RegexOptions.IgnoreCase)
                If subjMatch.Success Then subjectText = subjMatch.Groups(1).Value.Trim()
                Dim startTime = DateTime.Now.AddHours(1)
                Dim timeMatch = Regex.Match(userIntent, "(\d{1,2}(?::\d{2})?\s*(?:am|pm))", RegexOptions.IgnoreCase)
                If timeMatch.Success Then DateTime.TryParse(timeMatch.Value, startTime)
                Dim dateMatch = Regex.Match(userIntent, "(?:on\s+)?(\w+ \d{1,2}(?:st|nd|rd|th)?(?:,?\s+\d{4})?)", RegexOptions.IgnoreCase)
                If dateMatch.Success Then
                    Dim parsedDate As DateTime
                    If DateTime.TryParse(dateMatch.Groups(1).Value, parsedDate) Then
                        startTime = parsedDate.Date.Add(startTime.TimeOfDay)
                    End If
                End If
                sb.AppendLine("try {")
                sb.AppendLine("    $appt = $outlook.CreateItem(1)")
                sb.AppendLine("    $appt.Subject = '" & subjectText.Replace("'", "''") & "'")
                sb.AppendLine("    $appt.Start = '" & startTime.ToString("MM/dd/yyyy HH:mm") & "'")
                sb.AppendLine("    $appt.End = '" & startTime.AddHours(1).ToString("MM/dd/yyyy HH:mm") & "'")
                sb.AppendLine("    $appt.Body = '" & content.Replace("'", "''") & "'")
                sb.AppendLine("    $appt.ReminderSet = $true")
                sb.AppendLine("    $appt.ReminderMinutesBeforeStart = 15")
                sb.AppendLine("    $appt.Display()")
                sb.AppendLine("    Write-Output ('Appointment created: " & subjectText & " at " & startTime.ToString("MM/dd/yyyy HH:mm") & "')")
                sb.AppendLine("} catch { Write-Output ('Calendar error: ' + $_.Exception.Message) }")
            End If
            If wantsOutlookTask Then
                Dim taskSubject = "New Task"
                Dim taskMatch = Regex.Match(userIntent, "(?:task|todo|to-do)\s+(?:called|titled|named|about|to|for)?\s*['""]?(.+?)['""]?(?:\s+by|\s+due|\s+before|$)", RegexOptions.IgnoreCase)
                If taskMatch.Success Then taskSubject = taskMatch.Groups(1).Value.Trim()
                Dim dueDate = DateTime.Now.AddDays(7)
                Dim dueMatch = Regex.Match(userIntent, "(?:due|by|before)\s+(.+?)(?:\s|$)", RegexOptions.IgnoreCase)
                If dueMatch.Success Then DateTime.TryParse(dueMatch.Groups(1).Value, dueDate)
                sb.AppendLine("try {")
                sb.AppendLine("    $task = $outlook.CreateItem(3)")
                sb.AppendLine("    $task.Subject = '" & taskSubject.Replace("'", "''") & "'")
                sb.AppendLine("    $task.DueDate = '" & dueDate.ToString("MM/dd/yyyy") & "'")
                sb.AppendLine("    $task.Body = '" & content.Replace("'", "''") & "'")
                sb.AppendLine("    $task.Importance = 1")
                sb.AppendLine("    $task.Display()")
                sb.AppendLine("    Write-Output ('Task created: " & taskSubject & "')")
                sb.AppendLine("} catch { Write-Output ('Task error: ' + $_.Exception.Message) }")
            End If
            If wantsOutlookContact Then
                Dim firstName = ""
                Dim lastName = ""
                Dim email = ""
                Dim phone = ""
                Dim emailMatch = Regex.Match(userIntent, "([\w\.\-]+@[\w\.\-]+)")
                If emailMatch.Success Then email = emailMatch.Value
                Dim phoneMatch = Regex.Match(userIntent, "(\+?[\d\s\-\(\)]{7,15})")
                If phoneMatch.Success Then phone = phoneMatch.Value.Trim()
                Dim nameMatch = Regex.Match(userIntent, "(?:contact|person|for)\s+([A-Z][a-z]+)\s+([A-Z][a-z]+)", RegexOptions.IgnoreCase)
                If nameMatch.Success Then
                    firstName = nameMatch.Groups(1).Value
                    lastName = nameMatch.Groups(2).Value
                End If
                sb.AppendLine("try {")
                sb.AppendLine("    $contact = $outlook.CreateItem(2)")
                sb.AppendLine("    $contact.FirstName = '" & firstName.Replace("'", "''") & "'")
                sb.AppendLine("    $contact.LastName = '" & lastName.Replace("'", "''") & "'")
                sb.AppendLine("    $contact.Email1Address = '" & email.Replace("'", "''") & "'")
                sb.AppendLine("    $contact.BusinessTelephoneNumber = '" & phone.Replace("'", "''") & "'")
                sb.AppendLine("    $contact.Body = '" & content.Replace("'", "''") & "'")
                sb.AppendLine("    $contact.Display()")
                sb.AppendLine("    Write-Output ('Contact created: " & firstName & " " & lastName & "')")
                sb.AppendLine("} catch { Write-Output ('Contact error: ' + $_.Exception.Message) }")
            End If
            If wantsOutlookExport Then
                Dim exportPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                           "EmailExport_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".docx")
                sb.AppendLine("try {")
                sb.AppendLine("    $inbox = $ns.GetDefaultFolder(6)")
                sb.AppendLine("    $items = $inbox.Items")
                sb.AppendLine("    $items.Sort('[ReceivedTime]', $true)")
                sb.AppendLine("    $emails = $items | Select-Object -First 5")
                sb.AppendLine("    $word = New-Object -ComObject Word.Application")
                sb.AppendLine("    $word.Visible = $false")
                sb.AppendLine("    $doc = $word.Documents.Add()")
                sb.AppendLine("    $sel = $word.Selection")
                sb.AppendLine("    try { $sel.Style = $doc.Styles.Item('Title') } catch {}")
                sb.AppendLine("    $sel.TypeText('Email Export - ' + (Get-Date).ToString('MMMM dd, yyyy'))")
                sb.AppendLine("    $sel.TypeParagraph()")
                sb.AppendLine("    foreach ($mail in $emails) {")
                sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('Heading 1') } catch {}")
                sb.AppendLine("        $sel.TypeText($mail.Subject)")
                sb.AppendLine("        $sel.TypeParagraph()")
                sb.AppendLine("        try { $sel.Style = $doc.Styles.Item('Normal') } catch {}")
                sb.AppendLine("        $sel.TypeText('From: ' + $mail.SenderName + ' | ' + $mail.ReceivedTime.ToString('MM/dd/yyyy HH:mm'))")
                sb.AppendLine("        $sel.TypeParagraph()")
                sb.AppendLine("        $sel.TypeText($mail.Body.Substring(0, [Math]::Min(500, $mail.Body.Length)))")
                sb.AppendLine("        $sel.TypeParagraph()")
                sb.AppendLine("        $sel.TypeParagraph()")
                sb.AppendLine("    }")
                sb.AppendLine("    $doc.SaveAs2('" & exportPath.Replace("'", "''") & "')")
                If wantsOutlookExport AndAlso userIntent.Contains("pdf") Then
                    Dim pdfPath = Path.ChangeExtension(exportPath, ".pdf")
                    sb.AppendLine("    try { $doc.ExportAsFixedFormat('" & pdfPath.Replace("'", "''") & "', 17) } catch {}")
                End If
                sb.AppendLine("    $doc.Close()")
                sb.AppendLine("    $word.Quit()")
                sb.AppendLine("    Write-Output ('Emails exported to: " & exportPath & "')")
                sb.AppendLine("} catch { Write-Output ('Export error: ' + $_.Exception.Message) }")
            End If
            Return RunPowerShellSync(sb.ToString())
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function HandleOneNoteViaPS(content As String, userIntent As String) As String
        Try
            Dim tempTxt = GetUniqueTempPath("onenote_content_", ".txt")
            File.WriteAllText(tempTxt, content, Encoding.UTF8)
            Dim exportPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                   "NoteExport_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".docx")
            Dim noteTitle = "New Note"
            Dim titleMatch = Regex.Match(userIntent, "(?:titled|called|named|about)\s+['""]?(.+?)['""]?(?:\s+in|\s+to|\s+with|$)", RegexOptions.IgnoreCase)
            If titleMatch.Success Then noteTitle = titleMatch.Groups(1).Value.Trim()
            Dim sectionName = "Quick Notes"
            Dim sectionMatch = Regex.Match(userIntent, "(?:in\s+the\s+|in\s+|section\s+|notebook\s+)['""]?([A-Za-z][A-Za-z\s]{1,30}?)['""]?(?:\s+section|\s+notebook|$)", RegexOptions.IgnoreCase)
            If sectionMatch.Success Then sectionName = sectionMatch.Groups(1).Value.Trim()
            Dim sb As New StringBuilder()
            If wantsOneNoteSearch Then
                Dim searchTerm = ""
                Dim searchMatch = Regex.Match(userIntent, "(?:search|find|look for)\s+['""]?(.+?)['""]?(?:\s+in|\s+on|$)", RegexOptions.IgnoreCase)
                If searchMatch.Success Then searchTerm = searchMatch.Groups(1).Value.Trim()
                sb.AppendLine("try {")
                sb.AppendLine("    $on = New-Object -ComObject OneNote.Application")
                sb.AppendLine("    [xml]$results = ''")
                sb.AppendLine("    $on.FindPages('', '" & searchTerm.Replace("'", "''") & "', [ref]$results, $false, $false, 0)")
                sb.AppendLine("    $pages = $results.FindPagesResult.Page")
                sb.AppendLine("    $output = 'Search results for: " & searchTerm & "' + [char]13 + [char]10")
                sb.AppendLine("    foreach ($page in $pages | Select-Object -First 10) {")
                sb.AppendLine("        $output += '📄 ' + $page.name + [char]13 + [char]10")
                sb.AppendLine("    }")
                sb.AppendLine("    if (-not $pages) { $output += 'No results found.' }")
                sb.AppendLine("    Write-Output $output")
                sb.AppendLine("} catch { Write-Output ('Search error: ' + $_.Exception.Message) }")
                Return RunPowerShellSync(sb.ToString())
            End If
            If wantsOneNoteExport Then
                sb.AppendLine("try {")
                sb.AppendLine("    $on = New-Object -ComObject OneNote.Application")
                sb.AppendLine("    [xml]$hier = ''")
                sb.AppendLine("    $on.GetHierarchy('', 3, [ref]$hier)")
                sb.AppendLine("    $firstSection = $hier.Notebooks.Notebook.Section | Select-Object -First 1")
                sb.AppendLine("    $sectionId = $firstSection.ID")
                sb.AppendLine("    $on.Publish($sectionId, '" & exportPath.Replace("'", "''") & "', 4, '')")
                If userIntent.Contains("pdf") Then
                    Dim pdfPath = Path.ChangeExtension(exportPath, ".pdf")
                    sb.AppendLine("    $on.Publish($sectionId, '" & pdfPath.Replace("'", "''") & "', 1, '')")
                End If
                sb.AppendLine("    Write-Output ('Exported to: " & exportPath & "')")
                sb.AppendLine("} catch { Write-Output ('Export error: ' + $_.Exception.Message) }")
                Return RunPowerShellSync(sb.ToString())
            End If
            sb.AppendLine("try {")
            sb.AppendLine("    $on = New-Object -ComObject OneNote.Application")
            sb.AppendLine("")
            sb.AppendLine("    # Get hierarchy")
            sb.AppendLine("    [xml]$hier = ''")
            sb.AppendLine("    $on.GetHierarchy('', 3, [ref]$hier)")
            sb.AppendLine("")
            sb.AppendLine("    # Find target section")
            sb.AppendLine("    $targetSection = $null")
            sb.AppendLine("    foreach ($nb in $hier.Notebooks.Notebook) {")
            sb.AppendLine("        foreach ($sec in $nb.Section) {")
            sb.AppendLine("            if ($sec.name -like '*" & sectionName.Replace("'", "''") & "*') {")
            sb.AppendLine("                $targetSection = $sec")
            sb.AppendLine("                break")
            sb.AppendLine("            }")
            sb.AppendLine("        }")
            sb.AppendLine("        if ($targetSection) { break }")
            sb.AppendLine("    }")
            sb.AppendLine("    if (-not $targetSection) {")
            sb.AppendLine("        $targetSection = $hier.Notebooks.Notebook.Section | Select-Object -First 1")
            sb.AppendLine("    }")
            sb.AppendLine("    $sectionId = $targetSection.ID")
            sb.AppendLine("")
            sb.AppendLine("    # Create new page")
            sb.AppendLine("    $newPageId = ''")
            sb.AppendLine("    $on.CreateNewPage($sectionId, [ref]$newPageId, 1)")
            sb.AppendLine("")
            sb.AppendLine("    # Wait for page to be ready")
            sb.AppendLine("    Start-Sleep -Milliseconds 2000")
            sb.AppendLine("")
            sb.AppendLine("    # Get the existing page XML so we can inject into it")
            sb.AppendLine("    [xml]$pageXml = ''")
            sb.AppendLine("    $on.GetPageContent($newPageId, [ref]$pageXml, 7)")
            sb.AppendLine("    $pageId = $pageXml.Page.ID")
            sb.AppendLine("")
            sb.AppendLine("    # Read body lines from file")
            sb.AppendLine("    $bodyLines = [System.IO.File]::ReadAllLines('" & tempTxt.Replace("'", "''") & "', [System.Text.Encoding]::UTF8)")
            sb.AppendLine("")
            sb.AppendLine("    # Build outline content")
            sb.AppendLine("    $oeChildren = ''")
            sb.AppendLine("    foreach ($line in $bodyLines) {")
            sb.AppendLine("        $t = $line.Trim()")
            sb.AppendLine("        if ($t -eq '') {")
            sb.AppendLine("            $oeChildren += '<one:OE><one:T><![CDATA[]] ></one:T></one:OE>'")
            sb.AppendLine("            continue")
            sb.AppendLine("        }")
            If wantsOneNoteChecklist Then
                sb.AppendLine("        $oeChildren += '<one:OE><one:Tag index=""0"" completed=""false"" /><one:T><![CDATA[' + $t + ']] ></one:T></one:OE>'")
            ElseIf wantsOneNoteTable Then
            Else
                sb.AppendLine("        $isBullet = ($t -match '^[-\*•]\s+(.+)$')")
                sb.AppendLine("        $isH1 = ($t -match '^#\s+(.+)$')")
                sb.AppendLine("        $isH2 = ($t -match '^##\s+(.+)$')")
                sb.AppendLine("        if ($isH1 -or $isH2) {")
                sb.AppendLine("            $clean = $t -replace '^#+\s+', ''")
                sb.AppendLine("            $oeChildren += '<one:OE><one:T><![CDATA[' + $clean + ']] ></one:T></one:OE>'")
                sb.AppendLine("        } elseif ($isBullet) {")
                sb.AppendLine("            $clean = $t -replace '^[-\*•]\s+', ''")
                sb.AppendLine("            $oeChildren += '<one:OE spaceBefore=""4"" spaceAfter=""4""><one:List><one:Bullet>•</one:Bullet></one:List><one:T><![CDATA[' + $clean + ']] ></one:T></one:OE>'")
                sb.AppendLine("        } else {")
                sb.AppendLine("            $oeChildren += '<one:OE><one:T><![CDATA[' + $t + ']] ></one:T></one:OE>'")
                sb.AppendLine("        }")
            End If
            sb.AppendLine("    }")
            sb.AppendLine("")
            If wantsOneNoteTable Then
                sb.AppendLine("    $tableRows = ''")
                sb.AppendLine("    $tableRows += '<one:Row><one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Item]] ></one:T></one:OE></one:OEChildren></one:Cell>'")
                sb.AppendLine("    $tableRows += '<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Details]] ></one:T></one:OE></one:OEChildren></one:Cell>'")
                sb.AppendLine("    $tableRows += '<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Status]] ></one:T></one:OE></one:OEChildren></one:Cell></one:Row>'")
                sb.AppendLine("    foreach ($line in $bodyLines) {")
                sb.AppendLine("        $t = $line.Trim()")
                sb.AppendLine("        if ($t -eq '') { continue }")
                sb.AppendLine("        $tableRows += '<one:Row>'")
                sb.AppendLine("        $tableRows += '<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[' + $t + ']] ></one:T></one:OE></one:OEChildren></one:Cell>'")
                sb.AppendLine("        $tableRows += '<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[]] ></one:T></one:OE></one:OEChildren></one:Cell>'")
                sb.AppendLine("        $tableRows += '<one:Cell><one:OEChildren><one:OE><one:T><![CDATA[Pending]] ></one:T></one:OE></one:OEChildren></one:Cell>'")
                sb.AppendLine("        $tableRows += '</one:Row>'")
                sb.AppendLine("    }")
                sb.AppendLine("    $oeChildren = '<one:OE><one:Table bordersVisible=""true"">' + $tableRows + '</one:Table></one:OE>'")
            End If
            sb.AppendLine("    # Build full update XML using real page ID from GetPageContent")
            sb.AppendLine("    $ns = 'xmlns:one=""http://schemas.microsoft.com/office/onenote/2013/onenote""'")
            sb.AppendLine("    $updateXml  = '<?xml version=""1.0"" encoding=""utf-8""?>'")
            sb.AppendLine("    $updateXml += '<one:Page ' + $ns + ' ID=""' + $pageId + '"">'")
            sb.AppendLine("    $updateXml += '<one:Title><one:OE><one:T><![CDATA[" & noteTitle.Replace("'", "''") & "]] ></one:T></one:OE></one:Title>'")
            sb.AppendLine("    $updateXml += '<one:Outline>'")
            sb.AppendLine("    $updateXml += '<one:Position x=""36"" y=""86"" />'")
            sb.AppendLine("    $updateXml += '<one:Size width=""600"" height=""200"" />'")
            sb.AppendLine("    $updateXml += '<one:OEChildren>'")
            sb.AppendLine("    $updateXml += $oeChildren")
            sb.AppendLine("    $updateXml += '</one:OEChildren>'")
            sb.AppendLine("    $updateXml += '</one:Outline>'")
            sb.AppendLine("    $updateXml += '</one:Page>'")
            sb.AppendLine("")
            sb.AppendLine("    # Push content to OneNote")
            sb.AppendLine("    $on.UpdatePageContent($updateXml, [System.DateTime]::MinValue, 1, $false)")
            sb.AppendLine("    Start-Sleep -Milliseconds 1000")
            sb.AppendLine("")
            sb.AppendLine("    # Navigate to the new page")
            sb.AppendLine("    try { $on.NavigateTo($newPageId, '', $false) } catch {}")
            sb.AppendLine("    Write-Output ('Note created: " & noteTitle & " in section: ' + $targetSection.name)")
            sb.AppendLine("")
            sb.AppendLine("} catch {")
            sb.AppendLine("    Write-Output ('OneNote error: ' + $_.Exception.Message)")
            sb.AppendLine("    Write-Output ('Stack: ' + $_.Exception.StackTrace)")
            sb.AppendLine("}")
            sb.AppendLine("")
            sb.AppendLine("if (Test-Path '" & tempTxt.Replace("'", "''") & "') {")
            sb.AppendLine("    Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue")
            sb.AppendLine("}")
            Return RunPowerShellSync(sb.ToString())
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function HandleAccessViaPS(content As String, userIntent As String) As String
        Try
            Dim tempTxt = GetUniqueTempPath("access_content_", ".txt")
            If Not String.IsNullOrWhiteSpace(content) Then
                File.WriteAllText(tempTxt, content, Encoding.UTF8)
            End If
            Dim dbPath As String = ""
            For Each f In selectedFiles
                Dim ext = Path.GetExtension(f).ToLower()
                If ext = ".accdb" OrElse ext = ".mdb" Then
                    dbPath = f
                    Exit For
                End If
            Next
            If String.IsNullOrEmpty(dbPath) Then
                Dim pathMatch = Regex.Match(userIntent, "([a-zA-Z]:\\[^\s'""]+\.(?:accdb|mdb))", RegexOptions.IgnoreCase)
                If pathMatch.Success Then dbPath = pathMatch.Value
            End If
            Dim tableName = "Data"
            Dim tableMatch = Regex.Match(userIntent, "(?:table|called|named|into)\s+['""]?([A-Za-z][A-Za-z0-9_\s]{0,30}?)['""]?(?:\s+in|\s+from|\s+with|$)", RegexOptions.IgnoreCase)
            If tableMatch.Success Then tableName = tableMatch.Groups(1).Value.Trim()
            If String.IsNullOrEmpty(dbPath) Then
                dbPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                  "Database_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".accdb")
            End If
            Dim sqlQuery = "SELECT * FROM [" & tableName & "]"
            Dim sqlMatch = Regex.Match(userIntent, "(?:query|select|where|sql)\s+(.+?)(?:\s+from|\s+in|\s+and|$)", RegexOptions.IgnoreCase)
            If sqlMatch.Success Then
                Dim rawSql = sqlMatch.Groups(1).Value.Trim()
                If rawSql.ToUpper().StartsWith("SELECT") Then
                    sqlQuery = rawSql
                End If
            End If
            Dim exportPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                       "AccessExport_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".xlsx")
            Dim sb As New StringBuilder()
            sb.AppendLine("$dbPath = '" & dbPath.Replace("'", "''") & "'")
            sb.AppendLine("$tableName = '" & tableName.Replace("'", "''") & "'")
            sb.AppendLine("")
            sb.AppendLine("# Try COM first, fall back to ADODB")
            sb.AppendLine("$useADODB = $false")
            sb.AppendLine("$access = $null")
            sb.AppendLine("try {")
            sb.AppendLine("    $access = New-Object -ComObject Access.Application")
            sb.AppendLine("    $access.Visible = $true")
            sb.AppendLine("} catch {")
            sb.AppendLine("    Write-Output ('COM unavailable, using ADODB: ' + $_.Exception.Message)")
            sb.AppendLine("    $useADODB = $true")
            sb.AppendLine("}")
            sb.AppendLine("")
            If wantsAccessCreate Then
                sb.AppendLine("# CREATE DATABASE + TABLE")
                sb.AppendLine("try {")
                sb.AppendLine("    if (-not $useADODB) {")
                sb.AppendLine("        if (Test-Path $dbPath) { Remove-Item $dbPath -Force }")
                sb.AppendLine("        $access.NewCurrentDatabase($dbPath)")
                Dim columns = New List(Of String) From {"ID", "Name", "Value", "Notes"}
                If Not String.IsNullOrWhiteSpace(content) Then
                    Dim lines = content.Split({Environment.NewLine, Chr(10)}, StringSplitOptions.RemoveEmptyEntries)
                    If lines.Length > 0 Then
                        Dim firstLine = lines(0).Trim()
                        If firstLine.Contains(Chr(9)) OrElse firstLine.Contains(",") Then
                            Dim sep = If(firstLine.Contains(Chr(9)), Chr(9), ",")
                            Dim parts = firstLine.Split(sep).Select(Function(p) p.Trim().Trim(""""c)).ToArray()
                            If parts.Length > 0 Then columns = parts.ToList()
                        End If
                    End If
                End If
                sb.AppendLine("        $db = $access.CurrentDb()")
                sb.AppendLine("        $tbl = $db.CreateTableDef($tableName)")
                For Each col In columns
                    Dim colClean = Regex.Replace(col, "[^A-Za-z0-9_]", "_")
                    If colClean = "ID" Then
                        sb.AppendLine("        $fld = $tbl.CreateField('ID', 4)")
                        sb.AppendLine("        $fld.Attributes = 17")
                        sb.AppendLine("        $tbl.Fields.Append($fld)")
                    Else
                        sb.AppendLine("        $tbl.Fields.Append($tbl.CreateField('" & colClean & "', 10, 255))")
                    End If
                Next
                sb.AppendLine("        $db.TableDefs.Append($tbl)")
                sb.AppendLine("        Write-Output ('Database created: ' + $dbPath)")
                sb.AppendLine("    } else {")
                sb.AppendLine("        $cat = New-Object -ComObject ADOX.Catalog")
                sb.AppendLine("        $cat.Create('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + $dbPath)")
                sb.AppendLine("        $conn = New-Object -ComObject ADODB.Connection")
                sb.AppendLine("        $conn.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + $dbPath)")
                sb.AppendLine("        $cmd = New-Object -ComObject ADODB.Command")
                sb.AppendLine("        $cmd.ActiveConnection = $conn")
                sb.AppendLine("        $cmd.CommandText = 'CREATE TABLE [" & tableName & "] (ID AUTOINCREMENT PRIMARY KEY, " &
                          String.Join(", ", columns.Where(Function(c) c <> "ID").Select(Function(c) "[" & Regex.Replace(c, "[^A-Za-z0-9_]", "_") & "] TEXT(255)")) & ")'")
                sb.AppendLine("        $cmd.Execute() | Out-Null")
                sb.AppendLine("        $conn.Close()")
                sb.AppendLine("        Write-Output ('Database created via ADODB: ' + $dbPath)")
                sb.AppendLine("    }")
                sb.AppendLine("} catch { Write-Output ('Create error: ' + $_.Exception.Message) }")
                sb.AppendLine("")
            End If
            If wantsAccessImport Then
                Dim importSource = ""
                For Each f In selectedFiles
                    Dim ext = Path.GetExtension(f).ToLower()
                    If ext = ".xlsx" OrElse ext = ".csv" OrElse ext = ".xls" Then
                        importSource = f
                        Exit For
                    End If
                Next
                sb.AppendLine("# IMPORT FROM EXCEL/CSV")
                sb.AppendLine("try {")
                sb.AppendLine("    $importSource = '" & importSource.Replace("'", "''") & "'")
                sb.AppendLine("    if ($importSource -ne '' -and (Test-Path $importSource)) {")
                sb.AppendLine("        $ext = [System.IO.Path]::GetExtension($importSource).ToLower()")
                sb.AppendLine("        $conn = New-Object -ComObject ADODB.Connection")
                sb.AppendLine("        $conn.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + $dbPath)")
                sb.AppendLine("        if ($ext -eq '.csv') {")
                sb.AppendLine("            $csvLines = [System.IO.File]::ReadAllLines($importSource)")
                sb.AppendLine("            $headers = $csvLines[0].Split(',') | ForEach-Object { $_.Trim().Trim('""') }")
                sb.AppendLine("            $cmd = New-Object -ComObject ADODB.Command")
                sb.AppendLine("            $cmd.ActiveConnection = $conn")
                sb.AppendLine("            for ($i = 1; $i -lt $csvLines.Length; $i++) {")
                sb.AppendLine("                $vals = $csvLines[$i].Split(',') | ForEach-Object { $_.Trim().Trim('""').Replace(""'"", ""''"") }")
                sb.AppendLine("                $cols = ($headers | ForEach-Object { '[' + $_ + ']' }) -join ', '")
                sb.AppendLine("                $vstr = ($vals | ForEach-Object { ""'"" + $_ + ""'"" }) -join ', '")
                sb.AppendLine("                $cmd.CommandText = 'INSERT INTO [" & tableName & "] (' + $cols + ') VALUES (' + $vstr + ')'")
                sb.AppendLine("                try { $cmd.Execute() | Out-Null } catch {}")
                sb.AppendLine("            }")
                sb.AppendLine("        } else {")
                sb.AppendLine("            $xlConn = New-Object -ComObject ADODB.Connection")
                sb.AppendLine("            $xlConn.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + $importSource + ';Extended Properties=""Excel 12.0 Xml;HDR=YES""')")
                sb.AppendLine("            $rs = New-Object -ComObject ADODB.Recordset")
                sb.AppendLine("            $rs.Open('SELECT * FROM [Sheet1$]', $xlConn)")
                sb.AppendLine("            while (-not $rs.EOF) {")
                sb.AppendLine("                $cmd2 = New-Object -ComObject ADODB.Command")
                sb.AppendLine("                $cmd2.ActiveConnection = $conn")
                sb.AppendLine("                $fields = @()")
                sb.AppendLine("                $values = @()")
                sb.AppendLine("                for ($f = 0; $f -lt $rs.Fields.Count; $f++) {")
                sb.AppendLine("                    $fields += '[' + $rs.Fields.Item($f).Name + ']'")
                sb.AppendLine("                    $values += ""'"" + ($rs.Fields.Item($f).Value -replace ""'"",""''"") + ""'""")
                sb.AppendLine("                }")
                sb.AppendLine("                $cmd2.CommandText = 'INSERT INTO [" & tableName & "] (' + ($fields -join ',') + ') VALUES (' + ($values -join ',') + ')'")
                sb.AppendLine("                try { $cmd2.Execute() | Out-Null } catch {}")
                sb.AppendLine("                $rs.MoveNext()")
                sb.AppendLine("            }")
                sb.AppendLine("            $rs.Close()")
                sb.AppendLine("            $xlConn.Close()")
                sb.AppendLine("        }")
                sb.AppendLine("        $conn.Close()")
                sb.AppendLine("        Write-Output ('Import complete into table: " & tableName & "')")
                sb.AppendLine("    } else { Write-Output ('No import source file found.') }")
                sb.AppendLine("} catch { Write-Output ('Import error: ' + $_.Exception.Message) }")
                sb.AppendLine("")
            End If
            If wantsAccessQuery Then
                sb.AppendLine("# QUERY DATABASE")
                sb.AppendLine("try {")
                sb.AppendLine("    $conn = New-Object -ComObject ADODB.Connection")
                sb.AppendLine("    $conn.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + $dbPath)")
                sb.AppendLine("    $rs = New-Object -ComObject ADODB.Recordset")
                sb.AppendLine("    $rs.Open('" & sqlQuery.Replace("'", "''") & "', $conn)")
                sb.AppendLine("    $output = ''")
                sb.AppendLine("    $headers = @()")
                sb.AppendLine("    for ($f = 0; $f -lt $rs.Fields.Count; $f++) { $headers += $rs.Fields.Item($f).Name }")
                sb.AppendLine("    $output += ($headers -join ' | ') + [char]13 + [char]10")
                sb.AppendLine("    $output += ('-' * 60) + [char]13 + [char]10")
                sb.AppendLine("    $rowCount = 0")
                sb.AppendLine("    while (-not $rs.EOF -and $rowCount -lt 100) {")
                sb.AppendLine("        $row = @()")
                sb.AppendLine("        for ($f = 0; $f -lt $rs.Fields.Count; $f++) { $row += [string]$rs.Fields.Item($f).Value }")
                sb.AppendLine("        $output += ($row -join ' | ') + [char]13 + [char]10")
                sb.AppendLine("        $rs.MoveNext()")
                sb.AppendLine("        $rowCount++")
                sb.AppendLine("    }")
                sb.AppendLine("    $rs.Close()")
                sb.AppendLine("    $conn.Close()")
                sb.AppendLine("    Write-Output $output")
                sb.AppendLine("} catch { Write-Output ('Query error: ' + $_.Exception.Message) }")
                sb.AppendLine("")
            End If
            If wantsAccessModify Then
                Dim modifyAction = "INSERT"
                If userIntent.Contains("delete") Then modifyAction = "DELETE"
                If userIntent.Contains("update") OrElse userIntent.Contains("modify") OrElse userIntent.Contains("edit") Then modifyAction = "UPDATE"
                sb.AppendLine("# MODIFY RECORDS")
                sb.AppendLine("try {")
                sb.AppendLine("    $conn = New-Object -ComObject ADODB.Connection")
                sb.AppendLine("    $conn.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + $dbPath)")
                sb.AppendLine("    $cmd = New-Object -ComObject ADODB.Command")
                sb.AppendLine("    $cmd.ActiveConnection = $conn")
                If modifyAction = "INSERT" AndAlso Not String.IsNullOrWhiteSpace(content) Then
                    sb.AppendLine("    $lines = [System.IO.File]::ReadAllLines('" & tempTxt.Replace("'", "''") & "')")
                    sb.AppendLine("    foreach ($line in $lines) {")
                    sb.AppendLine("        $t = $line.Trim()")
                    sb.AppendLine("        if ($t -eq '') { continue }")
                    sb.AppendLine("        $cmd.CommandText = 'INSERT INTO [" & tableName & "] (Name) VALUES (''' + $t.Replace(""'"",""''"") + ''')'")
                    sb.AppendLine("        try { $cmd.Execute() | Out-Null } catch {}")
                    sb.AppendLine("    }")
                    sb.AppendLine("    Write-Output ('Records inserted into: " & tableName & "')")
                ElseIf modifyAction = "DELETE" Then
                    Dim whereClause = "1=1"
                    Dim whereMatch = Regex.Match(userIntent, "where\s+(.+?)(?:\s+in|\s+from|$)", RegexOptions.IgnoreCase)
                    If whereMatch.Success Then whereClause = whereMatch.Groups(1).Value.Trim()
                    sb.AppendLine("    $cmd.CommandText = 'DELETE FROM [" & tableName & "] WHERE " & whereClause.Replace("'", "''") & "'")
                    sb.AppendLine("    $cmd.Execute() | Out-Null")
                    sb.AppendLine("    Write-Output ('Records deleted from: " & tableName & "')")
                Else
                    sb.AppendLine("    Write-Output ('Update: please specify SET and WHERE clauses in your request.')")
                End If
                sb.AppendLine("    $conn.Close()")
                sb.AppendLine("} catch { Write-Output ('Modify error: ' + $_.Exception.Message) }")
                sb.AppendLine("")
            End If
            If wantsAccessExport Then
                sb.AppendLine("# EXPORT TO EXCEL/PDF")
                sb.AppendLine("try {")
                sb.AppendLine("    $conn = New-Object -ComObject ADODB.Connection")
                sb.AppendLine("    $conn.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + $dbPath)")
                sb.AppendLine("    $rs = New-Object -ComObject ADODB.Recordset")
                sb.AppendLine("    $rs.Open('SELECT * FROM [" & tableName & "]', $conn)")
                sb.AppendLine("    $excel = New-Object -ComObject Excel.Application")
                sb.AppendLine("    $excel.Visible = $false")
                sb.AppendLine("    $wb = $excel.Workbooks.Add()")
                sb.AppendLine("    $ws = $wb.Worksheets.Item(1)")
                sb.AppendLine("    $ws.Name = '" & tableName.Replace("'", "''") & "'")
                sb.AppendLine("    $col = 1")
                sb.AppendLine("    for ($f = 0; $f -lt $rs.Fields.Count; $f++) {")
                sb.AppendLine("        $ws.Cells.Item(1, $col).Value = $rs.Fields.Item($f).Name")
                sb.AppendLine("        $ws.Cells.Item(1, $col).Font.Bold = $true")
                sb.AppendLine("        $col++")
                sb.AppendLine("    }")
                sb.AppendLine("    $row = 2")
                sb.AppendLine("    while (-not $rs.EOF) {")
                sb.AppendLine("        for ($f = 0; $f -lt $rs.Fields.Count; $f++) {")
                sb.AppendLine("            $ws.Cells.Item($row, $f + 1).Value = [string]$rs.Fields.Item($f).Value")
                sb.AppendLine("        }")
                sb.AppendLine("        $rs.MoveNext()")
                sb.AppendLine("        $row++")
                sb.AppendLine("    }")
                sb.AppendLine("    $ws.UsedRange.Columns.AutoFit() | Out-Null")
                sb.AppendLine("    $rs.Close()")
                sb.AppendLine("    $conn.Close()")
                sb.AppendLine("    $wb.SaveAs('" & exportPath.Replace("'", "''") & "')")
                If userIntent.Contains("pdf") Then
                    Dim pdfExport = Path.ChangeExtension(exportPath, ".pdf")
                    sb.AppendLine("    $wb.ExportAsFixedFormat(0, '" & pdfExport.Replace("'", "''") & "')")
                End If
                sb.AppendLine("    $wb.Close($false)")
                sb.AppendLine("    $excel.Quit()")
                sb.AppendLine("    Write-Output ('Exported to: " & exportPath & "')")
                sb.AppendLine("} catch { Write-Output ('Export error: ' + $_.Exception.Message) }")
                sb.AppendLine("")
            End If
            If wantsAccessReport Then
                sb.AppendLine("# CREATE REPORT VIA WORD")
                sb.AppendLine("try {")
                sb.AppendLine("    $conn = New-Object -ComObject ADODB.Connection")
                sb.AppendLine("    $conn.Open('Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + $dbPath)")
                sb.AppendLine("    $rs = New-Object -ComObject ADODB.Recordset")
                sb.AppendLine("    $rs.Open('SELECT * FROM [" & tableName & "]', $conn)")
                sb.AppendLine("    $word = New-Object -ComObject Word.Application")
                sb.AppendLine("    $word.Visible = $true")
                sb.AppendLine("    $doc = $word.Documents.Add()")
                sb.AppendLine("    $sel = $word.Selection")
                sb.AppendLine("    try { $sel.Style = $doc.Styles.Item('Title') } catch {}")
                sb.AppendLine("    $sel.TypeText('Database Report: " & tableName.Replace("'", "''") & "')")
                sb.AppendLine("    $sel.TypeParagraph()")
                sb.AppendLine("    try { $sel.Style = $doc.Styles.Item('Normal') } catch {}")
                sb.AppendLine("    $sel.TypeText('Generated: ' + (Get-Date).ToString('MMMM dd, yyyy HH:mm'))")
                sb.AppendLine("    $sel.TypeParagraph()")
                sb.AppendLine("    $sel.TypeParagraph()")
                sb.AppendLine("    $colCount = $rs.Fields.Count")
                sb.AppendLine("    if ($colCount -gt 0 -and -not $rs.EOF) {")
                sb.AppendLine("        $tbl = $doc.Tables.Add($sel.Range, 1, $colCount)")
                sb.AppendLine("        $tbl.Style = 'Table Grid'")
                sb.AppendLine("        $tbl.Rows.Item(1).Shading.BackgroundPatternColor = 4626167")
                sb.AppendLine("        $tbl.Rows.Item(1).Range.Font.Bold = $true")
                sb.AppendLine("        $tbl.Rows.Item(1).Range.Font.Color = 16777215")
                sb.AppendLine("        for ($f = 0; $f -lt $colCount; $f++) {")
                sb.AppendLine("            $tbl.Cell(1, $f + 1).Range.Text = $rs.Fields.Item($f).Name")
                sb.AppendLine("        }")
                sb.AppendLine("        $rowNum = 2")
                sb.AppendLine("        while (-not $rs.EOF -and $rowNum -le 101) {")
                sb.AppendLine("            $tbl.Rows.Add() | Out-Null")
                sb.AppendLine("            for ($f = 0; $f -lt $colCount; $f++) {")
                sb.AppendLine("                $tbl.Cell($rowNum, $f + 1).Range.Text = [string]$rs.Fields.Item($f).Value")
                sb.AppendLine("            }")
                sb.AppendLine("            $rs.MoveNext()")
                sb.AppendLine("            $rowNum++")
                sb.AppendLine("        }")
                sb.AppendLine("        $tbl.Columns.AutoFit()")
                sb.AppendLine("    }")
                sb.AppendLine("    $rs.Close()")
                sb.AppendLine("    $conn.Close()")
                Dim reportPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                           "AccessReport_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".docx")
                sb.AppendLine("    $doc.SaveAs2('" & reportPath.Replace("'", "''") & "')")
                sb.AppendLine("    $word.ActiveWindow.WindowState = 1")
                sb.AppendLine("    Write-Output ('Report saved: " & reportPath & "')")
                sb.AppendLine("} catch { Write-Output ('Report error: ' + $_.Exception.Message) }")
                sb.AppendLine("")
            End If
            If wantsAccessCreate OrElse wantsAccessModify Then
                sb.AppendLine("# Open in Access")
                sb.AppendLine("if (-not $useADODB -and $access -ne $null) {")
                sb.AppendLine("    try {")
                sb.AppendLine("        if (-not $access.CurrentDb() -or $access.CurrentDb().Name -ne $dbPath) {")
                sb.AppendLine("            $access.OpenCurrentDatabase($dbPath)")
                sb.AppendLine("        }")
                sb.AppendLine("        $access.Visible = $true")
                sb.AppendLine("        $access.UserControl = $true")
                sb.AppendLine("    } catch {}")
                sb.AppendLine("} else {")
                sb.AppendLine("    try { Start-Process $dbPath } catch {}")
                sb.AppendLine("}")
            End If
            sb.AppendLine("")
            sb.AppendLine("if (Test-Path '" & tempTxt.Replace("'", "''") & "') {")
            sb.AppendLine("    Remove-Item '" & tempTxt.Replace("'", "''") & "' -Force -ErrorAction SilentlyContinue")
            sb.AppendLine("}")
            Return RunPowerShellSync(sb.ToString())
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try
    End Function
    Private Function GetAllWindowsUIATree() As String
        Try
            Dim scriptFile = Path.Combine(Path.GetTempPath(), "uia_scan.ps1")
            File.WriteAllText(scriptFile, UIAScript)
            Dim psi As New ProcessStartInfo("powershell.exe",
                $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -File ""{scriptFile}"" all") With {
                .RedirectStandardOutput = True, .RedirectStandardError = True,
                .UseShellExecute = False, .CreateNoWindow = True
            }
            Using proc = Process.Start(psi)
                Dim output = proc.StandardOutput.ReadToEnd().Trim()
                proc.WaitForExit(20000)
                If output.Length > 12000 Then output = output.Substring(0, 12000) & "... [TRUNCATED]"
                Return If(String.IsNullOrWhiteSpace(output), "(no UIA elements found)", output)
            End Using
        Catch ex As Exception
            Return "(all-windows UIA scan failed: " & ex.Message & ")"
        End Try
    End Function
    Private Function IsRecentlyCreatedFile(savePath As String, Optional maxAgeSeconds As Integer = 120) As Boolean
        Try
            If File.Exists(savePath) Then
                Dim fileAge = DateTime.Now - File.GetCreationTime(savePath)
                If fileAge.TotalSeconds < maxAgeSeconds Then
                    Return True
                End If
            End If
            Dim folder = Path.GetDirectoryName(savePath)
            Dim ext = Path.GetExtension(savePath)
            If Directory.Exists(folder) Then
                Dim recentFiles = Directory.GetFiles(folder, "*" & ext, SearchOption.TopDirectoryOnly).
                Where(Function(f) (DateTime.Now - File.GetCreationTime(f)).TotalSeconds < 60).
                ToArray()
                If recentFiles.Length > 0 Then
                    Return True
                End If
            End If
        Catch
        End Try
        Return False
    End Function
    Private Structure AppEntry
        Public Name As String
        Public AUMID As String
    End Structure
    Private Class PlaywrightInstallHelper
        Public Shared Async Function InstallBrowsersAsync() As Task(Of Boolean)
            Try
                Return Await Task.Run(Function()
                                          Dim exitCode = Microsoft.Playwright.Program.Main(New String() {"install", "chromium"})
                                          Return exitCode = 0
                                      End Function)
            Catch
                Return False
            End Try
        End Function
    End Class
    Private Class WebWorkflow
        <JsonProperty("steps")>
        Public Property Steps As List(Of WebStep)
        <JsonProperty("keepOpen")>
        Public Property KeepOpen As Boolean
    End Class
    Private Class WebStep
        <JsonProperty("action")>
        Public Property Action As String
        <JsonProperty("url")>
        Public Property Url As String
        <JsonProperty("selector")>
        Public Property Selector As String
        <JsonProperty("value")>
        Public Property Value As String
        <JsonProperty("timeoutMs")>
        Public Property TimeoutMs As Integer
        <JsonProperty("observeBefore")>
        Public Property ObserveBefore As Boolean
        <JsonProperty("description")>
        Public Property Description As String
    End Class
    Private Class WebAutomationResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property AccessibilityTree As String
        Public Property ElementMap As String
        Public Property KeepOpen As Boolean
        Public Property ScreenshotPath As String
        Public Property DomPath As String
        Public Property ConsoleLogPath As String
        Public Property OcrText As String
        Public Property ExtractedText As String
        Public Property CurrentUrl As String
        Public Property PageTitle As String
        Public Property WindowInFocus As Boolean
        Public Property BrowserProcessId As Integer
        Public Property StepsCompleted As Integer
        Public Property FailedStepDescription As String
    End Class
    Private Class FileAnalysisResult
        Public Property Text As String
        Public Property Metadata As Dictionary(Of String, String)
    End Class
    Private Class FileAnalysisHelper
        Public Shared Async Function AnalyzeFileAsync(filePath As String) As Task(Of FileAnalysisResult)
            Return Await Task.Run(Function()
                                      Dim ext = Path.GetExtension(filePath).ToLowerInvariant()
                                      Dim result As New FileAnalysisResult With {
                                          .Text = "",
                                          .Metadata = New Dictionary(Of String, String)()
                                      }
                                      Select Case ext
                                          Case ".docx" : result.Text = ExtractWordText(filePath)
                                          Case ".pptx" : result.Text = ExtractPptText(filePath)
                                          Case ".xlsx" : result.Text = ExtractExcelText(filePath)
                                          Case ".pdf" : result.Text = ExtractPdfText(filePath)
                                          Case ".txt", ".csv", ".json", ".xml", ".html", ".htm"
                                              result.Text = File.ReadAllText(filePath)
                                          Case ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".tif"
                                              result.Text = OcrHelper.ExtractTextFromImage(filePath)
                                          Case Else
                                              result.Text = $"Unsupported file type: {ext}"
                                      End Select
                                      Return result
                                  End Function)
        End Function
        Private Shared Function ExtractPdfText(filePath As String) As String
            Dim sb As New StringBuilder()
            Try
                Using document = UglyToad.PdfPig.PdfDocument.Open(filePath)
                    For pageNum = 1 To document.NumberOfPages
                        sb.AppendLine(document.GetPage(pageNum).Text)
                    Next
                End Using
            Catch ex As Exception
                Return $"(PDF Error: {ex.Message})"
            End Try
            Return sb.ToString()
        End Function
        Private Shared Function ExtractWordText(filePath As String) As String
            Using doc = WordprocessingDocument.Open(filePath, False)
                Return doc.MainDocumentPart.Document.Body.InnerText
            End Using
        End Function
        Private Shared Function ExtractPptText(filePath As String) As String
            Dim sb As New StringBuilder()
            Using ppt = PresentationDocument.Open(filePath, False)
                For Each slidePart In ppt.PresentationPart.SlideParts
                    For Each t In slidePart.Slide.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)()
                        sb.AppendLine(t.InnerText)
                    Next
                Next
            End Using
            Return sb.ToString()
        End Function
        Private Shared Function ExtractExcelText(filePath As String) As String
            Dim sb As New StringBuilder()
            Using doc = SpreadsheetDocument.Open(filePath, False)
                Dim wbPart = doc.WorkbookPart
                For Each wsPart In wbPart.WorksheetParts
                    Dim sheetData = wsPart.Worksheet.Elements(Of SheetData)().FirstOrDefault()
                    If sheetData Is Nothing Then Continue For
                    For Each row In sheetData.Elements(Of Row)()
                        Dim rowValues As New List(Of String)
                        For Each cell In row.Elements(Of Cell)()
                            rowValues.Add(GetCellValue(doc, cell))
                        Next
                        sb.AppendLine(String.Join(",", rowValues))
                    Next
                Next
            End Using
            Return sb.ToString()
        End Function
        Private Shared Function GetCellValue(doc As SpreadsheetDocument, cell As Cell) As String
            If cell Is Nothing Then Return ""
            Dim value = cell.InnerText
            If String.IsNullOrEmpty(value) Then Return ""
            If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString Then
                Dim sst = doc.WorkbookPart.SharedStringTablePart?.SharedStringTable
                If sst Is Nothing Then Return value
                Dim index As Integer
                If Integer.TryParse(value, index) AndAlso index < sst.ChildElements.Count Then
                    Return sst.ChildElements(index).InnerText
                End If
                Return value
            End If
            If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.InlineString Then
                Return cell.InnerText
            End If
            If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.Boolean Then
                Return If(value = "1", "TRUE", "FALSE")
            End If
            If cell.DataType Is Nothing Then
                Dim d As Double
                If Double.TryParse(value, d) Then
                    If cell.StyleIndex IsNot Nothing Then
                        Try
                            Dim stylesheet = doc.WorkbookPart.WorkbookStylesPart?.Stylesheet
                            If stylesheet IsNot Nothing Then
                                Dim styleIdx = CInt(cell.StyleIndex.Value)
                                Dim cellFormat = stylesheet.CellFormats?.ChildElements(styleIdx)
                                If cellFormat IsNot Nothing Then
                                    Dim nfType = cellFormat.GetType()
                                    Dim nfIdProp = nfType.GetProperty("NumberFormatId")
                                    If nfIdProp IsNot Nothing Then
                                        Dim nfId = CInt(nfIdProp.GetValue(cellFormat))
                                        Dim dateFormatIds = {14, 15, 16, 17, 18, 19, 20, 21, 22,
                                                            27, 28, 29, 30, 31, 32, 33, 34, 35, 36,
                                                            45, 46, 47, 50, 51, 52, 53, 54, 55, 56, 57, 58}
                                        If dateFormatIds.Contains(nfId) Then
                                            Try
                                                Return DateTime.FromOADate(d).ToShortDateString()
                                            Catch
                                            End Try
                                        End If
                                    End If
                                End If
                            End If
                        Catch
                        End Try
                    End If
                End If
                Return value
            End If
            Return value
        End Function
    End Class
    Private Class OcrHelper
        Private Shared ReadOnly TessDataPath As String = Path.Combine(Application.StartupPath, "tessdata")
        Public Shared Function IsOcrAvailable() As Boolean
            Return File.Exists(Path.Combine(TessDataPath, "eng.traineddata"))
        End Function
        Public Shared Function ExtractTextFromImage(imagePath As String) As String
            Try
                If Not IsOcrAvailable() Then
                    Return "(OCR not available - missing tessdata/eng.traineddata)"
                End If
                Using engine As New TesseractEngine(TessDataPath, "eng", EngineMode.Default)
                    Using img = Pix.LoadFromFile(imagePath)
                        Using ocrPage = engine.Process(img)
                            Return ocrPage.GetText()
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                Return $"(OCR Error: {ex.Message})"
            End Try
        End Function
    End Class
    Friend Class NativeMethods
        <DllImport("user32.dll")>
        Public Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
        End Function
        <DllImport("user32.dll")>
        Public Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function
        <DllImport("user32.dll")>
        Public Shared Function GetForegroundWindow() As IntPtr
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function GetWindowThreadProcessId(hwnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
        End Function
    End Class
End Class
Public Class ScenePreprocessor
    Public Shared Function CleanScene(rawScene As String) As String
        If String.IsNullOrWhiteSpace(rawScene) Then Return rawScene
        Dim lines = rawScene.Split({Environment.NewLine, Chr(10).ToString()}, StringSplitOptions.None).ToList()
        lines = lines.Where(Function(l)
                                Dim t = l.Trim()
                                If t.StartsWith("===") Then Return True
                                If t.Contains("App: 'AutonoMe'") OrElse t.Contains("App: ""AutonoMe""") Then Return False
                                Return True
                            End Function).ToList()
        lines = MergeConsecutiveTypeLines(lines)
        lines = MergeTypeAroundTaskbarSearch(lines)
        lines = MergeConsecutiveTypeLines(lines)
        lines = lines.Select(Function(l)
                                 If Not IsTypeLine(l) Then Return l
                                 Dim m = Regex.Match(l, "\[TYPE\] Typed Text: ""(.*?)""\s*\|.*$")
                                 If m.Success Then Return $"[TYPE] Typed Text: ""{m.Groups(1).Value}"""
                                 Return l
                             End Function).ToList()
        Return String.Join(Environment.NewLine, lines)
    End Function
    Private Shared Function MergeConsecutiveTypeLines(lines As List(Of String)) As List(Of String)
        Dim result As New List(Of String)
        Dim i As Integer = 0
        While i < lines.Count
            Dim line = lines(i)
            If IsTypeLine(line) Then
                Dim accumulated = ExtractTypedText(line)
                Dim currentApp = ExtractAppFromLine(line)
                While i + 1 < lines.Count AndAlso IsTypeLine(lines(i + 1))
                    Dim nextApp = ExtractAppFromLine(lines(i + 1))
                    If Not String.IsNullOrEmpty(currentApp) AndAlso
                       Not String.IsNullOrEmpty(nextApp) AndAlso
                       Not currentApp.Equals(nextApp, StringComparison.OrdinalIgnoreCase) Then
                        Exit While
                    End If
                    i += 1
                    accumulated &= ExtractTypedText(lines(i))
                    If String.IsNullOrEmpty(currentApp) Then currentApp = nextApp
                End While
                If Not String.IsNullOrEmpty(currentApp) Then
                    result.Add($"[TYPE] Typed Text: ""{accumulated}"" | App: '{currentApp}'")
                Else
                    result.Add($"[TYPE] Typed Text: ""{accumulated}""")
                End If
            Else
                result.Add(line)
            End If
            i += 1
        End While
        Return result
    End Function
    Private Shared Function MergeTypeAroundTaskbarSearch(lines As List(Of String)) As List(Of String)
        Dim result As New List(Of String)(lines)
        Dim changed As Boolean = True
        While changed
            changed = False
            For i = 0 To result.Count - 3
                If Not IsTypeLine(result(i)) Then Continue For
                For j = i + 1 To Math.Min(i + 4, result.Count - 1)
                    If result(j).Contains("[CLICK:TASKBAR_SEARCH]") Then
                        If j + 1 < result.Count AndAlso IsTypeLine(result(j + 1)) Then
                            Dim typeApp = ExtractAppFromLine(result(i))
                            Dim nextApp = ExtractAppFromLine(result(j + 1))
                            Dim searchApps = {"explorer", "searchhost", "searchapp", "searchui", ""}
                            Dim typeAppIsSarch = searchApps.Contains(typeApp.ToLower())
                            Dim nextAppIsSearch = searchApps.Contains(nextApp.ToLower())
                            If nextAppIsSearch OrElse String.IsNullOrEmpty(nextApp) Then
                                Dim appLaunched As Boolean = False
                                For k = j + 1 To Math.Min(j + 5, result.Count - 1)
                                    Dim kLine = result(k).Trim()
                                    If kLine.StartsWith("[CLICK") AndAlso
                                       Not kLine.Contains("[CLICK:TASKBAR_SEARCH]") AndAlso
                                       Not kLine.Contains("App: 'explorer'") AndAlso
                                       Not kLine.Contains("App: 'searchhost'") AndAlso
                                       Not kLine.Contains("App: 'searchapp'") Then
                                        appLaunched = True
                                        Exit For
                                    End If
                                Next
                                If Not appLaunched Then
                                    Dim mergedText = ExtractTypedText(result(i)) & ExtractTypedText(result(j + 1))
                                    Dim appTag = If(Not String.IsNullOrEmpty(typeApp), $" | App: '{typeApp}'", "")
                                    result(i) = $"[TYPE] Typed Text: ""{mergedText}""{appTag}"
                                    result.RemoveAt(j + 1)
                                    changed = True
                                    Exit For
                                End If
                            End If
                        End If
                        Exit For
                    End If
                    If IsActionLine(result(j)) AndAlso Not result(j).Contains("[CLICK:TASKBAR_SEARCH]") Then
                        Exit For
                    End If
                Next
                If changed Then Exit For
            Next
        End While
        Return result
    End Function
    Private Shared Function IsTypeLine(line As String) As Boolean
        Return line.TrimStart().StartsWith("[TYPE]")
    End Function
    Private Shared Function IsActionLine(line As String) As Boolean
        Dim t = line.TrimStart()
        Return t.StartsWith("[CLICK") OrElse t.StartsWith("[TYPE") OrElse t.StartsWith("[KEYPRESS")
    End Function
    Private Shared Function ExtractTypedText(typeLine As String) As String
        Dim m = Regex.Match(typeLine, "\[TYPE\].*?""(.*?)""")
        If m.Success Then Return m.Groups(1).Value
        Dim q1 = typeLine.IndexOf(""""c)
        Dim q2 = typeLine.LastIndexOf(""""c)
        If q1 >= 0 AndAlso q2 > q1 Then Return typeLine.Substring(q1 + 1, q2 - q1 - 1)
        Return ""
    End Function
    Private Shared Function ExtractAppFromLine(line As String) As String
        Dim m = Regex.Match(line, "App:\s*'([^']*)'")
        If m.Success Then Return m.Groups(1).Value
        Return ""
    End Function
End Class
Public Class MacroRecorder
    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const WH_MOUSE_LL As Integer = 14
    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const VK_SHIFT As Integer = &H10
    Private Const VK_CONTROL As Integer = &H11
    Private Const VK_MENU As Integer = &H12
    Private Const VK_CAPITAL As Integer = &H14
    Private Const VK_LWIN As Integer = &H5B
    Private Shared mouseHookID As IntPtr = IntPtr.Zero
    Private Shared keyboardHookID As IntPtr = IntPtr.Zero
    Private Shared ReadOnly mouseProc As LowLevelHookProc = AddressOf MouseHookCallback
    Private Shared ReadOnly keyboardProc As LowLevelHookProc = AddressOf KeyboardHookCallback
    Private Shared ReadOnly lockObj As New Object()
    Public Shared RecordedSteps As New List(Of String)
    Private Shared currentTypedText As New StringBuilder()
    Public Shared IsRecording As Boolean = False
    Private Shared ReadOnly _http As New HttpClient() With {.Timeout = TimeSpan.FromSeconds(5)}
    Private Shared currentSceneName As String = ""
    Private Delegate Function LowLevelHookProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetWindowsHookEx(idHook As Integer, lpfn As LowLevelHookProc, hMod As IntPtr, dwThreadId As UInteger) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function UnhookWindowsHookEx(hhk As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function CallNextHookEx(hhk As IntPtr, nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetModuleHandle(lpModuleName As String) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, ExactSpelling:=True)>
    Private Shared Function GetKeyState(keyCode As Integer) As Short
    End Function
    <DllImport("user32.dll")>
    Private Shared Function WindowFromPoint(p As POINT) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(hwnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowRect(hwnd As IntPtr, ByRef rect As NativeRect) As Boolean
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxCount As Integer) As Integer
    End Function
    Private Shared lastForegroundHwnd As IntPtr = IntPtr.Zero
    <StructLayout(LayoutKind.Sequential)>
    Private Structure POINT
        Public x As Integer
        Public y As Integer
    End Structure
    <StructLayout(LayoutKind.Sequential)>
    Private Structure MSLLHOOKSTRUCT
        Public pt As POINT
        Public mouseData As UInteger
        Public flags As UInteger
        Public time As UInteger
        Public dwExtraInfo As IntPtr
    End Structure
    <StructLayout(LayoutKind.Sequential)>
    Private Structure NativeRect
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure
    Public Shared Sub StartRecording(savePath As String)
        If IsRecording Then Return
        SyncLock lockObj
            RecordedSteps.Clear()
            currentTypedText.Clear()
            currentSceneName = savePath
            lastForegroundHwnd = IntPtr.Zero
            RecordedSteps.Add("=== AUTONO-ME SCENE RECORDING STARTED AT " & DateTime.Now & " ===")
        End SyncLock
        Using curProcess As Process = Process.GetCurrentProcess()
            Using curModule As ProcessModule = curProcess.MainModule
                Dim modHandle = GetModuleHandle(curModule.ModuleName)
                mouseHookID = SetWindowsHookEx(WH_MOUSE_LL, mouseProc, modHandle, 0)
                keyboardHookID = SetWindowsHookEx(WH_KEYBOARD_LL, keyboardProc, modHandle, 0)
            End Using
        End Using
        IsRecording = True
    End Sub
    Public Shared Sub StopRecording(savePath As String)
        If Not IsRecording Then Return
        UnhookWindowsHookEx(mouseHookID)
        UnhookWindowsHookEx(keyboardHookID)
        IsRecording = False
        SyncLock lockObj
            FlushTypedText()
            RecordedSteps.Add("=== SCENE RECORDING STOPPED ===")
            File.WriteAllLines(savePath, RecordedSteps)
        End SyncLock
    End Sub
    Private Shared Sub FlushTypedText()
        If currentTypedText.Length > 0 Then
            Dim appName As String = ""
            Try
                If lastForegroundHwnd <> IntPtr.Zero Then
                    Dim pid As UInteger = 0
                    GetWindowThreadProcessId(lastForegroundHwnd, pid)
                    If pid > 0 Then
                        Dim proc = Process.GetProcessById(CInt(pid))
                        appName = proc.ProcessName
                    End If
                End If
            Catch
            End Try
            If Not String.IsNullOrEmpty(appName) Then
                RecordedSteps.Add($"[TYPE] Typed Text: ""{currentTypedText.ToString()}"" | App: '{appName}'")
            Else
                RecordedSteps.Add($"[TYPE] Typed Text: ""{currentTypedText.ToString()}""")
            End If
            currentTypedText.Clear()
        End If
    End Sub
    Private Shared Function IsOwnProcess(hwnd As IntPtr) As Boolean
        Dim pid As UInteger = 0
        GetWindowThreadProcessId(hwnd, pid)
        If pid = CUInt(Process.GetCurrentProcess().Id) Then Return True
        Try
            Dim sb As New StringBuilder(256)
            GetWindowText(hwnd, sb, 256)
            Dim title = sb.ToString()
            If title.Contains("Autono-Me") OrElse title.Contains("AutonoMe") Then Return True
        Catch
        End Try
        Return False
    End Function
    Private Shared lastSearchQuery As String = ""
    Private Shared lastSearchTime As DateTime = DateTime.MinValue
    Private Shared Function MouseHookCallback(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        If nCode >= 0 AndAlso wParam = CType(WM_LBUTTONDOWN, IntPtr) Then
            Dim hookStruct = Marshal.PtrToStructure(Of MSLLHOOKSTRUCT)(lParam)
            Dim hwnd = WindowFromPoint(hookStruct.pt)
            If IsOwnProcess(hwnd) Then
                Return CallNextHookEx(mouseHookID, nCode, wParam, lParam)
            End If
            Task.Run(Sub()
                         Try
                             Dim pt As New System.Windows.Point(hookStruct.pt.x, hookStruct.pt.y)
                             Dim el = AutomationElement.FromPoint(pt)
                             Dim name = el.Current.Name
                             Dim controlType = el.Current.LocalizedControlType
                             Dim autoId = el.Current.AutomationId
                             Dim procName = "UnknownApp"
                             Dim windowTitle = "UnknownWindow"
                             Try
                                 Dim p = Process.GetProcessById(el.Current.ProcessId)
                                 procName = p.ProcessName
                                 windowTitle = p.MainWindowTitle
                             Catch
                             End Try
                             If procName.Equals("AutonoMe", StringComparison.OrdinalIgnoreCase) OrElse
                            windowTitle.Contains("Autono-Me") OrElse windowTitle.Contains("AutonoMe") Then Return
                             If String.IsNullOrWhiteSpace(name) Then name = "(unnamed)"
                             Dim isTaskbarSearch = hookStruct.pt.y > Screen.PrimaryScreen.Bounds.Height - 80 AndAlso
                             (procName.ToLower() = "explorer" OrElse procName.ToLower() = "searchhost" OrElse
                              procName.ToLower() = "searchapp" OrElse String.IsNullOrWhiteSpace(windowTitle))
                             Dim isSearchResultClick = (procName.ToLower() = "searchhost" OrElse
                             procName.ToLower() = "searchapp") AndAlso
                             hookStruct.pt.y < Screen.PrimaryScreen.Bounds.Height - 80
                             Dim isBrowser = {"chrome", "msedge", "brave"}.Contains(procName.ToLower())
                             Dim stepLog As String
                             If isTaskbarSearch Then
                                 Threading.Thread.Sleep(2000)
                                 SyncLock lockObj
                                     lastSearchQuery = currentTypedText.ToString()
                                     lastSearchTime = DateTime.Now
                                 End SyncLock
                                 stepLog = $"[CLICK:TASKBAR_SEARCH] {controlType} | Name: '{name}' | App: '{procName}' | X:{hookStruct.pt.x} Y:{hookStruct.pt.y}"
                             ElseIf isSearchResultClick Then
                                 Dim clickContext = ""
                                 If Not String.IsNullOrWhiteSpace(name) AndAlso name <> "(unnamed)" Then
                                     clickContext = name
                                 End If
                                 Dim isAdminClick = name.IndexOf("administrator", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                 name.IndexOf("admin", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                 autoId.IndexOf("admin", StringComparison.OrdinalIgnoreCase) >= 0
                                 Dim isOpenClick = name.IndexOf("Open", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                 name.IndexOf("Run", StringComparison.OrdinalIgnoreCase) >= 0
                                 SyncLock lockObj
                                     If isAdminClick Then
                                         stepLog = $"[CLICK:SEARCH_RESULT_ADMIN] {controlType} | Name: '{name}' | SearchQuery: '{lastSearchQuery}' | App: '{procName}' | X:{hookStruct.pt.x} Y:{hookStruct.pt.y}"
                                         FlushTypedText()
                                         RecordedSteps.Add(stepLog)
                                     ElseIf isOpenClick OrElse Not String.IsNullOrWhiteSpace(clickContext) Then
                                         stepLog = $"[CLICK:SEARCH_RESULT] {controlType} | Name: '{name}' | SearchQuery: '{lastSearchQuery}' | App: '{procName}' | X:{hookStruct.pt.x} Y:{hookStruct.pt.y}"
                                         FlushTypedText()
                                         RecordedSteps.Add(stepLog)
                                     Else
                                         stepLog = $"[CLICK:SEARCH_RESULT] {controlType} | Name: '{name}' | ID: '{autoId}' | SearchQuery: '{lastSearchQuery}' | App: '{procName}' | X:{hookStruct.pt.x} Y:{hookStruct.pt.y}"
                                         FlushTypedText()
                                         RecordedSteps.Add(stepLog)
                                     End If
                                 End SyncLock
                                 If isBrowser Then DomExtractor.ExtractAndSave(currentSceneName)
                                 Return
                             ElseIf isBrowser Then
                                 Threading.Thread.Sleep(150)
                                 Dim webInfo = GetBrowserElementAtPoint(procName, hookStruct.pt.x, hookStruct.pt.y, windowTitle)
                                 If Not String.IsNullOrWhiteSpace(webInfo) Then
                                     stepLog = $"[CLICK:WEB] {webInfo} | App: '{procName}' | X:{hookStruct.pt.x} Y:{hookStruct.pt.y}"
                                 Else
                                     stepLog = $"[CLICK] {controlType} | Name: '{name}' | ID: '{autoId}' | App: '{procName}' | X:{hookStruct.pt.x} Y:{hookStruct.pt.y}"
                                 End If
                             Else
                                 If (DateTime.Now - lastSearchTime).TotalSeconds < 15 AndAlso
                                Not String.IsNullOrWhiteSpace(lastSearchQuery) AndAlso
                                procName.ToLower() <> "searchhost" AndAlso
                                procName.ToLower() <> "searchapp" AndAlso
                                procName.ToLower() <> "explorer" Then
                                     SyncLock lockObj
                                         FlushTypedText()
                                         RecordedSteps.Add($"[SEARCH_LAUNCH] Launched '{procName}' from search query '{lastSearchQuery}' | Window: '{windowTitle}'")
                                     End SyncLock
                                     lastSearchQuery = ""
                                     lastSearchTime = DateTime.MinValue
                                 End If
                                 stepLog = $"[CLICK] {controlType} | Name: '{name}' | ID: '{autoId}' | App: '{procName}' | Window: '{windowTitle}' | X:{hookStruct.pt.x} Y:{hookStruct.pt.y}"
                             End If
                             SyncLock lockObj
                                 FlushTypedText()
                                 RecordedSteps.Add(stepLog)
                             End SyncLock
                             If isBrowser Then DomExtractor.ExtractAndSave(currentSceneName)
                         Catch
                             SyncLock lockObj
                                 FlushTypedText()
                                 RecordedSteps.Add($"[CLICK] X:{hookStruct.pt.x} Y:{hookStruct.pt.y} (UIA element not readable)")
                             End SyncLock
                         End Try
                     End Sub)
        End If
        Return CallNextHookEx(mouseHookID, nCode, wParam, lParam)
    End Function
    Private Shared Function KeyboardHookCallback(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        If nCode >= 0 AndAlso wParam = CType(WM_KEYDOWN, IntPtr) Then
            Dim fgHwnd As IntPtr = GetForegroundWindow()
            If IsOwnProcess(fgHwnd) Then
                Return CallNextHookEx(keyboardHookID, nCode, wParam, lParam)
            End If
            If fgHwnd <> lastForegroundHwnd AndAlso lastForegroundHwnd <> IntPtr.Zero Then
                SyncLock lockObj
                    FlushTypedText()
                End SyncLock
            End If
            lastForegroundHwnd = fgHwnd
            Dim vkCode As Integer = Marshal.ReadInt32(lParam)
            Dim key As Keys = CType(vkCode, Keys)
            Dim isCtrlDown As Boolean = (GetKeyState(VK_CONTROL) And &H8000) <> 0
            Dim isAltDown As Boolean = (GetKeyState(VK_MENU) And &H8000) <> 0
            Dim isShiftDown As Boolean = (GetKeyState(VK_SHIFT) And &H8000) <> 0
            Dim isWinDown As Boolean = (GetKeyState(VK_LWIN) And &H8000) <> 0
            Dim isCapsLockOn As Boolean = (GetKeyState(VK_CAPITAL) And &H1) <> 0
            If key = Keys.ControlKey OrElse key = Keys.ShiftKey OrElse
           key = Keys.Menu OrElse key = Keys.LWin OrElse key = Keys.RWin Then
                Return CallNextHookEx(keyboardHookID, nCode, wParam, lParam)
            End If
            If isCtrlDown OrElse isAltDown OrElse isWinDown Then
                SyncLock lockObj
                    FlushTypedText()
                    Dim combo As String = ""
                    If isCtrlDown Then combo &= "Ctrl+"
                    If isAltDown Then combo &= "Alt+"
                    If isShiftDown Then combo &= "Shift+"
                    If isWinDown Then combo &= "Win+"
                    combo &= key.ToString()
                    RecordedSteps.Add($"[KEYPRESS] {combo}")
                End SyncLock
                Return CallNextHookEx(keyboardHookID, nCode, wParam, lParam)
            End If
            If key = Keys.Enter OrElse key = Keys.Tab OrElse key = Keys.Escape OrElse
           key = Keys.Back OrElse key = Keys.Delete Then
                SyncLock lockObj
                    FlushTypedText()
                    RecordedSteps.Add($"[KEYPRESS] {key.ToString()}")
                End SyncLock
            Else
                Dim isUpper As Boolean = isShiftDown Xor isCapsLockOn
                Dim keyChar As String = ""
                If key >= Keys.D0 AndAlso key <= Keys.D9 Then
                    If isShiftDown Then
                        Dim shiftNums As String = ")!@#$%^&*("
                        keyChar = shiftNums(key - Keys.D0).ToString()
                    Else
                        keyChar = ChrW(vkCode).ToString()
                    End If
                ElseIf key >= Keys.A AndAlso key <= Keys.Z Then
                    keyChar = If(isUpper, ChrW(vkCode).ToString().ToUpper(), ChrW(vkCode).ToString().ToLower())
                ElseIf key = Keys.Space Then
                    keyChar = " "
                ElseIf key = Keys.OemPeriod Then
                    keyChar = If(isShiftDown, ">", ".")
                ElseIf key = Keys.Oemcomma Then
                    keyChar = If(isShiftDown, "<", ",")
                ElseIf key = Keys.OemQuestion Then
                    keyChar = If(isShiftDown, "?", "/")
                ElseIf key = Keys.OemMinus Then
                    keyChar = If(isShiftDown, "_", "-")
                ElseIf key = Keys.Oemplus Then
                    keyChar = If(isShiftDown, "+", "=")
                ElseIf key = Keys.OemSemicolon Then
                    keyChar = If(isShiftDown, ":", ";")
                ElseIf key = Keys.OemQuotes Then
                    keyChar = If(isShiftDown, """", "'")
                ElseIf key = Keys.OemOpenBrackets Then
                    keyChar = If(isShiftDown, "{", "[")
                ElseIf key = Keys.OemCloseBrackets Then
                    keyChar = If(isShiftDown, "}", "]")
                ElseIf key = Keys.OemPipe Then
                    keyChar = If(isShiftDown, "|", "\")
                ElseIf key = Keys.Oemtilde Then
                    keyChar = If(isShiftDown, "~", "`")
                End If
                If Not String.IsNullOrEmpty(keyChar) Then
                    SyncLock lockObj
                        currentTypedText.Append(keyChar)
                    End SyncLock
                End If
            End If
        End If
        Return CallNextHookEx(keyboardHookID, nCode, wParam, lParam)
    End Function
    Private Shared Function GetBrowserElementAtPoint(browserProc As String, screenX As Integer,
                                                      screenY As Integer, windowTitle As String) As String
        Try
            Dim ports = {9222, 9223, 9224}
            Dim cdpBase = ""
            For Each port In ports
                Try
                    Dim response = _http.GetStringAsync($"http://localhost:{port}/json/version").Result
                    If Not String.IsNullOrWhiteSpace(response) Then
                        cdpBase = $"http://localhost:{port}"
                        Exit For
                    End If
                Catch
                End Try
            Next
            If String.IsNullOrWhiteSpace(cdpBase) Then Return ""
            Return ""
        Catch
            Return ""
        End Try
    End Function
    Public Class DomExtractor
        Private Shared ReadOnly _http As New HttpClient() With {.Timeout = TimeSpan.FromSeconds(5)}
        Public Shared Sub ExtractAndSave(sceneSavePath As String)
            Task.Run(Sub()
                         Try
                         Catch
                         End Try
                     End Sub)
        End Sub
    End Class
    Public Class ScreenRecorder
        Private Shared recordingThread As Threading.Thread
        Private Shared isRecording As Boolean = False
        Private Shared currentVideoPath As String = ""
        Private Shared frameList As New List(Of String)
        Private Shared ReadOnly frameLock As New Object()
        Private Shared frameFolder As String = ""
        Private Const MaxFrames As Integer = 600
        Private Const FrameIntervalMs As Integer = 500
        Private Const CaptureScale As Double = 0.5
        Public Shared Sub StartRecording(savePath As String)
            isRecording = False
            SyncLock frameLock
                frameList.Clear()
            End SyncLock
            currentVideoPath = savePath
            frameFolder = savePath & "_frames"
            If Directory.Exists(frameFolder) Then Directory.Delete(frameFolder, True)
            Directory.CreateDirectory(frameFolder)
            isRecording = True
            recordingThread = New Threading.Thread(AddressOf CaptureLoop)
            recordingThread.IsBackground = True
            recordingThread.Priority = Threading.ThreadPriority.BelowNormal
            recordingThread.Start()
        End Sub
        Public Shared Sub StopRecording()
            isRecording = False
            If recordingThread IsNot Nothing Then recordingThread.Join(3000)
            CompileVideo()
        End Sub
        Private Shared Sub CaptureLoop()
            Dim frameIndex As Integer = 0
            Dim bounds = Screen.PrimaryScreen.Bounds
            Dim captureWidth = CInt(bounds.Width * CaptureScale)
            Dim captureHeight = CInt(bounds.Height * CaptureScale)
            While isRecording AndAlso frameIndex < MaxFrames
                Try
                    Using fullBmp As New Bitmap(bounds.Width, bounds.Height)
                        Using g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(fullBmp)
                            g.CopyFromScreen(bounds.Location, System.Drawing.Point.Empty, bounds.Size)
                        End Using
                        Using scaled As New Bitmap(captureWidth, captureHeight)
                            Using g2 As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(scaled)
                                g2.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Bilinear
                                g2.DrawImage(fullBmp, 0, 0, captureWidth, captureHeight)
                            End Using
                            Dim framePath = Path.Combine(frameFolder, $"frame_{frameIndex:D6}.jpg")
                            Dim encoder = Imaging.ImageCodecInfo.GetImageEncoders().First(Function(enc) enc.MimeType = "image/jpeg")
                            Dim encoderParams As New Imaging.EncoderParameters(1)
                            encoderParams.Param(0) = New Imaging.EncoderParameter(Imaging.Encoder.Quality, 50L)
                            scaled.Save(framePath, encoder, encoderParams)
                            SyncLock frameLock
                                frameList.Add(framePath)
                            End SyncLock
                            frameIndex += 1
                        End Using
                    End Using
                Catch
                End Try
                System.Threading.Thread.Sleep(FrameIntervalMs)
            End While
        End Sub
        Private Shared Sub CompileVideo()
            Try
                Dim listFile = Path.Combine(frameFolder, "frames.txt")
                Dim sb As New StringBuilder()
                SyncLock frameLock
                    For Each f In frameList
                        sb.AppendLine($"file '{f.Replace("'", "'\\''")}'")
                        sb.AppendLine("duration 0.5")
                    Next
                End SyncLock
                File.WriteAllText(listFile, sb.ToString())
                Dim ffmpegPath = FindFFmpeg()
                If Not String.IsNullOrWhiteSpace(ffmpegPath) Then
                    Dim args = $"-f concat -safe 0 -i ""{listFile}"" -vf ""scale=1280:-2"" -r 2 -c:v libx264 -preset ultrafast -pix_fmt yuv420p ""{currentVideoPath}.mp4"" -y"
                    Dim psi As New ProcessStartInfo(ffmpegPath, args) With {
                        .UseShellExecute = False, .CreateNoWindow = True, .RedirectStandardError = True
                    }
                    Dim proc = Process.Start(psi)
                    proc.WaitForExit(30000)
                    If File.Exists(currentVideoPath & ".mp4") Then
                        Directory.Delete(frameFolder, True)
                    End If
                End If
            Catch
            End Try
        End Sub
        Private Shared Function FindFFmpeg() As String
            Dim candidates = {
                "ffmpeg.exe",
                Path.Combine(Application.StartupPath, "ffmpeg.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ffmpeg", "bin", "ffmpeg.exe")
            }
            For Each c In candidates
                If File.Exists(c) Then Return c
            Next
            Try
                Dim psi As New ProcessStartInfo("where", "ffmpeg") With {
                    .UseShellExecute = False, .RedirectStandardOutput = True, .CreateNoWindow = True
                }
                Dim proc = Process.Start(psi)
                Dim output = proc.StandardOutput.ReadLine()
                proc.WaitForExit(2000)
                If Not String.IsNullOrWhiteSpace(output) AndAlso File.Exists(output.Trim()) Then Return output.Trim()
            Catch
            End Try
            Return ""
        End Function
        Public Shared Function GetVideoPath() As String
            Return If(File.Exists(currentVideoPath & ".mp4"), currentVideoPath & ".mp4", "")
        End Function
        Public Shared Function GetFrameFolder() As String
            Return frameFolder
        End Function
        Public Shared Function GetFramePaths() As List(Of String)
            SyncLock frameLock
                Return New List(Of String)(frameList)
            End SyncLock
        End Function
    End Class
    Public Class ConversationEntry
        Public Property Role As String
        Public Property Content As String
        Public Sub New(r As String, c As String)
            Role = r
            Content = c
        End Sub
    End Class
End Class

'Autono-Me
'© 2026 Copyright Oranyx Labs/Elliot Monteverde All Rights Reserved
'GNU General Public License v3.0