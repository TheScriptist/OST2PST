<#
.SYNOPSIS
    Convertit un fichier OST orphelin (Outlook 2013/2016/2019/365) en fichier PST.

.DESCRIPTION
    1. Extrait le contenu de l'OST avec libpff (open source), via un Python portable
       téléchargé automatiquement dans %LOCALAPPDATA%\ost2pst (aucun droit administrateur requis).
    2. Recrée un fichier PST via Outlook (COM) : dossiers, e-mails (expéditeur,
       destinataires, dates, lu/non lu, pièces jointes), contacts, rendez-vous,
       tâches et notes.

    Sans paramètre, une fenêtre s'ouvre pour choisir les fichiers.

.PARAMETER OstPath
    Fichier .ost à convertir.

.PARAMETER PstPath
    Fichier .pst à créer (ne doit pas exister).

.PARAMETER IncludeRecovered
    Ajoute aussi les éléments supprimés que pffexport arrive à récupérer,
    dans un dossier "Éléments récupérés".

.PARAMETER ExportDir
    Réutilise une extraction existante (chemin du dossier ".export")
    au lieu de relancer l'extraction.

.PARAMETER KeepExport
    Conserve les fichiers extraits (supprimés par défaut à la fin).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Convertir-OST-en-PST.ps1 -OstPath C:\ost\boite.ost -PstPath C:\ost\boite.pst
#>
[CmdletBinding()]
param(
    [string]$OstPath,
    [string]$PstPath,
    [switch]$IncludeRecovered,
    [string]$ExportDir,
    [switch]$KeepExport
)

$ErrorActionPreference = 'Stop'

# --- Constantes Outlook -------------------------------------------------------
$olAppointmentItem = 1; $olContactItem = 2; $olTaskItem = 3; $olNoteItem = 5; $olPostItem = 6
$olFolderCalendar = 9; $olFolderContacts = 10; $olFolderNotes = 12; $olFolderTasks = 13; $olFolderInbox = 6
$olStoreUnicode = 3; $olMSGUnicode = 9; $olByValue = 1; $olDiscard = 1

$PT = 'http://schemas.microsoft.com/mapi/proptag/'
$PR = @{
    DeliveryTime      = $PT + '0x0E060040'
    SubmitTime        = $PT + '0x00390040'
    SenderName        = $PT + '0x0C1A001F'
    SenderEmail       = $PT + '0x0C1F001F'
    SenderAddrType    = $PT + '0x0C1E001F'
    SentRepName       = $PT + '0x0042001F'
    SentRepEmail      = $PT + '0x0065001F'
    SentRepAddrType   = $PT + '0x0064001F'
    DisplayTo         = $PT + '0x0E04001F'
    DisplayCc         = $PT + '0x0E03001F'
    TransportHeaders  = $PT + '0x007D001F'
}

$script:LogAction = { param($msg) Write-Host $msg }
$script:ProgressAction = { param($done, $total) }
$script:LogFile = $null

function Write-Log([string]$msg) {
    $line = '{0:HH:mm:ss}  {1}' -f (Get-Date), $msg
    & $script:LogAction $line
    if ($script:LogFile) { try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch {} }
}

# --- Lecture des fichiers produits par pffexport ------------------------------
function Read-TextAuto([string]$path) {
    $b = [IO.File]::ReadAllBytes($path)
    if ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF) { return [Text.Encoding]::UTF8.GetString($b, 3, $b.Length - 3) }
    if ($b.Length -ge 2 -and $b[0] -eq 0xFF -and $b[1] -eq 0xFE) { return [Text.Encoding]::Unicode.GetString($b, 2, $b.Length - 2) }
    try { return (New-Object Text.UTF8Encoding($false, $true)).GetString($b) }
    catch { return [Text.Encoding]::GetEncoding(1252).GetString($b) }
}

function Read-KeyValues([string]$path) {
    $h = @{}
    if (-not (Test-Path -LiteralPath $path)) { return $h }
    foreach ($line in ((Read-TextAuto $path) -split "`r?`n")) {
        if ($line -match '^([^\t:][^:]*?):\t+(.*)$') {
            $k = $matches[1].Trim(); $v = $matches[2].Trim()
            if ($v -and -not $h.ContainsKey($k)) { $h[$k] = $v }
        }
    }
    return $h
}

function Read-Recipients([string]$path) {
    $list = @()
    if (-not (Test-Path -LiteralPath $path)) { return $list }
    foreach ($block in ((Read-TextAuto $path) -split '(?:\r?\n){2,}')) {
        $r = @{}
        foreach ($line in ($block -split "`r?`n")) {
            if ($line -match '^([^\t:][^:]*?):\t+(.*)$') { $r[$matches[1].Trim()] = $matches[2].Trim() }
        }
        if ($r.Count -gt 0) {
            $name = $r['Display name']; if (-not $name) { $name = $r['Recipient display name'] }
            if (-not $name) { $name = $r['Email address'] }
            $type = $r['Recipient type']; if (-not $type) { $type = 'To' }
            $list += [pscustomobject]@{ Name = $name; Email = $r['Email address']; Type = $type }
        }
    }
    return $list
}

# "Jun 05, 2014 16:22:32.000000000 UTC" -> DateTime UTC
function ConvertFrom-PffDate([string]$s) {
    if ($s -and $s -match '^([A-Z][a-z]{2} \d{2}, \d{4} \d{2}:\d{2}:\d{2})') {
        $d = [datetime]::ParseExact($matches[1], 'MMM dd, yyyy HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
        if ($d.Year -lt 1980 -or $d.Year -gt 2200) { return $null }
        return [datetime]::SpecifyKind($d, 'Utc')
    }
    return $null
}

function Test-IsItemDir($dir) {
    return ($dir.Name -match '^(Message|Contact|Appointment|Meeting|Task|Note|Activity|Document|DistributionList|RSSFeed)\d{5}$')
}

function Get-ItemKind($dir) { return ($dir.Name -replace '\d{5}$', '') }

# Fusionne les fichiers d'information de l'élément (OutlookHeaders.txt, Contact.txt, Appointment.txt...)
function Get-ItemInfo([string]$dirPath) {
    $info = @{}
    foreach ($f in (Get-ChildItem -LiteralPath $dirPath -File -Filter '*.txt')) {
        if ($f.Name -in @('Message.txt', 'InternetHeaders.txt', 'ConversationIndex.txt', 'Recipients.txt', 'ItemValues.txt')) { continue }
        $kv = Read-KeyValues $f.FullName
        foreach ($k in $kv.Keys) { if (-not $info.ContainsKey($k)) { $info[$k] = $kv[$k] } }
    }
    return $info
}

function Get-Body([string]$dirPath) {
    $html = Join-Path $dirPath 'Message.html'
    $rtf = Join-Path $dirPath 'Message.rtf'
    $txt = Join-Path $dirPath 'Message.txt'
    $b = @{ Html = $null; Rtf = $null; Text = $null }
    if (Test-Path -LiteralPath $html) { $b.Html = Read-TextAuto $html }
    if (Test-Path -LiteralPath $rtf) { $b.Rtf = [IO.File]::ReadAllBytes($rtf) }
    if (Test-Path -LiteralPath $txt) { $b.Text = Read-TextAuto $txt }
    if (-not $b.Text -and $b.Html) {
        $t = $b.Html -replace '(?is)<(script|style)[^>]*>.*?</\1>', '' -replace '(?i)<br\s*/?>|</p>|</div>|</tr>', "`r`n" -replace '<[^>]+>', ''
        $b.Text = [Net.WebUtility]::HtmlDecode($t).Trim()
    }
    return $b
}

# --- Création des éléments dans Outlook ---------------------------------------
function New-OlItem($folder, [int]$type) {
    if ($folder) { return $folder.Items.Add($type) }
    return $script:Outlook.CreateItem($type)
}

function Set-Prop($item, [string]$tag, $value) {
    if ($null -eq $value -or ($value -is [string] -and -not $value)) { return }
    try { $item.PropertyAccessor.SetProperty($tag, $value) } catch {}
}

function Set-IfValue($item, [string]$prop, $value) {
    if ($null -eq $value -or $value -eq '' -or $value -eq 'Unknown') { return }
    try { $item.$prop = $value } catch {}
}

function Get-SafeFileName([string]$name) {
    if (-not $name) { $name = 'element' }
    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $name = $name.Replace([string]$c, '_') }
    if ($name.Length -gt 80) { $name = $name.Substring(0, 80) }
    return $name.Trim()
}

function Add-Attachments($item, [string]$dirPath) {
    $attDir = Join-Path $dirPath 'Attachments'
    if (-not (Test-Path -LiteralPath $attDir)) { return }

    foreach ($f in (Get-ChildItem -LiteralPath $attDir -File)) {
        # pffexport préfixe les noms par "N_" : on restaure le nom d'origine
        $name = $f.Name -replace '^\d+_', ''
        $tmp = Join-Path $script:TempAttDir ([guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path $tmp -Force | Out-Null
        $dest = Join-Path $tmp $name
        Copy-Item -LiteralPath $f.FullName -Destination $dest
        try { $null = $item.Attachments.Add($dest, $olByValue) }
        catch { Write-Log ("   Pièce jointe ignorée ({0}) : {1}" -f $name, $_.Exception.Message) }
    }

    # Éléments joints (mail transféré en pièce jointe, invitation...) : recréés en .msg
    foreach ($sub in (Get-ChildItem -LiteralPath $attDir -Directory)) {
        foreach ($emb in (Get-ChildItem -LiteralPath $sub.FullName -Directory | Where-Object { Test-IsItemDir $_ })) {
            $e = $null
            try {
                $e = New-ItemFromDir $emb $null
                if (-not $e) { continue }
                $tmp = Join-Path $script:TempAttDir ([guid]::NewGuid().ToString('N').Substring(0, 8))
                New-Item -ItemType Directory -Path $tmp -Force | Out-Null
                $msgPath = Join-Path $tmp ((Get-SafeFileName $e.Subject) + '.msg')
                $e.SaveAs($msgPath, $olMSGUnicode)
                $null = $item.Attachments.Add($msgPath, $olByValue)
            }
            catch { Write-Log ("   Élément joint ignoré : {0}" -f $_.Exception.Message) }
            finally {
                if ($e) { try { $e.Close($olDiscard) } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($e) }
            }
        }
    }
}

function Set-MailContent($item, $info, $body, [string]$dirPath) {
    Set-IfValue $item 'Subject' $info['Subject']
    if ($body.Html) { $item.HTMLBody = $body.Html }
    elseif ($body.Rtf) { try { $item.RTFBody = $body.Rtf } catch { $item.Body = $body.Text } }
    elseif ($body.Text) { $item.Body = $body.Text }
    switch ($info['Importance']) { 'High' { $item.Importance = 2 } 'Low' { $item.Importance = 0 } }
    if ($info['Sensitivity'] -eq 'Private') { try { $item.Sensitivity = 2 } catch {} }
    Add-Attachments $item $dirPath
}

# E-mail : on crée un "Post" puis on le transforme en IPM.Note.
# Ainsi le message apparaît comme reçu (et non comme un brouillon "pas encore envoyé").
function New-MailItem($folder, $dirPath, $info, $body) {
    $item = New-OlItem $folder $olPostItem
    $item.MessageClass = 'IPM.Note'
    Set-MailContent $item $info $body $dirPath
    $flags = $info['Flags']
    $item.UnRead = -not ($flags -and $flags -match '\bRead\b')
    if ($folder) { $item.Save() }

    $recips = Read-Recipients (Join-Path $dirPath 'Recipients.txt')
    $to = ($recips | Where-Object { $_.Type -eq 'To' } | ForEach-Object { $_.Name }) -join '; '
    $cc = ($recips | Where-Object { $_.Type -eq 'CC' } | ForEach-Object { $_.Name }) -join '; '

    $senderName = $info['Sender name']; $senderMail = $info['Sender email address']
    $repName = $info['Sent representing name']; $repMail = $info['Sent representing email address']
    if (-not $repName) { $repName = $senderName }; if (-not $repMail) { $repMail = $senderMail }
    if ($senderName -eq 'Unknown') { $senderName = $null }; if ($repName -eq 'Unknown') { $repName = $null }
    if ($senderMail -eq 'Unknown') { $senderMail = $null }; if ($repMail -eq 'Unknown') { $repMail = $null }

    Set-Prop $item $PR.SenderName $senderName
    Set-Prop $item $PR.SenderEmail $senderMail
    Set-Prop $item $PR.SentRepName $repName
    Set-Prop $item $PR.SentRepEmail $repMail
    if ($senderMail -and $senderMail -match '@') { Set-Prop $item $PR.SenderAddrType 'SMTP' }
    if ($repMail -and $repMail -match '@') { Set-Prop $item $PR.SentRepAddrType 'SMTP' }
    Set-Prop $item $PR.DisplayTo $to
    Set-Prop $item $PR.DisplayCc $cc

    $hdr = Join-Path $dirPath 'InternetHeaders.txt'
    if (Test-Path -LiteralPath $hdr) { Set-Prop $item $PR.TransportHeaders (Read-TextAuto $hdr) }

    $delivery = ConvertFrom-PffDate $info['Delivery time']
    $submit = ConvertFrom-PffDate $info['Client submit time']
    if (-not $delivery) { $delivery = $submit }; if (-not $submit) { $submit = $delivery }
    Set-Prop $item $PR.DeliveryTime $delivery
    Set-Prop $item $PR.SubmitTime $submit

    if ($folder) { $item.Save() }
    return $item
}

function New-ContactItem($folder, $info, $body) {
    $item = New-OlItem $folder $olContactItem
    $map = [ordered]@{
        'Given name' = 'FirstName'; 'Surname' = 'LastName'; 'Initials' = 'Initials'; 'Title' = 'Title'
        'Generational abbreviation' = 'Suffix'; 'File under' = 'FileAs'; 'Company name' = 'CompanyName'
        'Job title' = 'JobTitle'; 'Department name' = 'Department'; 'Office location' = 'OfficeLocation'
        'Email address 1' = 'Email1Address'; 'Email address 2' = 'Email2Address'; 'Email address 3' = 'Email3Address'
        'Business phone number 1' = 'BusinessTelephoneNumber'; 'Business phone number 2' = 'Business2TelephoneNumber'
        'Business fax number' = 'BusinessFaxNumber'; 'Home phone number' = 'HomeTelephoneNumber'
        'Mobile phone number' = 'MobileTelephoneNumber'; 'Primary phone number' = 'PrimaryTelephoneNumber'
        'Callback phone number' = 'CallbackTelephoneNumber'; 'Locality' = 'BusinessAddressCity'
        'Country' = 'BusinessAddressCountry'; 'Postal address' = 'MailingAddress'
    }
    foreach ($k in $map.Keys) { Set-IfValue $item $map[$k] $info[$k] }
    if (-not $info['Given name'] -and -not $info['Surname']) { Set-IfValue $item 'FullName' $info['Display name'] }
    if ($body.Text) { $item.Body = $body.Text }
    if ($folder) { $item.Save() }
    return $item
}

function New-AppointmentItem($folder, $dirPath, $info, $body) {
    $item = New-OlItem $folder $olAppointmentItem
    Set-IfValue $item 'Subject' $info['Subject']
    Set-IfValue $item 'Location' $info['Location']
    $start = ConvertFrom-PffDate $info['Start time']
    $end = ConvertFrom-PffDate $info['End time']
    if ($start) {
        $item.StartUTC = $start
        if ($end) { $item.EndUTC = $end }
        elseif ($info['Duration'] -match '^(\d+)') { $item.Duration = [int]$matches[1] }
    }
    $text = $body.Text
    if ($info['Recurrence pattern']) { $text = ("[Récurrence d'origine : {0}]`r`n`r`n{1}" -f $info['Recurrence pattern'], $text) }
    if ($body.Rtf -and -not $info['Recurrence pattern']) { try { $item.RTFBody = $body.Rtf } catch { $item.Body = $text } }
    elseif ($text) { $item.Body = $text }
    $item.ReminderSet = ($info['Is a reminder'] -eq 'yes')
    if ($info['Busy status'] -match '^0x([0-9a-fA-F]+)') {
        $bs = [Convert]::ToInt32($matches[1], 16); if ($bs -le 5) { try { $item.BusyStatus = $bs } catch {} }
    }
    if ($info['Is private'] -eq 'yes') { try { $item.Sensitivity = 2 } catch {} }
    Add-Attachments $item $dirPath
    if ($folder) { $item.Save() }
    return $item
}

function New-TaskItem($folder, $info, $body) {
    $item = New-OlItem $folder $olTaskItem
    Set-IfValue $item 'Subject' $info['Subject']
    if ($body.Text) { $item.Body = $body.Text }
    $s = ConvertFrom-PffDate $info['Start date']; if ($s) { try { $item.StartDate = $s.ToLocalTime().Date } catch {} }
    $d = ConvertFrom-PffDate $info['Due date']; if ($d) { try { $item.DueDate = $d.ToLocalTime().Date } catch {} }
    if ($info['Percentage complete'] -match '^(\d+)') { try { $item.PercentComplete = [int]$matches[1] } catch {} }
    if ($info['Is complete'] -eq 'yes') { try { $item.Complete = $true } catch {} }
    if ($folder) { $item.Save() }
    return $item
}

function New-NoteItem($folder, $info, $body) {
    $item = New-OlItem $folder $olNoteItem
    $text = $body.Text
    if ($info['Subject'] -and ($null -eq $text -or -not $text.StartsWith($info['Subject']))) { $text = $info['Subject'] + "`r`n" + $text }
    $item.Body = $text
    if ($folder) { $item.Save() }
    return $item
}

# Crée l'élément correspondant au dossier pffexport. $folder = $null : élément non enregistré (pièce jointe).
function New-ItemFromDir($dir, $folder) {
    $kind = Get-ItemKind $dir
    $info = Get-ItemInfo $dir.FullName
    $body = Get-Body $dir.FullName
    switch ($kind) {
        'Contact' { return New-ContactItem $folder $info $body }
        'Appointment' { return New-AppointmentItem $folder $dir.FullName $info $body }
        'Meeting' { return New-AppointmentItem $folder $dir.FullName $info $body }
        'Task' { return New-TaskItem $folder $info $body }
        'Note' { return New-NoteItem $folder $info $body }
        'DistributionList' { $script:Stats.Skipped++; Write-Log ("   Liste de distribution ignorée : {0}" -f $info['Display name']); return $null }
        default { return New-MailItem $folder $dir.FullName $info $body }
    }
}

# --- Parcours de l'arborescence -----------------------------------------------
function Get-FolderType($dirPath) {
    $kinds = Get-ChildItem -LiteralPath $dirPath -Directory | Where-Object { Test-IsItemDir $_ } | ForEach-Object { Get-ItemKind $_ } | Group-Object | Sort-Object Count -Descending
    if (-not $kinds) { return $null }
    switch ($kinds[0].Name) {
        'Contact' { return $olFolderContacts }
        'Appointment' { return $olFolderCalendar }
        'Meeting' { return $olFolderCalendar }
        'Task' { return $olFolderTasks }
        'Note' { return $olFolderNotes }
        default { return $olFolderInbox }
    }
}

function Test-HasItems($dirPath) {
    foreach ($d in (Get-ChildItem -LiteralPath $dirPath -Directory)) {
        if (Test-IsItemDir $d) { return $true }
        if (Test-HasItems $d.FullName) { return $true }
    }
    return $false
}

function Get-OrAddFolder($parent, [string]$name, $type) {
    foreach ($f in $parent.Folders) { if ($f.Name -eq $name) { return $f } }
    if ($type) { return $parent.Folders.Add($name, $type) }
    return $parent.Folders.Add($name)
}

$script:CancelRequested = $false
function Test-Cancel {
    if ($script:CancelRequested) { throw (New-Object OperationCanceledException 'Conversion abandonnée.') }
}

function Import-ItemDir($dir, $olFolder) {
    Test-Cancel
    $item = $null
    try {
        $item = New-ItemFromDir $dir $olFolder
        if ($item) { $script:Stats.Imported++ }
    }
    catch {
        $script:Stats.Errors++
        Write-Log ("   Erreur sur {0} : {1}" -f $dir.FullName, $_.Exception.Message)
    }
    finally {
        if ($item) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($item) }
        $script:Stats.Done++
        if ($script:Stats.Done % 50 -eq 0) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
        & $script:ProgressAction $script:Stats.Done $script:Stats.Total
    }
}

function Import-Folder([string]$dirPath, $olFolder, [string]$displayPath) {
    $subs = Get-ChildItem -LiteralPath $dirPath -Directory
    $items = @($subs | Where-Object { Test-IsItemDir $_ })
    if ($items.Count -gt 0) { Write-Log ("{0} : {1} élément(s)" -f $displayPath, $items.Count) }
    foreach ($it in $items) { Import-ItemDir $it $olFolder }

    foreach ($sub in ($subs | Where-Object { -not (Test-IsItemDir $_) })) {
        if (-not (Test-HasItems $sub.FullName)) { continue }
        $child = Get-OrAddFolder $olFolder $sub.Name (Get-FolderType $sub.FullName)
        Import-Folder $sub.FullName $child ($displayPath + '\' + $sub.Name)
    }
}

function Count-Items([string]$dirPath) {
    $n = 0
    foreach ($d in (Get-ChildItem -LiteralPath $dirPath -Directory)) {
        if (Test-IsItemDir $d) { $n++ } else { $n += Count-Items $d.FullName }
    }
    return $n
}

# Point de départ : le dossier IPM_SUBTREE (dossiers visibles de la boîte), sinon la racine du fichier.
function Get-MailboxRoot([string]$exportDir) {
    $tops = @(Get-ChildItem -LiteralPath $exportDir -Directory | Where-Object { $_.Name -ne 'Root - Public' })
    foreach ($t in $tops) {
        $ipm = Join-Path $t.FullName 'IPM_SUBTREE'
        if (Test-Path -LiteralPath $ipm) { return $ipm }
    }
    if ($tops.Count -eq 1) { return $tops[0].FullName }
    return $exportDir
}

# --- Extraction (Python portable + libpff, sans droits administrateur) -------
$PythonVersion = '3.12.10'
$script:CacheDir = Join-Path $env:LOCALAPPDATA 'ost2pst'

function Get-Download([string]$url, [string]$dest) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $url -OutFile $dest -UseBasicParsing
}

# Installe (une seule fois) un Python portable et la bibliothèque libpff dans %LOCALAPPDATA%\ost2pst
function Initialize-Python {
    $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64' -or $env:PROCESSOR_ARCHITEW6432 -eq 'ARM64') { 'arm64' } elseif ([Environment]::Is64BitOperatingSystem) { 'amd64' } else { 'win32' }
    $pyDir = Join-Path $script:CacheDir 'python'
    $pyExe = Join-Path $pyDir 'python.exe'
    New-Item -ItemType Directory -Path $script:CacheDir -Force | Out-Null

    if (-not (Test-Path -LiteralPath $pyExe)) {
        Write-Log "Téléchargement de Python portable (une seule fois, environ 10 Mo)..."
        $zip = Join-Path $script:CacheDir 'python.zip'
        $embed = if ($arch -eq 'win32') { 'embed-win32' } else { "embed-$arch" }
        try { Get-Download "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-$embed.zip" $zip }
        catch { throw "Téléchargement de Python impossible (connexion Internet ou proxy d'entreprise ?) : $($_.Exception.Message)" }
        Expand-Archive -LiteralPath $zip -DestinationPath $pyDir -Force
        Remove-Item -LiteralPath $zip -Force
    }

    if (-not (Get-ChildItem -LiteralPath $pyDir -Filter 'pypff*.pyd' -ErrorAction SilentlyContinue)) {
        Write-Log "Téléchargement de la bibliothèque libpff (une seule fois)..."
        try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
            $meta = Invoke-RestMethod -Uri 'https://pypi.org/pypi/libpff-python/json' -UseBasicParsing
            $tag = 'cp312-cp312-win_' + $arch + '.whl'
            $wheel = $meta.urls | Where-Object { $_.filename -like "*$tag" } | Select-Object -First 1
            if (-not $wheel) { throw "aucune version de libpff pour $tag" }
            $whl = Join-Path $script:CacheDir 'libpff.zip'
            Get-Download $wheel.url $whl
        }
        catch { throw "Téléchargement de libpff impossible : $($_.Exception.Message)" }
        $tmp = Join-Path $script:CacheDir 'libpff_tmp'
        Expand-Archive -LiteralPath $whl -DestinationPath $tmp -Force
        Get-ChildItem -LiteralPath $tmp -Recurse -Filter '*.pyd' | Copy-Item -Destination $pyDir -Force
        Remove-Item -LiteralPath $tmp, $whl -Recurse -Force
    }

    $script:PythonExe = $pyExe
    $script:Extractor = Join-Path $script:CacheDir 'extraire_ost.py'
    [IO.File]::WriteAllText($script:Extractor, $ExtractorSource, (New-Object Text.UTF8Encoding($false)))
}

function Invoke-OstExtraction([string]$ost, [string]$workDir) {
    Initialize-Python
    $target = Join-Path $workDir 'boite'
    $errLog = Join-Path $workDir 'extraction.log'
    Write-Log "Extraction de l'OST (peut être long pour une grosse boîte)..."

    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = $script:PythonExe
    $psi.Arguments = '"{0}" "{1}" "{2}" "{3}"' -f $script:Extractor, $ost, $target, $errLog
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $p = [Diagnostics.Process]::Start($psi)
    while (-not $p.HasExited) {
        if ($script:Gui) { [System.Windows.Forms.Application]::DoEvents() }
        if ($script:CancelRequested) {
            try { $p.Kill() } catch {}
            $p.WaitForExit()
            Test-Cancel
        }
        Start-Sleep -Milliseconds 200
    }
    $p.WaitForExit()
    $rc = $p.ExitCode

    if (Test-Path -LiteralPath $errLog) {
        $lines = @(Get-Content -LiteralPath $errLog -Encoding UTF8)
        if ($lines.Count -gt 0) {
            Write-Log ("Avertissements de l'extraction ({0}) :" -f $lines.Count)
            $lines | Select-Object -Last 15 | ForEach-Object { Write-Log ('   ' + $_) }
            if ($script:LogFile) { Add-Content -LiteralPath $script:LogFile -Value $lines -Encoding UTF8 }
        }
    }
    $exp = $target + '.export'
    if ($rc -ne 0 -or -not (Test-Path -LiteralPath $exp)) {
        throw "L'extraction a échoué (code $rc). Le fichier OST est peut-être endommagé ou chiffré. Détails dans le journal."
    }
    return $exp
}

# Script Python d'extraction (écrit dans %LOCALAPPDATA%\ost2pst au lancement)
$ExtractorSource = @'
# Extrait un fichier OST/PST avec libpff (pypff) dans la même arborescence que pffexport :
#   <cible>.export\<dossiers>\Message00001\{OutlookHeaders.txt, Recipients.txt, InternetHeaders.txt,
#                                         Message.html|rtf|txt, Attachments\...}
#   <cible>.orphans\...   (éléments orphelins, récupérables)
# Usage : python extraire_ost.py <fichier.ost> <cible> [journal]
import os
import struct
import sys
import uuid

import pypff

PSETID_ADDRESS = '00062004-0000-0000-c000-000000000046'
PSETID_APPOINTMENT = '00062002-0000-0000-c000-000000000046'
PSETID_TASK = '00062003-0000-0000-c000-000000000046'
PSETID_COMMON = '00062008-0000-0000-c000-000000000046'

SKIP_CLASSES = ('IPM.CONFIGURATION', 'IPM.MICROSOFT', 'IPM.MESSAGEMANAGER', 'IPM.EXTENDEDRULE',
                'IPM.RULEORGANIZER', 'IPM.AGGREGATION', 'IPC.')
INVALID = '<>:"/\\|?*'

stats = {'items': 0, 'errors': 0}


def load_named_map(f):
    """Table des propriétés nommées : tag (0x8000+) -> (guid, id)."""
    result = {}
    try:
        rs = f.name_to_id_map.get_record_set(0)
        data = {e.entry_type: e.data for e in rs.entries}
        g, es = data.get(2) or b'', data.get(3) or b''
        guids = [str(uuid.UUID(bytes_le=g[i:i + 16])) for i in range(0, len(g) - 15, 16)]
        for i in range(0, len(es) - 7, 8):
            pid, gi, idx = struct.unpack('<IHH', es[i:i + 8])
            if gi & 1:
                continue
            gx = gi >> 1
            guid = {1: 'mapi', 2: 'public'}.get(gx) or (guids[gx - 3] if 0 <= gx - 3 < len(guids) else None)
            result[0x8000 + idx] = (guid, pid)
    except Exception:
        pass
    return result


def props(item, named):
    """Propriétés de l'élément : {tag: entrée} et {(guid, id): entrée}."""
    fixed, nm = {}, {}
    try:
        if item.number_of_record_sets == 0:
            return fixed, nm
        for e in item.get_record_set(0).entries:
            fixed[e.entry_type] = e
            if e.entry_type >= 0x8000 and e.entry_type in named:
                nm[named[e.entry_type]] = e
    except Exception:
        pass
    return fixed, nm


def val(e):
    if e is None:
        return None
    try:
        t = e.value_type
        if t in (0x1f, 0x1e):
            return e.data_as_string
        if t == 0x40:
            return e.data_as_datetime
        if t == 0x0b:
            return bool(e.data and e.data[0])
        if t in (0x02, 0x03, 0x14):
            return e.data_as_integer
        if t == 0x05:
            return struct.unpack('<d', e.data[:8])[0]
        return e.data
    except Exception:
        return None


def fmt_date(d):
    if d is None:
        return None
    try:
        return d.strftime('%b %d, %Y %H:%M:%S') + '.000000000 UTC'
    except Exception:
        return None


def safe(name, maxlen=60):
    name = ''.join('_' if (c in INVALID or ord(c) < 32) else c for c in (name or ''))
    name = name.strip().rstrip('. ')
    return (name[:maxlen].rstrip('. ') or 'sans_nom')


def write_kv(path, sections):
    with open(path, 'w', encoding='utf-8', newline='\n') as fh:
        for title, pairs in sections:
            fh.write(title + ':\n')
            for k, v in pairs:
                if v is None or v == '':
                    continue
                fh.write('%s:\t\t\t%s\n' % (k, v))
            fh.write('\n')


def kind_of(cls):
    c = (cls or 'IPM.Note').upper()
    if c.startswith(SKIP_CLASSES):
        return None
    if c.startswith('IPM.CONTACT'):
        return 'Contact'
    if c.startswith('IPM.DISTLIST'):
        return 'DistributionList'
    if c.startswith('IPM.APPOINTMENT'):
        return 'Appointment'
    if c.startswith('IPM.TASK') and not c.startswith('IPM.TASKREQUEST'):
        return 'Task'
    if c.startswith('IPM.STICKYNOTE'):
        return 'Note'
    return 'Message'


def decode_html(raw, cpid):
    if not raw:
        return None
    enc = {65001: 'utf-8', 1252: 'cp1252', 28591: 'latin-1', 20127: 'ascii', 28605: 'iso8859-15'}.get(cpid)
    if not enc and cpid:
        enc = 'cp%d' % cpid
    for e in (enc, 'utf-8', 'cp1252'):
        if not e:
            continue
        try:
            return raw.decode(e)
        except Exception:
            continue
    return raw.decode('cp1252', errors='replace')


def export_item(m, path, named, kind):
    os.makedirs(path, exist_ok=True)
    fx, nm = props(m, named)
    g = lambda tag: val(fx.get(tag))
    n = lambda guid, pid: val(nm.get((guid, pid)))

    flags = g(0x0E07) or 0
    importance = {0: 'Low', 1: 'Normal', 2: 'High'}.get(g(0x0017))
    sensitivity = 'Private' if g(0x0036) == 2 else 'None'
    message = [
        ('Client submit time', fmt_date(g(0x0039))),
        ('Delivery time', fmt_date(g(0x0E06))),
        ('Creation time', fmt_date(g(0x3007))),
        ('Flags', '0x%08x (%s)' % (flags, 'Read' if flags & 1 else 'Unread')),
        ('Display name', g(0x3001)),
        ('Subject', g(0x0037)),
        ('Sender name', g(0x0C1A)),
        ('Sender email address', g(0x5D01) or g(0x0C1F)),
        ('Sent representing name', g(0x0042)),
        ('Sent representing email address', g(0x5D02) or g(0x0065)),
        ('Importance', importance),
        ('Sensitivity', sensitivity),
        ('Is a reminder', 'yes' if n(PSETID_COMMON, 0x8503) else 'no'),
        ('Is private', 'yes' if sensitivity == 'Private' else 'no'),
    ]

    if kind == 'Contact':
        details = [
            ('File under', n(PSETID_ADDRESS, 0x8005)), ('Given name', g(0x3A06)), ('Initials', g(0x3A0A)),
            ('Surname', g(0x3A11)), ('Generational abbreviation', g(0x3A05)), ('Company name', g(0x3A16)),
            ('Job title', g(0x3A17)), ('Department name', g(0x3A18)), ('Office location', g(0x3A19)),
            ('Country', g(0x3A26)), ('Locality', g(0x3A27)), ('Postal address', g(0x3A15)),
            ('Business phone number 1', g(0x3A08)), ('Business phone number 2', g(0x3A1B)),
            ('Business fax number', g(0x3A24)), ('Home phone number', g(0x3A09)),
            ('Mobile phone number', g(0x3A1C)), ('Primary phone number', g(0x3A1A)),
            ('Callback phone number', g(0x3A02)),
            ('Email address 1', n(PSETID_ADDRESS, 0x8083)), ('Email address 2', n(PSETID_ADDRESS, 0x8093)),
            ('Email address 3', n(PSETID_ADDRESS, 0x80A3)),
        ]
        write_kv(os.path.join(path, 'Contact.txt'), [('Message', message), ('Contact', details)])
    elif kind == 'Appointment':
        busy = n(PSETID_APPOINTMENT, 0x8205)
        details = [
            ('Start time', fmt_date(n(PSETID_APPOINTMENT, 0x820D) or g(0x0060))),
            ('End time', fmt_date(n(PSETID_APPOINTMENT, 0x820E) or g(0x0061))),
            ('Location', n(PSETID_APPOINTMENT, 0x8208)),
            ('Recurrence pattern', n(PSETID_APPOINTMENT, 0x8232) if n(PSETID_APPOINTMENT, 0x8223) else None),
            ('Busy status', '0x%08x' % busy if isinstance(busy, int) else None),
        ]
        write_kv(os.path.join(path, 'Appointment.txt'), [('Message', message), ('Appointment', details)])
    elif kind == 'Task':
        pct = n(PSETID_TASK, 0x8102)
        details = [
            ('Start date', fmt_date(n(PSETID_TASK, 0x8104))),
            ('Due date', fmt_date(n(PSETID_TASK, 0x8105))),
            ('Percentage complete', '%d' % round(pct * 100) if isinstance(pct, float) else None),
            ('Is complete', 'yes' if n(PSETID_TASK, 0x811C) else 'no'),
        ]
        write_kv(os.path.join(path, 'Task.txt'), [('Message', message), ('Task', details)])
    elif kind == 'DistributionList':
        write_kv(os.path.join(path, 'DistributionList.txt'), [('Message', message)])
    else:
        write_kv(os.path.join(path, 'OutlookHeaders.txt'), [('Message', message)])

    # Destinataires
    try:
        r = m.recipients if hasattr(m, 'recipients') else None
        if r is not None and r.number_of_record_sets:
            with open(os.path.join(path, 'Recipients.txt'), 'w', encoding='utf-8', newline='\n') as fh:
                for i in range(r.number_of_record_sets):
                    rp = {e.entry_type: e for e in r.get_record_set(i).entries}
                    rtype = {1: 'To', 2: 'CC', 3: 'BCC'}.get(val(rp.get(0x0C15)) if rp.get(0x0C15) is not None else 1, 'To')
                    email = val(rp.get(0x39FE)) or val(rp.get(0x3003))
                    fh.write('Display name:\t\t%s\nEmail address:\t\t%s\nRecipient type:\t\t%s\n\n'
                             % (val(rp.get(0x3001)) or email or '', email or '', rtype))
    except Exception:
        pass

    # En-têtes Internet
    try:
        h = m.transport_headers
        if h:
            with open(os.path.join(path, 'InternetHeaders.txt'), 'w', encoding='utf-8', newline='\r\n') as fh:
                fh.write(h.replace('\r\n', '\n'))
    except Exception:
        pass

    # Corps du message
    try:
        html = decode_html(m.html_body, g(0x3FDE))
        if html:
            with open(os.path.join(path, 'Message.html'), 'w', encoding='utf-8') as fh:
                fh.write(html)
        else:
            rtf = None
            try:
                rtf = m.rtf_body
            except Exception:
                pass
            if rtf:
                with open(os.path.join(path, 'Message.rtf'), 'wb') as fh:
                    fh.write(rtf)
        txt = m.plain_text_body
        if txt:
            with open(os.path.join(path, 'Message.txt'), 'wb') as fh:
                fh.write(txt if isinstance(txt, bytes) else txt.encode('utf-8'))
    except Exception:
        pass

    # Pièces jointes
    try:
        count = m.number_of_attachments
    except Exception:
        count = 0
    for i in range(count):
        try:
            a = m.get_attachment(i)
            ap, _ = props(a, named)
            method = val(ap.get(0x3705))
            att_dir = os.path.join(path, 'Attachments')
            os.makedirs(att_dir, exist_ok=True)
            if method == 5:
                if not a.number_of_sub_items:
                    sys.stderr.write('Message joint non pris en charge dans %s\n' % path)
                    continue
                sub = a.get_sub_item(0)
                sfx, _ = props(sub, named)
                skind = kind_of(val(sfx.get(0x001A))) or 'Message'
                export_item(sub, os.path.join(att_dir, 'Attachment%05d' % (i + 1), '%s00001' % skind), named, skind)
                continue
            name = None
            try:
                name = a.long_filename
            except Exception:
                pass
            name = safe(name or val(ap.get(0x3707)) or val(ap.get(0x3704)) or val(ap.get(0x3001)) or 'Piece_jointe_%d' % (i + 1), 120)
            size = a.size or 0
            with open(os.path.join(att_dir, '%d_%s' % (i + 1, name)), 'wb') as fh:
                done = 0
                while done < size:
                    chunk = a.read_buffer(min(1 << 20, size - done))
                    if not chunk:
                        break
                    fh.write(chunk)
                    done += len(chunk)
        except Exception as ex:
            sys.stderr.write('Pièce jointe %d de %s : %s\n' % (i + 1, path, ex))


def export_folder(folder, path, named):
    os.makedirs(path, exist_ok=True)
    try:
        n_items = folder.number_of_sub_messages
    except Exception:
        n_items = 0
    for i in range(n_items):
        try:
            m = folder.get_sub_message(i)
            fx, _ = props(m, named)
            kind = kind_of(val(fx.get(0x001A)))
            if not kind:
                continue
            export_item(m, os.path.join(path, '%s%05d' % (kind, i + 1)), named, kind)
            stats['items'] += 1
            if stats['items'] % 200 == 0:
                print('PROGRESS %d' % stats['items'], flush=True)
        except Exception as ex:
            stats['errors'] += 1
            sys.stderr.write('Élément %d de %s : %s\n' % (i + 1, path, ex))
    used = set()
    for i in range(folder.number_of_sub_folders):
        try:
            sub = folder.get_sub_folder(i)
            name = safe(sub.name or 'Dossier')
            base, k = name, 2
            while name.lower() in used:
                name = '%s (%d)' % (base, k)
                k += 1
            used.add(name.lower())
            export_folder(sub, os.path.join(path, name), named)
        except Exception as ex:
            stats['errors'] += 1
            sys.stderr.write('Dossier %d de %s : %s\n' % (i + 1, path, ex))


def main():
    src, target = sys.argv[1], sys.argv[2]
    if len(sys.argv) > 3:
        sys.stderr = open(sys.argv[3], 'a', encoding='utf-8')
    f = pypff.file()
    f.open(src)
    named = load_named_map(f)
    export_folder(f.root_folder, target + '.export', named)

    orphans = 0
    try:
        orphans = f.number_of_orphan_items
    except Exception:
        pass
    for i in range(orphans):
        try:
            m = f.get_orphan_item(i)
            if not isinstance(m, pypff.message):
                continue
            fx, _ = props(m, named)
            kind = kind_of(val(fx.get(0x001A)))
            if kind:
                export_item(m, os.path.join(target + '.orphans', '%s%05d' % (kind, i + 1)), named, kind)
        except Exception:
            pass
    f.close()
    print('DONE %d %d' % (stats['items'], stats['errors']), flush=True)


if __name__ == '__main__':
    main()
'@

# --- Conversion complète ------------------------------------------------------
function Invoke-Conversion([string]$ost, [string]$pst, [bool]$withRecovered, [string]$existingExport, [bool]$keep) {
    $pst = [IO.Path]::GetFullPath($pst)
    if ([IO.Path]::GetExtension($pst) -ne '.pst') { $pst += '.pst' }
    if (Test-Path -LiteralPath $pst) { throw "Le fichier $pst existe déjà. Choisissez un autre nom." }
    $script:LogFile = $pst + '.log.txt'
    $script:Stats = @{ Total = 0; Done = 0; Imported = 0; Errors = 0; Skipped = 0 }

    $workDir = $null
    $script:TempAttDir = $null
    $script:CancelRequested = $false
    $pstRoot = $null
    try {
        if ($existingExport) {
            $exportDir = [IO.Path]::GetFullPath($existingExport)
        }
        else {
            $ost = [IO.Path]::GetFullPath($ost)
            if (-not (Test-Path -LiteralPath $ost)) { throw "Fichier introuvable : $ost" }
            # Dossier de travail court, pour limiter les chemins trop longs
            $workDir = Join-Path ($env:SystemDrive + '\') ('ost2pst_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
            New-Item -ItemType Directory -Path $workDir -Force | Out-Null
            $exportDir = Invoke-OstExtraction $ost $workDir
        }
        $base = $exportDir -replace '\.export$', ''
        $recoveredDirs = @()
        if ($withRecovered) {
            foreach ($suffix in '.recovered', '.orphans') { if (Test-Path -LiteralPath ($base + $suffix)) { $recoveredDirs += ($base + $suffix) } }
        }

        $script:TempAttDir = Join-Path $env:TEMP ('ost2pst_att_' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path $script:TempAttDir -Force | Out-Null

        $root = Get-MailboxRoot $exportDir
        $script:Stats.Total = Count-Items $root
        foreach ($r in $recoveredDirs) { $script:Stats.Total += Count-Items $r }
        Write-Log ("{0} élément(s) à importer." -f $script:Stats.Total)

        Write-Log "Ouverture d'Outlook et création du PST..."
        $script:Outlook = New-Object -ComObject Outlook.Application
        $ns = $script:Outlook.GetNamespace('MAPI')
        $ns.AddStoreEx($pst, $olStoreUnicode)
        $store = $null
        foreach ($s in $ns.Stores) {
            try { if ($s.FilePath -and ([IO.Path]::GetFullPath($s.FilePath) -ieq $pst)) { $store = $s; break } } catch {}
        }
        if (-not $store) { throw "Outlook n'a pas pu créer le fichier $pst" }
        $pstRoot = $store.GetRootFolder()
        try { $pstRoot.Name = 'Converti - ' + [IO.Path]::GetFileNameWithoutExtension($pst) } catch {}

        Import-Folder $root $pstRoot 'Boîte'
        if ($recoveredDirs.Count -gt 0) {
            $recFolder = Get-OrAddFolder $pstRoot 'Éléments récupérés' $olFolderInbox
            foreach ($r in $recoveredDirs) { Import-Folder $r $recFolder 'Éléments récupérés' }
        }
    }
    catch [OperationCanceledException] {
        Write-Log ("Conversion abandonnée après {0} élément(s) importé(s)." -f $script:Stats.Imported)
        if ($pstRoot) { Write-Log ("Le PST partiel est conservé : {0}. Vous pouvez le supprimer depuis Outlook (clic droit > Fermer), puis supprimer le fichier." -f $pst) }
        throw
    }
    finally {
        if ($script:TempAttDir) { Remove-Item -LiteralPath $script:TempAttDir -Recurse -Force -ErrorAction SilentlyContinue }
        if ($workDir -and -not $keep) {
            Write-Log "Suppression des fichiers temporaires..."
            cmd.exe /c rd /s /q $workDir 2>$null
        }
        elseif ($workDir) { Write-Log "Fichiers extraits conservés dans : $workDir" }
    }

    Write-Log ("Terminé : {0} élément(s) importé(s), {1} erreur(s), {2} ignoré(s)." -f $script:Stats.Imported, $script:Stats.Errors, $script:Stats.Skipped)
    Write-Log ("PST créé : {0} (il est déjà ouvert dans Outlook, dossier « {1} »)." -f $pst, $pstRoot.Name)
    Write-Log ("Journal : {0}" -f $script:LogFile)
}

# --- Interface graphique --------------------------------------------------------
function Show-Gui {
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing
    try {
        Add-Type -Namespace Ost2Pst -Name Native -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
[DllImport("kernel32.dll")] public static extern System.IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
'@
        [void][Ost2Pst.Native]::SetProcessDPIAware()
        # Masque la fenêtre console noire derrière l'application
        [void][Ost2Pst.Native]::ShowWindow([Ost2Pst.Native]::GetConsoleWindow(), 0)
    } catch {}
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $script:Gui = $true

    $accent = [Drawing.Color]::FromArgb(15, 108, 189)
    $accentDark = [Drawing.Color]::FromArgb(17, 94, 163)
    $bg = [Drawing.Color]::FromArgb(243, 245, 248)
    $text = [Drawing.Color]::FromArgb(32, 32, 32)
    $muted = [Drawing.Color]::FromArgb(96, 96, 96)
    $font = New-Object Drawing.Font('Segoe UI', 10)
    $fontBold = New-Object Drawing.Font('Segoe UI Semibold', 10)

    $form = New-Object Windows.Forms.Form
    $form.Text = 'Convertisseur OST vers PST'
    $form.AutoScaleDimensions = New-Object Drawing.SizeF(96, 96)
    $form.AutoScaleMode = 'Dpi'
    $form.ClientSize = New-Object Drawing.Size(760, 600)
    $form.MinimumSize = New-Object Drawing.Size(640, 520)
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = $bg
    $form.ForeColor = $text
    $form.Font = $font
    $form.AllowDrop = $true
    try {
        $olPath = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' -ErrorAction Stop).'(default)'
        $form.Icon = [Drawing.Icon]::ExtractAssociatedIcon($olPath)
    } catch {}

    # Bandeau
    $header = New-Object Windows.Forms.Panel
    $header.Dock = 'Top'; $header.Height = 92; $header.BackColor = $accent
    $title = New-Object Windows.Forms.Label
    $title.Text = 'Convertisseur OST vers PST'
    $title.Font = New-Object Drawing.Font('Segoe UI Semibold', 18)
    $title.ForeColor = [Drawing.Color]::White
    $title.AutoSize = $true; $title.Location = New-Object Drawing.Point(24, 14)
    $subtitle = New-Object Windows.Forms.Label
    $subtitle.Text = 'Récupérez le contenu d''un fichier Outlook hors connexion dans un fichier PST.'
    $subtitle.ForeColor = [Drawing.Color]::FromArgb(222, 236, 249)
    $subtitle.AutoSize = $true; $subtitle.Location = New-Object Drawing.Point(26, 56)
    $header.Controls.AddRange(@($title, $subtitle))

    # Zone principale
    $main = New-Object Windows.Forms.TableLayoutPanel
    $main.Dock = 'Fill'; $main.Padding = New-Object Windows.Forms.Padding(24, 18, 24, 18)
    $main.ColumnCount = 1; $main.RowCount = 4
    [void]$main.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
    [void]$main.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
    [void]$main.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
    [void]$main.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', 100)))

    function New-Card {
        $p = New-Object Windows.Forms.TableLayoutPanel
        $p.Dock = 'Fill'; $p.AutoSize = $true; $p.BackColor = [Drawing.Color]::White
        $p.Padding = New-Object Windows.Forms.Padding(18, 14, 18, 14)
        $p.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 14)
        return $p
    }

    function New-FlatButton([string]$label, [bool]$primary) {
        $b = New-Object Windows.Forms.Button
        $b.Text = $label; $b.FlatStyle = 'Flat'; $b.Cursor = 'Hand'; $b.AutoSize = $true
        $b.Padding = New-Object Windows.Forms.Padding(10, 3, 10, 3)
        if ($primary) {
            $b.BackColor = $accent; $b.ForeColor = [Drawing.Color]::White; $b.Font = $fontBold
            $b.FlatAppearance.BorderSize = 0; $b.FlatAppearance.MouseOverBackColor = $accentDark
        }
        else {
            $b.BackColor = [Drawing.Color]::White; $b.ForeColor = $text
            $b.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(200, 200, 200)
            $b.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(235, 243, 251)
        }
        return $b
    }

    # Carte 1 : fichiers
    $card = New-Card
    $card.ColumnCount = 2; $card.RowCount = 5
    [void]$card.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$card.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('AutoSize')))

    function Add-FileRow($card, [string]$label, [string]$hint, [int]$row) {
        $l = New-Object Windows.Forms.Label
        $l.Text = $label; $l.Font = $fontBold; $l.AutoSize = $true
        $l.Margin = New-Object Windows.Forms.Padding(0, 4, 0, 4)
        $card.Controls.Add($l, 0, $row); $card.SetColumnSpan($l, 2)
        $t = New-Object Windows.Forms.TextBox
        $t.Dock = 'Fill'; $t.BorderStyle = 'FixedSingle'; $t.Margin = New-Object Windows.Forms.Padding(0, 2, 8, 10)
        try { $t.PlaceholderText = $hint } catch {}
        $b = New-FlatButton 'Parcourir…' $false
        $b.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 10)
        $card.Controls.Add($t, 0, $row + 1); $card.Controls.Add($b, 1, $row + 1)
        return @($t, $b)
    }

    $ostRow = Add-FileRow $card '1.  Fichier OST à convertir' 'Glissez le fichier .ost ici ou cliquez sur Parcourir' 0
    $pstRow = Add-FileRow $card '2.  Fichier PST à créer' '' 2
    $txtOst = $ostRow[0]; $txtPst = $pstRow[0]

    $chkRec = New-Object Windows.Forms.CheckBox
    $chkRec.Text = 'Récupérer aussi les éléments supprimés (dossier « Éléments récupérés »)'
    $chkRec.AutoSize = $true; $chkRec.ForeColor = $muted
    $chkRec.Margin = New-Object Windows.Forms.Padding(0, 2, 0, 0)
    $card.Controls.Add($chkRec, 0, 4); $card.SetColumnSpan($chkRec, 2)
    $main.Controls.Add($card, 0, 0)

    # Boutons Convertir / Abandonner
    $actions = New-Object Windows.Forms.TableLayoutPanel
    $actions.Dock = 'Fill'; $actions.AutoSize = $true; $actions.ColumnCount = 2; $actions.RowCount = 1
    $actions.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 14)
    [void]$actions.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$actions.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('AutoSize')))
    $btn = New-FlatButton 'Convertir' $true
    $btn.AutoSize = $false; $btn.Dock = 'Fill'; $btn.Height = 46
    $btn.Font = New-Object Drawing.Font('Segoe UI Semibold', 12)
    $btn.Margin = New-Object Windows.Forms.Padding(0)
    $btnCancel = New-FlatButton 'Abandonner' $false
    $btnCancel.AutoSize = $false; $btnCancel.Size = New-Object Drawing.Size(150, 46)
    $btnCancel.Font = New-Object Drawing.Font('Segoe UI Semibold', 11)
    $btnCancel.ForeColor = [Drawing.Color]::FromArgb(196, 43, 28)
    $btnCancel.FlatAppearance.BorderColor = [Drawing.Color]::FromArgb(196, 43, 28)
    $btnCancel.FlatAppearance.MouseOverBackColor = [Drawing.Color]::FromArgb(253, 231, 233)
    $btnCancel.Margin = New-Object Windows.Forms.Padding(10, 0, 0, 0)
    $btnCancel.Visible = $false
    $actions.Controls.Add($btn, 0, 0); $actions.Controls.Add($btnCancel, 1, 0)
    $main.Controls.Add($actions, 0, 1)
    $script:Running = $false
    $script:CloseAfter = $false

    # Carte 2 : progression
    $prog = New-Card
    $prog.ColumnCount = 2; $prog.RowCount = 2
    [void]$prog.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('Percent', 100)))
    [void]$prog.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle('AutoSize')))
    $status = New-Object Windows.Forms.Label
    $status.Text = 'Prêt.'; $status.AutoSize = $false; $status.Dock = 'Fill'; $status.Height = 26
    $status.AutoEllipsis = $true; $status.ForeColor = $muted
    $status.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 8)
    $count = New-Object Windows.Forms.Label
    $count.Text = ''; $count.AutoSize = $true; $count.ForeColor = $muted; $count.Anchor = 'Right'
    $bar = New-Object Windows.Forms.ProgressBar
    $bar.Dock = 'Fill'; $bar.Height = 10; $bar.Style = 'Continuous'
    $bar.Margin = New-Object Windows.Forms.Padding(0)
    $prog.Controls.Add($status, 0, 0); $prog.Controls.Add($count, 1, 0)
    $prog.Controls.Add($bar, 0, 1); $prog.SetColumnSpan($bar, 2)
    $main.Controls.Add($prog, 0, 2)

    # Détails (journal), repliable
    $detailsPanel = New-Object Windows.Forms.TableLayoutPanel
    $detailsPanel.Dock = 'Fill'; $detailsPanel.ColumnCount = 1; $detailsPanel.RowCount = 2
    [void]$detailsPanel.RowStyles.Add((New-Object Windows.Forms.RowStyle('AutoSize')))
    [void]$detailsPanel.RowStyles.Add((New-Object Windows.Forms.RowStyle('Percent', 100)))
    $toggle = New-Object Windows.Forms.LinkLabel
    $toggle.Text = '▸ Afficher les détails'; $toggle.AutoSize = $true
    $toggle.LinkColor = $accent; $toggle.ActiveLinkColor = $accentDark; $toggle.LinkBehavior = 'HoverUnderline'
    $log = New-Object Windows.Forms.TextBox
    $log.Multiline = $true; $log.ScrollBars = 'Vertical'; $log.ReadOnly = $true
    $log.Dock = 'Fill'; $log.BorderStyle = 'None'; $log.BackColor = [Drawing.Color]::White
    $log.Font = New-Object Drawing.Font('Consolas', 9); $log.Visible = $false
    $log.Margin = New-Object Windows.Forms.Padding(0, 6, 0, 0)
    $detailsPanel.Controls.Add($toggle, 0, 0); $detailsPanel.Controls.Add($log, 0, 1)
    $main.Controls.Add($detailsPanel, 0, 3)

    $form.Controls.Add($main)
    $form.Controls.Add($header)

    $toggle.Add_LinkClicked({
        $log.Visible = -not $log.Visible
        if ($log.Visible) { $toggle.Text = '▾ Masquer les détails' } else { $toggle.Text = '▸ Afficher les détails' }
    })

    $setOst = {
        param([string]$path)
        $txtOst.Text = $path
        if (-not $txtPst.Text) { $txtPst.Text = [IO.Path]::ChangeExtension($path, '.pst') }
    }.GetNewClosure()

    $ostRow[1].Add_Click({
        $d = New-Object Windows.Forms.OpenFileDialog
        $d.Title = 'Choisir le fichier OST'
        $d.Filter = 'Fichier Outlook hors connexion (*.ost)|*.ost|Tous les fichiers (*.*)|*.*'
        $d.InitialDirectory = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
        if ($d.ShowDialog() -eq 'OK') { & $setOst $d.FileName }
    })
    $pstRow[1].Add_Click({
        $d = New-Object Windows.Forms.SaveFileDialog
        $d.Title = 'Enregistrer le fichier PST'
        $d.Filter = 'Fichier de données Outlook (*.pst)|*.pst'
        $d.OverwritePrompt = $false
        if ($txtPst.Text) {
            $d.FileName = [IO.Path]::GetFileName($txtPst.Text)
            try { $d.InitialDirectory = [IO.Path]::GetDirectoryName($txtPst.Text) } catch {}
        }
        if ($d.ShowDialog() -eq 'OK') { $txtPst.Text = $d.FileName }
    })

    # Glisser-déposer du fichier OST sur la fenêtre
    $form.Add_DragEnter({
        if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = 'Copy' }
    })
    $form.Add_DragDrop({
        $files = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
        if ($files -and $files.Count -gt 0) { & $setOst $files[0] }
    })

    $script:LogAction = {
        param($msg)
        $log.AppendText($msg + "`r`n")
        $s = $msg -replace '^\d{2}:\d{2}:\d{2}\s+', ''
        if ($s -notmatch '^\s') { $status.Text = $s }
        [System.Windows.Forms.Application]::DoEvents()
    }.GetNewClosure()
    $script:ProgressAction = {
        param($done, $total)
        if ($total -gt 0) {
            $bar.Value = [Math]::Min(100, [int](100 * $done / $total))
            $count.Text = ('{0:N0} / {1:N0} éléments' -f $done, $total)
        }
        [System.Windows.Forms.Application]::DoEvents()
    }.GetNewClosure()

    $requestCancel = {
        $answer = [Windows.Forms.MessageBox]::Show('Abandonner la conversion en cours ?' + "`n`n" + 'Les éléments déjà importés resteront dans le PST partiel.', 'Abandonner', 'YesNo', 'Question')
        if ($answer -eq 'Yes') {
            $script:CancelRequested = $true
            $btnCancel.Enabled = $false; $btnCancel.Text = 'Abandon…'
            $status.Text = 'Abandon en cours…'
            return $true
        }
        return $false
    }  # pas de GetNewClosure : $script:CancelRequested doit viser la portée du script

    $btnCancel.Add_Click({ [void](& $requestCancel) })

    $form.Add_FormClosing({
        if ($script:Running) {
            $_.Cancel = $true
            if (& $requestCancel) { $script:CloseAfter = $true }
        }
    })

    $btn.Add_Click({
        if (-not $txtOst.Text -or -not $txtPst.Text) {
            [Windows.Forms.MessageBox]::Show('Choisissez le fichier OST et le fichier PST à créer.', 'Convertisseur OST vers PST', 'OK', 'Information') | Out-Null
            return
        }
        $btn.Enabled = $false; $btn.Text = 'Conversion en cours…'; $btn.BackColor = $muted
        $btnCancel.Text = 'Abandonner'; $btnCancel.Enabled = $true; $btnCancel.Visible = $true
        $status.ForeColor = $muted
        $script:Running = $true
        $bar.Value = 0; $count.Text = ''; $log.Clear()
        foreach ($c in @($txtOst, $txtPst, $ostRow[1], $pstRow[1], $chkRec)) { $c.Enabled = $false }
        try {
            Invoke-Conversion $txtOst.Text $txtPst.Text $chkRec.Checked $null $false
            $bar.Value = 100
            $status.ForeColor = [Drawing.Color]::FromArgb(16, 124, 16)
            $status.Text = ('Terminé : {0:N0} élément(s) importé(s), {1} erreur(s).' -f $script:Stats.Imported, $script:Stats.Errors)
            [Windows.Forms.MessageBox]::Show(("Conversion terminée.`n`n{0:N0} élément(s) importé(s).`nLe PST est déjà ouvert dans Outlook." -f $script:Stats.Imported), 'Convertisseur OST vers PST', 'OK', 'Information') | Out-Null
        }
        catch [OperationCanceledException] {
            $status.ForeColor = [Drawing.Color]::FromArgb(157, 93, 0)
            $status.Text = ('Conversion abandonnée ({0:N0} élément(s) déjà importé(s)).' -f $script:Stats.Imported)
        }
        catch {
            Write-Log ('ERREUR : ' + $_.Exception.Message)
            $status.ForeColor = [Drawing.Color]::FromArgb(196, 43, 28)
            $status.Text = 'Erreur : ' + $_.Exception.Message
            [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Erreur', 'OK', 'Error') | Out-Null
        }
        finally {
            $script:Running = $false
            $btnCancel.Visible = $false
            $btn.Enabled = $true; $btn.Text = 'Convertir'; $btn.BackColor = $accent
            foreach ($c in @($txtOst, $txtPst, $ostRow[1], $pstRow[1], $chkRec)) { $c.Enabled = $true }
            if ($script:CloseAfter) { $form.Close() }
        }
    })

    $form.Add_Shown({ $btn.Focus() })
    [void]$form.ShowDialog()
}

# --- Point d'entrée -------------------------------------------------------------
if (-not $OstPath -and -not $ExportDir) {
    try { Show-Gui }
    catch {
        # Réaffiche la console pour que l'erreur reste lisible
        try { [void][Ost2Pst.Native]::ShowWindow([Ost2Pst.Native]::GetConsoleWindow(), 5) } catch {}
        throw
    }
}
else {
    if (-not $PstPath) {
        if ($OstPath) { $PstPath = [IO.Path]::ChangeExtension($OstPath, '.pst') }
        else { throw 'Indiquez -PstPath.' }
    }
    Invoke-Conversion $OstPath $PstPath $IncludeRecovered.IsPresent $ExportDir $KeepExport.IsPresent
}
