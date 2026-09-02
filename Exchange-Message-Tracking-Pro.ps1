#requires -Version 5.1
#requires -PSEdition Desktop

<#
.SYNOPSIS
    Exchange Server ileti izleme günlükleri için Türkçe, güvenli ve anlaşılır WinForms arayüzü.

.DESCRIPTION
    Yalnızca şirket içi (on-premises) Exchange Server ortamlarındaki
    Get-MessageTrackingLog cmdlet'ini kullanır. Exchange Management Shell içinde doğrudan
    çalışabilir veya standart Windows PowerShell 5.1 üzerinden bir Exchange sunucusuna
    uzak PowerShell bağlantısı kurabilir.

    Betik hiçbir Exchange ayarını değiştirmez; yalnızca günlük sorgular ve isteğe bağlı
    olarak sonuçları yerel CSV/HTML dosyasına aktarır.

.PARAMETER ExchangeServerFqdn
    GUI açıldığında bağlantı için kullanılacak Exchange sunucusunun FQDN veya NetBIOS adı.

.PARAMETER UseCredential
    Uzak bağlantıda geçerli Windows oturumu yerine kimlik bilgisi sorulmasını sağlar.

.EXAMPLE
    .\Exchange-Message-Tracking-Pro.ps1

.EXAMPLE
    .\Exchange-Message-Tracking-Pro.ps1 -ExchangeServerFqdn ex01.contoso.local

.NOTES
    Gereksinimler: Windows PowerShell 5.1, Exchange Server 2016/2019/SE yönetim erişimi
    ve ileti izleme günlüklerini okuma yetkisi.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$ExchangeServerFqdn,

    [Parameter()]
    [switch]$UseCredential
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:ExchangeSession       = $null
$script:ImportedExchangeModule = $null
$script:TransportServers      = @()
$script:NormalizedResults     = New-Object System.Collections.ArrayList
$script:CancelRequested       = $false
$script:IsSearching           = $false

$script:Colors = @{
    Navy       = [System.Drawing.Color]::FromArgb(24, 45, 76)
    Blue       = [System.Drawing.Color]::FromArgb(32, 111, 180)
    LightBlue  = [System.Drawing.Color]::FromArgb(232, 242, 252)
    Background = [System.Drawing.Color]::FromArgb(244, 247, 250)
    Border     = [System.Drawing.Color]::FromArgb(210, 218, 227)
    Text       = [System.Drawing.Color]::FromArgb(35, 45, 55)
    Muted      = [System.Drawing.Color]::FromArgb(100, 112, 124)
    Success    = [System.Drawing.Color]::FromArgb(31, 122, 67)
    SuccessBg  = [System.Drawing.Color]::FromArgb(229, 246, 235)
    Warning    = [System.Drawing.Color]::FromArgb(157, 102, 0)
    WarningBg  = [System.Drawing.Color]::FromArgb(255, 245, 214)
    Error      = [System.Drawing.Color]::FromArgb(180, 45, 45)
    ErrorBg    = [System.Drawing.Color]::FromArgb(253, 232, 232)
    Info       = [System.Drawing.Color]::FromArgb(70, 86, 101)
    InfoBg     = [System.Drawing.Color]::FromArgb(238, 242, 246)
}

$script:EventDescriptions = @{
    AGENTINFO             = 'Taşıma aracısı özel verileri günlüğe kaydetti.'
    BADMAIL               = 'Pickup veya Replay dizininden sunulan ileti teslim edilemedi ve gönderene iade edilemedi.'
    CLIENTSUBMISSION      = 'İleti, bir posta kutusunun Giden Kutusu üzerinden gönderim için sunuldu.'
    DEFER                 = 'İleti teslimi ertelendi.'
    DELIVER               = 'İleti yerel posta kutusuna teslim edildi.'
    DELIVERFAIL           = 'Bir aracı, iletiyi posta kutusunda bulunmayan bir klasöre teslim etmeyi denedi.'
    DROP                  = 'İleti, teslim durumu bildirimi (DSN/NDR) üretilmeden sistemden çıkarıldı; örneğin sessizce bırakılan spam veya tamamlanmış denetim isteği.'
    DSN                   = 'Bir teslim durumu bildirimi (DSN) üretildi. Bu olay tek başına başarısızlık anlamına gelmez; ayrıntı alıcı durumu ve bağlam alanlarındadır.'
    DUPLICATEDELIVER      = 'Aynı alıcı için yinelenen teslim algılandı; Exchange bilgi deposu yinelenen kopyayı kaldırdı.'
    DUPLICATEEXPAND       = 'Dağıtım grubu genişletilirken yinelenen bir alıcı algılandı.'
    DUPLICATEREDIRECT     = 'İletinin alternatif alıcısı zaten alıcılar arasındaydı.'
    EXPAND                = 'Dağıtım grubu üyelerine genişletildi.'
    FAIL                  = 'İleti teslimi SMTP, DNS, kuyruk veya yönlendirme aşamasında başarısız oldu.'
    HADISCARD             = 'Birincil kopya sonraki atlama noktasına teslim edildiği için gölge ileti kopyası silindi.'
    HARECEIVE             = 'Sunucu, yerel DAG veya Active Directory sitesi içinden bir gölge ileti kopyası aldı.'
    HAREDIRECT            = 'Gölge ileti kopyası oluşturuldu.'
    HAREDIRECTFAIL        = 'Gölge ileti kopyası oluşturulamadı; ayrıntı SourceContext alanındadır.'
    INITMESSAGECREATED    = 'İleti denetimli bir alıcıya gönderildi ve onay için tahkim posta kutusuna aktarıldı.'
    LOAD                  = 'İleti, sunucu başlangıcında başarıyla yüklendi.'
    MODERATIONEXPIRE      = 'Denetleyici iletiyi onaylamadığı veya reddetmediği için ileti zaman aşımına uğradı.'
    MODERATORAPPROVE      = 'Denetleyici iletiyi onayladı ve ileti denetimli alıcıya teslim edildi.'
    MODERATORREJECT       = 'Denetleyici iletiyi reddetti; ileti denetimli alıcıya teslim edilmedi.'
    MODERATORSALLNDR      = 'Denetleyicilere gönderilen tüm onay istekleri teslim edilemedi ve NDR ile sonuçlandı.'
    NOTIFYMAPI            = 'Yerel sunucudaki bir posta kutusunun Giden Kutusunda ileti algılandı.'
    NOTIFYSHADOW          = 'Yerel Giden Kutusunda ileti algılandı ve ileti için gölge kopya oluşturulması gerekiyor.'
    POISONMESSAGE         = 'İleti zehirli ileti kuyruğuna alındı veya kuyruktan çıkarıldı.'
    PROCESS               = 'İleti başarıyla işlendi.'
    PROCESSMEETINGMESSAGE = 'Toplantı iletisi Mailbox Transport Delivery tarafından işlendi.'
    RECEIVE               = 'İleti; SMTP Receive, Pickup/Replay veya posta kutusu gönderimi yoluyla taşıma hizmetine alındı.'
    REDIRECT              = 'İleti, Active Directory aramasından sonra alternatif bir alıcıya yönlendirildi.'
    RESOLVE               = 'İletinin alıcı adresi, Active Directory aramasından sonra farklı bir e-posta adresine çözümlendi.'
    RESUBMIT              = 'İleti Safety Net üzerinden otomatik olarak yeniden gönderildi.'
    RESUBMITDEFER         = 'Safety Net üzerinden yeniden gönderilen iletinin işlemi ertelendi.'
    RESUBMITFAIL          = 'Safety Net üzerinden yeniden gönderilen iletinin işlemi başarısız oldu.'
    SEND                  = 'İleti SMTP ile taşıma hizmetleri arasında gönderildi. Bu olay, son alıcıya teslim edildiği anlamına gelmez.'
    SUBMIT                = 'Mailbox Transport Submission hizmeti iletiyi Transport hizmetine başarıyla aktardı; bu olay son teslim değildir.'
    SUBMITDEFER           = 'Mailbox Transport Submission hizmetinden Transport hizmetine aktarım ertelendi.'
    SUBMITFAIL            = 'Mailbox Transport Submission hizmetinden Transport hizmetine aktarım başarısız oldu.'
    SUPPRESSED            = 'İleti aktarımı bastırıldı.'
    THROTTLE              = 'İleti hız sınırlamasına tabi tutuldu; neden Reference alanında bulunur.'
    TRANSFER              = 'İçerik dönüştürme, alıcı sınırı veya bir aracı nedeniyle alıcılar iletinin çatallanmış bir kopyasına taşındı.'
}

$script:EventOutcomes = @{
    AGENTINFO             = 'Bilgi'
    BADMAIL               = 'Başarısız'
    CLIENTSUBMISSION      = 'Aktarım'
    DEFER                 = 'Gecikme'
    DELIVER               = 'Teslim edildi'
    DELIVERFAIL           = 'Başarısız'
    DROP                  = 'Teslim edilmedi'
    DSN                   = 'Bildirim'
    DUPLICATEDELIVER      = 'Bilgi'
    DUPLICATEEXPAND       = 'Bilgi'
    DUPLICATEREDIRECT     = 'Bilgi'
    EXPAND                = 'Yönlendirme'
    FAIL                  = 'Başarısız'
    HADISCARD             = 'Bilgi'
    HARECEIVE             = 'Alındı'
    HAREDIRECT            = 'Aktarım'
    HAREDIRECTFAIL        = 'Başarısız'
    INITMESSAGECREATED    = 'Onay süreci'
    LOAD                  = 'İşlendi'
    MODERATIONEXPIRE      = 'Teslim edilmedi'
    MODERATORAPPROVE      = 'Teslim edildi'
    MODERATORREJECT       = 'Teslim edilmedi'
    MODERATORSALLNDR      = 'Başarısız'
    NOTIFYMAPI            = 'Alındı'
    NOTIFYSHADOW          = 'Alındı'
    POISONMESSAGE         = 'İnceleme gerekli'
    PROCESS               = 'İşlendi'
    PROCESSMEETINGMESSAGE = 'İşlendi'
    RECEIVE               = 'Alındı'
    REDIRECT              = 'Yönlendirme'
    RESOLVE               = 'Yönlendirme'
    RESUBMIT              = 'Aktarım'
    RESUBMITDEFER         = 'Gecikme'
    RESUBMITFAIL          = 'Başarısız'
    SEND                  = 'Aktarım'
    SUBMIT                = 'Aktarım'
    SUBMITDEFER           = 'Gecikme'
    SUBMITFAIL            = 'Başarısız'
    SUPPRESSED            = 'Teslim edilmedi'
    THROTTLE              = 'Gecikme'
    TRANSFER              = 'Aktarım'
}

$script:SuccessEvents = @('DELIVER', 'MODERATORAPPROVE')
$script:WarningEvents = @(
    'DEFER', 'DROP', 'MODERATIONEXPIRE', 'MODERATORREJECT', 'RESUBMITDEFER',
    'SUBMITDEFER', 'SUPPRESSED', 'THROTTLE', 'POISONMESSAGE'
)
$script:ErrorEvents = @(
    'BADMAIL', 'DELIVERFAIL', 'FAIL', 'HAREDIRECTFAIL', 'MODERATORSALLNDR',
    'RESUBMITFAIL', 'SUBMITFAIL'
)

function Show-AppMessage {
    param(
        [Parameter(Mandatory)]
        [string]$Text,

        [Parameter()]
        [string]$Title = 'Exchange Message Tracking Pro',

        [Parameter()]
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information,

        [Parameter()]
        [System.Windows.Forms.MessageBoxButtons]$Buttons = [System.Windows.Forms.MessageBoxButtons]::OK
    )

    return [System.Windows.Forms.MessageBox]::Show($script:MainForm, $Text, $Title, $Buttons, $Icon)
}

function Get-ObjectPropertyValue {
    param(
        [Parameter(Mandatory)]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [string]$Name
    )

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function ConvertTo-FlatString {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [string]) {
        return $Value.Trim()
    }

    if ($Value -is [System.Collections.IDictionary]) {
        return (($Value.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value }) -join '; ')
    }

    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
        return (($Value | ForEach-Object { [string]$_ }) -join '; ')
    }

    return ([string]$Value).Trim()
}

function Get-EventPresentation {
    param(
        [AllowNull()][string]$EventId,
        [AllowNull()][string]$RecipientStatus,
        [AllowNull()][string]$SourceContext
    )

    $normalizedEvent = if ([string]::IsNullOrWhiteSpace($EventId)) { 'UNKNOWN' } else { $EventId.Trim().ToUpperInvariant() }
    $diagnosticText = '{0} {1}' -f $RecipientStatus, $SourceContext

    $status = $script:EventOutcomes[$normalizedEvent]
    if ([string]::IsNullOrWhiteSpace($status)) {
        $status = 'Bilgi'
    }
    $severity = 'Info'

    if ($diagnosticText -match '(^|[\s;])5\d\d([\s.;-]|$)') {
        $status = 'Başarısız'
        $severity = 'Error'
    }
    elseif ($diagnosticText -match '(^|[\s;])4\d\d([\s.;-]|$)') {
        $status = 'Gecikme'
        $severity = 'Warning'
    }
    elseif ($script:ErrorEvents -contains $normalizedEvent) {
        $severity = 'Error'
    }
    elseif ($script:WarningEvents -contains $normalizedEvent) {
        $severity = 'Warning'
    }
    elseif ($script:SuccessEvents -contains $normalizedEvent) {
        $severity = 'Success'
    }

    $description = $script:EventDescriptions[$normalizedEvent]
    if ([string]::IsNullOrWhiteSpace($description)) {
        $description = 'Exchange tarafından kaydedilmiş ileti akışı olayı.'
    }

    [pscustomobject]@{
        Status      = $status
        Severity    = $severity
        Description = $description
    }
}

function Test-IsHealthMessage {
    param([Parameter(Mandatory)][object]$InputObject)

    $sender = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'Sender')
    $recipients = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'Recipients')

    return (($sender -match '(?i)(^|@)HealthMailbox') -or ($recipients -match '(?i)(^|[;\s])HealthMailbox'))
}

function ConvertTo-NormalizedResult {
    param(
        [Parameter(Mandatory)][object]$InputObject,
        [AllowNull()][string]$QueryServer
    )

    $eventId = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'EventId')
    $recipientStatus = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'RecipientStatus')
    $sourceContext = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'SourceContext')
    $presentation = Get-EventPresentation -EventId $eventId -RecipientStatus $recipientStatus -SourceContext $sourceContext
    $timestampValue = Get-ObjectPropertyValue -InputObject $InputObject -Name 'Timestamp'
    $totalBytes = Get-ObjectPropertyValue -InputObject $InputObject -Name 'TotalBytes'
    $serverHostname = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'ServerHostname')

    $timestamp = if ($null -eq $timestampValue) { [datetime]::MinValue } else { [datetime]$timestampValue }
    $sizeKB = if ($null -eq $totalBytes -or [string]::IsNullOrWhiteSpace([string]$totalBytes)) {
        $null
    }
    else {
        [math]::Round(([double]$totalBytes / 1KB), 2)
    }

    if ([string]::IsNullOrWhiteSpace($serverHostname)) {
        $serverHostname = $QueryServer
    }

    [pscustomobject][ordered]@{
        Timestamp               = $timestamp
        Status                  = $presentation.Status
        Severity                = $presentation.Severity
        EventId                 = $eventId.ToUpperInvariant()
        Description             = $presentation.Description
        Source                  = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'Source')
        Server                  = $serverHostname
        Sender                  = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'Sender')
        Recipients              = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'Recipients')
        Subject                 = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'MessageSubject')
        RecipientStatus         = $recipientStatus
        SizeKB                  = $sizeKB
        MessageId               = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'MessageId')
        NetworkMessageId        = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'NetworkMessageId')
        InternalMessageId       = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'InternalMessageId')
        ClientHostname          = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'ClientHostname')
        ClientIp                = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'ClientIp')
        ServerIp                = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'ServerIp')
        ConnectorId             = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'ConnectorId')
        Directionality          = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'Directionality')
        RelatedRecipientAddress = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'RelatedRecipientAddress')
        ReturnPath              = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'ReturnPath')
        SourceContext           = $sourceContext
        Reference               = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'Reference')
        MessageInfo             = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'MessageInfo')
        CustomData              = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'CustomData')
        EventData               = ConvertTo-FlatString (Get-ObjectPropertyValue -InputObject $InputObject -Name 'EventData')
        QueryServer             = $QueryServer
    }
}

function New-ResultTable {
    $table = New-Object System.Data.DataTable 'MessageTrackingResults'
    [void]$table.Columns.Add('Timestamp', [datetime])
    [void]$table.Columns.Add('Status', [string])
    [void]$table.Columns.Add('Severity', [string])
    [void]$table.Columns.Add('EventId', [string])
    [void]$table.Columns.Add('Description', [string])
    [void]$table.Columns.Add('Source', [string])
    [void]$table.Columns.Add('Server', [string])
    [void]$table.Columns.Add('Sender', [string])
    [void]$table.Columns.Add('Recipients', [string])
    [void]$table.Columns.Add('Subject', [string])
    [void]$table.Columns.Add('RecipientStatus', [string])
    [void]$table.Columns.Add('SizeKB', [double])
    [void]$table.Columns.Add('MessageId', [string])
    [void]$table.Columns.Add('NetworkMessageId', [string])
    [void]$table.Columns.Add('InternalMessageId', [string])
    [void]$table.Columns.Add('ClientHostname', [string])
    [void]$table.Columns.Add('ClientIp', [string])
    [void]$table.Columns.Add('ServerIp', [string])
    [void]$table.Columns.Add('ConnectorId', [string])
    [void]$table.Columns.Add('Directionality', [string])
    [void]$table.Columns.Add('RelatedRecipientAddress', [string])
    [void]$table.Columns.Add('ReturnPath', [string])
    [void]$table.Columns.Add('SourceContext', [string])
    [void]$table.Columns.Add('Reference', [string])
    [void]$table.Columns.Add('MessageInfo', [string])
    [void]$table.Columns.Add('CustomData', [string])
    [void]$table.Columns.Add('EventData', [string])
    [void]$table.Columns.Add('QueryServer', [string])
    # DataTable IEnumerable uyguladığı için virgül, boş tablonun pipeline'da kaybolmasını önler.
    return ,$table
}

function Add-NormalizedResultToTable {
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$Table,
        [Parameter(Mandatory)][object]$Result
    )

    $row = $Table.NewRow()
    foreach ($column in $Table.Columns) {
        $value = Get-ObjectPropertyValue -InputObject $Result -Name $column.ColumnName
        if ($null -eq $value -or ([string]$value).Length -eq 0) {
            $row[$column.ColumnName] = [DBNull]::Value
        }
        else {
            $row[$column.ColumnName] = $value
        }
    }
    [void]$Table.Rows.Add($row)
}

function Add-AppLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    $line = '[{0:HH:mm:ss}] [{1}] {2}{3}' -f (Get-Date), $Level, $Message, [Environment]::NewLine
    $txtLog.AppendText($line)
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
}

function Set-ConnectionState {
    param(
        [Parameter(Mandatory)][bool]$Connected,
        [AllowNull()][string]$Text
    )

    if ($Connected) {
        $lblConnection.Text = if ([string]::IsNullOrWhiteSpace($Text)) { 'Bağlı' } else { $Text }
        $lblConnection.ForeColor = $script:Colors.Success
        $btnSearch.Enabled = $true
    }
    else {
        $lblConnection.Text = if ([string]::IsNullOrWhiteSpace($Text)) { 'Exchange bağlantısı yok' } else { $Text }
        $lblConnection.ForeColor = $script:Colors.Error
        $btnSearch.Enabled = $false
    }
}

function Test-ExchangeCommandAvailable {
    return ($null -ne (Get-Command -Name Get-MessageTrackingLog -ErrorAction SilentlyContinue))
}

function Update-ServerList {
    $cmbServer.BeginUpdate()
    try {
        $cmbServer.Items.Clear()
        [void]$cmbServer.Items.Add('[Tümü]')
        [void]$cmbServer.Items.Add('[Geçerli sunucu]')

        try {
            $servers = @(
                Get-TransportService -ErrorAction Stop |
                    Where-Object { $null -eq $_.MessageTrackingLogEnabled -or $_.MessageTrackingLogEnabled } |
                    ForEach-Object { [string]$_.Name } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique
            )
            $script:TransportServers = $servers
            foreach ($serverName in $servers) {
                [void]$cmbServer.Items.Add($serverName)
            }
            Add-AppLog -Message ('{0} taşıma sunucusu bulundu.' -f $servers.Count) -Level OK
        }
        catch {
            $script:TransportServers = @()
            Add-AppLog -Message ('Sunucu listesi alınamadı; geçerli sunucuda arama yapılabilir. {0}' -f $_.Exception.Message) -Level WARN
        }

        $cmbServer.SelectedIndex = if ($script:TransportServers.Count -gt 0) { 0 } else { 1 }
    }
    finally {
        $cmbServer.EndUpdate()
    }
}

function Connect-ExchangeRemote {
    param(
        [Parameter(Mandatory)][string]$ServerOrUri,
        [Parameter()][switch]$PromptForCredential
    )

    if ([string]::IsNullOrWhiteSpace($ServerOrUri)) {
        throw 'Exchange sunucu adı veya bağlantı URI değeri boş olamaz.'
    }

    $connectionUri = if ($ServerOrUri -match '^https?://') {
        $ServerOrUri.TrimEnd('/') + '/'
    }
    else {
        'http://{0}/PowerShell/' -f $ServerOrUri.Trim()
    }

    if ($null -ne $script:ExchangeSession) {
        Remove-PSSession -Session $script:ExchangeSession -ErrorAction SilentlyContinue
        $script:ExchangeSession = $null
    }

    $sessionParameters = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = $connectionUri
        Authentication    = 'Kerberos'
        ErrorAction       = 'Stop'
    }

    if ($PromptForCredential) {
        $credential = Get-Credential -Message 'Exchange yönetim yetkisine sahip hesabı girin.'
        if ($null -eq $credential) {
            throw 'Kimlik bilgisi girişi iptal edildi.'
        }
        $sessionParameters.Credential = $credential
    }

    Add-AppLog -Message ('Uzak Exchange oturumu açılıyor: {0}' -f $connectionUri)
    try {
        $script:ExchangeSession = New-PSSession @sessionParameters
        $script:ImportedExchangeModule = Import-PSSession -Session $script:ExchangeSession -DisableNameChecking -AllowClobber -ErrorAction Stop

        if (-not (Test-ExchangeCommandAvailable)) {
            throw 'Oturum açıldı ancak Get-MessageTrackingLog komutu RBAC izinleri nedeniyle içe aktarılamadı.'
        }
    }
    catch {
        if ($null -ne $script:ExchangeSession) {
            Remove-PSSession -Session $script:ExchangeSession -ErrorAction SilentlyContinue
            $script:ExchangeSession = $null
        }
        $script:ImportedExchangeModule = $null
        throw
    }
}

function ConvertTo-PreviewValue {
    param([AllowNull()][object]$Value)

    if ($Value -is [datetime]) {
        return "'{0:yyyy-MM-dd HH:mm:ss}'" -f $Value
    }

    if (($Value -is [System.Collections.IEnumerable]) -and -not ($Value -is [string])) {
        return '@({0})' -f (($Value | ForEach-Object { "'{0}'" -f ([string]$_).Replace("'", "''") }) -join ', ')
    }

    return "'{0}'" -f ([string]$Value).Replace("'", "''")
}

function Get-SearchParameters {
    $parameters = @{
        Start      = $dtpStart.Value
        End        = $dtpEnd.Value
        ResultSize = if ($cmbResultSize.SelectedItem -eq 'Sınırsız') { 'Unlimited' } else { [int]$cmbResultSize.SelectedItem }
    }

    $textParameters = @{
        Sender            = $txtSender.Text
        MessageSubject    = $txtSubject.Text
        EventId           = if ($cmbEvent.SelectedIndex -gt 0) { [string]$cmbEvent.SelectedItem } else { '' }
        Source            = if ($cmbSource.SelectedIndex -gt 0) { [string]$cmbSource.SelectedItem } else { '' }
        MessageId         = $txtMessageId.Text
        NetworkMessageId  = $txtNetworkMessageId.Text
        InternalMessageId = $txtInternalMessageId.Text
        Reference         = $txtReference.Text
    }

    foreach ($entry in $textParameters.GetEnumerator()) {
        if (-not [string]::IsNullOrWhiteSpace([string]$entry.Value)) {
            $parameters[$entry.Key] = ([string]$entry.Value).Trim()
        }
    }

    $recipients = @(
        $txtRecipients.Text -split '[,;\r\n]+' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )
    if ($recipients.Count -gt 0) {
        $parameters.Recipients = $recipients
    }

    return $parameters
}

function Test-SearchParameters {
    param([Parameter(Mandatory)][hashtable]$Parameters)

    if ([datetime]$Parameters.Start -ge [datetime]$Parameters.End) {
        throw 'Başlangıç tarihi bitiş tarihinden önce olmalıdır.'
    }

    $command = Get-Command -Name Get-MessageTrackingLog -ErrorAction Stop
    foreach ($parameterName in @($Parameters.Keys)) {
        if (-not $command.Parameters.ContainsKey($parameterName)) {
            throw ('Bağlı Exchange sürümü -{0} parametresini desteklemiyor. İlgili alanı boş bırakın.' -f $parameterName)
        }
    }
}

function Get-QueryServers {
    $selection = [string]$cmbServer.SelectedItem
    if ($selection -eq '[Tümü]') {
        return @($script:TransportServers)
    }
    if ($selection -eq '[Geçerli sunucu]' -or [string]::IsNullOrWhiteSpace($selection)) {
        return @($null)
    }
    return @($selection)
}

function Update-QueryPreview {
    if (-not (Test-ExchangeCommandAvailable)) {
        $txtLog.Text = 'Get-MessageTrackingLog henüz kullanılabilir değil.'
        return
    }

    try {
        $parameters = Get-SearchParameters
        $parts = New-Object System.Collections.Generic.List[string]
        $parts.Add('Get-MessageTrackingLog')
        foreach ($key in ($parameters.Keys | Sort-Object)) {
            $parts.Add(('-{0} {1}' -f $key, (ConvertTo-PreviewValue -Value $parameters[$key])))
        }
        if ([string]$cmbServer.SelectedItem -eq '[Tümü]') {
            $parts.Add('-Server <her taşıma sunucusu>')
        }
        elseif ([string]$cmbServer.SelectedItem -ne '[Geçerli sunucu]') {
            $parts.Add(('-Server {0}' -f (ConvertTo-PreviewValue -Value $cmbServer.SelectedItem)))
        }
        $txtQueryPreview.Text = $parts -join ' '
    }
    catch {
        $txtQueryPreview.Text = $_.Exception.Message
    }
}

function Update-Summary {
    $results = @($script:NormalizedResults)
    $uniqueIds = @(
        $results |
            ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_.MessageId)) { $_.MessageId }
                elseif (-not [string]::IsNullOrWhiteSpace($_.NetworkMessageId)) { $_.NetworkMessageId }
            } |
            Sort-Object -Unique
    )
    $delivered = @($results | Where-Object { $_.EventId -eq 'DELIVER' }).Count
    $problems = @($results | Where-Object { $_.Severity -in @('Warning', 'Error') }).Count

    $lblRowsValue.Text = [string]$results.Count
    $lblMessagesValue.Text = [string]$uniqueIds.Count
    $lblDeliveredValue.Text = [string]$delivered
    $lblProblemsValue.Text = [string]$problems
}

function ConvertTo-RowFilterLiteral {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrEmpty($Value)) {
        return ''
    }

    $builder = New-Object System.Text.StringBuilder
    foreach ($character in $Value.ToCharArray()) {
        switch ($character) {
            "'" { [void]$builder.Append("''") }
            '*' { [void]$builder.Append('[*]') }
            '%' { [void]$builder.Append('[%]') }
            '[' { [void]$builder.Append('[[]') }
            ']' { [void]$builder.Append('[]]') }
            default { [void]$builder.Append($character) }
        }
    }
    return $builder.ToString()
}

function Get-VisibleResults {
    $visibleResults = New-Object System.Collections.ArrayList
    foreach ($viewRow in $script:ResultTable.DefaultView) {
        $properties = [ordered]@{}
        foreach ($column in $script:ResultTable.Columns) {
            $value = $viewRow.Row[$column.ColumnName]
            $properties[$column.ColumnName] = if ($value -is [DBNull]) { '' } else { $value }
        }
        [void]$visibleResults.Add([pscustomobject]$properties)
    }
    return @($visibleResults)
}

function Update-ResultEventFilterChoices {
    $currentSelection = [string]$cmbResultEvent.SelectedItem
    $events = @(
        $script:NormalizedResults |
            ForEach-Object { $_.EventId } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )

    $cmbResultEvent.BeginUpdate()
    try {
        $cmbResultEvent.Items.Clear()
        [void]$cmbResultEvent.Items.Add('[Tüm olaylar]')
        foreach ($eventName in $events) {
            [void]$cmbResultEvent.Items.Add($eventName)
        }

        if (-not [string]::IsNullOrWhiteSpace($currentSelection) -and $cmbResultEvent.Items.Contains($currentSelection)) {
            $cmbResultEvent.SelectedItem = $currentSelection
        }
        else {
            $cmbResultEvent.SelectedIndex = 0
        }
    }
    finally {
        $cmbResultEvent.EndUpdate()
    }
}

function Update-ResultFilter {
    $clauses = New-Object System.Collections.Generic.List[string]
    $searchColumns = @(
        'Status', 'EventId', 'Description', 'Source', 'Server', 'Sender', 'Recipients',
        'Subject', 'RecipientStatus', 'MessageId', 'NetworkMessageId',
        'InternalMessageId', 'ClientHostname', 'ClientIp', 'ConnectorId',
        'SourceContext', 'Reference'
    )

    $searchTerms = @(
        $txtResultFilter.Text -split '\s+' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )

    foreach ($term in $searchTerms) {
        $escapedTerm = ConvertTo-RowFilterLiteral -Value $term
        $termClauses = @($searchColumns | ForEach-Object { '[{0}] LIKE ''%{1}%''' -f $_, $escapedTerm })
        $clauses.Add(('({0})' -f ($termClauses -join ' OR ')))
    }

    if ($cmbResultStatus.SelectedIndex -gt 0) {
        $statusValue = ConvertTo-RowFilterLiteral -Value ([string]$cmbResultStatus.SelectedItem)
        $clauses.Add(('[Status] = ''{0}''' -f $statusValue))
    }

    if ($cmbResultEvent.SelectedIndex -gt 0) {
        $eventValue = ConvertTo-RowFilterLiteral -Value ([string]$cmbResultEvent.SelectedItem)
        $clauses.Add(('[EventId] = ''{0}''' -f $eventValue))
    }

    try {
        $script:ResultTable.DefaultView.RowFilter = $clauses -join ' AND '
        $visibleCount = $script:ResultTable.DefaultView.Count
        $totalCount = $script:ResultTable.Rows.Count
        $lblFilterCount.Text = 'Gösterilen: {0} / {1}' -f $visibleCount, $totalCount
        $lblFilterCount.ForeColor = if ($visibleCount -eq 0 -and $totalCount -gt 0) { $script:Colors.Error } else { $script:Colors.Muted }
        $hasVisibleRows = ($visibleCount -gt 0)
        $btnExportCsv.Enabled = $hasVisibleRows
        $btnExportHtml.Enabled = $hasVisibleRows
        $btnCopy.Enabled = $hasVisibleRows
    }
    catch {
        $lblFilterCount.Text = 'Filtre uygulanamadı'
        $lblFilterCount.ForeColor = $script:Colors.Error
        Add-AppLog -Message ('Sonuç filtresi uygulanamadı: {0}' -f $_.Exception.Message) -Level ERROR
    }
}

function Clear-ResultFilter {
    $resultFilterTimer.Stop()
    $txtResultFilter.Clear()
    $cmbResultStatus.SelectedIndex = 0
    $cmbResultEvent.SelectedIndex = 0
    $resultFilterTimer.Stop()
    Update-ResultFilter
    $txtResultFilter.Focus()
}

function Update-ResultsGrid {
    $script:ResultTable.BeginLoadData()
    try {
        $script:ResultTable.Clear()
        foreach ($result in @($script:NormalizedResults | Sort-Object Timestamp -Descending)) {
            Add-NormalizedResultToTable -Table $script:ResultTable -Result $result
        }
    }
    finally {
        $script:ResultTable.EndLoadData()
    }

    Update-ResultEventFilterChoices
    Update-Summary
    Update-ResultFilter
}

function Clear-Results {
    $script:NormalizedResults.Clear()
    $script:ResultTable.Clear()
    $txtDetail.Clear()
    $resultFilterTimer.Stop()
    $txtResultFilter.Clear()
    $cmbResultStatus.SelectedIndex = 0
    Update-ResultEventFilterChoices
    $resultFilterTimer.Stop()
    Update-Summary
    Update-ResultFilter
    $statusLabel.Text = 'Sonuçlar temizlendi.'
}

function Set-SearchUiState {
    param([Parameter(Mandatory)][bool]$Searching)

    $script:IsSearching = $Searching
    $btnSearch.Enabled = (-not $Searching) -and (Test-ExchangeCommandAvailable)
    $btnCancel.Enabled = $Searching
    $btnConnect.Enabled = -not $Searching
    $leftPanel.Enabled = -not $Searching
    $progressBar.Visible = $Searching
    $script:MainForm.UseWaitCursor = $Searching
    [System.Windows.Forms.Application]::DoEvents()
}

function Invoke-TrackingSearch {
    if (-not (Test-ExchangeCommandAvailable)) {
        [void](Show-AppMessage -Text 'Önce Exchange bağlantısını kurun veya betiği Exchange Management Shell içinde çalıştırın.' -Icon Warning)
        return
    }

    try {
        $parameters = Get-SearchParameters
        Test-SearchParameters -Parameters $parameters
        $servers = @(Get-QueryServers)

        if ($servers.Count -eq 0) {
        throw 'Aranacak taşıma sunucusu bulunamadı. Sunucu listesini yenileyin veya Geçerli sunucu seçeneğini kullanın.'
        }

        $timeSpan = ([datetime]$parameters.End - [datetime]$parameters.Start)
        $broadSearch = $servers.Count -gt 1 -and $timeSpan.TotalDays -gt 7
        $unlimitedSearch = ([string]$parameters.ResultSize -eq 'Unlimited') -and ($servers.Count -gt 1 -or $timeSpan.TotalHours -gt 24)
        if ($broadSearch -or $unlimitedSearch) {
            $answer = Show-AppMessage -Text 'Bu sorgu birden çok sunucuda geniş bir zaman aralığını veya sınırsız sonucu kapsıyor. İşlem uzun sürebilir ve Exchange sunucularında yük oluşturabilir. Devam edilsin mi?' -Title 'Geniş sorgu onayı' -Icon Warning -Buttons YesNo
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
                return
            }
        }

        Clear-Results
        $script:CancelRequested = $false
        Set-SearchUiState -Searching $true
        $txtLog.Clear()
        Update-QueryPreview
        Add-AppLog -Message ('Arama başladı. Zaman aralığı: {0:dd.MM.yyyy HH:mm} - {1:dd.MM.yyyy HH:mm}' -f $parameters.Start, $parameters.End)
        $startedAt = Get-Date
        $serverErrors = New-Object System.Collections.ArrayList

        for ($serverIndex = 0; $serverIndex -lt $servers.Count; $serverIndex++) {
            if ($script:CancelRequested) { break }

            $serverName = $servers[$serverIndex]
            $serverLabel = if ([string]::IsNullOrWhiteSpace([string]$serverName)) { 'Geçerli sunucu' } else { [string]$serverName }
            $statusLabel.Text = '{0} sorgulanıyor ({1}/{2})...' -f $serverLabel, ($serverIndex + 1), $servers.Count
            $progressBar.Value = [math]::Min(99, [int](($serverIndex / [math]::Max(1, $servers.Count)) * 100))
            [System.Windows.Forms.Application]::DoEvents()

            $serverParameters = @{} + $parameters
            if (-not [string]::IsNullOrWhiteSpace([string]$serverName)) {
                $serverParameters.Server = $serverName
            }

            try {
                Add-AppLog -Message ('Sorgulanıyor: {0}' -f $serverLabel)
                $rawResults = @(Get-MessageTrackingLog @serverParameters -WarningAction SilentlyContinue -ErrorAction Stop)
                $acceptedCount = 0

                foreach ($rawResult in $rawResults) {
                    if ($script:CancelRequested) { break }
                    if (-not $chkIncludeHealth.Checked -and (Test-IsHealthMessage -InputObject $rawResult)) {
                        continue
                    }

                    $normalized = ConvertTo-NormalizedResult -InputObject $rawResult -QueryServer $serverLabel
                    [void]$script:NormalizedResults.Add($normalized)
                    $acceptedCount++

                    if (($acceptedCount % 100) -eq 0) {
                        $statusLabel.Text = '{0}: {1} kayıt işlendi...' -f $serverLabel, $acceptedCount
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                }
                Add-AppLog -Message ('{0}: {1} kayıt eklendi.' -f $serverLabel, $acceptedCount) -Level OK
            }
            catch {
                $errorText = '{0}: {1}' -f $serverLabel, $_.Exception.Message
                [void]$serverErrors.Add($errorText)
                Add-AppLog -Message $errorText -Level ERROR
            }
        }

        Update-ResultsGrid
        $elapsed = (Get-Date) - $startedAt
        $progressBar.Value = 100

        if ($script:CancelRequested) {
            $statusLabel.Text = 'Arama kullanıcı tarafından iptal edildi. Bulunan kısmi sonuçlar gösteriliyor.'
            Add-AppLog -Message $statusLabel.Text -Level WARN
        }
        elseif ($serverErrors.Count -gt 0) {
            $statusLabel.Text = '{0} kayıt bulundu; {1} sunucuda hata oluştu. Ayrıntı için Günlük sekmesine bakın.' -f $script:NormalizedResults.Count, $serverErrors.Count
        }
        else {
            $statusLabel.Text = '{0} kayıt bulundu. Süre: {1:N1} sn.' -f $script:NormalizedResults.Count, $elapsed.TotalSeconds
            Add-AppLog -Message $statusLabel.Text -Level OK
        }

        if ($script:NormalizedResults.Count -eq 0 -and $serverErrors.Count -eq 0) {
            [void](Show-AppMessage -Text 'Seçilen ölçütlerle eşleşen bir ileti izleme kaydı bulunamadı. Tarih aralığını ve filtreleri kontrol edin.')
        }
        elseif ($script:NormalizedResults.Count -eq 0 -and $serverErrors.Count -gt 0) {
            [void](Show-AppMessage -Text ('Arama tamamlanamadı. Günlükteki ilk hata:{0}{1}' -f [Environment]::NewLine, $serverErrors[0]) -Icon Error)
        }
    }
    catch {
        $statusLabel.Text = 'Arama başlatılamadı.'
        Add-AppLog -Message $_.Exception.Message -Level ERROR
        [void](Show-AppMessage -Text $_.Exception.Message -Title 'Arama hatası' -Icon Error)
    }
    finally {
        Set-SearchUiState -Searching $false
    }
}

function Get-FilterSummary {
    $parameters = Get-SearchParameters
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(('Tarih: {0:dd.MM.yyyy HH:mm} - {1:dd.MM.yyyy HH:mm}' -f $parameters.Start, $parameters.End))
    $lines.Add(('Sunucu: {0}' -f $cmbServer.SelectedItem))
    foreach ($name in @('Sender', 'Recipients', 'MessageSubject', 'EventId', 'Source', 'MessageId', 'NetworkMessageId', 'InternalMessageId', 'Reference', 'ResultSize')) {
        if ($parameters.ContainsKey($name)) {
            $lines.Add(('{0}: {1}' -f $name, (ConvertTo-FlatString $parameters[$name])))
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($txtResultFilter.Text)) {
        $lines.Add(('Sonuç metin filtresi: {0}' -f $txtResultFilter.Text.Trim()))
    }
    if ($cmbResultStatus.SelectedIndex -gt 0) {
        $lines.Add(('Sonuç durum filtresi: {0}' -f $cmbResultStatus.SelectedItem))
    }
    if ($cmbResultEvent.SelectedIndex -gt 0) {
        $lines.Add(('Sonuç olay filtresi: {0}' -f $cmbResultEvent.SelectedItem))
    }
    return ($lines -join ' | ')
}

function Export-ResultsCsv {
    $exportResults = @(Get-VisibleResults)
    if ($exportResults.Count -eq 0) { return }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = 'Sonuçları CSV olarak kaydet'
    $dialog.Filter = 'CSV dosyası (*.csv)|*.csv'
    $dialog.FileName = 'MessageTracking_{0:yyyyMMdd_HHmmss}.csv' -f (Get-Date)
    $dialog.AddExtension = $true
    $dialog.OverwritePrompt = $true

    if ($dialog.ShowDialog($script:MainForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        $exportResults |
            Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding UTF8 -UseCulture
        $statusLabel.Text = 'CSV kaydedildi: {0}' -f $dialog.FileName
        Add-AppLog -Message $statusLabel.Text -Level OK
    }
    catch {
        [void](Show-AppMessage -Text ('CSV kaydedilemedi: {0}' -f $_.Exception.Message) -Icon Error)
    }
}

function Export-ResultsHtml {
    $exportResults = @(Get-VisibleResults)
    if ($exportResults.Count -eq 0) { return }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = 'Sonuçları HTML raporu olarak kaydet'
    $dialog.Filter = 'HTML dosyası (*.html)|*.html'
    $dialog.FileName = 'MessageTracking_{0:yyyyMMdd_HHmmss}.html' -f (Get-Date)
    $dialog.AddExtension = $true
    $dialog.OverwritePrompt = $true

    if ($dialog.ShowDialog($script:MainForm) -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        $encodedSummary = [System.Net.WebUtility]::HtmlEncode((Get-FilterSummary))
        $head = @'
<meta charset="utf-8">
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#23303d;background:#f4f7fa}
h1{color:#182d4c;margin-bottom:4px} .meta{color:#64707c;margin-bottom:20px}
table{border-collapse:collapse;width:100%;background:white;font-size:12px}
th{background:#182d4c;color:white;text-align:left;position:sticky;top:0}
th,td{padding:7px;border:1px solid #d2dae3;vertical-align:top}
tr:nth-child(even){background:#f8fafc}
</style>
'@
        $preContent = '<h1>Exchange Message Tracking Raporu</h1><div class="meta">Oluşturulma: {0:dd.MM.yyyy HH:mm:ss}<br>{1}<br>Dışa aktarılan kayıt: {2}</div>' -f (Get-Date), $encodedSummary, $exportResults.Count
        $reportRows = $exportResults | Select-Object Timestamp, Status, EventId, Description, Source, Server, Sender, Recipients, Subject, RecipientStatus, SizeKB, MessageId, NetworkMessageId, InternalMessageId, ClientHostname, ConnectorId, Directionality, SourceContext, Reference
        $html = $reportRows | ConvertTo-Html -Title 'Exchange Message Tracking Raporu' -Head $head -PreContent $preContent
        $html | Set-Content -LiteralPath $dialog.FileName -Encoding UTF8
        $statusLabel.Text = 'HTML raporu kaydedildi: {0}' -f $dialog.FileName
        Add-AppLog -Message $statusLabel.Text -Level OK
    }
    catch {
        [void](Show-AppMessage -Text ('HTML raporu kaydedilemedi: {0}' -f $_.Exception.Message) -Icon Error)
    }
}

function Add-GridTextColumn {
    param(
        [Parameter(Mandatory)][string]$Property,
        [Parameter(Mandatory)][string]$Header,
        [Parameter(Mandatory)][int]$Width,
        [Parameter()][bool]$Visible = $true,
        [Parameter()][string]$Format = ''
    )

    $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $column.Name = $Property
    $column.DataPropertyName = $Property
    $column.HeaderText = $Header
    $column.Width = $Width
    $column.Visible = $Visible
    if (-not [string]::IsNullOrWhiteSpace($Format)) {
        $column.DefaultCellStyle.Format = $Format
    }
    [void]$gridResults.Columns.Add($column)
}

#region Form and controls
$script:MainForm = New-Object System.Windows.Forms.Form
$script:MainForm.Text = 'Exchange Message Tracking Pro | emrahtolu'
$script:MainForm.StartPosition = 'CenterScreen'
$script:MainForm.MinimumSize = New-Object System.Drawing.Size(1180, 760)
$script:MainForm.Size = New-Object System.Drawing.Size(1480, 900)
$script:MainForm.BackColor = $script:Colors.Background
$script:MainForm.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$script:MainForm.AutoScaleMode = 'Dpi'
$script:MainForm.KeyPreview = $true

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = 'Top'
$headerPanel.Height = 82
$headerPanel.BackColor = $script:Colors.Navy
$headerPanel.Padding = New-Object System.Windows.Forms.Padding(18, 10, 18, 8)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'Exchange Message Tracking Pro'
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 17)
$lblTitle.Location = New-Object System.Drawing.Point(18, 9)
$lblTitle.AutoSize = $true

$lblBrand = New-Object System.Windows.Forms.LinkLabel
$lblBrand.Text = 'emrahtolu'
$lblBrand.Location = New-Object System.Drawing.Point(382, 18)
$lblBrand.Size = New-Object System.Drawing.Size(92, 23)
$lblBrand.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$lblBrand.LinkColor = [System.Drawing.Color]::FromArgb(116, 190, 255)
$lblBrand.ActiveLinkColor = [System.Drawing.Color]::White
$lblBrand.VisitedLinkColor = [System.Drawing.Color]::FromArgb(116, 190, 255)
$lblBrand.LinkBehavior = 'HoverUnderline'
$lblBrand.Cursor = [System.Windows.Forms.Cursors]::Hand
$lblBrand.BackColor = $script:Colors.Navy

$lblSubtitle = New-Object System.Windows.Forms.Label
$lblSubtitle.Text = 'On-premises Exchange ileti akışı arama ve olay analizi'
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(194, 211, 230)
$lblSubtitle.Location = New-Object System.Drawing.Point(21, 45)
$lblSubtitle.AutoSize = $true

$txtConnection = New-Object System.Windows.Forms.TextBox
$txtConnection.Anchor = 'Top,Right'
$txtConnection.Location = New-Object System.Drawing.Point(920, 18)
$txtConnection.Size = New-Object System.Drawing.Size(260, 25)
$txtConnection.Text = $ExchangeServerFqdn

$chkCredential = New-Object System.Windows.Forms.CheckBox
$chkCredential.Anchor = 'Top,Right'
$chkCredential.Location = New-Object System.Drawing.Point(920, 48)
$chkCredential.Size = New-Object System.Drawing.Size(155, 22)
$chkCredential.Text = 'Farklı kullanıcı kullan'
$chkCredential.Checked = [bool]$UseCredential
$chkCredential.ForeColor = [System.Drawing.Color]::White
$chkCredential.BackColor = $script:Colors.Navy

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Anchor = 'Top,Right'
$btnConnect.Location = New-Object System.Drawing.Point(1190, 17)
$btnConnect.Size = New-Object System.Drawing.Size(105, 29)
$btnConnect.Text = 'Bağlan / Yenile'
$btnConnect.FlatStyle = 'Flat'
$btnConnect.BackColor = $script:Colors.Blue
$btnConnect.ForeColor = [System.Drawing.Color]::White
$btnConnect.FlatAppearance.BorderSize = 0

$lblConnection = New-Object System.Windows.Forms.Label
$lblConnection.Anchor = 'Top,Right'
$lblConnection.Location = New-Object System.Drawing.Point(1080, 50)
$lblConnection.Size = New-Object System.Drawing.Size(355, 20)
$lblConnection.TextAlign = 'MiddleRight'
$lblConnection.Text = 'Kontrol ediliyor...'
$lblConnection.ForeColor = [System.Drawing.Color]::White

$headerPanel.Controls.AddRange(@($lblTitle, $lblBrand, $lblSubtitle, $txtConnection, $chkCredential, $btnConnect, $lblConnection))

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.SizingGrip = $false
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Spring = $true
$statusLabel.TextAlign = 'MiddleLeft'
$statusLabel.Text = 'Hazır.'
$progressBar = New-Object System.Windows.Forms.ToolStripProgressBar
$progressBar.Size = New-Object System.Drawing.Size(180, 16)
$progressBar.Visible = $false
[void]$statusStrip.Items.Add($statusLabel)
[void]$statusStrip.Items.Add($progressBar)

$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = 'Fill'
$mainPanel.BackColor = $script:Colors.Background

$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Dock = 'Left'
$leftPanel.Width = 370
$leftPanel.AutoScroll = $true
$leftPanel.BackColor = [System.Drawing.Color]::White
$leftPanel.Padding = New-Object System.Windows.Forms.Padding(12)

$grpTime = New-Object System.Windows.Forms.GroupBox
$grpTime.Text = 'Zaman aralığı'
$grpTime.Location = New-Object System.Drawing.Point(12, 10)
$grpTime.Size = New-Object System.Drawing.Size(340, 112)

$lblStart = New-Object System.Windows.Forms.Label
$lblStart.Text = 'Başlangıç'
$lblStart.Location = New-Object System.Drawing.Point(14, 26)
$lblStart.Size = New-Object System.Drawing.Size(75, 22)
$dtpStart = New-Object System.Windows.Forms.DateTimePicker
$dtpStart.Location = New-Object System.Drawing.Point(92, 24)
$dtpStart.Size = New-Object System.Drawing.Size(230, 25)
$dtpStart.Format = 'Custom'
$dtpStart.CustomFormat = 'dd.MM.yyyy HH:mm'
$dtpStart.Value = (Get-Date).AddHours(-24)

$lblEnd = New-Object System.Windows.Forms.Label
$lblEnd.Text = 'Bitiş'
$lblEnd.Location = New-Object System.Drawing.Point(14, 65)
$lblEnd.Size = New-Object System.Drawing.Size(75, 22)
$dtpEnd = New-Object System.Windows.Forms.DateTimePicker
$dtpEnd.Location = New-Object System.Drawing.Point(92, 63)
$dtpEnd.Size = New-Object System.Drawing.Size(230, 25)
$dtpEnd.Format = 'Custom'
$dtpEnd.CustomFormat = 'dd.MM.yyyy HH:mm'
$dtpEnd.Value = Get-Date
$grpTime.Controls.AddRange(@($lblStart, $dtpStart, $lblEnd, $dtpEnd))

$grpBasic = New-Object System.Windows.Forms.GroupBox
$grpBasic.Text = 'Temel filtreler'
$grpBasic.Location = New-Object System.Drawing.Point(12, 130)
$grpBasic.Size = New-Object System.Drawing.Size(340, 225)

$basicLabels = @('Gönderen', 'Alıcılar', 'Konu', 'Olay', 'Kaynak')
for ($i = 0; $i -lt $basicLabels.Count; $i++) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $basicLabels[$i]
    $label.Location = New-Object System.Drawing.Point(14, (27 + ($i * 38)))
    $label.Size = New-Object System.Drawing.Size(75, 23)
    $grpBasic.Controls.Add($label)
}

$txtSender = New-Object System.Windows.Forms.TextBox
$txtSender.Location = New-Object System.Drawing.Point(92, 24)
$txtSender.Size = New-Object System.Drawing.Size(230, 25)

$txtRecipients = New-Object System.Windows.Forms.TextBox
$txtRecipients.Location = New-Object System.Drawing.Point(92, 62)
$txtRecipients.Size = New-Object System.Drawing.Size(230, 25)
$txtRecipients.Tag = 'Birden çok adres için ; kullanın'

$txtSubject = New-Object System.Windows.Forms.TextBox
$txtSubject.Location = New-Object System.Drawing.Point(92, 100)
$txtSubject.Size = New-Object System.Drawing.Size(230, 25)

$cmbEvent = New-Object System.Windows.Forms.ComboBox
$cmbEvent.Location = New-Object System.Drawing.Point(92, 138)
$cmbEvent.Size = New-Object System.Drawing.Size(230, 25)
$cmbEvent.DropDownStyle = 'DropDownList'
[void]$cmbEvent.Items.Add('[Tümü]')
foreach ($eventName in ($script:EventDescriptions.Keys | Sort-Object)) { [void]$cmbEvent.Items.Add($eventName) }
$cmbEvent.SelectedIndex = 0

$cmbSource = New-Object System.Windows.Forms.ComboBox
$cmbSource.Location = New-Object System.Drawing.Point(92, 176)
$cmbSource.Size = New-Object System.Drawing.Size(230, 25)
$cmbSource.DropDownStyle = 'DropDownList'
[void]$cmbSource.Items.AddRange(@('[Tümü]', 'ADMIN', 'AGENT', 'APPROVAL', 'BOOTLOADER', 'DNS', 'DSN', 'GATEWAY', 'MAILBOXRULE', 'MEETINGMESSAGEPROCESSOR', 'ORAR', 'PICKUP', 'POISONMESSAGE', 'PUBLICFOLDER', 'QUEUE', 'REDUNDANCY', 'RESOLVER', 'ROUTING', 'SAFETYNET', 'SMTP', 'STOREDRIVER'))
$cmbSource.SelectedIndex = 0
$grpBasic.Controls.AddRange(@($txtSender, $txtRecipients, $txtSubject, $cmbEvent, $cmbSource))

$grpAdvanced = New-Object System.Windows.Forms.GroupBox
$grpAdvanced.Text = 'Gelişmiş kimlik filtreleri'
$grpAdvanced.Location = New-Object System.Drawing.Point(12, 363)
$grpAdvanced.Size = New-Object System.Drawing.Size(340, 188)

$advancedLabels = @('Message-ID', 'Network ID', 'Internal ID', 'Reference')
for ($i = 0; $i -lt $advancedLabels.Count; $i++) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $advancedLabels[$i]
    $label.Location = New-Object System.Drawing.Point(14, (27 + ($i * 38)))
    $label.Size = New-Object System.Drawing.Size(75, 23)
    $grpAdvanced.Controls.Add($label)
}

$txtMessageId = New-Object System.Windows.Forms.TextBox
$txtMessageId.Location = New-Object System.Drawing.Point(92, 24)
$txtMessageId.Size = New-Object System.Drawing.Size(230, 25)
$txtNetworkMessageId = New-Object System.Windows.Forms.TextBox
$txtNetworkMessageId.Location = New-Object System.Drawing.Point(92, 62)
$txtNetworkMessageId.Size = New-Object System.Drawing.Size(230, 25)
$txtInternalMessageId = New-Object System.Windows.Forms.TextBox
$txtInternalMessageId.Location = New-Object System.Drawing.Point(92, 100)
$txtInternalMessageId.Size = New-Object System.Drawing.Size(230, 25)
$txtReference = New-Object System.Windows.Forms.TextBox
$txtReference.Location = New-Object System.Drawing.Point(92, 138)
$txtReference.Size = New-Object System.Drawing.Size(230, 25)
$grpAdvanced.Controls.AddRange(@($txtMessageId, $txtNetworkMessageId, $txtInternalMessageId, $txtReference))

$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Text = 'Kapsam ve seçenekler'
$grpOptions.Location = New-Object System.Drawing.Point(12, 559)
$grpOptions.Size = New-Object System.Drawing.Size(340, 142)

$lblServer = New-Object System.Windows.Forms.Label
$lblServer.Text = 'Sunucu'
$lblServer.Location = New-Object System.Drawing.Point(14, 27)
$lblServer.Size = New-Object System.Drawing.Size(75, 23)
$cmbServer = New-Object System.Windows.Forms.ComboBox
$cmbServer.Location = New-Object System.Drawing.Point(92, 24)
$cmbServer.Size = New-Object System.Drawing.Size(230, 25)
$cmbServer.DropDownStyle = 'DropDownList'
[void]$cmbServer.Items.Add('[Geçerli sunucu]')
$cmbServer.SelectedIndex = 0

$lblResultSize = New-Object System.Windows.Forms.Label
$lblResultSize.Text = 'Sunucu başı'
$lblResultSize.Location = New-Object System.Drawing.Point(14, 65)
$lblResultSize.Size = New-Object System.Drawing.Size(75, 23)
$cmbResultSize = New-Object System.Windows.Forms.ComboBox
$cmbResultSize.Location = New-Object System.Drawing.Point(92, 62)
$cmbResultSize.Size = New-Object System.Drawing.Size(230, 25)
$cmbResultSize.DropDownStyle = 'DropDownList'
[void]$cmbResultSize.Items.AddRange(@('100', '500', '1000', '5000', 'Sınırsız'))
$cmbResultSize.SelectedItem = '1000'

$chkIncludeHealth = New-Object System.Windows.Forms.CheckBox
$chkIncludeHealth.Text = 'HealthMailbox iletilerini göster'
$chkIncludeHealth.Location = New-Object System.Drawing.Point(92, 101)
$chkIncludeHealth.Size = New-Object System.Drawing.Size(230, 25)
$grpOptions.Controls.AddRange(@($lblServer, $cmbServer, $lblResultSize, $cmbResultSize, $chkIncludeHealth))

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = 'ARA'
$btnSearch.Location = New-Object System.Drawing.Point(12, 714)
$btnSearch.Size = New-Object System.Drawing.Size(224, 42)
$btnSearch.FlatStyle = 'Flat'
$btnSearch.BackColor = $script:Colors.Blue
$btnSearch.ForeColor = [System.Drawing.Color]::White
$btnSearch.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 10)
$btnSearch.FlatAppearance.BorderSize = 0
$btnSearch.Enabled = $false

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'İptal'
$btnCancel.Location = New-Object System.Drawing.Point(244, 714)
$btnCancel.Size = New-Object System.Drawing.Size(108, 42)
$btnCancel.FlatStyle = 'Flat'
$btnCancel.BackColor = [System.Drawing.Color]::White
$btnCancel.ForeColor = $script:Colors.Error
$btnCancel.FlatAppearance.BorderColor = $script:Colors.Error
$btnCancel.Enabled = $false

$leftPanel.Controls.AddRange(@($grpTime, $grpBasic, $grpAdvanced, $grpOptions, $btnSearch, $btnCancel))

$rightPanel = New-Object System.Windows.Forms.Panel
$rightPanel.Dock = 'Fill'
$rightPanel.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 8)
$rightPanel.BackColor = $script:Colors.Background

$summaryPanel = New-Object System.Windows.Forms.Panel
$summaryPanel.Dock = 'Top'
$summaryPanel.Height = 70
$summaryPanel.BackColor = $script:Colors.Background

function New-SummaryCard {
    param([int]$X, [string]$Title, [System.Drawing.Color]$Accent)
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Location = New-Object System.Drawing.Point($X, 2)
    $panel.Size = New-Object System.Drawing.Size(180, 58)
    $panel.BackColor = [System.Drawing.Color]::White
    $panel.BorderStyle = 'FixedSingle'
    $bar = New-Object System.Windows.Forms.Panel
    $bar.Dock = 'Left'
    $bar.Width = 5
    $bar.BackColor = $Accent
    $value = New-Object System.Windows.Forms.Label
    $value.Text = '0'
    $value.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 16)
    $value.ForeColor = $script:Colors.Text
    $value.Location = New-Object System.Drawing.Point(16, 5)
    $value.Size = New-Object System.Drawing.Size(150, 28)
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = $Title
    $titleLabel.ForeColor = $script:Colors.Muted
    $titleLabel.Location = New-Object System.Drawing.Point(17, 35)
    $titleLabel.Size = New-Object System.Drawing.Size(150, 18)
    $panel.Controls.AddRange(@($value, $titleLabel, $bar))
    $summaryPanel.Controls.Add($panel)
    return $value
}

$lblRowsValue = New-SummaryCard -X 0 -Title 'Toplam olay kaydı' -Accent $script:Colors.Blue
$lblMessagesValue = New-SummaryCard -X 190 -Title 'Benzersiz ileti' -Accent $script:Colors.Info
$lblDeliveredValue = New-SummaryCard -X 380 -Title 'Teslim (DELIVER)' -Accent $script:Colors.Success
$lblProblemsValue = New-SummaryCard -X 570 -Title 'Sorun / inceleme' -Accent $script:Colors.Error

$toolbarPanel = New-Object System.Windows.Forms.Panel
$toolbarPanel.Dock = 'Top'
$toolbarPanel.Height = 82
$toolbarPanel.BackColor = $script:Colors.Background

function New-ToolbarButton {
    param([string]$Text, [int]$X, [int]$Width, [int]$Y = 4)
    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Location = New-Object System.Drawing.Point($X, $Y)
    $button.Size = New-Object System.Drawing.Size($Width, 31)
    $button.FlatStyle = 'Flat'
    $button.BackColor = [System.Drawing.Color]::White
    $button.ForeColor = $script:Colors.Text
    $button.FlatAppearance.BorderColor = $script:Colors.Border
    $toolbarPanel.Controls.Add($button)
    return $button
}

$btnExportCsv = New-ToolbarButton -Text 'CSV dışa aktar' -X 0 -Width 120
$btnExportHtml = New-ToolbarButton -Text 'HTML raporu' -X 128 -Width 110
$btnCopy = New-ToolbarButton -Text 'Seçileni kopyala' -X 246 -Width 125
$btnClear = New-ToolbarButton -Text 'Temizle' -X 379 -Width 85
$btnExportCsv.Enabled = $false
$btnExportHtml.Enabled = $false
$btnCopy.Enabled = $false

$lblFilterCount = New-Object System.Windows.Forms.Label
$lblFilterCount.Anchor = 'Top,Left'
$lblFilterCount.Location = New-Object System.Drawing.Point(820, 10)
$lblFilterCount.Size = New-Object System.Drawing.Size(235, 22)
$lblFilterCount.Text = 'Gösterilen: 0 / 0'
$lblFilterCount.TextAlign = 'MiddleRight'
$lblFilterCount.ForeColor = $script:Colors.Muted

$lblResultFilter = New-Object System.Windows.Forms.Label
$lblResultFilter.Location = New-Object System.Drawing.Point(0, 49)
$lblResultFilter.Size = New-Object System.Drawing.Size(94, 23)
$lblResultFilter.Text = 'Sonuçlarda ara'
$lblResultFilter.TextAlign = 'MiddleLeft'

$txtResultFilter = New-Object System.Windows.Forms.TextBox
$txtResultFilter.Location = New-Object System.Drawing.Point(96, 47)
$txtResultFilter.Size = New-Object System.Drawing.Size(250, 25)

$cmbResultStatus = New-Object System.Windows.Forms.ComboBox
$cmbResultStatus.Location = New-Object System.Drawing.Point(354, 47)
$cmbResultStatus.Size = New-Object System.Drawing.Size(145, 25)
$cmbResultStatus.DropDownStyle = 'DropDownList'
[void]$cmbResultStatus.Items.AddRange(@('[Tüm sonuçlar]', 'Teslim edildi', 'Başarısız', 'Gecikme', 'Teslim edilmedi', 'Aktarım', 'Alındı', 'İşlendi', 'Bildirim', 'Yönlendirme', 'Onay süreci', 'İnceleme gerekli', 'Bilgi'))
$cmbResultStatus.SelectedIndex = 0

$cmbResultEvent = New-Object System.Windows.Forms.ComboBox
$cmbResultEvent.Location = New-Object System.Drawing.Point(507, 47)
$cmbResultEvent.Size = New-Object System.Drawing.Size(125, 25)
$cmbResultEvent.DropDownStyle = 'DropDownList'
[void]$cmbResultEvent.Items.Add('[Tüm olaylar]')
$cmbResultEvent.SelectedIndex = 0

$btnClearFilter = New-ToolbarButton -Text 'Filtreyi sıfırla' -X 640 -Width 112 -Y 43
$toolbarPanel.Controls.AddRange(@($lblFilterCount, $lblResultFilter, $txtResultFilter, $cmbResultStatus, $cmbResultEvent))
$toolbarPanel.Add_Resize({
    $lblFilterCount.Left = [math]::Max(470, $toolbarPanel.ClientSize.Width - $lblFilterCount.Width)
})

$resultFilterTimer = New-Object System.Windows.Forms.Timer
$resultFilterTimer.Interval = 250

$splitResults = New-Object System.Windows.Forms.SplitContainer
$splitResults.Dock = 'Fill'
$splitResults.Orientation = 'Horizontal'
$splitResults.SplitterDistance = 510
$splitResults.Panel1MinSize = 240
$splitResults.Panel2MinSize = 130
$splitResults.BackColor = $script:Colors.Border

$gridResults = New-Object System.Windows.Forms.DataGridView
$gridResults.Dock = 'Fill'
$gridResults.BackgroundColor = [System.Drawing.Color]::White
$gridResults.BorderStyle = 'FixedSingle'
$gridResults.AutoGenerateColumns = $false
$gridResults.AllowUserToAddRows = $false
$gridResults.AllowUserToDeleteRows = $false
$gridResults.AllowUserToResizeRows = $false
$gridResults.ReadOnly = $true
$gridResults.MultiSelect = $true
$gridResults.SelectionMode = 'FullRowSelect'
$gridResults.RowHeadersVisible = $false
$gridResults.AutoSizeRowsMode = 'None'
$gridResults.RowTemplate.Height = 25
$gridResults.ColumnHeadersHeight = 32
$gridResults.EnableHeadersVisualStyles = $false
$gridResults.ColumnHeadersDefaultCellStyle.BackColor = $script:Colors.Navy
$gridResults.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$gridResults.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font('Segoe UI Semibold', 9)
$gridResults.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(204, 225, 247)
$gridResults.DefaultCellStyle.SelectionForeColor = $script:Colors.Text
$gridResults.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$gridResults.ClipboardCopyMode = 'EnableAlwaysIncludeHeaderText'

Add-GridTextColumn -Property Timestamp -Header 'Zaman' -Width 130 -Format 'dd.MM.yyyy HH:mm:ss'
Add-GridTextColumn -Property Status -Header 'Sonuç' -Width 125
Add-GridTextColumn -Property Severity -Header 'Önem' -Width 75 -Visible $false
Add-GridTextColumn -Property EventId -Header 'Olay' -Width 110
Add-GridTextColumn -Property Description -Header 'Olayın anlamı' -Width 245
Add-GridTextColumn -Property Source -Header 'Kaynak' -Width 100
Add-GridTextColumn -Property Server -Header 'Sunucu' -Width 120
Add-GridTextColumn -Property Sender -Header 'Gönderen' -Width 190
Add-GridTextColumn -Property Recipients -Header 'Alıcılar' -Width 230
Add-GridTextColumn -Property Subject -Header 'Konu' -Width 260
Add-GridTextColumn -Property RecipientStatus -Header 'Alıcı durumu' -Width 230
Add-GridTextColumn -Property SizeKB -Header 'KB' -Width 70 -Format 'N2'
Add-GridTextColumn -Property MessageId -Header 'Message-ID' -Width 250
Add-GridTextColumn -Property NetworkMessageId -Header 'Network ID' -Width 220 -Visible $false
Add-GridTextColumn -Property InternalMessageId -Header 'Internal ID' -Width 100 -Visible $false
Add-GridTextColumn -Property ClientHostname -Header 'İstemci' -Width 150 -Visible $false
Add-GridTextColumn -Property ConnectorId -Header 'Bağlayıcı' -Width 180 -Visible $false

$script:ResultTable = New-ResultTable
$gridResults.DataSource = $script:ResultTable.DefaultView
$splitResults.Panel1.Controls.Add($gridResults)

$detailTabs = New-Object System.Windows.Forms.TabControl
$detailTabs.Dock = 'Fill'

$tabDetail = New-Object System.Windows.Forms.TabPage
$tabDetail.Text = 'Seçili kayıt ayrıntısı'
$txtDetail = New-Object System.Windows.Forms.RichTextBox
$txtDetail.Dock = 'Fill'
$txtDetail.ReadOnly = $true
$txtDetail.BorderStyle = 'None'
$txtDetail.BackColor = [System.Drawing.Color]::White
$txtDetail.Font = New-Object System.Drawing.Font('Consolas', 9)
$tabDetail.Controls.Add($txtDetail)

$tabLog = New-Object System.Windows.Forms.TabPage
$tabLog.Text = 'Sorgu ve günlük'
$logSplit = New-Object System.Windows.Forms.SplitContainer
$logSplit.Dock = 'Fill'
$logSplit.Orientation = 'Horizontal'
$logSplit.SplitterDistance = 48
$txtQueryPreview = New-Object System.Windows.Forms.TextBox
$txtQueryPreview.Dock = 'Fill'
$txtQueryPreview.ReadOnly = $true
$txtQueryPreview.Multiline = $true
$txtQueryPreview.BackColor = [System.Drawing.Color]::FromArgb(248, 250, 252)
$txtQueryPreview.Font = New-Object System.Drawing.Font('Consolas', 8.5)
$txtLog = New-Object System.Windows.Forms.RichTextBox
$txtLog.Dock = 'Fill'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$txtLog.Font = New-Object System.Drawing.Font('Consolas', 8.5)
$logSplit.Panel1.Controls.Add($txtQueryPreview)
$logSplit.Panel2.Controls.Add($txtLog)
$tabLog.Controls.Add($logSplit)

[void]$detailTabs.TabPages.Add($tabDetail)
[void]$detailTabs.TabPages.Add($tabLog)
$splitResults.Panel2.Controls.Add($detailTabs)

$rightPanel.Controls.Add($splitResults)
$rightPanel.Controls.Add($toolbarPanel)
$rightPanel.Controls.Add($summaryPanel)
$mainPanel.Controls.Add($rightPanel)
$mainPanel.Controls.Add($leftPanel)

$script:MainForm.Controls.Add($mainPanel)
$script:MainForm.Controls.Add($statusStrip)
$script:MainForm.Controls.Add($headerPanel)
#endregion Form and controls

#region Events
$lblBrand.Add_LinkClicked({
    try {
        Start-Process -FilePath 'https://www.emrahtolu.com'
    }
    catch {
        [void](Show-AppMessage -Text ('Web adresi açılamadı: {0}' -f $_.Exception.Message) -Icon Error)
    }
})

$btnConnect.Add_Click({
    if ($script:IsSearching) { return }

    try {
        $btnConnect.Enabled = $false
        $script:MainForm.UseWaitCursor = $true

        if (Test-ExchangeCommandAvailable) {
            Update-ServerList
            Set-ConnectionState -Connected $true -Text 'Exchange komutları hazır'
            $statusLabel.Text = 'Exchange bağlantısı yenilendi.'
        }
        else {
            Connect-ExchangeRemote -ServerOrUri $txtConnection.Text -PromptForCredential:$chkCredential.Checked
            Update-ServerList
            Set-ConnectionState -Connected $true -Text 'Uzak Exchange oturumu bağlı'
            $statusLabel.Text = 'Uzak Exchange bağlantısı kuruldu.'
            Add-AppLog -Message $statusLabel.Text -Level OK
        }
        Update-QueryPreview
    }
    catch {
        Set-ConnectionState -Connected $false -Text 'Bağlantı başarısız'
        $statusLabel.Text = 'Exchange bağlantısı kurulamadı.'
        Add-AppLog -Message $_.Exception.Message -Level ERROR
        [void](Show-AppMessage -Text ('Exchange bağlantısı kurulamadı.{0}{0}{1}{0}{0}Betik Exchange Management Shell içinde de doğrudan çalıştırılabilir.' -f [Environment]::NewLine, $_.Exception.Message) -Title 'Bağlantı hatası' -Icon Error)
    }
    finally {
        $script:MainForm.UseWaitCursor = $false
        $btnConnect.Enabled = $true
    }
})

$btnSearch.Add_Click({ Invoke-TrackingSearch })
$btnCancel.Add_Click({
    $script:CancelRequested = $true
    $btnCancel.Enabled = $false
    $statusLabel.Text = 'İptal isteği alındı; geçerli kayıt grubu tamamlanıyor...'
})
$btnExportCsv.Add_Click({ Export-ResultsCsv })
$btnExportHtml.Add_Click({ Export-ResultsHtml })
$btnClear.Add_Click({ Clear-Results })
$btnClearFilter.Add_Click({ Clear-ResultFilter })
$txtResultFilter.Add_TextChanged({
    $resultFilterTimer.Stop()
    $resultFilterTimer.Start()
})
$cmbResultStatus.Add_SelectedIndexChanged({ Update-ResultFilter })
$cmbResultEvent.Add_SelectedIndexChanged({ Update-ResultFilter })
$resultFilterTimer.Add_Tick({
    $resultFilterTimer.Stop()
    Update-ResultFilter
})

$btnCopy.Add_Click({
    try {
        if ($gridResults.SelectedRows.Count -eq 0) {
            [void](Show-AppMessage -Text 'Önce kopyalanacak bir veya daha fazla satırı seçin.')
            return
        }
        $clipboardContent = $gridResults.GetClipboardContent()
        if ($null -ne $clipboardContent) {
            [System.Windows.Forms.Clipboard]::SetDataObject($clipboardContent, $true)
            $statusLabel.Text = '{0} satır panoya kopyalandı.' -f $gridResults.SelectedRows.Count
        }
    }
    catch {
        [void](Show-AppMessage -Text ('Panoya kopyalanamadı: {0}' -f $_.Exception.Message) -Icon Error)
    }
})

$gridResults.Add_CellFormatting({
    param($sender, $eventArgs)
    if ($eventArgs.RowIndex -lt 0) { return }

    $severityValue = [string]$gridResults.Rows[$eventArgs.RowIndex].Cells['Severity'].Value
    switch ($severityValue) {
        'Success' {
            $gridResults.Rows[$eventArgs.RowIndex].DefaultCellStyle.BackColor = $script:Colors.SuccessBg
            $gridResults.Rows[$eventArgs.RowIndex].Cells['Status'].Style.ForeColor = $script:Colors.Success
        }
        'Warning' {
            $gridResults.Rows[$eventArgs.RowIndex].DefaultCellStyle.BackColor = $script:Colors.WarningBg
            $gridResults.Rows[$eventArgs.RowIndex].Cells['Status'].Style.ForeColor = $script:Colors.Warning
        }
        'Error' {
            $gridResults.Rows[$eventArgs.RowIndex].DefaultCellStyle.BackColor = $script:Colors.ErrorBg
            $gridResults.Rows[$eventArgs.RowIndex].Cells['Status'].Style.ForeColor = $script:Colors.Error
        }
        default {
            $gridResults.Rows[$eventArgs.RowIndex].DefaultCellStyle.BackColor = if (($eventArgs.RowIndex % 2) -eq 0) { [System.Drawing.Color]::White } else { [System.Drawing.Color]::FromArgb(248, 250, 252) }
            $gridResults.Rows[$eventArgs.RowIndex].Cells['Status'].Style.ForeColor = $script:Colors.Info
        }
    }
})

$gridResults.Add_SelectionChanged({
    if ($null -eq $gridResults.CurrentRow -or $null -eq $gridResults.CurrentRow.DataBoundItem) { return }
    $view = $gridResults.CurrentRow.DataBoundItem
    if (-not ($view -is [System.Data.DataRowView])) { return }

    $detailMap = [ordered]@{
        'Zaman'             = 'Timestamp'
        'Sonuç'             = 'Status'
        'Olay'              = 'EventId'
        'Olayın anlamı'     = 'Description'
        'Kaynak'            = 'Source'
        'Sunucu'            = 'Server'
        'Sorgulanan sunucu' = 'QueryServer'
        'Gönderen'          = 'Sender'
        'Alıcılar'          = 'Recipients'
        'Konu'              = 'Subject'
        'Alıcı durumu'      = 'RecipientStatus'
        'Boyut (KB)'        = 'SizeKB'
        'Message-ID'        = 'MessageId'
        'Network Message ID'= 'NetworkMessageId'
        'Internal Message ID'= 'InternalMessageId'
        'İstemci adı'       = 'ClientHostname'
        'İstemci IP'        = 'ClientIp'
        'Sunucu IP'         = 'ServerIp'
        'Bağlayıcı'         = 'ConnectorId'
        'Yön'               = 'Directionality'
        'İlişkili alıcı'    = 'RelatedRecipientAddress'
        'Return-Path'       = 'ReturnPath'
        'Source Context'    = 'SourceContext'
        'Reference'         = 'Reference'
        'Message Info'      = 'MessageInfo'
        'Custom Data'       = 'CustomData'
        'Event Data'        = 'EventData'
    }

    $builder = New-Object System.Text.StringBuilder
    foreach ($entry in $detailMap.GetEnumerator()) {
        $value = $view.Row[$entry.Value]
        if ($value -is [DBNull]) { $value = '' }
        [void]$builder.AppendLine(('{0,-21}: {1}' -f $entry.Key, $value))
    }
    $txtDetail.Text = $builder.ToString()
})

$previewControls = @($dtpStart, $dtpEnd, $txtSender, $txtRecipients, $txtSubject, $cmbEvent, $cmbSource, $txtMessageId, $txtNetworkMessageId, $txtInternalMessageId, $txtReference, $cmbServer, $cmbResultSize)
foreach ($control in $previewControls) {
    if ($control -is [System.Windows.Forms.TextBox]) {
        $control.Add_TextChanged({ Update-QueryPreview })
    }
    elseif ($control -is [System.Windows.Forms.DateTimePicker]) {
        $control.Add_ValueChanged({ Update-QueryPreview })
    }
    else {
        $control.Add_SelectedIndexChanged({ Update-QueryPreview })
    }
}

$script:MainForm.Add_KeyDown({
    param($sender, $eventArgs)
    if ($eventArgs.Control -and $eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::F) {
        $txtResultFilter.Focus()
        $txtResultFilter.SelectAll()
        $eventArgs.SuppressKeyPress = $true
    }
    elseif ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::F5 -and $btnSearch.Enabled) {
        Invoke-TrackingSearch
        $eventArgs.SuppressKeyPress = $true
    }
    elseif ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Escape -and $script:IsSearching) {
        $script:CancelRequested = $true
        $eventArgs.SuppressKeyPress = $true
    }
    elseif ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Escape -and
        (-not [string]::IsNullOrWhiteSpace($txtResultFilter.Text) -or $cmbResultStatus.SelectedIndex -gt 0 -or $cmbResultEvent.SelectedIndex -gt 0)) {
        Clear-ResultFilter
        $eventArgs.SuppressKeyPress = $true
    }
})

$script:MainForm.Add_Shown({
    Add-AppLog -Message 'Uygulama başlatıldı.'
    if (Test-ExchangeCommandAvailable) {
        Set-ConnectionState -Connected $true -Text 'Exchange komutları hazır'
        Update-ServerList
        Update-QueryPreview
        $statusLabel.Text = 'Hazır. F5 ile arama yapabilirsiniz.'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($ExchangeServerFqdn)) {
        $btnConnect.PerformClick()
    }
    else {
        Set-ConnectionState -Connected $false -Text 'Exchange bağlantısı gerekli'
        $txtLog.AppendText('Exchange komutları bulunamadı. Üst alana Exchange sunucu adını yazıp Bağlan / Yenile düğmesini kullanın veya betiği Exchange Management Shell içinde çalıştırın.' + [Environment]::NewLine)
        $statusLabel.Text = 'Exchange bağlantısı bekleniyor.'
    }
})

$script:MainForm.Add_FormClosing({
    param($sender, $eventArgs)
    if ($script:IsSearching) {
        $answer = Show-AppMessage -Text 'Arama devam ediyor. Uygulama kapatılsın mı?' -Title 'Çıkış onayı' -Icon Warning -Buttons YesNo
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $eventArgs.Cancel = $true
            return
        }
        $script:CancelRequested = $true
    }
})

$script:MainForm.Add_FormClosed({
    $resultFilterTimer.Stop()
    $resultFilterTimer.Dispose()
    if ($null -ne $script:ImportedExchangeModule) {
        Remove-Module -ModuleInfo $script:ImportedExchangeModule -Force -ErrorAction SilentlyContinue
    }
    if ($null -ne $script:ExchangeSession) {
        Remove-PSSession -Session $script:ExchangeSession -ErrorAction SilentlyContinue
    }
})
#endregion Events

try {
    [void]$script:MainForm.ShowDialog()
}
finally {
    if ($null -ne $script:ImportedExchangeModule) {
        Remove-Module -ModuleInfo $script:ImportedExchangeModule -Force -ErrorAction SilentlyContinue
    }
    if ($null -ne $script:ExchangeSession) {
        Remove-PSSession -Session $script:ExchangeSession -ErrorAction SilentlyContinue
    }
    $script:MainForm.Dispose()
}
