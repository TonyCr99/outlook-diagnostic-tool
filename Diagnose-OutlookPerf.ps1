#requires -Version 5.1
<#
.SYNOPSIS
    OUTLOOK DIAGNOSTIC TOOL  ::  Herramienta de diagnostico de lentitud/sincronizacion de Outlook (M365).

.DESCRIPTION
    CLI con estetica ciberpunk para diagnosticar los problemas mas comunes reportados por usuarios:
      - "No entran / no salen correos"        -> traza de flujo de correo (message trace)
      - "No se sincroniza"                     -> cuotas, dispositivos moviles, conectividad
      - "Outlook de escritorio tarda en cargar"-> tamano OST, add-ins, modo cache, sistema
      - Bucles / cuelgues                       -> carpetas con demasiados items (>100K) [caso conocido]

    Combina en una sola app:
      * Diagnostico REMOTO del buzon via Exchange Online PowerShell.
      * Diagnostico LOCAL del cliente Outlook (registro, OST, add-ins, red).

    Cada comprobacion registra "hallazgos" con severidad (OK / INFO / ALERTA / CRITICO),
    detalle y recomendacion. Al final puedes exportar un informe HTML con el mismo estilo.

    Disponible en espanol, ingles y portugues (-Language, o desde el menu).

.PARAMETER Mailbox
    Buzon objetivo para las comprobaciones remotas (ej. usuario@dominio.com).

.PARAMETER LocalOnly
    Ejecuta solo el diagnostico del cliente local (no se conecta a Exchange Online).

.PARAMETER RemoteOnly
    Ejecuta solo el diagnostico remoto del buzon.

.PARAMETER Auto
    Ejecuta el diagnostico completo sin menu (modo desatendido) y sale.

.PARAMETER ReportPath
    Ruta del informe HTML a generar al terminar (con -Auto).

.PARAMETER Language
    Idioma de la interfaz: es (default), en, pt. Si se omite en modo interactivo, se pregunta al inicio.

.EXAMPLE
    .\Diagnose-OutlookPerf.ps1
    Abre el selector de idioma y luego el menu interactivo.

.EXAMPLE
    .\Diagnose-OutlookPerf.ps1 -Mailbox juan@clx.com -Language en -Auto -ReportPath .\report.html
    Diagnostico completo desatendido en ingles + informe HTML.

.NOTES
    Autor : IT Ops  |  Requiere: PowerShell 5.1+, modulo ExchangeOnlineManagement (para remoto).
#>
[CmdletBinding()]
param(
    [string]$Mailbox,
    [switch]$LocalOnly,
    [switch]$RemoteOnly,
    [switch]$Auto,
    [string]$ReportPath,
    [ValidateSet('es','en','pt')]
    [string]$Language
)

# =====================================================================================
#  CONFIGURACION  ::  Umbrales (ajustables)
# =====================================================================================
$Script:Thresholds = @{
    FolderItemsCritical = 100000   # Microsoft recomienda <100K items por carpeta -> el caso del bucle
    FolderItemsWarning  = 50000
    MailboxQuotaWarnPct = 85       # % de cuota usado -> puede dejar de recibir
    MailboxQuotaCritPct = 95
    OstSizeWarnGB       = 40
    OstSizeCritGB       = 50       # >50GB rendimiento degradado / corrupcion probable
    InboxRulesWarn      = 40       # demasiadas reglas ralentizan el buzon
    AddinsWarn          = 8        # muchos add-ins activos -> arranque lento
    DiskFreeWarnGB      = 10
    DiskFreeCritGB      = 3
    RamWarnGB           = 4        # RAM total baja para cliente pesado
    TraceHours          = 48       # ventana de traza de correo
    AutoCompleteWarnMB  = 5        # cache de autocompletado grande
}

# =====================================================================================
#  COMPATIBILIDAD  ::  PowerShell 5.1 (Windows PowerShell) no tiene $IsWindows
# =====================================================================================
$Script:OnWindows = if (Test-Path variable:IsWindows) { $IsWindows } else { $true }

# =====================================================================================
#  IDIOMA  ::  es (default) / en / pt
# =====================================================================================
$Script:Lang = if ($Language) { $Language } else { 'es' }

# =====================================================================================
#  MOTOR VISUAL  ::  Paleta neon / ANSI truecolor
# =====================================================================================
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch {}
$Script:E = [char]27
$Script:Palette = @{
    Cyan   = '0;255;255'
    Pink   = '255;0;153'
    Green  = '57;255;20'
    Yellow = '255;214;0'
    Red    = '255;45;85'
    Purple = '157;78;221'
    Blue   = '0;170;255'
    Gray   = '120;120;140'
    White  = '235;235;245'
    Orange = '255;140;0'
    Dark   = '80;80;100'
}

function Ink {
    param([string]$Text, [string]$Color = 'White', [switch]$Bold)
    $rgb = $Script:Palette[$Color]
    if (-not $rgb) { $rgb = $Script:Palette['White'] }
    $b = if ($Bold) { '1;' } else { '' }
    "$($Script:E)[${b}38;2;${rgb}m$Text$($Script:E)[0m"
}

function Write-Rule {
    param([string]$Color = 'Purple', [int]$Width = 78)
    Write-Host (Ink ('═' * $Width) $Color)
}

function Write-Section {
    param([string]$Title, [string]$Color = 'Cyan')
    Write-Host ''
    Write-Host (Ink "╺━━ " $Color) -NoNewline
    Write-Host (Ink " $Title " $Color -Bold) -NoNewline
    Write-Host (Ink (' ' + ('━' * [Math]::Max(4, 68 - $Title.Length))) $Color)
}

# =====================================================================================
#  I18N  ::  T (traduce) + selector de idioma
# =====================================================================================
function T {
    param([string]$Key, [object[]]$FormatArgs = @())
    $table = $null
    if ($Script:I18N.ContainsKey($Script:Lang)) { $table = $Script:I18N[$Script:Lang] }
    $text = $null
    if ($table -and $table.ContainsKey($Key)) { $text = $table[$Key] }
    if (-not $text -and $Script:I18N['es'].ContainsKey($Key)) { $text = $Script:I18N['es'][$Key] }
    if (-not $text) { return $Key }
    if ($FormatArgs.Count -gt 0) { return ($text -f $FormatArgs) }
    return $text
}

function Read-YesNo {
    param([string]$Key, [object[]]$FormatArgs = @())
    $suffix = if ($Script:Lang -eq 'en') { ' [y/N] ' } else { ' [s/N] ' }
    Write-Host (Ink "   $(T $Key $FormatArgs)$suffix" 'Cyan') -NoNewline
    $ans = Read-Host
    if ($Script:Lang -eq 'en') { return $ans -match '^[yY]' }
    return $ans -match '^[sS]'
}

function Show-LanguagePicker {
    Write-Host ''
    Write-Host (Ink '  ┌──────────────────────────────────────────────────┐' 'Purple')
    Write-Host (Ink '  │  ' 'Purple') -NoNewline
    Write-Host (Ink '[1] Espanol' 'Cyan') -NoNewline
    Write-Host (Ink '     ' 'Purple') -NoNewline
    Write-Host (Ink '[2] English' 'Cyan') -NoNewline
    Write-Host (Ink '     ' 'Purple') -NoNewline
    Write-Host (Ink '[3] Portugues' 'Cyan') -NoNewline
    Write-Host (Ink '  │' 'Purple')
    Write-Host (Ink '  └──────────────────────────────────────────────────┘' 'Purple')
    Write-Host (Ink '  └─▶ ' 'Purple') -NoNewline
    $sel = Read-Host
    $Script:Lang = switch ($sel) { '2' { 'en' } '3' { 'pt' } default { 'es' } }
}

function Show-Banner {
    try { Clear-Host } catch { Write-Host "$($Script:E)[2J$($Script:E)[H" -NoNewline }
    $letters = @{
        'O' = @(' ███ ','█   █','█   █','█   █',' ███ ')
        'U' = @('█   █','█   █','█   █','█   █',' ███ ')
        'T' = @('█████','  █  ','  █  ','  █  ','  █  ')
        'L' = @('█    ','█    ','█    ','█    ','█████')
        'K' = @('█   █','█  █ ','███  ','█  █ ','█   █')
    }
    $word = 'OUTLOOK'
    $colors = @('Cyan','Blue','Purple','Pink','Purple','Blue','Cyan')
    Write-Host ''
    for ($r = 0; $r -lt 5; $r++) {
        $line = ''
        for ($i = 0; $i -lt $word.Length; $i++) {
            $glyph = $letters[$word[$i].ToString()][$r]
            $line += (Ink ($glyph + ' ') $colors[$i] -Bold)
        }
        Write-Host "  $line"
    }
    Write-Host ''
    Write-Host (Ink "        $(T 'banner.tagline')" 'Green' -Bold)
    Write-Host (Ink "        $(T 'banner.subtitle')" 'Cyan')
    Write-Host (Ink "        $(T 'banner.footer')" 'Gray')
    Write-Host ''
}

# =====================================================================================
#  DICCIONARIO DE TRADUCCIONES
# =====================================================================================
$Script:I18N = @{
    es = @{
        'banner.tagline' = 'HERRAMIENTA DE DIAGNOSTICO'
        'banner.subtitle' = 'Diagnostico de Rendimiento y Sincronizacion'
        'banner.footer' = 'M365 / Exchange Online  ·  cliente local  ·  v1.2'
        'menu.opt1' = 'Diagnostico COMPLETO (buzon + cliente local)'
        'menu.opt2' = 'Solo BUZON (Exchange Online)'
        'menu.opt3' = 'Solo CLIENTE local'
        'menu.opt4' = 'Conectar / verificar Exchange Online'
        'menu.opt5' = 'Ver RESUMEN de hallazgos'
        'menu.opt6' = 'Exportar INFORME HTML'
        'menu.opt7' = 'Cambiar buzon objetivo'
        'menu.opt8' = 'REMEDIAR (archivo / mover / borrar / reparar)'
        'menu.opt9' = 'Cambiar idioma / Language'
        'menu.opt0' = 'Salir'
        'menu.mailboxLabel' = 'buzon:'
        'menu.undefined' = '(sin definir)'
        'menu.connected' = 'CONECTADO'
        'menu.disconnected' = 'desconectado'
        'prompt.selectOption' = 'selecciona opcion'
        'prompt.mailboxToDiagnose' = 'buzon a diagnosticar (UPN/SMTP)'
        'prompt.mailbox' = 'buzon (UPN/SMTP)'
        'prompt.newMailbox' = 'nuevo buzon objetivo (UPN/SMTP)'
        'prompt.remediationAction' = 'accion de remediacion'
        'prompt.cutoffDate' = 'borrar items RECIBIDOS ANTES de (YYYY-MM-DD)'
        'prompt.maxBatches' = 'maximo de lotes a procesar (cada lote ~10 items/buzon; ENTER = 20)'
        'section.remoteDiag' = 'DIAGNOSTICO REMOTO DEL BUZON  ::  {0}'
        'section.localDiag' = 'DIAGNOSTICO LOCAL DEL CLIENTE  ::  {0}'
        'section.summary' = 'RESUMEN DE HALLAZGOS'
        'section.remediation' = 'REMEDIACION  ::  buzon: {0}'
        'section.remediate1' = 'REMEDIAR :: Activar archivo en linea  ::  {0}'
        'section.remediate2' = 'REMEDIAR :: Mover items antiguos al archivo  ::  {0}'
        'section.remediate3' = 'REMEDIAR :: Borrado guiado por antiguedad (soft-delete)  ::  {0}'
        'section.remediate4' = 'REMEDIAR :: Runbook de reparacion del cliente Outlook'
        'exo.alreadyConnected' = 'Ya existe una sesion de Exchange Online activa.'
        'exo.moduleNotInstalled' = 'Modulo ExchangeOnlineManagement no instalado.'
        'exo.connecting' = 'Conectando a Exchange Online (se abrira el login)...'
        'exo.connected' = 'Conexion establecida.'
        'exo.connectError' = 'Error al conectar: {0}'
    }
    en = @{
        'banner.tagline' = 'DIAGNOSTIC TOOL'
        'banner.subtitle' = 'Performance & Sync Diagnostics'
        'banner.footer' = 'M365 / Exchange Online  ·  local client  ·  v1.2'
        'menu.opt1' = 'FULL diagnostic (mailbox + local client)'
        'menu.opt2' = 'MAILBOX only (Exchange Online)'
        'menu.opt3' = 'Local CLIENT only'
        'menu.opt4' = 'Connect / verify Exchange Online'
        'menu.opt5' = 'View findings SUMMARY'
        'menu.opt6' = 'Export HTML REPORT'
        'menu.opt7' = 'Change target mailbox'
        'menu.opt8' = 'REMEDIATE (archive / move / delete / repair)'
        'menu.opt9' = 'Change language'
        'menu.opt0' = 'Exit'
        'menu.mailboxLabel' = 'mailbox:'
        'menu.undefined' = '(not set)'
        'menu.connected' = 'CONNECTED'
        'menu.disconnected' = 'disconnected'
        'prompt.selectOption' = 'select option'
        'prompt.mailboxToDiagnose' = 'mailbox to diagnose (UPN/SMTP)'
        'prompt.mailbox' = 'mailbox (UPN/SMTP)'
        'prompt.newMailbox' = 'new target mailbox (UPN/SMTP)'
        'prompt.remediationAction' = 'remediation action'
        'prompt.cutoffDate' = 'delete items RECEIVED BEFORE (YYYY-MM-DD)'
        'prompt.maxBatches' = 'maximum batches to process (each batch ~10 items/mailbox; ENTER = 20)'
        'section.remoteDiag' = 'REMOTE MAILBOX DIAGNOSTIC  ::  {0}'
        'section.localDiag' = 'LOCAL CLIENT DIAGNOSTIC  ::  {0}'
        'section.summary' = 'FINDINGS SUMMARY'
        'section.remediation' = 'REMEDIATION  ::  mailbox: {0}'
        'section.remediate1' = 'REMEDIATE :: Enable online archive  ::  {0}'
        'section.remediate2' = 'REMEDIATE :: Move old items to archive  ::  {0}'
        'section.remediate3' = 'REMEDIATE :: Age-guided deletion (soft-delete)  ::  {0}'
        'section.remediate4' = 'REMEDIATE :: Outlook client repair runbook'
        'exo.alreadyConnected' = 'An active Exchange Online session already exists.'
        'exo.moduleNotInstalled' = 'ExchangeOnlineManagement module not installed.'
        'exo.connecting' = 'Connecting to Exchange Online (sign-in will open)...'
        'exo.connected' = 'Connection established.'
        'exo.connectError' = 'Connection error: {0}'
    }
    pt = @{
        'banner.tagline' = 'FERRAMENTA DE DIAGNOSTICO'
        'banner.subtitle' = 'Diagnostico de Desempenho e Sincronizacao'
        'banner.footer' = 'M365 / Exchange Online  ·  cliente local  ·  v1.2'
        'menu.opt1' = 'Diagnostico COMPLETO (caixa + cliente local)'
        'menu.opt2' = 'Somente CAIXA (Exchange Online)'
        'menu.opt3' = 'Somente CLIENTE local'
        'menu.opt4' = 'Conectar / verificar Exchange Online'
        'menu.opt5' = 'Ver RESUMO de descobertas'
        'menu.opt6' = 'Exportar RELATORIO HTML'
        'menu.opt7' = 'Alterar caixa alvo'
        'menu.opt8' = 'REMEDIAR (arquivar / mover / excluir / reparar)'
        'menu.opt9' = 'Alterar idioma'
        'menu.opt0' = 'Sair'
        'menu.mailboxLabel' = 'caixa:'
        'menu.undefined' = '(nao definida)'
        'menu.connected' = 'CONECTADO'
        'menu.disconnected' = 'desconectado'
        'prompt.selectOption' = 'selecione opcao'
        'prompt.mailboxToDiagnose' = 'caixa a diagnosticar (UPN/SMTP)'
        'prompt.mailbox' = 'caixa (UPN/SMTP)'
        'prompt.newMailbox' = 'nova caixa alvo (UPN/SMTP)'
        'prompt.remediationAction' = 'acao de remediacao'
        'prompt.cutoffDate' = 'excluir itens RECEBIDOS ANTES de (YYYY-MM-DD)'
        'prompt.maxBatches' = 'maximo de lotes a processar (cada lote ~10 itens/caixa; ENTER = 20)'
        'section.remoteDiag' = 'DIAGNOSTICO REMOTO DA CAIXA  ::  {0}'
        'section.localDiag' = 'DIAGNOSTICO LOCAL DO CLIENTE  ::  {0}'
        'section.summary' = 'RESUMO DE DESCOBERTAS'
        'section.remediation' = 'REMEDIACAO  ::  caixa: {0}'
        'section.remediate1' = 'REMEDIAR :: Ativar arquivo online  ::  {0}'
        'section.remediate2' = 'REMEDIAR :: Mover itens antigos para o arquivo  ::  {0}'
        'section.remediate3' = 'REMEDIAR :: Exclusao guiada por antiguidade (soft-delete)  ::  {0}'
        'section.remediate4' = 'REMEDIAR :: Runbook de reparo do cliente Outlook'
        'exo.alreadyConnected' = 'Ja existe uma sessao ativa do Exchange Online.'
        'exo.moduleNotInstalled' = 'Modulo ExchangeOnlineManagement nao instalado.'
        'exo.connecting' = 'Conectando ao Exchange Online (o login sera aberto)...'
        'exo.connected' = 'Conexao estabelecida.'
        'exo.connectError' = 'Erro ao conectar: {0}'
    }
}

$Script:I18N.es += @{
    'cat.quota' = 'Cuota'; 'cat.folders' = 'Carpetas'; 'cat.rules' = 'Reglas'; 'cat.flow' = 'Flujo'
    'cat.sync' = 'Sync'; 'cat.mailbox' = 'Buzon'; 'cat.remote' = 'Remoto'
    'remote.noConnectionTitle' = 'No hay conexion con Exchange Online'
    'remote.noConnectionDetail' = 'Las comprobaciones remotas no se pueden ejecutar.'
    'remote.noConnectionRec' = 'Conecta con la opcion [4] del menu.'
    'remote.notFoundTitle' = 'Buzon no encontrado: {0}'
    'remote.notFoundRec' = 'Verifica la direccion (UPN/SMTP) o si es un buzon compartido/on-prem.'
    'quota.detail' = "Tamano total : {0}`nCuota bloqueo: {1}`nUso          : {2}%`nItems totales: {3}`nUltimo logon : {4}"
    'quota.unlimited' = 'Ilimitada/no definida'
    'quota.critTitle' = 'Buzon al {0}% de su cuota'
    'quota.critRec' = 'Cerca del limite: puede DEJAR DE RECIBIR correo. Archiva, vacia Elementos eliminados o amplia cuota.'
    'quota.warnRec' = 'Uso alto de cuota. Habilita/usa el archivo en linea o depura carpetas grandes.'
    'quota.okTitle' = 'Uso de cuota normal ({0}%)'
    'quota.errorTitle' = 'No se pudieron obtener estadisticas del buzon'
    'folders.analyzing' = 'Analizando estadisticas de carpetas (puede tardar)...'
    'folders.topTitle' = 'Top carpetas por numero de items'
    'folders.critTitle' = 'Carpeta con {0} items: {1}'
    'folders.critDetail' = 'Supera el limite recomendado de {0} items por carpeta.'
    'folders.critRec' = 'Causa clasica de bucles/cuelgues y sincronizacion lenta (Outlook Mac/cache). Mueve items a subcarpetas o archivo, o vacia la carpeta.'
    'folders.warnTitle' = 'Carpeta grande ({0} items): {1}'
    'folders.warnDetail' = 'Se acerca al limite recomendado de {0} items.'
    'folders.warnRec' = 'Reparte en subcarpetas antes de que degrade el rendimiento de sincronizacion.'
    'folders.okTitle' = 'Ninguna carpeta supera el umbral de items'
    'folders.okDetail' = 'Maximo detectado: {0} items.'
    'folders.errorTitle' = 'No se pudo obtener estadisticas de carpetas'
    'rules.detail' = "Reglas totales: {0}`nReglas de reenvio/redireccion: {1}"
    'rules.warnTitle' = 'Muchas reglas de bandeja ({0})'
    'rules.warnRec' = 'Un exceso de reglas (o reglas pesadas) ralentiza el procesamiento y puede provocar que "no lleguen" correos que se mueven/borran solos. Revisa y consolida.'
    'rules.okTitle' = 'Reglas de bandeja: {0}'
    'rules.fwdTitle' = 'Existen reglas de reenvio/redireccion'
    'rules.fwdRec' = 'Si el usuario dice que "faltan" correos, una regla podria estar moviendolos o reenviandolos.'
    'rules.errorTitle' = 'No se pudieron leer las reglas de bandeja'
    'fwd.title' = 'Reenvio configurado a nivel de buzon'
    'fwd.detail' = "ForwardingSmtpAddress: {0}`nForwardingAddress: {1}`nEntregar y reenviar: {2}"
    'fwd.rec' = 'Si el correo "no aparece", puede estar reenviandose fuera. Confirma que es intencionado.'
    'devices.warnTitle' = '{0} dispositivo(s) con estado de sync anomalo'
    'devices.warnRec' = 'Un dispositivo en bucle de sync puede saturar el buzon. Considera quitar/re-emparejar dispositivos obsoletos.'
    'devices.okTitle' = '{0} dispositivo(s) movil(es), sync correcto'
    'devices.noneTitle' = 'Sin dispositivos moviles (EAS) registrados'
    'devices.errorTitle' = 'No se pudo consultar dispositivos moviles'
    'trace.tracing' = 'Trazando flujo de correo ultimas {0}h...'
    'trace.detail' = "ENTRANTES ({0}h): {1}  ·  fallidos/spam: {2}`nSALIENTES: {3}  ·  fallidos: {4}"
    'trace.noInboundTitle' = 'Sin correo ENTRANTE en la ventana analizada'
    'trace.noInboundRec' = 'Confirma que es normal para este usuario. Si esperaba correo, revisa reglas, cuarentena y conectores.'
    'trace.inFailTitle' = '{0} correos entrantes fallidos/en spam'
    'trace.inFailRec' = 'Revisa cuarentena (Get-QuarantineMessage) y politicas anti-spam.'
    'trace.okTitle' = 'Flujo de correo entrante normal'
    'trace.outFailTitle' = '{0} correos salientes fallidos'
    'trace.outFailRec' = 'Revisa NDRs, limites de envio y reputacion del dominio.'
    'trace.unavailableTitle' = 'Cmdlet de traza no disponible en esta sesion'
    'trace.unavailableDetail' = 'Se necesitan permisos de "Message Tracking" en Exchange Online.'
    'trace.errorTitle' = 'No se pudo trazar el flujo de correo'
    'archive.enabled' = 'Habilitado'
    'archive.disabled' = 'Deshabilitado'
    'archive.title' = 'Estado de archivo y retencion'
    'archive.detail' = "Archivo en linea: {0}`nLitigation hold : {1}`nTipo de buzon   : {2}"
}
$Script:I18N.en += @{
    'cat.quota' = 'Quota'; 'cat.folders' = 'Folders'; 'cat.rules' = 'Rules'; 'cat.flow' = 'Flow'
    'cat.sync' = 'Sync'; 'cat.mailbox' = 'Mailbox'; 'cat.remote' = 'Remote'
    'remote.noConnectionTitle' = 'No connection to Exchange Online'
    'remote.noConnectionDetail' = 'Remote checks cannot run.'
    'remote.noConnectionRec' = 'Connect using menu option [4].'
    'remote.notFoundTitle' = 'Mailbox not found: {0}'
    'remote.notFoundRec' = 'Check the address (UPN/SMTP) or whether it is a shared/on-prem mailbox.'
    'quota.detail' = "Total size    : {0}`nBlock quota   : {1}`nUsage         : {2}%`nTotal items   : {3}`nLast logon    : {4}"
    'quota.unlimited' = 'Unlimited/not defined'
    'quota.critTitle' = 'Mailbox at {0}% of quota'
    'quota.critRec' = 'Near the limit: mail may STOP BEING RECEIVED. Archive, empty Deleted Items, or increase the quota.'
    'quota.warnRec' = 'High quota usage. Enable/use the online archive or clean up large folders.'
    'quota.okTitle' = 'Normal quota usage ({0}%)'
    'quota.errorTitle' = 'Could not retrieve mailbox statistics'
    'folders.analyzing' = 'Analyzing folder statistics (may take a while)...'
    'folders.topTitle' = 'Top folders by item count'
    'folders.critTitle' = 'Folder with {0} items: {1}'
    'folders.critDetail' = 'Exceeds the recommended limit of {0} items per folder.'
    'folders.critRec' = 'Classic cause of loops/hangs and slow sync (Outlook Mac/cache mode). Move items to subfolders or the archive, or empty the folder.'
    'folders.warnTitle' = 'Large folder ({0} items): {1}'
    'folders.warnDetail' = 'Approaching the recommended limit of {0} items.'
    'folders.warnRec' = 'Split into subfolders before it degrades sync performance.'
    'folders.okTitle' = 'No folder exceeds the item threshold'
    'folders.okDetail' = 'Maximum detected: {0} items.'
    'folders.errorTitle' = 'Could not retrieve folder statistics'
    'rules.detail' = "Total rules: {0}`nForward/redirect rules: {1}"
    'rules.warnTitle' = 'Many inbox rules ({0})'
    'rules.warnRec' = 'Too many (or heavy) rules slow down processing and can make mail seem to "not arrive" if it is being auto-moved/deleted. Review and consolidate.'
    'rules.okTitle' = 'Inbox rules: {0}'
    'rules.fwdTitle' = 'Forward/redirect rules exist'
    'rules.fwdRec' = 'If the user says mail is "missing", a rule may be moving or forwarding it.'
    'rules.errorTitle' = 'Could not read inbox rules'
    'fwd.title' = 'Mailbox-level forwarding configured'
    'fwd.detail' = "ForwardingSmtpAddress: {0}`nForwardingAddress: {1}`nDeliver and forward: {2}"
    'fwd.rec' = 'If mail "does not appear", it may be forwarded elsewhere. Confirm it is intentional.'
    'devices.warnTitle' = '{0} device(s) with abnormal sync status'
    'devices.warnRec' = 'A device stuck in a sync loop can overload the mailbox. Consider removing/re-pairing stale devices.'
    'devices.okTitle' = '{0} mobile device(s), sync OK'
    'devices.noneTitle' = 'No mobile devices (EAS) registered'
    'devices.errorTitle' = 'Could not query mobile devices'
    'trace.tracing' = 'Tracing mail flow for the last {0}h...'
    'trace.detail' = "INBOUND ({0}h): {1}  ·  failed/spam: {2}`nOUTBOUND: {3}  ·  failed: {4}"
    'trace.noInboundTitle' = 'No INBOUND mail in the analyzed window'
    'trace.noInboundRec' = 'Confirm this is normal for the user. If mail was expected, check rules, quarantine, and connectors.'
    'trace.inFailTitle' = '{0} failed/spam inbound messages'
    'trace.inFailRec' = 'Check quarantine (Get-QuarantineMessage) and anti-spam policies.'
    'trace.okTitle' = 'Normal inbound mail flow'
    'trace.outFailTitle' = '{0} failed outbound messages'
    'trace.outFailRec' = 'Check NDRs, sending limits, and domain reputation.'
    'trace.unavailableTitle' = 'Trace cmdlet not available in this session'
    'trace.unavailableDetail' = '"Message Tracking" permissions are needed in Exchange Online.'
    'trace.errorTitle' = 'Could not trace mail flow'
    'archive.enabled' = 'Enabled'
    'archive.disabled' = 'Disabled'
    'archive.title' = 'Archive and retention status'
    'archive.detail' = "Online archive : {0}`nLitigation hold: {1}`nMailbox type   : {2}"
}
$Script:I18N.pt += @{
    'cat.quota' = 'Cota'; 'cat.folders' = 'Pastas'; 'cat.rules' = 'Regras'; 'cat.flow' = 'Fluxo'
    'cat.sync' = 'Sync'; 'cat.mailbox' = 'Caixa'; 'cat.remote' = 'Remoto'
    'remote.noConnectionTitle' = 'Sem conexao com o Exchange Online'
    'remote.noConnectionDetail' = 'As verificacoes remotas nao podem ser executadas.'
    'remote.noConnectionRec' = 'Conecte-se usando a opcao [4] do menu.'
    'remote.notFoundTitle' = 'Caixa nao encontrada: {0}'
    'remote.notFoundRec' = 'Verifique o endereco (UPN/SMTP) ou se e uma caixa compartilhada/on-prem.'
    'quota.detail' = "Tamanho total : {0}`nCota bloqueio : {1}`nUso           : {2}%`nItens totais  : {3}`nUltimo logon  : {4}"
    'quota.unlimited' = 'Ilimitada/nao definida'
    'quota.critTitle' = 'Caixa em {0}% da cota'
    'quota.critRec' = 'Perto do limite: pode PARAR DE RECEBER e-mails. Arquive, esvazie Itens Excluidos ou aumente a cota.'
    'quota.warnRec' = 'Uso alto de cota. Habilite/use o arquivo online ou limpe pastas grandes.'
    'quota.okTitle' = 'Uso de cota normal ({0}%)'
    'quota.errorTitle' = 'Nao foi possivel obter estatisticas da caixa'
    'folders.analyzing' = 'Analisando estatisticas de pastas (pode demorar)...'
    'folders.topTitle' = 'Principais pastas por numero de itens'
    'folders.critTitle' = 'Pasta com {0} itens: {1}'
    'folders.critDetail' = 'Excede o limite recomendado de {0} itens por pasta.'
    'folders.critRec' = 'Causa classica de loops/travamentos e sincronizacao lenta (Outlook Mac/cache). Mova itens para subpastas ou arquivo, ou esvazie a pasta.'
    'folders.warnTitle' = 'Pasta grande ({0} itens): {1}'
    'folders.warnDetail' = 'Aproximando-se do limite recomendado de {0} itens.'
    'folders.warnRec' = 'Distribua em subpastas antes que o desempenho de sincronizacao seja afetado.'
    'folders.okTitle' = 'Nenhuma pasta excede o limite de itens'
    'folders.okDetail' = 'Maximo detectado: {0} itens.'
    'folders.errorTitle' = 'Nao foi possivel obter estatisticas de pastas'
    'rules.detail' = "Regras totais: {0}`nRegras de encaminhamento/redirecionamento: {1}"
    'rules.warnTitle' = 'Muitas regras de caixa de entrada ({0})'
    'rules.warnRec' = 'Excesso de regras (ou regras pesadas) torna o processamento lento e pode fazer parecer que e-mails "nao chegam" quando estao sendo movidos/excluidos automaticamente. Revise e consolide.'
    'rules.okTitle' = 'Regras de caixa de entrada: {0}'
    'rules.fwdTitle' = 'Existem regras de encaminhamento/redirecionamento'
    'rules.fwdRec' = 'Se o usuario disser que e-mails "estao faltando", uma regra pode estar movendo-os ou encaminhando-os.'
    'rules.errorTitle' = 'Nao foi possivel ler as regras da caixa de entrada'
    'fwd.title' = 'Encaminhamento configurado no nivel da caixa'
    'fwd.detail' = "ForwardingSmtpAddress: {0}`nForwardingAddress: {1}`nEntregar e encaminhar: {2}"
    'fwd.rec' = 'Se o e-mail "nao aparece", pode estar sendo encaminhado para fora. Confirme se e intencional.'
    'devices.warnTitle' = '{0} dispositivo(s) com status de sincronizacao anormal'
    'devices.warnRec' = 'Um dispositivo em loop de sincronizacao pode sobrecarregar a caixa. Considere remover/re-parear dispositivos obsoletos.'
    'devices.okTitle' = '{0} dispositivo(s) movel(is), sincronizacao OK'
    'devices.noneTitle' = 'Nenhum dispositivo movel (EAS) registrado'
    'devices.errorTitle' = 'Nao foi possivel consultar dispositivos moveis'
    'trace.tracing' = 'Rastreando fluxo de e-mail das ultimas {0}h...'
    'trace.detail' = "ENTRADA ({0}h): {1}  ·  falhas/spam: {2}`nSAIDA: {3}  ·  falhas: {4}"
    'trace.noInboundTitle' = 'Sem e-mail de ENTRADA na janela analisada'
    'trace.noInboundRec' = 'Confirme se isso e normal para este usuario. Se esperava e-mail, verifique regras, quarentena e conectores.'
    'trace.inFailTitle' = '{0} e-mails de entrada com falha/spam'
    'trace.inFailRec' = 'Verifique a quarentena (Get-QuarantineMessage) e as politicas anti-spam.'
    'trace.okTitle' = 'Fluxo de e-mail de entrada normal'
    'trace.outFailTitle' = '{0} e-mails de saida com falha'
    'trace.outFailRec' = 'Verifique NDRs, limites de envio e reputacao do dominio.'
    'trace.unavailableTitle' = 'Cmdlet de rastreamento nao disponivel nesta sessao'
    'trace.unavailableDetail' = 'Sao necessarias permissoes de "Message Tracking" no Exchange Online.'
    'trace.errorTitle' = 'Nao foi possivel rastrear o fluxo de e-mail'
    'archive.enabled' = 'Habilitado'
    'archive.disabled' = 'Desabilitado'
    'archive.title' = 'Status de arquivo e retencao'
    'archive.detail' = "Arquivo online  : {0}`nLitigation hold: {1}`nTipo de caixa   : {2}"
}

$Script:I18N.es += @{
    'cat.system' = 'Sistema'; 'cat.outlook' = 'Outlook'; 'cat.ost' = 'OST'; 'cat.pst' = 'PST'
    'cat.cache' = 'Cache'; 'cat.addins' = 'Add-ins'; 'cat.network' = 'Red'
    'local.onlyWindowsTitle' = 'Diagnostico local disponible solo en Windows'
    'local.onlyWindowsRec' = 'Ejecuta la parte local en el equipo del usuario afectado (Windows).'
    'sys.detail' = "SO         : {0} (build {1})`nRAM total  : {2} GB  (libre {3} GB)`nDisco {4} : {5} GB libres"
    'sys.ramWarnTitle' = 'RAM baja ({0} GB)'
    'sys.ramWarnRec' = 'Outlook en modo cache con OST grande necesita memoria; considera ampliar RAM.'
    'sys.okTitle' = 'Sistema: {0} GB RAM'
    'sys.diskCritTitle' = 'Disco casi lleno ({0} GB libres)'
    'sys.diskCritRec' = 'Poco espacio impide que el OST crezca/repare -> cuelgues y errores de sincronizacion. Libera espacio.'
    'sys.diskWarnTitle' = 'Poco espacio en disco ({0} GB libres)'
    'sys.diskWarnRec' = 'El OST necesita margen para crecer. Libera espacio.'
    'sys.errorTitle' = 'No se pudo leer info del sistema'
    'outlook.unknown' = 'Desconocido'
    'outlook.detail' = "Version   : {0}`nPlataforma: {1}`nCanal     : {2}`nProductos : {3}"
    'outlook.title' = 'Outlook / Microsoft 365 Apps  ({0})'
    'outlook.legacyTitle' = 'Outlook detectado (MSI o version legacy)'
    'outlook.legacyDetail' = 'Rama de Office: {0}'
    'outlook.errorTitle' = 'No se pudo determinar la version de Outlook'
    'proc.detail' = "Procesos outlook.exe: {0}`nMemoria (working set): {1} MB`nSin responder: {2}"
    'proc.multiTitle' = 'Multiples procesos de Outlook ({0})'
    'proc.multiRec' = 'Instancias duplicadas/zombis pueden bloquear el OST. Cierra Outlook por completo (o mata outlook.exe) y reabre.'
    'proc.notRespondingTitle' = 'Outlook no responde actualmente'
    'proc.notRespondingRec' = 'Coincide con el sintoma de "cuelgue/bucle". Revisa carpetas grandes y add-ins.'
    'proc.okTitle' = 'Outlook en ejecucion ({0} MB)'
    'proc.notRunningTitle' = 'Outlook no esta en ejecucion'
    'proc.notRunningDetail' = 'Algunas comprobaciones (OST en uso) reflejan el ultimo estado en disco.'
    'ost.noneTitle' = 'No se encontraron archivos OST en la ruta por defecto'
    'ost.noneDetail' = "Ruta: {0}`n(Perfil online o ruta personalizada.)"
    'ost.detail' = "Archivo: {0}`nTamano : {1} GB`nModif. : {2}"
    'ost.critTitle' = 'OST muy grande: {0} GB ({1})'
    'ost.critRec' = 'Por encima de ~50GB el rendimiento cae y aumenta el riesgo de corrupcion -> lentitud de carga y bucles. Reduce la ventana de sincronizacion (Cached Mode slider) o activa archivo en linea.'
    'ost.warnTitle' = 'OST grande: {0} GB ({1})'
    'ost.warnRec' = 'Reduce el "Mail to keep offline" a 3-6 meses para acelerar carga/sincronizacion.'
    'ost.okTitle' = 'OST dentro de lo normal: {0} GB'
    'ost.errorTitle' = 'No se pudo analizar archivos de datos'
    'pst.title' = '{0} archivo(s) PST local(es)'
    'pst.rec' = 'Los PST montados aumentan el tiempo de carga y son fragiles en red.'
    'cache.warnTitle' = 'Cache de autocompletado grande ({0} MB)'
    'cache.warnDetail' = 'Ruta: {0}'
    'cache.warnRec' = 'Un stream de autocompletado corrupto/enorme puede ralentizar el arranque. Renombra la carpeta RoamCache con Outlook cerrado para regenerarla.'
    'cache.okTitle' = 'Cache de autocompletado normal ({0} MB)'
    'addins.warnTitle' = '{0} complementos activos al inicio'
    'addins.warnRec' = 'Muchos add-ins retrasan el arranque de Outlook. Deshabilita los no esenciales (COM Add-ins) y prueba con "outlook.exe /safe".'
    'addins.okTitle' = '{0} complementos activos'
    'addins.noneTitle' = 'Sin complementos de terceros activos'
    'addins.errorTitle' = 'No se pudieron enumerar los complementos'
    'net.testing' = 'Probando conectividad a endpoints de Microsoft 365...'
    'net.desc.exo' = 'Exchange Online (correo)'
    'net.desc.owa' = 'OWA / servicio'
    'net.desc.autodiscover' = 'Autodiscover'
    'net.desc.auth' = 'Autenticacion (Entra ID)'
    'net.critTitle' = 'Sin conectividad: {0}:{1}'
    'net.critDetail' = '{0} — no responde ({1} ms timeout).'
    'net.critRec' = 'Bloqueo probable de firewall/proxy/VPN. Sin este endpoint no hay sincronizacion.'
    'net.warnTitle' = 'Latencia alta a {0} ({1} ms)'
    'net.warnRec' = 'Latencia alta ralentiza la sincronizacion. Revisa proxy/VPN/red.'
    'net.okTitle' = '{0} accesible ({1} ms)'
    'net.allOkTitle' = 'Todos los endpoints M365 accesibles'
}
$Script:I18N.en += @{
    'cat.system' = 'System'; 'cat.outlook' = 'Outlook'; 'cat.ost' = 'OST'; 'cat.pst' = 'PST'
    'cat.cache' = 'Cache'; 'cat.addins' = 'Add-ins'; 'cat.network' = 'Network'
    'local.onlyWindowsTitle' = 'Local diagnostics available on Windows only'
    'local.onlyWindowsRec' = "Run the local part on the affected user's computer (Windows)."
    'sys.detail' = "OS         : {0} (build {1})`nTotal RAM  : {2} GB  (free {3} GB)`nDrive {4} : {5} GB free"
    'sys.ramWarnTitle' = 'Low RAM ({0} GB)'
    'sys.ramWarnRec' = 'Outlook in cached mode with a large OST needs memory; consider adding RAM.'
    'sys.okTitle' = 'System: {0} GB RAM'
    'sys.diskCritTitle' = 'Disk almost full ({0} GB free)'
    'sys.diskCritRec' = 'Low space prevents the OST from growing/repairing -> hangs and sync errors. Free up space.'
    'sys.diskWarnTitle' = 'Low disk space ({0} GB free)'
    'sys.diskWarnRec' = 'The OST needs room to grow. Free up space.'
    'sys.errorTitle' = 'Could not read system info'
    'outlook.unknown' = 'Unknown'
    'outlook.detail' = "Version   : {0}`nPlatform  : {1}`nChannel   : {2}`nProducts  : {3}"
    'outlook.title' = 'Outlook / Microsoft 365 Apps  ({0})'
    'outlook.legacyTitle' = 'Outlook detected (MSI or legacy version)'
    'outlook.legacyDetail' = 'Office branch: {0}'
    'outlook.errorTitle' = 'Could not determine the Outlook version'
    'proc.detail' = "outlook.exe processes: {0}`nMemory (working set): {1} MB`nNot responding: {2}"
    'proc.multiTitle' = 'Multiple Outlook processes ({0})'
    'proc.multiRec' = 'Duplicate/zombie instances can lock the OST. Fully close Outlook (or kill outlook.exe) and reopen.'
    'proc.notRespondingTitle' = 'Outlook is currently not responding'
    'proc.notRespondingRec' = 'Matches the "hang/loop" symptom. Check large folders and add-ins.'
    'proc.okTitle' = 'Outlook running ({0} MB)'
    'proc.notRunningTitle' = 'Outlook is not running'
    'proc.notRunningDetail' = 'Some checks (OST in use) reflect the last state on disk.'
    'ost.noneTitle' = 'No OST files found in the default path'
    'ost.noneDetail' = "Path: {0}`n(Online profile or custom path.)"
    'ost.detail' = "File    : {0}`nSize    : {1} GB`nModified: {2}"
    'ost.critTitle' = 'OST very large: {0} GB ({1})'
    'ost.critRec' = 'Above ~50GB performance drops and corruption risk increases -> slow loading and hangs. Reduce the sync window (Cached Mode slider) or enable the online archive.'
    'ost.warnTitle' = 'Large OST: {0} GB ({1})'
    'ost.warnRec' = 'Reduce "Mail to keep offline" to 3-6 months to speed up loading/sync.'
    'ost.okTitle' = 'OST within normal range: {0} GB'
    'ost.errorTitle' = 'Could not analyze data files'
    'pst.title' = '{0} local PST file(s)'
    'pst.rec' = 'Mounted PSTs increase load time and are fragile over the network.'
    'cache.warnTitle' = 'Large autocomplete cache ({0} MB)'
    'cache.warnDetail' = 'Path: {0}'
    'cache.warnRec' = 'A corrupt/oversized autocomplete stream can slow startup. Rename the RoamCache folder with Outlook closed to regenerate it.'
    'cache.okTitle' = 'Normal autocomplete cache ({0} MB)'
    'addins.warnTitle' = '{0} add-ins active at startup'
    'addins.warnRec' = 'Many add-ins delay Outlook startup. Disable non-essential ones (COM Add-ins) and try "outlook.exe /safe".'
    'addins.okTitle' = '{0} active add-ins'
    'addins.noneTitle' = 'No third-party add-ins active'
    'addins.errorTitle' = 'Could not enumerate add-ins'
    'net.testing' = 'Testing connectivity to Microsoft 365 endpoints...'
    'net.desc.exo' = 'Exchange Online (mail)'
    'net.desc.owa' = 'OWA / service'
    'net.desc.autodiscover' = 'Autodiscover'
    'net.desc.auth' = 'Authentication (Entra ID)'
    'net.critTitle' = 'No connectivity: {0}:{1}'
    'net.critDetail' = '{0} — not responding ({1} ms timeout).'
    'net.critRec' = 'Likely firewall/proxy/VPN block. Without this endpoint there is no sync.'
    'net.warnTitle' = 'High latency to {0} ({1} ms)'
    'net.warnRec' = 'High latency slows down sync. Check proxy/VPN/network.'
    'net.okTitle' = '{0} reachable ({1} ms)'
    'net.allOkTitle' = 'All M365 endpoints reachable'
}
$Script:I18N.pt += @{
    'cat.system' = 'Sistema'; 'cat.outlook' = 'Outlook'; 'cat.ost' = 'OST'; 'cat.pst' = 'PST'
    'cat.cache' = 'Cache'; 'cat.addins' = 'Add-ins'; 'cat.network' = 'Rede'
    'local.onlyWindowsTitle' = 'Diagnostico local disponivel apenas no Windows'
    'local.onlyWindowsRec' = 'Execute a parte local no computador do usuario afetado (Windows).'
    'sys.detail' = "SO         : {0} (build {1})`nRAM total  : {2} GB  (livre {3} GB)`nDisco {4} : {5} GB livres"
    'sys.ramWarnTitle' = 'RAM baixa ({0} GB)'
    'sys.ramWarnRec' = 'O Outlook em modo cache com um OST grande precisa de memoria; considere aumentar a RAM.'
    'sys.okTitle' = 'Sistema: {0} GB RAM'
    'sys.diskCritTitle' = 'Disco quase cheio ({0} GB livres)'
    'sys.diskCritRec' = 'Pouco espaco impede que o OST cresca/repare -> travamentos e erros de sincronizacao. Libere espaco.'
    'sys.diskWarnTitle' = 'Pouco espaco em disco ({0} GB livres)'
    'sys.diskWarnRec' = 'O OST precisa de espaco para crescer. Libere espaco.'
    'sys.errorTitle' = 'Nao foi possivel ler informacoes do sistema'
    'outlook.unknown' = 'Desconhecido'
    'outlook.detail' = "Versao    : {0}`nPlataforma: {1}`nCanal     : {2}`nProdutos  : {3}"
    'outlook.title' = 'Outlook / Microsoft 365 Apps  ({0})'
    'outlook.legacyTitle' = 'Outlook detectado (MSI ou versao legada)'
    'outlook.legacyDetail' = 'Ramo do Office: {0}'
    'outlook.errorTitle' = 'Nao foi possivel determinar a versao do Outlook'
    'proc.detail' = "Processos outlook.exe: {0}`nMemoria (working set): {1} MB`nNao respondendo: {2}"
    'proc.multiTitle' = 'Multiplos processos do Outlook ({0})'
    'proc.multiRec' = 'Instancias duplicadas/zumbis podem bloquear o OST. Feche o Outlook completamente (ou finalize outlook.exe) e reabra.'
    'proc.notRespondingTitle' = 'O Outlook nao esta respondendo no momento'
    'proc.notRespondingRec' = 'Corresponde ao sintoma de "travamento/loop". Verifique pastas grandes e complementos.'
    'proc.okTitle' = 'Outlook em execucao ({0} MB)'
    'proc.notRunningTitle' = 'O Outlook nao esta em execucao'
    'proc.notRunningDetail' = 'Algumas verificacoes (OST em uso) refletem o ultimo estado em disco.'
    'ost.noneTitle' = 'Nenhum arquivo OST encontrado no caminho padrao'
    'ost.noneDetail' = "Caminho: {0}`n(Perfil online ou caminho personalizado.)"
    'ost.detail' = "Arquivo: {0}`nTamanho: {1} GB`nModif. : {2}"
    'ost.critTitle' = 'OST muito grande: {0} GB ({1})'
    'ost.critRec' = 'Acima de ~50GB o desempenho cai e o risco de corrupcao aumenta -> carregamento lento e travamentos. Reduza a janela de sincronizacao (slider do Modo Cache) ou ative o arquivo online.'
    'ost.warnTitle' = 'OST grande: {0} GB ({1})'
    'ost.warnRec' = 'Reduza "Mail to keep offline" para 3-6 meses para acelerar o carregamento/sincronizacao.'
    'ost.okTitle' = 'OST dentro do normal: {0} GB'
    'ost.errorTitle' = 'Nao foi possivel analisar arquivos de dados'
    'pst.title' = '{0} arquivo(s) PST local(is)'
    'pst.rec' = 'PSTs montados aumentam o tempo de carregamento e sao frageis em rede.'
    'cache.warnTitle' = 'Cache de autocompletar grande ({0} MB)'
    'cache.warnDetail' = 'Caminho: {0}'
    'cache.warnRec' = 'Um stream de autocompletar corrompido/enorme pode tornar a inicializacao lenta. Renomeie a pasta RoamCache com o Outlook fechado para regenera-la.'
    'cache.okTitle' = 'Cache de autocompletar normal ({0} MB)'
    'addins.warnTitle' = '{0} complementos ativos na inicializacao'
    'addins.warnRec' = 'Muitos complementos atrasam a inicializacao do Outlook. Desabilite os nao essenciais (COM Add-ins) e tente "outlook.exe /safe".'
    'addins.okTitle' = '{0} complementos ativos'
    'addins.noneTitle' = 'Nenhum complemento de terceiros ativo'
    'addins.errorTitle' = 'Nao foi possivel enumerar os complementos'
    'net.testing' = 'Testando conectividade com endpoints do Microsoft 365...'
    'net.desc.exo' = 'Exchange Online (e-mail)'
    'net.desc.owa' = 'OWA / servico'
    'net.desc.autodiscover' = 'Autodiscover'
    'net.desc.auth' = 'Autenticacao (Entra ID)'
    'net.critTitle' = 'Sem conectividade: {0}:{1}'
    'net.critDetail' = '{0} — nao responde ({1} ms timeout).'
    'net.critRec' = 'Provavel bloqueio de firewall/proxy/VPN. Sem este endpoint nao ha sincronizacao.'
    'net.warnTitle' = 'Latencia alta para {0} ({1} ms)'
    'net.warnRec' = 'Latencia alta torna a sincronizacao lenta. Verifique proxy/VPN/rede.'
    'net.okTitle' = '{0} acessivel ({1} ms)'
    'net.allOkTitle' = 'Todos os endpoints M365 acessiveis'
}

$Script:I18N.es += @{
    'summary.none' = '(Aun no se ha ejecutado ningun diagnostico.)'
    'summary.critical' = 'CRITICO'
    'summary.warning' = 'ALERTA'
    'summary.ok' = 'OK'
    'summary.info' = 'INFO'
    'summary.priorityActions' = '▶ ACCIONES PRIORITARIAS:'
    'summary.allHealthy' = '✓ Sin problemas criticos ni alertas. Buzon/cliente saludables.'
    'report.title' = 'Informe de diagnostico'
    'report.mailboxLabel' = 'Buzon'
    'report.computerLabel' = 'Equipo'
    'report.analystLabel' = 'Analista'
    'report.dateLabel' = 'Fecha'
    'report.headerSev' = 'SEV'
    'report.headerArea' = 'AREA'
    'report.headerFinding' = 'HALLAZGO'
    'report.footer' = 'Generado por'
}
$Script:I18N.en += @{
    'summary.none' = '(No diagnostic has been run yet.)'
    'summary.critical' = 'CRITICAL'
    'summary.warning' = 'WARNING'
    'summary.ok' = 'OK'
    'summary.info' = 'INFO'
    'summary.priorityActions' = '▶ PRIORITY ACTIONS:'
    'summary.allHealthy' = '✓ No critical issues or warnings. Mailbox/client healthy.'
    'report.title' = 'Diagnostic Report'
    'report.mailboxLabel' = 'Mailbox'
    'report.computerLabel' = 'Computer'
    'report.analystLabel' = 'Analyst'
    'report.dateLabel' = 'Date'
    'report.headerSev' = 'SEV'
    'report.headerArea' = 'AREA'
    'report.headerFinding' = 'FINDING'
    'report.footer' = 'Generated by'
}
$Script:I18N.pt += @{
    'summary.none' = '(Nenhum diagnostico foi executado ainda.)'
    'summary.critical' = 'CRITICO'
    'summary.warning' = 'ALERTA'
    'summary.ok' = 'OK'
    'summary.info' = 'INFO'
    'summary.priorityActions' = '▶ ACOES PRIORITARIAS:'
    'summary.allHealthy' = '✓ Sem problemas criticos ou alertas. Caixa/cliente saudaveis.'
    'report.title' = 'Relatorio de diagnostico'
    'report.mailboxLabel' = 'Caixa'
    'report.computerLabel' = 'Computador'
    'report.analystLabel' = 'Analista'
    'report.dateLabel' = 'Data'
    'report.headerSev' = 'SEV'
    'report.headerArea' = 'AREA'
    'report.headerFinding' = 'DESCOBERTA'
    'report.footer' = 'Gerado por'
}

$Script:I18N.es += @{
    'confirm.toConfirmType' = 'Para confirmar, escribe exactamente:'
    'confirm.cancelled' = 'Cancelado (la confirmacion no coincide).'
    'confirm.word.activate' = 'ACTIVAR'
    'confirm.word.apply' = 'APLICAR'
    'assert.noMailbox' = 'No hay buzon objetivo. Define uno con la opcion [7].'
    'assert.needConnection' = 'Necesitas conexion a Exchange Online.'
    'gap.noAccessTitle' = 'Tu sesion NO tiene acceso a estos cmdlets:'
    'gap.rbacExplain1' = 'Esto casi siempre es un tema de ROL (RBAC), no de conexion:'
    'gap.rbacExplain2' = 'Connect-ExchangeOnline solo crea en tu sesion los cmdlets para los'
    'gap.rbacExplain3' = 'que tu cuenta tiene permiso. Si Get-Mailbox funciona pero esto no,'
    'gap.rbacExplain4' = 'te falta el rol de escritura correspondiente.'
    'gap.roleLabel' = 'Rol necesario probable :'
    'gap.whereLabel' = 'Donde verificarlo       :'
    'gap.whereValue' = 'Entra ID > Roles y administradores  (o EAC > Roles > Grupos de roles)'
    'gap.altLabel' = 'Alternativa manual (GUI):'
    'scc.moduleMissing' = 'Falta el modulo ExchangeOnlineManagement (Connect-IPPSSession).'
    'scc.connecting' = 'Conectando a Security & Compliance (se abrira el login)...'
    'scc.connected' = 'Conectado a Security & Compliance.'
    'scc.connectError' = 'No se pudo conectar: {0}'
    'role.recipientMgmt' = 'Recipient Management (grupo de roles de Exchange) o rol Exchange Administrator en Entra ID'
    'role.compliance' = 'eDiscovery Manager o Organization Management (grupos de roles en Microsoft Purview / Security & Compliance)'
    'archive1.alreadyHas' = 'El buzon YA tiene archivo en linea habilitado.'
    'archive1.confirmWarn' = 'Se HABILITARA el archivo en linea para {0} (no destructivo).'
    'archive1.enabled' = 'Archivo en linea habilitado.'
    'archive1.askAutoExpand' = '¿Habilitar tambien el ARCHIVO DE EXPANSION AUTOMATICA (hasta 1.5TB)?'
    'archive1.autoExpandEnabled' = 'Auto-expanding archive habilitado.'
    'archive1.autoExpandError' = 'No se pudo habilitar auto-expanding: {0}'
    'archive1.provisionNote' = 'Nota: el buzon de archivo puede tardar unos minutos en aprovisionarse.'
    'archive1.error' = 'Error: {0}'
    'archive1.manualAlt' = 'Centro de admin de Exchange > Destinatarios > Buzones > (usuario) > Buzon de archivo > Habilitar'
    'move2.manualAlt' = 'Centro de admin de Exchange > Directivas de cumplimiento > Retencion'
    'move2.needsArchive' = 'El buzon no tiene archivo en linea. Se necesita para mover items.'
    'move2.askEnableNow' = '¿Activarlo ahora?'
    'move2.explain1' = 'La correccion aplica una politica de retencion que MUEVE (no borra) al archivo'
    'move2.explain2' = 'los correos mas antiguos, y ejecuta el Managed Folder Assistant.'
    'move2.currentPolicyLabel' = 'Politica de retencion actual:'
    'move2.none' = '(ninguna)'
    'move2.noPolicyTitle' = 'No hay politicas de retencion en el tenant. Crea una con un tag "Mover a archivo".'
    'move2.confirmApply' = "Se asignara la politica '{0}' al buzon {1} (mueve items antiguos al archivo)."
    'move2.policyAssigned' = "Politica '{0}' asignada."
    'move2.runningMFA' = 'Ejecutando Managed Folder Assistant (procesa el buzon en segundo plano)...'
    'move2.mfaStarted' = 'Asistente iniciado. El movimiento al archivo se completa de forma gradual.'
    'move2.suggestion' = 'Sugerencia: revisa el conteo de items con la opcion [2] tras unas horas.'
    'move2.error' = 'Error: {0}'
}
$Script:I18N.en += @{
    'confirm.toConfirmType' = 'To confirm, type exactly:'
    'confirm.cancelled' = 'Cancelled (confirmation did not match).'
    'confirm.word.activate' = 'ENABLE'
    'confirm.word.apply' = 'APPLY'
    'assert.noMailbox' = 'No target mailbox. Set one using option [7].'
    'assert.needConnection' = 'You need a connection to Exchange Online.'
    'gap.noAccessTitle' = 'Your session does NOT have access to these cmdlets:'
    'gap.rbacExplain1' = 'This is almost always a ROLE (RBAC) issue, not a connection issue:'
    'gap.rbacExplain2' = 'Connect-ExchangeOnline only creates the cmdlets in your session that'
    'gap.rbacExplain3' = "your account has permission for. If Get-Mailbox works but this doesn't,"
    'gap.rbacExplain4' = 'you are missing the corresponding write role.'
    'gap.roleLabel' = 'Likely required role   :'
    'gap.whereLabel' = 'Where to check          :'
    'gap.whereValue' = 'Entra ID > Roles and administrators  (or EAC > Roles > Role groups)'
    'gap.altLabel' = 'Manual alternative (GUI):'
    'scc.moduleMissing' = 'The ExchangeOnlineManagement module is missing (Connect-IPPSSession).'
    'scc.connecting' = 'Connecting to Security & Compliance (sign-in will open)...'
    'scc.connected' = 'Connected to Security & Compliance.'
    'scc.connectError' = 'Could not connect: {0}'
    'role.recipientMgmt' = 'Recipient Management (Exchange role group) or the Exchange Administrator role in Entra ID'
    'role.compliance' = 'eDiscovery Manager or Organization Management (role groups in Microsoft Purview / Security & Compliance)'
    'archive1.alreadyHas' = 'The mailbox ALREADY has the online archive enabled.'
    'archive1.confirmWarn' = 'The online archive will be ENABLED for {0} (non-destructive).'
    'archive1.enabled' = 'Online archive enabled.'
    'archive1.askAutoExpand' = 'Also enable AUTO-EXPANDING ARCHIVE (up to 1.5TB)?'
    'archive1.autoExpandEnabled' = 'Auto-expanding archive enabled.'
    'archive1.autoExpandError' = 'Could not enable auto-expanding: {0}'
    'archive1.provisionNote' = 'Note: the archive mailbox may take a few minutes to provision.'
    'archive1.error' = 'Error: {0}'
    'archive1.manualAlt' = 'Exchange admin center > Recipients > Mailboxes > (user) > Archive mailbox > Enable'
    'move2.manualAlt' = 'Exchange admin center > Compliance policies > Retention'
    'move2.needsArchive' = 'The mailbox has no online archive. It is required to move items.'
    'move2.askEnableNow' = 'Enable it now?'
    'move2.explain1' = 'The fix applies a retention policy that MOVES (does not delete) to the archive'
    'move2.explain2' = 'the oldest mail, and runs the Managed Folder Assistant.'
    'move2.currentPolicyLabel' = 'Current retention policy:'
    'move2.none' = '(none)'
    'move2.noPolicyTitle' = 'No retention policies exist in the tenant. Create one with a "Move to archive" tag.'
    'move2.confirmApply' = "Policy '{0}' will be assigned to mailbox {1} (moves old items to the archive)."
    'move2.policyAssigned' = "Policy '{0}' assigned."
    'move2.runningMFA' = 'Running Managed Folder Assistant (processes the mailbox in the background)...'
    'move2.mfaStarted' = 'Assistant started. The move to archive completes gradually.'
    'move2.suggestion' = 'Tip: check the item count with option [2] after a few hours.'
    'move2.error' = 'Error: {0}'
}
$Script:I18N.pt += @{
    'confirm.toConfirmType' = 'Para confirmar, digite exatamente:'
    'confirm.cancelled' = 'Cancelado (a confirmacao nao corresponde).'
    'confirm.word.activate' = 'ATIVAR'
    'confirm.word.apply' = 'APLICAR'
    'assert.noMailbox' = 'Nenhuma caixa alvo. Defina uma com a opcao [7].'
    'assert.needConnection' = 'Voce precisa de uma conexao com o Exchange Online.'
    'gap.noAccessTitle' = 'Sua sessao NAO tem acesso a estes cmdlets:'
    'gap.rbacExplain1' = 'Isso quase sempre e uma questao de FUNCAO (RBAC), nao de conexao:'
    'gap.rbacExplain2' = 'O Connect-ExchangeOnline so cria na sua sessao os cmdlets para os quais'
    'gap.rbacExplain3' = 'sua conta tem permissao. Se o Get-Mailbox funciona mas isso nao,'
    'gap.rbacExplain4' = 'falta a funcao de escrita correspondente.'
    'gap.roleLabel' = 'Funcao provavelmente necessaria:'
    'gap.whereLabel' = 'Onde verificar          :'
    'gap.whereValue' = 'Entra ID > Funcoes e administradores (ou EAC > Funcoes > Grupos de funcoes)'
    'gap.altLabel' = 'Alternativa manual (GUI):'
    'scc.moduleMissing' = 'Falta o modulo ExchangeOnlineManagement (Connect-IPPSSession).'
    'scc.connecting' = 'Conectando ao Security & Compliance (o login sera aberto)...'
    'scc.connected' = 'Conectado ao Security & Compliance.'
    'scc.connectError' = 'Nao foi possivel conectar: {0}'
    'role.recipientMgmt' = 'Recipient Management (grupo de funcoes do Exchange) ou funcao Exchange Administrator no Entra ID'
    'role.compliance' = 'eDiscovery Manager ou Organization Management (grupos de funcoes no Microsoft Purview / Security & Compliance)'
    'archive1.alreadyHas' = 'A caixa JA tem o arquivo online habilitado.'
    'archive1.confirmWarn' = 'O arquivo online sera HABILITADO para {0} (nao destrutivo).'
    'archive1.enabled' = 'Arquivo online habilitado.'
    'archive1.askAutoExpand' = 'Tambem habilitar o ARQUIVO DE EXPANSAO AUTOMATICA (ate 1,5TB)?'
    'archive1.autoExpandEnabled' = 'Arquivo de expansao automatica habilitado.'
    'archive1.autoExpandError' = 'Nao foi possivel habilitar a expansao automatica: {0}'
    'archive1.provisionNote' = 'Nota: a caixa de arquivo pode levar alguns minutos para ser provisionada.'
    'archive1.error' = 'Erro: {0}'
    'archive1.manualAlt' = 'Centro de admin do Exchange > Destinatarios > Caixas de correio > (usuario) > Caixa de arquivo > Habilitar'
    'move2.manualAlt' = 'Centro de admin do Exchange > Politicas de conformidade > Retencao'
    'move2.needsArchive' = 'A caixa nao tem arquivo online. E necessario para mover itens.'
    'move2.askEnableNow' = 'Ativar agora?'
    'move2.explain1' = 'A correcao aplica uma politica de retencao que MOVE (nao exclui) para o arquivo'
    'move2.explain2' = 'os e-mails mais antigos, e executa o Managed Folder Assistant.'
    'move2.currentPolicyLabel' = 'Politica de retencao atual:'
    'move2.none' = '(nenhuma)'
    'move2.noPolicyTitle' = 'Nao ha politicas de retencao no tenant. Crie uma com uma tag "Mover para arquivo".'
    'move2.confirmApply' = "A politica '{0}' sera atribuida a caixa {1} (move itens antigos para o arquivo)."
    'move2.policyAssigned' = "Politica '{0}' atribuida."
    'move2.runningMFA' = 'Executando o Managed Folder Assistant (processa a caixa em segundo plano)...'
    'move2.mfaStarted' = 'Assistente iniciado. A movimentacao para o arquivo e concluida gradualmente.'
    'move2.suggestion' = 'Sugestao: verifique a contagem de itens com a opcao [2] apos algumas horas.'
    'move2.error' = 'Erro: {0}'
}

$Script:I18N.es += @{
    'del3.explain1' = 'Este borrado usa Content Search (Security & Compliance) y es SOFT-DELETE:'
    'del3.explain2' = 'los items van a "Elementos recuperables" y se pueden restaurar durante el'
    'del3.explain3' = 'periodo de retencion. NUNCA se hace purga permanente.'
    'del3.note1' = 'Nota: el purgado procesa ~10 items por buzon y lote; para volumenes muy'
    'del3.note2' = 'grandes (100K+) usa mejor la opcion "Mover al archivo".'
    'del3.manualAlt' = 'portal Microsoft Purview (compliance.microsoft.com) > eDiscovery > Content search'
    'del3.invalidDate' = 'Fecha no valida (esperado: YYYY-MM-DD). {0}'
    'del3.kqlLabel' = 'Consulta KQL:'
    'del3.searching' = 'Creando y ejecutando la busqueda (dry-run / estimacion)...'
    'del3.status' = '... estado: {0}'
    'del3.searchIncomplete' = 'La busqueda no se completo (estado: {0}). Abortando.'
    'del3.estimationHeader' = 'ESTIMACION (DRY-RUN)'
    'del3.matchingItems' = 'Items que coinciden :'
    'del3.estimatedSize' = 'Tamano estimado     :'
    'del3.nothingToDelete' = 'Nada que borrar con ese criterio.'
    'del3.confirmDelete' = 'Se hara SOFT-DELETE (recuperable) de ~{0} items de {1} recibidos antes de {2}.'
    'del3.purging' = 'Purgando (SoftDelete) en lotes, maximo {0}...'
    'del3.batchError' = 'Lote {0}: {1}'
    'del3.batchStatus' = '... lote {0}  estado: {1}'
    'del3.remainingEstimated' = 'restantes estimados: {0}'
    'del3.finished' = 'Purga finalizada. Lotes: {0}. Restantes estimados: {1}.'
    'del3.stillRemaining' = 'Aun quedan items. Vuelve a ejecutar la accion o usa "Mover al archivo".'
    'del3.recoverable' = 'Recuperables: el usuario puede restaurar desde "Elementos recuperables".'
    'del3.error' = 'Error en el borrado: {0}'
    'runbook4.instructions' = 'Ejecuta estos pasos EN EL EQUIPO DEL USUARIO afectado (Windows):'
    'runbook4.step1' = '1. Arrancar en modo seguro (descarta add-ins):'
    'runbook4.step2' = '2. Reconstruir vistas corruptas (util en bucles/cuelgues):'
    'runbook4.step3' = '3. Restablecer el panel de navegacion:'
    'runbook4.step4' = '4. Forzar re-descarga del OST (con Outlook CERRADO, renombrar el .ost para regenerarlo):'
    'runbook4.step5' = '5. Reducir la ventana de sincronizacion en cache (Archivo > Config. cuenta > Cambiar):'
    'runbook4.step5sub' = 'Mover el slider "Correo para conservar sin conexion" a 3-6 meses'
    'runbook4.step6' = '6. Reparar Office si persiste:'
    'runbook4.step6sub' = 'Panel de control > Programas > Microsoft 365 > Modificar > Reparacion rapida'
    'runbook4.reminder1' = 'Recuerda: el arreglo de raiz para "carpeta >100K" es reducir items'
    'runbook4.reminder2' = '(mover a subcarpetas/archivo). El cliente solo se recupera si baja el conteo.'
    'remMenu.opt1' = 'Activar ARCHIVO en linea (+ auto-expanding)   [seguro]'
    'remMenu.opt2' = 'MOVER items antiguos al archivo (politica + MFA) [recomendado]'
    'remMenu.opt3' = 'BORRADO guiado por antiguedad (soft-delete)   [reversible]'
    'remMenu.opt4' = 'Runbook de REPARACION del cliente Outlook'
    'remMenu.opt0' = 'Volver al menu principal'
    'remMenu.auditLabel' = '(auditoria -> {0})'
    'remMenu.invalidOption' = 'Opcion no valida.'
    'remMenu.pressEnter' = 'Pulsa ENTER para continuar...'
    'main.pressEnterReturn' = 'Pulsa ENTER para volver al menu...'
    'main.sessionEnded' = 'Sesion finalizada. Hasta pronto.'
    'main.invalidOption' = 'Opcion no valida.'
}
$Script:I18N.en += @{
    'del3.explain1' = 'This deletion uses Content Search (Security & Compliance) and is SOFT-DELETE:'
    'del3.explain2' = 'items go to "Recoverable Items" and can be restored during the'
    'del3.explain3' = 'retention period. A permanent purge is NEVER performed.'
    'del3.note1' = 'Note: purging processes ~10 items per mailbox per batch; for very'
    'del3.note2' = 'large volumes (100K+) use the "Move to archive" option instead.'
    'del3.manualAlt' = 'Microsoft Purview portal (compliance.microsoft.com) > eDiscovery > Content search'
    'del3.invalidDate' = 'Invalid date (expected: YYYY-MM-DD). {0}'
    'del3.kqlLabel' = 'KQL query:'
    'del3.searching' = 'Creating and running the search (dry-run / estimate)...'
    'del3.status' = '... status: {0}'
    'del3.searchIncomplete' = 'The search did not complete (status: {0}). Aborting.'
    'del3.estimationHeader' = 'ESTIMATE (DRY-RUN)'
    'del3.matchingItems' = 'Matching items      :'
    'del3.estimatedSize' = 'Estimated size      :'
    'del3.nothingToDelete' = 'Nothing to delete with that criteria.'
    'del3.confirmDelete' = 'A SOFT-DELETE (recoverable) of ~{0} items from {1} received before {2} will be performed.'
    'del3.purging' = 'Purging (SoftDelete) in batches, maximum {0}...'
    'del3.batchError' = 'Batch {0}: {1}'
    'del3.batchStatus' = '... batch {0}  status: {1}'
    'del3.remainingEstimated' = 'estimated remaining: {0}'
    'del3.finished' = 'Purge finished. Batches: {0}. Estimated remaining: {1}.'
    'del3.stillRemaining' = 'Items still remain. Re-run the action or use "Move to archive".'
    'del3.recoverable' = 'Recoverable: the user can restore from "Recoverable Items".'
    'del3.error' = 'Error during deletion: {0}'
    'runbook4.instructions' = "Run these steps ON THE AFFECTED USER'S COMPUTER (Windows):"
    'runbook4.step1' = '1. Start in safe mode (skips add-ins):'
    'runbook4.step2' = '2. Rebuild corrupt views (useful for loops/hangs):'
    'runbook4.step3' = '3. Reset the navigation pane:'
    'runbook4.step4' = '4. Force OST re-download (with Outlook CLOSED, rename the .ost to regenerate it):'
    'runbook4.step5' = '5. Reduce the cached sync window (File > Account Settings > Change):'
    'runbook4.step5sub' = 'Move the "Mail to keep offline" slider to 3-6 months'
    'runbook4.step6' = '6. Repair Office if it persists:'
    'runbook4.step6sub' = 'Control Panel > Programs > Microsoft 365 > Modify > Quick Repair'
    'runbook4.reminder1' = 'Remember: the root fix for a ">100K folder" is reducing items'
    'runbook4.reminder2' = '(move to subfolders/archive). The client only recovers if the count drops.'
    'remMenu.opt1' = 'Enable online ARCHIVE (+ auto-expanding)   [safe]'
    'remMenu.opt2' = 'MOVE old items to archive (policy + MFA) [recommended]'
    'remMenu.opt3' = 'Age-guided DELETION (soft-delete)   [reversible]'
    'remMenu.opt4' = 'Outlook client REPAIR runbook'
    'remMenu.opt0' = 'Back to main menu'
    'remMenu.auditLabel' = '(audit log -> {0})'
    'remMenu.invalidOption' = 'Invalid option.'
    'remMenu.pressEnter' = 'Press ENTER to continue...'
    'main.pressEnterReturn' = 'Press ENTER to return to the menu...'
    'main.sessionEnded' = 'Session ended. See you soon.'
    'main.invalidOption' = 'Invalid option.'
}
$Script:I18N.pt += @{
    'del3.explain1' = 'Esta exclusao usa Content Search (Security & Compliance) e e SOFT-DELETE:'
    'del3.explain2' = 'os itens vao para "Itens Recuperaveis" e podem ser restaurados durante o'
    'del3.explain3' = 'periodo de retencao. NUNCA e feita uma exclusao permanente.'
    'del3.note1' = 'Nota: a exclusao processa ~10 itens por caixa por lote; para volumes muito'
    'del3.note2' = 'grandes (100K+) use a opcao "Mover para o arquivo".'
    'del3.manualAlt' = 'portal Microsoft Purview (compliance.microsoft.com) > eDiscovery > Content search'
    'del3.invalidDate' = 'Data invalida (esperado: YYYY-MM-DD). {0}'
    'del3.kqlLabel' = 'Consulta KQL:'
    'del3.searching' = 'Criando e executando a pesquisa (dry-run / estimativa)...'
    'del3.status' = '... status: {0}'
    'del3.searchIncomplete' = 'A pesquisa nao foi concluida (status: {0}). Abortando.'
    'del3.estimationHeader' = 'ESTIMATIVA (DRY-RUN)'
    'del3.matchingItems' = 'Itens correspondentes:'
    'del3.estimatedSize' = 'Tamanho estimado    :'
    'del3.nothingToDelete' = 'Nada a excluir com esse criterio.'
    'del3.confirmDelete' = 'Sera feito um SOFT-DELETE (recuperavel) de ~{0} itens de {1} recebidos antes de {2}.'
    'del3.purging' = 'Excluindo (SoftDelete) em lotes, maximo {0}...'
    'del3.batchError' = 'Lote {0}: {1}'
    'del3.batchStatus' = '... lote {0}  status: {1}'
    'del3.remainingEstimated' = 'restantes estimados: {0}'
    'del3.finished' = 'Exclusao finalizada. Lotes: {0}. Restantes estimados: {1}.'
    'del3.stillRemaining' = 'Ainda restam itens. Execute a acao novamente ou use "Mover para o arquivo".'
    'del3.recoverable' = 'Recuperavel: o usuario pode restaurar a partir de "Itens Recuperaveis".'
    'del3.error' = 'Erro na exclusao: {0}'
    'runbook4.instructions' = 'Execute estas etapas NO COMPUTADOR DO USUARIO afetado (Windows):'
    'runbook4.step1' = '1. Iniciar em modo de seguranca (ignora complementos):'
    'runbook4.step2' = '2. Reconstruir visualizacoes corrompidas (util em loops/travamentos):'
    'runbook4.step3' = '3. Redefinir o painel de navegacao:'
    'runbook4.step4' = '4. Forcar novo download do OST (com o Outlook FECHADO, renomeie o .ost para regenera-lo):'
    'runbook4.step5' = '5. Reduzir a janela de sincronizacao em cache (Arquivo > Config. da conta > Alterar):'
    'runbook4.step5sub' = 'Mover o controle deslizante "Emails a manter offline" para 3-6 meses'
    'runbook4.step6' = '6. Reparar o Office se persistir:'
    'runbook4.step6sub' = 'Painel de Controle > Programas > Microsoft 365 > Modificar > Reparo rapido'
    'runbook4.reminder1' = 'Lembre-se: a correcao raiz para "pasta >100K" e reduzir itens'
    'runbook4.reminder2' = '(mover para subpastas/arquivo). O cliente so se recupera se a contagem diminuir.'
    'remMenu.opt1' = 'Ativar ARQUIVO online (+ expansao automatica)   [seguro]'
    'remMenu.opt2' = 'MOVER itens antigos para o arquivo (politica + MFA) [recomendado]'
    'remMenu.opt3' = 'EXCLUSAO guiada por antiguidade (soft-delete)   [reversivel]'
    'remMenu.opt4' = 'Runbook de REPARO do cliente Outlook'
    'remMenu.opt0' = 'Voltar ao menu principal'
    'remMenu.auditLabel' = '(auditoria -> {0})'
    'remMenu.invalidOption' = 'Opcao invalida.'
    'remMenu.pressEnter' = 'Pressione ENTER para continuar...'
    'main.pressEnterReturn' = 'Pressione ENTER para voltar ao menu...'
    'main.sessionEnded' = 'Sessao encerrada. Ate breve.'
    'main.invalidOption' = 'Opcao invalida.'
}

# =====================================================================================
#  MOTOR DE HALLAZGOS
# =====================================================================================
$Script:Findings = [System.Collections.Generic.List[object]]::new()
$Script:SessionInfo = [ordered]@{
    Started   = Get-Date
    Mailbox   = $Mailbox
    Host      = $env:COMPUTERNAME
    User      = $env:USERNAME
    Connected = $false
}

function Write-Finding {
    param($F)
    switch ($F.Severity) {
        'CRIT' { $tag = ' ' + (T 'summary.critical').PadRight(8) + ' '; $c = 'Red';    $g = '✕' }
        'WARN' { $tag = ' ' + (T 'summary.warning').PadRight(8) + ' '; $c = 'Yellow'; $g = '▲' }
        'OK'   { $tag = ' ' + (T 'summary.ok').PadRight(8) + ' ';      $c = 'Green';  $g = '✓' }
        default{ $tag = ' ' + (T 'summary.info').PadRight(8) + ' ';    $c = 'Cyan';   $g = '◆' }
    }
    Write-Host ('  ' + (Ink "$($Script:E)[48;2;$($Script:Palette[$c])m$($Script:E)[38;2;10;10;15m$tag$($Script:E)[0m")) -NoNewline
    Write-Host (Ink " $g " $c) -NoNewline
    Write-Host (Ink $F.Title 'White')
    if ($F.Detail) {
        foreach ($line in ($F.Detail -split "`n")) {
            Write-Host (Ink "         │ " 'Dark') -NoNewline
            Write-Host (Ink $line 'Gray')
        }
    }
    if ($F.Severity -in @('CRIT','WARN') -and $F.Recommendation) {
        Write-Host (Ink "         └▶ " $c) -NoNewline
        Write-Host (Ink $F.Recommendation 'Orange')
    }
}

function Add-Finding {
    param(
        [string]$Category,
        [ValidateSet('OK','INFO','WARN','CRIT')][string]$Severity,
        [string]$Title,
        [string]$Detail = '',
        [string]$Recommendation = ''
    )
    $f = [pscustomobject]@{
        Time = Get-Date; Category = $Category; Severity = $Severity
        Title = $Title; Detail = $Detail; Recommendation = $Recommendation
    }
    $Script:Findings.Add($f)
    Write-Finding $f
}

# =====================================================================================
#  UTILIDADES
# =====================================================================================
function Get-Bytes {
    param($SizeValue)
    if ($null -eq $SizeValue) { return [int64]0 }
    # Objetos EXO suelen exponer ToBytes()
    try { if ($SizeValue.PSObject.Methods['ToBytes']) { return [int64]$SizeValue.ToBytes() } } catch {}
    $s = "$SizeValue"
    if ($s -match '\(([0-9,\.]+)\s*bytes\)') { return [int64]([double]($matches[1] -replace '[,\.]','')) }
    if ($s -match '([0-9\.]+)\s*(B|KB|MB|GB|TB)') {
        $n = [double]$matches[1]
        switch ($matches[2]) {
            'B'  { return [int64]$n }
            'KB' { return [int64]($n * 1KB) }
            'MB' { return [int64]($n * 1MB) }
            'GB' { return [int64]($n * 1GB) }
            'TB' { return [int64]($n * 1TB) }
        }
    }
    return [int64]0
}

function Format-Size {
    param([int64]$Bytes)
    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ('{0:N2} GB' -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ('{0:N1} MB' -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ('{0:N0} KB' -f ($Bytes / 1KB)) }
    return "$Bytes B"
}

function Test-Tcp {
    param([string]$HostName, [int]$Port = 443, [int]$TimeoutMs = 3500)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $client = [System.Net.Sockets.TcpClient]::new()
        $iar = $client.BeginConnect($HostName, $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs)
        if ($ok -and $client.Connected) {
            $client.EndConnect($iar); $client.Close(); $sw.Stop()
            return [pscustomobject]@{ Ok = $true; Ms = $sw.ElapsedMilliseconds }
        }
        $client.Close(); $sw.Stop()
        return [pscustomobject]@{ Ok = $false; Ms = $sw.ElapsedMilliseconds }
    } catch { $sw.Stop(); return [pscustomobject]@{ Ok = $false; Ms = $sw.ElapsedMilliseconds } }
}

function Read-Prompt {
    param([string]$Text)
    Write-Host (Ink "  ┌─[" 'Purple') -NoNewline
    Write-Host (Ink 'outlook' 'Green') -NoNewline
    Write-Host (Ink "]─[" 'Purple') -NoNewline
    Write-Host (Ink $Text 'Cyan') -NoNewline
    Write-Host (Ink "]" 'Purple')
    Write-Host (Ink "  └─▶ " 'Purple') -NoNewline
    return Read-Host
}

# =====================================================================================
#  CONEXION EXCHANGE ONLINE
# =====================================================================================
function Test-EXOConnected {
    try {
        $c = Get-ConnectionInformation -ErrorAction SilentlyContinue
        return [bool]($c | Where-Object { $_.State -eq 'Connected' })
    } catch { return $false }
}

function Connect-EXO {
    if (Test-EXOConnected) {
        $Script:SessionInfo.Connected = $true
        Write-Host (Ink "  [~] $(T 'exo.alreadyConnected')" 'Green')
        return $true
    }
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host (Ink "  [x] $(T 'exo.moduleNotInstalled')" 'Red')
        Write-Host (Ink '      Install-Module ExchangeOnlineManagement -Scope CurrentUser' 'Gray')
        return $false
    }
    try {
        Write-Host (Ink "  [~] $(T 'exo.connecting')" 'Cyan')
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        # Forzar recarga del modulo para que carguen todos los cmdlets de EXO
        Import-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue | Out-Null
        $Script:SessionInfo.Connected = $true
        Write-Host (Ink "  [+] $(T 'exo.connected')" 'Green')
        return $true
    } catch {
        Write-Host (Ink "  [x] $(T 'exo.connectError' @($_.Exception.Message))" 'Red')
        return $false
    }
}

# =====================================================================================
#  DIAGNOSTICO REMOTO  ::  Exchange Online
# =====================================================================================
function Invoke-RemoteDiagnostics {
    param([string]$Mbx)

    Write-Section (T 'section.remoteDiag' @($Mbx)) 'Pink'

    if (-not (Test-EXOConnected)) {
        if (-not (Connect-EXO)) {
            Add-Finding (T 'cat.remote') 'CRIT' (T 'remote.noConnectionTitle') (T 'remote.noConnectionDetail') (T 'remote.noConnectionRec')
            return
        }
    }

    # --- 1. Existencia + overview del buzon ---------------------------------------
    try {
        $mb = Get-Mailbox -Identity $Mbx -ErrorAction Stop
    } catch {
        Add-Finding (T 'cat.remote') 'CRIT' (T 'remote.notFoundTitle' @($Mbx)) $_.Exception.Message (T 'remote.notFoundRec')
        return
    }
    $Script:SessionInfo.Mailbox = $mb.PrimarySmtpAddress

    # --- 2. Cuota y tamano --------------------------------------------------------
    try {
        $st = Get-MailboxStatistics -Identity $Mbx -ErrorAction Stop
        $usedB = Get-Bytes $st.TotalItemSize
        $quotaB = Get-Bytes $mb.ProhibitSendReceiveQuota
        if ($quotaB -le 0) { $quotaB = Get-Bytes $mb.ProhibitSendQuota }
        $items = [int64]$st.ItemCount
        $pct = if ($quotaB -gt 0) { [Math]::Round(($usedB / $quotaB) * 100, 1) } else { 0 }
        $quotaText = if ($quotaB -gt 0) { Format-Size $quotaB } else { T 'quota.unlimited' }
        $detail = T 'quota.detail' @((Format-Size $usedB), $quotaText, $pct, ('{0:N0}' -f $items), $st.LastLogonTime)
        if ($pct -ge $Script:Thresholds.MailboxQuotaCritPct) {
            Add-Finding (T 'cat.quota') 'CRIT' (T 'quota.critTitle' @($pct)) $detail (T 'quota.critRec')
        } elseif ($pct -ge $Script:Thresholds.MailboxQuotaWarnPct) {
            Add-Finding (T 'cat.quota') 'WARN' (T 'quota.critTitle' @($pct)) $detail (T 'quota.warnRec')
        } else {
            Add-Finding (T 'cat.quota') 'OK' (T 'quota.okTitle' @($pct)) $detail
        }
    } catch {
        Add-Finding (T 'cat.quota') 'WARN' (T 'quota.errorTitle') $_.Exception.Message
    }

    # --- 3. CARPETAS CON DEMASIADOS ITEMS (el caso del bucle) ---------------------
    try {
        Write-Host (Ink "  [~] $(T 'folders.analyzing')" 'Gray')
        $folders = Get-MailboxFolderStatistics -Identity $Mbx -ErrorAction Stop |
            Select-Object Name, FolderPath, ItemsInFolder, FolderAndSubfolderSize, FolderSize
        $top = $folders | Sort-Object ItemsInFolder -Descending | Select-Object -First 8
        $lines = $top | ForEach-Object { '{0,10:N0} items  ·  {1}' -f $_.ItemsInFolder, $_.FolderPath }
        Add-Finding (T 'cat.folders') 'INFO' (T 'folders.topTitle') ($lines -join "`n")

        $crit = $folders | Where-Object { $_.ItemsInFolder -ge $Script:Thresholds.FolderItemsCritical }
        $warn = $folders | Where-Object { $_.ItemsInFolder -ge $Script:Thresholds.FolderItemsWarning -and $_.ItemsInFolder -lt $Script:Thresholds.FolderItemsCritical }
        foreach ($f in $crit) {
            Add-Finding (T 'cat.folders') 'CRIT' (T 'folders.critTitle' @(('{0:N0}' -f $f.ItemsInFolder), $f.FolderPath)) `
                (T 'folders.critDetail' @(('{0:N0}' -f $Script:Thresholds.FolderItemsCritical))) `
                (T 'folders.critRec')
        }
        foreach ($f in $warn) {
            Add-Finding (T 'cat.folders') 'WARN' (T 'folders.warnTitle' @(('{0:N0}' -f $f.ItemsInFolder), $f.FolderPath)) `
                (T 'folders.warnDetail' @(('{0:N0}' -f $Script:Thresholds.FolderItemsCritical))) `
                (T 'folders.warnRec')
        }
        if (-not $crit -and -not $warn) {
            Add-Finding (T 'cat.folders') 'OK' (T 'folders.okTitle') (T 'folders.okDetail' @(('{0:N0}' -f (($folders | Measure-Object ItemsInFolder -Maximum).Maximum))))
        }
    } catch {
        Add-Finding (T 'cat.folders') 'WARN' (T 'folders.errorTitle') $_.Exception.Message
    }

    # --- 4. Reglas de bandeja -----------------------------------------------------
    try {
        $rules = @(Get-InboxRule -Mailbox $Mbx -ErrorAction Stop)
        $n = $rules.Count
        $fwd = @($rules | Where-Object { $_.ForwardTo -or $_.RedirectTo -or $_.ForwardAsAttachmentTo })
        $detail = T 'rules.detail' @($n, $fwd.Count)
        if ($n -ge $Script:Thresholds.InboxRulesWarn) {
            Add-Finding (T 'cat.rules') 'WARN' (T 'rules.warnTitle' @($n)) $detail (T 'rules.warnRec')
        } else {
            Add-Finding (T 'cat.rules') 'OK' (T 'rules.okTitle' @($n)) $detail
        }
        if ($fwd.Count -gt 0) {
            $fwdLines = $fwd | ForEach-Object { "· $($_.Name) -> $([string]::Join(',', @($_.ForwardTo + $_.RedirectTo)))" }
            Add-Finding (T 'cat.rules') 'INFO' (T 'rules.fwdTitle') ($fwdLines -join "`n") (T 'rules.fwdRec')
        }
    } catch {
        Add-Finding (T 'cat.rules') 'INFO' (T 'rules.errorTitle') $_.Exception.Message
    }

    # --- 5. Reenvio a nivel de buzon ---------------------------------------------
    if ($mb.ForwardingSmtpAddress -or $mb.ForwardingAddress) {
        Add-Finding (T 'cat.flow') 'WARN' (T 'fwd.title') `
            (T 'fwd.detail' @($mb.ForwardingSmtpAddress, $mb.ForwardingAddress, $mb.DeliverToMailboxAndForward)) `
            (T 'fwd.rec')
    }

    # --- 6. Dispositivos moviles / sincronizacion --------------------------------
    try {
        $devs = @(Get-MobileDeviceStatistics -Mailbox $Mbx -ErrorAction Stop)
        if ($devs.Count -gt 0) {
            $lines = $devs | Sort-Object LastSuccessSync -Descending | Select-Object -First 6 | ForEach-Object {
                "· $($_.DeviceModel) [$($_.DeviceType)]  estado:$($_.Status)  ultimo sync:$($_.LastSuccessSync)"
            }
            $broken = @($devs | Where-Object { $_.Status -ne 'DeviceOk' })
            if ($broken.Count -gt 0) {
                Add-Finding (T 'cat.sync') 'WARN' (T 'devices.warnTitle' @($broken.Count)) ($lines -join "`n") (T 'devices.warnRec')
            } else {
                Add-Finding (T 'cat.sync') 'OK' (T 'devices.okTitle' @($devs.Count)) ($lines -join "`n")
            }
        } else {
            Add-Finding (T 'cat.sync') 'INFO' (T 'devices.noneTitle') ''
        }
    } catch {
        Add-Finding (T 'cat.sync') 'INFO' (T 'devices.errorTitle') $_.Exception.Message
    }

    # --- 7. Traza de flujo de correo (no entran / no salen) ----------------------
    try {
        $end = (Get-Date)
        $start = $end.AddHours(-$Script:Thresholds.TraceHours)
        $smtp = "$($mb.PrimarySmtpAddress)"
        Write-Host (Ink "  [~] $(T 'trace.tracing' @($Script:Thresholds.TraceHours))" 'Gray')

        $traceCmd = if (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue) { 'Get-MessageTraceV2' }
                    elseif (Get-Command Get-MessageTrace -ErrorAction SilentlyContinue) { 'Get-MessageTrace' }
                    else { $null }

        if ($traceCmd) {
            $inbound  = @(& $traceCmd -RecipientAddress $smtp -StartDate $start -EndDate $end -ErrorAction SilentlyContinue)
            $outbound = @(& $traceCmd -SenderAddress $smtp -StartDate $start -EndDate $end -ErrorAction SilentlyContinue)

            $inFail  = @($inbound  | Where-Object { $_.Status -match 'Fail|Quarantine|FilteredAsSpam' })
            $outFail = @($outbound | Where-Object { $_.Status -match 'Fail|Quarantine' })

            $detail = T 'trace.detail' @($Script:Thresholds.TraceHours, $inbound.Count, $inFail.Count, $outbound.Count, $outFail.Count)

            if ($inbound.Count -eq 0) {
                Add-Finding (T 'cat.flow') 'WARN' (T 'trace.noInboundTitle') $detail (T 'trace.noInboundRec')
            } elseif ($inFail.Count -gt 0) {
                $ex = ($inFail | Select-Object -First 4 | ForEach-Object { "· $($_.Received)  de:$($_.SenderAddress)  estado:$($_.Status)" }) -join "`n"
                Add-Finding (T 'cat.flow') 'WARN' (T 'trace.inFailTitle' @($inFail.Count)) "$detail`n$ex" (T 'trace.inFailRec')
            } else {
                Add-Finding (T 'cat.flow') 'OK' (T 'trace.okTitle') $detail
            }
            if ($outFail.Count -gt 0) {
                $outEx = ($outFail | Select-Object -First 4 | ForEach-Object { "· para:$($_.RecipientAddress)  estado:$($_.Status)" }) -join "`n"
                Add-Finding (T 'cat.flow') 'WARN' (T 'trace.outFailTitle' @($outFail.Count)) $outEx (T 'trace.outFailRec')
            }
        } else {
            Add-Finding (T 'cat.flow') 'INFO' (T 'trace.unavailableTitle') (T 'trace.unavailableDetail')
        }
    } catch {
        Add-Finding (T 'cat.flow') 'INFO' (T 'trace.errorTitle') $_.Exception.Message
    }

    # --- 8. Archivo en linea + litigation hold -----------------------------------
    try {
        $arch = if ($mb.ArchiveStatus -eq 'Active' -or $mb.ArchiveDatabase) { T 'archive.enabled' } else { T 'archive.disabled' }
        $hold = if ($mb.LitigationHoldEnabled) { 'SI' } else { 'No' }
        Add-Finding (T 'cat.mailbox') 'INFO' (T 'archive.title') (T 'archive.detail' @($arch, $hold, $mb.RecipientTypeDetails))
    } catch {}
}

# =====================================================================================
#  DIAGNOSTICO LOCAL  ::  Cliente Outlook
# =====================================================================================
function Invoke-LocalDiagnostics {
    Write-Section (T 'section.localDiag' @($env:COMPUTERNAME)) 'Green'

    if (-not $Script:OnWindows) {
        Add-Finding (T 'cat.system') 'INFO' (T 'local.onlyWindowsTitle') (T 'local.onlyWindowsRec')
        return
    }

    # --- 1. Sistema (RAM / disco / OS) -------------------------------------------
    try {
        $os  = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $cs  = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $ramGB = [Math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
        $freeRamGB = [Math]::Round($os.FreePhysicalMemory * 1KB / 1GB, 1)
        $sysDrive = ($env:SystemDrive)
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$sysDrive'" -ErrorAction SilentlyContinue
        $freeGB = if ($disk) { [Math]::Round($disk.FreeSpace / 1GB, 1) } else { 0 }
        $detail = T 'sys.detail' @($os.Caption, $os.BuildNumber, $ramGB, $freeRamGB, $sysDrive, $freeGB)

        if ($ramGB -lt $Script:Thresholds.RamWarnGB) {
            Add-Finding (T 'cat.system') 'WARN' (T 'sys.ramWarnTitle' @($ramGB)) $detail (T 'sys.ramWarnRec')
        } else {
            Add-Finding (T 'cat.system') 'OK' (T 'sys.okTitle' @($ramGB)) $detail
        }
        if ($freeGB -le $Script:Thresholds.DiskFreeCritGB) {
            Add-Finding (T 'cat.system') 'CRIT' (T 'sys.diskCritTitle' @($freeGB)) $detail (T 'sys.diskCritRec')
        } elseif ($freeGB -le $Script:Thresholds.DiskFreeWarnGB) {
            Add-Finding (T 'cat.system') 'WARN' (T 'sys.diskWarnTitle' @($freeGB)) $detail (T 'sys.diskWarnRec')
        }
    } catch {
        Add-Finding (T 'cat.system') 'INFO' (T 'sys.errorTitle') $_.Exception.Message
    }

    # --- 2. Instalacion de Outlook (Click-to-Run) --------------------------------
    try {
        $c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue
        if ($c2r) {
            $chanMap = @{
                'CurrentChannel'='Current (Monthly)'; 'Current'='Current (Monthly)';
                'MonthlyEnterprise'='Monthly Enterprise'; 'FirstReleaseCurrent'='Current (Preview)';
                'SemiAnnual'='Semi-Annual'; 'SemiAnnualPreview'='Semi-Annual (Preview)';
                'Deferred'='Semi-Annual'; 'InsiderFast'='Beta'
            }
            $chanRaw = "$($c2r.UpdateChannel)$($c2r.CDNBaseUrl)"
            $chan = ($chanMap.GetEnumerator() | Where-Object { $chanRaw -match $_.Key } | Select-Object -First 1).Value
            if (-not $chan) { $chan = T 'outlook.unknown' }
            $detail = T 'outlook.detail' @($c2r.VersionToReport, $c2r.Platform, $chan, $c2r.ProductReleaseIds)
            Add-Finding (T 'cat.outlook') 'INFO' (T 'outlook.title' @($c2r.Platform)) $detail
        } else {
            # Outlook MSI / version por registro Office
            $ver = (Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^\d+\.\d+$' } | Sort-Object PSChildName -Descending | Select-Object -First 1).PSChildName
            Add-Finding (T 'cat.outlook') 'INFO' (T 'outlook.legacyTitle') (T 'outlook.legacyDetail' @($ver))
        }
    } catch {
        Add-Finding (T 'cat.outlook') 'INFO' (T 'outlook.errorTitle') $_.Exception.Message
    }

    # --- 3. Procesos de Outlook en ejecucion -------------------------------------
    try {
        $procs = @(Get-Process outlook -ErrorAction SilentlyContinue)
        if ($procs.Count -gt 0) {
            $memMB = [Math]::Round((($procs | Measure-Object WorkingSet64 -Sum).Sum) / 1MB, 0)
            $notResp = @($procs | Where-Object { $_.Responding -eq $false })
            $detail = T 'proc.detail' @($procs.Count, $memMB, $notResp.Count)
            if ($procs.Count -gt 1) {
                Add-Finding (T 'cat.outlook') 'WARN' (T 'proc.multiTitle' @($procs.Count)) $detail (T 'proc.multiRec')
            } elseif ($notResp.Count -gt 0) {
                Add-Finding (T 'cat.outlook') 'WARN' (T 'proc.notRespondingTitle') $detail (T 'proc.notRespondingRec')
            } else {
                Add-Finding (T 'cat.outlook') 'OK' (T 'proc.okTitle' @($memMB)) $detail
            }
        } else {
            Add-Finding (T 'cat.outlook') 'INFO' (T 'proc.notRunningTitle') (T 'proc.notRunningDetail')
        }
    } catch {}

    # --- 4. Archivos OST / PST ---------------------------------------------------
    try {
        $outlookData = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook'
        $osts = @(Get-ChildItem -Path $outlookData -Filter '*.ost' -ErrorAction SilentlyContinue)
        if ($osts.Count -eq 0) {
            Add-Finding (T 'cat.ost') 'INFO' (T 'ost.noneTitle') (T 'ost.noneDetail' @($outlookData))
        }
        foreach ($ost in $osts) {
            $gb = [Math]::Round($ost.Length / 1GB, 2)
            $detail = T 'ost.detail' @($ost.Name, $gb, $ost.LastWriteTime)
            if ($gb -ge $Script:Thresholds.OstSizeCritGB) {
                Add-Finding (T 'cat.ost') 'CRIT' (T 'ost.critTitle' @($gb, $ost.Name)) $detail (T 'ost.critRec')
            } elseif ($gb -ge $Script:Thresholds.OstSizeWarnGB) {
                Add-Finding (T 'cat.ost') 'WARN' (T 'ost.warnTitle' @($gb, $ost.Name)) $detail (T 'ost.warnRec')
            } else {
                Add-Finding (T 'cat.ost') 'OK' (T 'ost.okTitle' @($gb)) $detail
            }
        }
        $psts = @(Get-ChildItem -Path $outlookData -Filter '*.pst' -ErrorAction SilentlyContinue)
        if ($psts.Count -gt 0) {
            $lines = $psts | ForEach-Object { "· $($_.Name)  $([Math]::Round($_.Length/1GB,2)) GB" }
            Add-Finding (T 'cat.pst') 'INFO' (T 'pst.title' @($psts.Count)) ($lines -join "`n") (T 'pst.rec')
        }
    } catch {
        Add-Finding (T 'cat.ost') 'INFO' (T 'ost.errorTitle') $_.Exception.Message
    }

    # --- 5. Cache de autocompletado (RoamCache) ----------------------------------
    try {
        $roam = Join-Path $env:LOCALAPPDATA 'Microsoft\Outlook\RoamCache'
        $ac = @(Get-ChildItem -Path $roam -Filter 'Stream_Autocomplete_*' -ErrorAction SilentlyContinue)
        if ($ac.Count -gt 0) {
            $mb = [Math]::Round((($ac | Measure-Object Length -Sum).Sum) / 1MB, 1)
            if ($mb -ge $Script:Thresholds.AutoCompleteWarnMB) {
                Add-Finding (T 'cat.cache') 'WARN' (T 'cache.warnTitle' @($mb)) (T 'cache.warnDetail' @($roam)) (T 'cache.warnRec')
            } else {
                Add-Finding (T 'cat.cache') 'OK' (T 'cache.okTitle' @($mb)) ''
            }
        }
    } catch {}

    # --- 6. Complementos (Add-ins) -----------------------------------------------
    try {
        $addinRoots = @(
            'HKCU:\Software\Microsoft\Office\Outlook\Addins',
            'HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins'
        )
        $active = [System.Collections.Generic.List[object]]::new()
        foreach ($root in $addinRoots) {
            if (Test-Path $root) {
                foreach ($k in (Get-ChildItem $root -ErrorAction SilentlyContinue)) {
                    $p = Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue
                    $lb = $p.LoadBehavior
                    if ($lb -in 3, 9) {  # cargado al inicio
                        $active.Add([pscustomobject]@{
                            Name = if ($p.FriendlyName) { $p.FriendlyName } else { $k.PSChildName }
                            Prog = $k.PSChildName
                        })
                    }
                }
            }
        }
        $uniq = $active | Sort-Object Prog -Unique
        $lines = $uniq | ForEach-Object { "· $($_.Name)" }
        if ($uniq.Count -ge $Script:Thresholds.AddinsWarn) {
            Add-Finding (T 'cat.addins') 'WARN' (T 'addins.warnTitle' @($uniq.Count)) ($lines -join "`n") (T 'addins.warnRec')
        } elseif ($uniq.Count -gt 0) {
            Add-Finding (T 'cat.addins') 'OK' (T 'addins.okTitle' @($uniq.Count)) ($lines -join "`n") ''
        } else {
            Add-Finding (T 'cat.addins') 'OK' (T 'addins.noneTitle') ''
        }
    } catch {
        Add-Finding (T 'cat.addins') 'INFO' (T 'addins.errorTitle') $_.Exception.Message
    }

    # --- 7. Conectividad a endpoints M365 ----------------------------------------
    Write-Host (Ink "  [~] $(T 'net.testing')" 'Gray')
    $endpoints = @(
        @{ Host = 'outlook.office365.com';        Port = 443; Desc = (T 'net.desc.exo') }
        @{ Host = 'outlook.office.com';           Port = 443; Desc = (T 'net.desc.owa') }
        @{ Host = 'autodiscover-s.outlook.com';   Port = 443; Desc = (T 'net.desc.autodiscover') }
        @{ Host = 'login.microsoftonline.com';    Port = 443; Desc = (T 'net.desc.auth') }
    )
    $failed = 0
    foreach ($ep in $endpoints) {
        $r = Test-Tcp -HostName $ep.Host -Port $ep.Port
        if (-not $r.Ok) {
            $failed++
            Add-Finding (T 'cat.network') 'CRIT' (T 'net.critTitle' @($ep.Host, $ep.Port)) (T 'net.critDetail' @($ep.Desc, $r.Ms)) (T 'net.critRec')
        } elseif ($r.Ms -gt 800) {
            Add-Finding (T 'cat.network') 'WARN' (T 'net.warnTitle' @($ep.Host, $r.Ms)) $ep.Desc (T 'net.warnRec')
        } else {
            Add-Finding (T 'cat.network') 'OK' (T 'net.okTitle' @($ep.Host, $r.Ms)) $ep.Desc
        }
    }
    if ($failed -eq 0) {
        Add-Finding (T 'cat.network') 'OK' (T 'net.allOkTitle') ''
    }
}

# =====================================================================================
#  RESUMEN / DASHBOARD
# =====================================================================================
function Show-Summary {
    Write-Section (T 'section.summary') 'Yellow'
    if ($Script:Findings.Count -eq 0) {
        Write-Host (Ink "  $(T 'summary.none')" 'Gray')
        return
    }
    $crit = @($Script:Findings | Where-Object Severity -eq 'CRIT')
    $warn = @($Script:Findings | Where-Object Severity -eq 'WARN')
    $ok   = @($Script:Findings | Where-Object Severity -eq 'OK')
    $info = @($Script:Findings | Where-Object Severity -eq 'INFO')

    function Bar { param([int]$n,[string]$c) (Ink ('█' * [Math]::Min($n, 40)) $c) + (Ink " $n" $c -Bold) }
    Write-Host ''
    Write-Host (Ink "   $((T 'summary.critical').PadRight(9)) " 'Red' -Bold)    -NoNewline; Write-Host (Bar $crit.Count 'Red')
    Write-Host (Ink "   $((T 'summary.warning').PadRight(9)) " 'Yellow' -Bold) -NoNewline; Write-Host (Bar $warn.Count 'Yellow')
    Write-Host (Ink "   $((T 'summary.ok').PadRight(9)) " 'Green' -Bold)  -NoNewline; Write-Host (Bar $ok.Count 'Green')
    Write-Host (Ink "   $((T 'summary.info').PadRight(9)) " 'Cyan' -Bold)   -NoNewline; Write-Host (Bar $info.Count 'Cyan')
    Write-Host ''

    if ($crit.Count -gt 0 -or $warn.Count -gt 0) {
        Write-Host (Ink "   $(T 'summary.priorityActions')" 'Orange' -Bold)
        foreach ($f in ($crit + $warn)) {
            $c = if ($f.Severity -eq 'CRIT') { 'Red' } else { 'Yellow' }
            Write-Host (Ink "     [$($f.Severity)] " $c) -NoNewline
            Write-Host (Ink $f.Title 'White')
            if ($f.Recommendation) { Write-Host (Ink "            $($f.Recommendation)" 'Gray') }
        }
    } else {
        Write-Host (Ink "   $(T 'summary.allHealthy')" 'Green' -Bold)
    }
    Write-Host ''
}

# =====================================================================================
#  EXPORTAR INFORME HTML (ciberpunk)
# =====================================================================================
function Export-Report {
    param([string]$Path)
    try { Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue } catch {}
    if (-not $Path) {
        $safe = ($Script:SessionInfo.Mailbox -replace '[^\w.@-]','_')
        if (-not $safe) { $safe = $env:COMPUTERNAME }
        $Path = Join-Path (Get-Location) "OutlookDiag_${safe}_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    }
    $rows = foreach ($f in $Script:Findings) {
        $cls = switch ($f.Severity) { 'CRIT' {'crit'} 'WARN' {'warn'} 'OK' {'ok'} default {'info'} }
        $det = [System.Web.HttpUtility]::HtmlEncode($f.Detail) -replace "`n",'<br>'
        $rec = if ($f.Recommendation) { "<div class='rec'>▶ " + [System.Web.HttpUtility]::HtmlEncode($f.Recommendation) + "</div>" } else { '' }
        @"
<tr class="$cls">
  <td class="sev">$($f.Severity)</td>
  <td class="cat">$([System.Web.HttpUtility]::HtmlEncode($f.Category))</td>
  <td><div class="title">$([System.Web.HttpUtility]::HtmlEncode($f.Title))</div><div class="det">$det</div>$rec</td>
</tr>
"@
    }
    $c = @($Script:Findings | Where-Object Severity -eq 'CRIT').Count
    $w = @($Script:Findings | Where-Object Severity -eq 'WARN').Count
    $o = @($Script:Findings | Where-Object Severity -eq 'OK').Count
    $i = @($Script:Findings | Where-Object Severity -eq 'INFO').Count
    $htmlLang = if ($Script:Lang -eq 'pt') { 'pt-BR' } else { $Script:Lang }

    $html = @"
<!DOCTYPE html><html lang="$htmlLang"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>OUTLOOK :: $(T 'report.title')</title>
<style>
  :root{--bg:#0a0a12;--panel:#12121f;--cyan:#00ffff;--pink:#ff0099;--green:#39ff14;--yellow:#ffd600;--red:#ff2d55;--gray:#8a8aa0;}
  *{box-sizing:border-box;}
  body{margin:0;background:radial-gradient(circle at 20% 0%,#160c22 0,var(--bg) 55%);color:#e6e6f0;font-family:'Consolas','Cascadia Code',monospace;padding:28px;}
  h1{font-size:28px;letter-spacing:3px;margin:0;color:var(--cyan);text-shadow:0 0 12px rgba(0,255,255,.6);}
  .sub{color:var(--pink);letter-spacing:2px;margin:4px 0 18px;text-shadow:0 0 8px rgba(255,0,153,.5);}
  .meta{color:var(--gray);font-size:13px;margin-bottom:22px;border-left:3px solid var(--pink);padding-left:12px;}
  .cards{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:24px;}
  .card{flex:1;min-width:120px;background:var(--panel);border:1px solid #2a2a45;border-radius:10px;padding:14px 18px;text-align:center;}
  .card .n{font-size:34px;font-weight:bold;}
  .card .l{font-size:12px;letter-spacing:2px;color:var(--gray);}
  .card.c{border-color:var(--red);box-shadow:0 0 18px rgba(255,45,85,.25);} .card.c .n{color:var(--red);text-shadow:0 0 10px rgba(255,45,85,.7);}
  .card.w{border-color:var(--yellow);} .card.w .n{color:var(--yellow);text-shadow:0 0 10px rgba(255,214,0,.6);}
  .card.o{border-color:var(--green);} .card.o .n{color:var(--green);text-shadow:0 0 10px rgba(57,255,20,.6);}
  .card.ii{border-color:var(--cyan);} .card.ii .n{color:var(--cyan);text-shadow:0 0 10px rgba(0,255,255,.6);}
  table{width:100%;border-collapse:collapse;background:var(--panel);border:1px solid #2a2a45;border-radius:10px;overflow:hidden;}
  th{background:#1a1a2e;color:var(--cyan);text-align:left;padding:10px 14px;letter-spacing:2px;font-size:12px;border-bottom:2px solid var(--pink);}
  td{padding:12px 14px;border-bottom:1px solid #22223a;vertical-align:top;font-size:13px;}
  td.sev{font-weight:bold;letter-spacing:1px;white-space:nowrap;}
  td.cat{color:var(--gray);white-space:nowrap;}
  tr.crit td.sev{color:var(--red);} tr.crit{background:rgba(255,45,85,.06);}
  tr.warn td.sev{color:var(--yellow);}
  tr.ok td.sev{color:var(--green);}
  tr.info td.sev{color:var(--cyan);}
  .title{font-weight:bold;color:#fff;}
  .det{color:var(--gray);margin-top:4px;white-space:pre-wrap;}
  .rec{margin-top:6px;color:var(--yellow);border-left:2px solid var(--yellow);padding-left:8px;}
  footer{margin-top:22px;color:var(--gray);font-size:11px;text-align:center;letter-spacing:2px;}
</style></head><body>
<h1>OUTLOOK</h1>
<div class="sub">:: $((T 'report.title').ToUpper()) ::</div>
<div class="meta">
  $(T 'report.mailboxLabel'): <b>$([System.Web.HttpUtility]::HtmlEncode("$($Script:SessionInfo.Mailbox)"))</b> &nbsp;|&nbsp;
  $(T 'report.computerLabel'): <b>$($Script:SessionInfo.Host)</b> &nbsp;|&nbsp;
  $(T 'report.analystLabel'): <b>$($Script:SessionInfo.User)</b> &nbsp;|&nbsp;
  $(T 'report.dateLabel'): <b>$(Get-Date -Format 'yyyy-MM-dd HH:mm')</b>
</div>
<div class="cards">
  <div class="card c"><div class="n">$c</div><div class="l">$(T 'summary.critical')</div></div>
  <div class="card w"><div class="n">$w</div><div class="l">$(T 'summary.warning')</div></div>
  <div class="card o"><div class="n">$o</div><div class="l">$(T 'summary.ok')</div></div>
  <div class="card ii"><div class="n">$i</div><div class="l">$(T 'summary.info')</div></div>
</div>
<table>
  <thead><tr><th>$(T 'report.headerSev')</th><th>$(T 'report.headerArea')</th><th>$(T 'report.headerFinding')</th></tr></thead>
  <tbody>
$($rows -join "`n")
  </tbody>
</table>
<footer>$(T 'report.footer') OUTLOOK $(T 'banner.tagline') · $(Get-Date -Format o)</footer>
</body></html>
"@
    try {
        $html | Out-File -FilePath $Path -Encoding utf8
        Write-Host (Ink "  [+] Informe HTML generado: $Path" 'Green')
        return $Path
    } catch {
        Write-Host (Ink "  [x] No se pudo escribir el informe: $($_.Exception.Message)" 'Red')
    }
}

# =====================================================================================
#  ORQUESTACION
# =====================================================================================
function Invoke-FullDiagnostic {
    param([string]$Mbx, [switch]$SkipRemote, [switch]$SkipLocal)
    $Script:Findings.Clear()
    if (-not $SkipRemote) {
        if (-not $Mbx) { $Mbx = Read-Prompt (T 'prompt.mailboxToDiagnose') }
        if ($Mbx) { Invoke-RemoteDiagnostics -Mbx $Mbx }
    }
    if (-not $SkipLocal) { Invoke-LocalDiagnostics }
    Show-Summary
}

# =====================================================================================
#  REMEDIACION  ::  acciones correctivas (con confirmacion + auditoria)
# =====================================================================================
$Script:AuditPath = Join-Path (Get-Location) "OutlookDiag_Remediation_$(Get-Date -Format 'yyyyMMdd').log"

function Write-AuditLog {
    param([string]$Text)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  [$($env:USERNAME)@$($env:COMPUTERNAME)]  $Text"
    try { Add-Content -Path $Script:AuditPath -Value $line -Encoding utf8 -ErrorAction SilentlyContinue } catch {}
}

function Confirm-Typed {
    # Requiere que el usuario escriba EXACTAMENTE la frase esperada. Devuelve $true/$false.
    param([string]$Warning, [string]$Phrase)
    Write-Host ''
    Write-Host (Ink '   ⚠  ' 'Yellow' -Bold) -NoNewline
    Write-Host (Ink $Warning 'Yellow')
    Write-Host (Ink "   $(T 'confirm.toConfirmType') " 'Gray') -NoNewline
    Write-Host (Ink $Phrase 'Pink' -Bold)
    Write-Host (Ink '   └─▶ ' 'Purple') -NoNewline
    $ans = Read-Host
    if ($ans -eq $Phrase) { return $true }
    Write-Host (Ink "   [·] $(T 'confirm.cancelled')" 'Gray')
    return $false
}

function Assert-RemediationReady {
    param([string]$Mbx)
    if (-not $Mbx) { Write-Host (Ink "   [x] $(T 'assert.noMailbox')" 'Red'); return $false }
    if (-not (Test-EXOConnected)) {
        Write-Host (Ink "   [~] $(T 'assert.needConnection')" 'Cyan')
        if (-not (Connect-EXO)) { return $false }
    }
    return $true
}

# --- Verificacion de permisos/RBAC: los cmdlets de escritura solo se cargan en la ----
# --- sesion si tu cuenta tiene el rol adecuado. Si faltan, Enable-Mailbox (u otro) ---
# --- da "no se reconoce" aunque la conexion este bien (los Get-* si funcionan). -----
function Test-RequiredCmdlets {
    param([string[]]$Names)
    return @($Names | Where-Object { -not (Get-Command -Name $_ -ErrorAction SilentlyContinue) })
}

function Show-CmdletGapWarning {
    param([string[]]$Missing, [string]$RoleHint, [string]$ManualAlternative = '')
    Write-Host ''
    Write-Host (Ink "   [x] $(T 'gap.noAccessTitle')" 'Red' -Bold)
    foreach ($m in $Missing) { Write-Host (Ink "       · $m" 'Yellow') }
    Write-Host ''
    Write-Host (Ink "   $(T 'gap.rbacExplain1')" 'Gray')
    Write-Host (Ink "   $(T 'gap.rbacExplain2')" 'Gray')
    Write-Host (Ink "   $(T 'gap.rbacExplain3')" 'Gray')
    Write-Host (Ink "   $(T 'gap.rbacExplain4')" 'Gray')
    Write-Host ''
    Write-Host (Ink "   $(T 'gap.roleLabel') " 'Cyan') -NoNewline; Write-Host (Ink $RoleHint 'White' -Bold)
    Write-Host (Ink "   $(T 'gap.whereLabel') " 'Cyan') -NoNewline; Write-Host (Ink (T 'gap.whereValue') 'White')
    if ($ManualAlternative) {
        Write-Host (Ink "   $(T 'gap.altLabel') " 'Cyan') -NoNewline; Write-Host (Ink $ManualAlternative 'White')
    }
    Write-AuditLog "Cmdlet(s) no disponibles en sesion: $($Missing -join ', ')  ::  posible falta de rol RBAC ($RoleHint)"
}

# --- Conexion a Security & Compliance (para el borrado por cumplimiento) -------------
function Connect-SCC {
    if (Get-Command Get-ComplianceSearch -ErrorAction SilentlyContinue) {
        try { Get-ComplianceSearch -ErrorAction Stop | Out-Null; return $true } catch {}
    }
    if (-not (Get-Command Connect-IPPSSession -ErrorAction SilentlyContinue)) {
        Write-Host (Ink "   [x] $(T 'scc.moduleMissing')" 'Red')
        return $false
    }
    try {
        Write-Host (Ink "   [~] $(T 'scc.connecting')" 'Cyan')
        Connect-IPPSSession -ErrorAction Stop
        # Forzar la importacion del modulo para que carguen todos los cmdlets de Compliance
        Import-Module ExchangeOnlineManagement -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Host (Ink "   [+] $(T 'scc.connected')" 'Green')
        return $true
    } catch {
        Write-Host (Ink "   [x] $(T 'scc.connectError' @($_.Exception.Message))" 'Red')
        return $false
    }
}

# --- ACCION 1: Activar archivo en linea ---------------------------------------------
function Enable-MailboxArchive {
    param([string]$Mbx)
    Write-Section (T 'section.remediate1' @($Mbx)) 'Blue'
    if (-not (Assert-RemediationReady -Mbx $Mbx)) { return }
    $missing = Test-RequiredCmdlets -Names 'Enable-Mailbox'
    if ($missing.Count -gt 0) {
        Show-CmdletGapWarning -Missing $missing -RoleHint (T 'role.recipientMgmt') -ManualAlternative (T 'archive1.manualAlt')
        return
    }
    try {
        $mb = Get-Mailbox -Identity $Mbx -ErrorAction Stop
        $hasArchive = ($mb.ArchiveStatus -eq 'Active') -or ($mb.ArchiveGuid -and $mb.ArchiveGuid -ne '00000000-0000-0000-0000-000000000000')
        if ($hasArchive) {
            Write-Host (Ink "   [~] $(T 'archive1.alreadyHas')" 'Green')
        } else {
            if (-not (Confirm-Typed (T 'archive1.confirmWarn' @($Mbx)) (T 'confirm.word.activate'))) { return }
            Enable-Mailbox -Identity $Mbx -Archive -ErrorAction Stop | Out-Null
            Write-AuditLog "Enable-Mailbox -Archive  =>  $Mbx"
            Write-Host (Ink "   [+] $(T 'archive1.enabled')" 'Green')
        }
        # Auto-expanding archive (opcional)
        if (Read-YesNo 'archive1.askAutoExpand') {
            try {
                Enable-Mailbox -Identity $Mbx -AutoExpandingArchive -ErrorAction Stop | Out-Null
                Write-AuditLog "Enable-Mailbox -AutoExpandingArchive  =>  $Mbx"
                Write-Host (Ink "   [+] $(T 'archive1.autoExpandEnabled')" 'Green')
            } catch {
                Write-Host (Ink "   [!] $(T 'archive1.autoExpandError' @($_.Exception.Message))" 'Yellow')
            }
        }
        Write-Host (Ink "   $(T 'archive1.provisionNote')" 'Gray')
    } catch {
        Write-Host (Ink "   [x] $(T 'archive1.error' @($_.Exception.Message))" 'Red')
    }
}

# --- ACCION 2: Mover items antiguos al archivo (politica + MFA) ----------------------
function Move-OldItemsToArchive {
    param([string]$Mbx)
    Write-Section (T 'section.remediate2' @($Mbx)) 'Purple'
    if (-not (Assert-RemediationReady -Mbx $Mbx)) { return }
    $missing = Test-RequiredCmdlets -Names 'Set-Mailbox','Start-ManagedFolderAssistant'
    if ($missing.Count -gt 0) {
        Show-CmdletGapWarning -Missing $missing -RoleHint (T 'role.recipientMgmt') -ManualAlternative (T 'move2.manualAlt')
        return
    }
    try {
        $mb = Get-Mailbox -Identity $Mbx -ErrorAction Stop
        $hasArchive = ($mb.ArchiveStatus -eq 'Active') -or ($mb.ArchiveGuid -and $mb.ArchiveGuid -ne '00000000-0000-0000-0000-000000000000')
        if (-not $hasArchive) {
            Write-Host (Ink "   [!] $(T 'move2.needsArchive')" 'Yellow')
            if (Read-YesNo 'move2.askEnableNow') { Enable-MailboxArchive -Mbx $Mbx } else { return }
        }

        Write-Host (Ink "   $(T 'move2.explain1')" 'Gray')
        Write-Host (Ink "   $(T 'move2.explain2')" 'Gray')
        Write-Host ''
        $pol = $mb.RetentionPolicy
        Write-Host (Ink "   $(T 'move2.currentPolicyLabel') " 'Gray') -NoNewline
        Write-Host (Ink ($(if($pol){$pol}else{T 'move2.none'})) 'Cyan')

        if (-not $pol) {
            $avail = @(Get-RetentionPolicy -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
            $default = if ($avail -contains 'Default MRM Policy') { 'Default MRM Policy' } else { $avail | Select-Object -First 1 }
            if (-not $default) {
                Write-Host (Ink "   [x] $(T 'move2.noPolicyTitle')" 'Red')
                return
            }
            if (Confirm-Typed (T 'move2.confirmApply' @($default, $Mbx)) (T 'confirm.word.apply')) {
                Set-Mailbox -Identity $Mbx -RetentionPolicy $default -ErrorAction Stop
                Write-AuditLog "Set-Mailbox -RetentionPolicy '$default'  =>  $Mbx"
                Write-Host (Ink "   [+] $(T 'move2.policyAssigned' @($default))" 'Green')
            } else { return }
        }

        Write-Host (Ink "   [~] $(T 'move2.runningMFA')" 'Cyan')
        Start-ManagedFolderAssistant -Identity $Mbx -ErrorAction Stop
        Write-AuditLog "Start-ManagedFolderAssistant  =>  $Mbx"
        Write-Host (Ink "   [+] $(T 'move2.mfaStarted')" 'Green')
        Write-Host (Ink "   $(T 'move2.suggestion')" 'Gray')
    } catch {
        Write-Host (Ink "   [x] $(T 'move2.error' @($_.Exception.Message))" 'Red')
    }
}

# --- ACCION 3: Borrado guiado por antiguedad (SOFT-DELETE recuperable) --------------
function Remove-OldMailboxItems {
    param([string]$Mbx)
    Write-Section (T 'section.remediate3' @($Mbx)) 'Red'
    if (-not (Assert-RemediationReady -Mbx $Mbx)) { return }

    Write-Host (Ink "   $(T 'del3.explain1')" 'Gray')
    Write-Host (Ink "   $(T 'del3.explain2')" 'Gray')
    Write-Host (Ink "   $(T 'del3.explain3')" 'Gray')
    Write-Host (Ink "   $(T 'del3.note1')" 'Yellow')
    Write-Host (Ink "   $(T 'del3.note2')" 'Yellow')
    Write-Host ''

    if (-not (Connect-SCC)) { return }
    $missing = Test-RequiredCmdlets -Names 'New-ComplianceSearch','Start-ComplianceSearch','New-ComplianceSearchAction'
    if ($missing.Count -gt 0) {
        Show-CmdletGapWarning -Missing $missing -RoleHint (T 'role.compliance') -ManualAlternative (T 'del3.manualAlt')
        return
    }

    # Fecha de corte
    $cutoffStr = Read-Prompt (T 'prompt.cutoffDate')
    try {
        $cutoff = [datetime]::ParseExact($cutoffStr, 'yyyy-MM-dd', $null)
    } catch {
        Write-Host (Ink "   [x] $(T 'del3.invalidDate' @($_.Exception.Message))" 'Red'); return
    }
    $kql = "received<=$($cutoff.ToString('yyyy-MM-dd'))"
    Write-Host (Ink "   $(T 'del3.kqlLabel') " 'Gray') -NoNewline; Write-Host (Ink $kql 'Cyan')

    $searchName = "OLNR-Purge-$(Get-Date -Format 'yyyyMMddHHmmss')"
    try {
        Write-Host (Ink "   [~] $(T 'del3.searching')" 'Cyan')
        New-ComplianceSearch -Name $searchName -ExchangeLocation $Mbx -ContentMatchQuery $kql -ErrorAction Stop | Out-Null
        Start-ComplianceSearch -Identity $searchName -ErrorAction Stop

        # Poll hasta completar
        $tries = 0
        do {
            Start-Sleep -Seconds 5
            $s = Get-ComplianceSearch -Identity $searchName -ErrorAction Stop
            Write-Host (Ink "   $(T 'del3.status' @($s.Status))" 'Dark')
            $tries++
        } while ($s.Status -notin 'Completed','Failed' -and $tries -lt 60)

        if ($s.Status -ne 'Completed') {
            Write-Host (Ink "   [x] $(T 'del3.searchIncomplete' @($s.Status))" 'Red')
            Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction SilentlyContinue
            return
        }

        $itemCount = [int]$s.Items
        Write-Host ''
        Write-Host (Ink "   ┌─ $(T 'del3.estimationHeader') ────────────────────────────" 'Yellow')
        Write-Host (Ink "   │ $(T 'del3.matchingItems') " 'Yellow') -NoNewline; Write-Host (Ink ('{0:N0}' -f $itemCount) 'White' -Bold)
        Write-Host (Ink "   │ $(T 'del3.estimatedSize') $($s.Size)" 'Yellow')
        Write-Host (Ink "   └───────────────────────────────────────────────────" 'Yellow')

        if ($itemCount -eq 0) {
            Write-Host (Ink "   [·] $(T 'del3.nothingToDelete')" 'Gray')
            Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction SilentlyContinue
            return
        }

        # Confirmacion doble: escribir la direccion del buzon
        if (-not (Confirm-Typed (T 'del3.confirmDelete' @($itemCount, $Mbx, $cutoff.ToString('yyyy-MM-dd'))) $Mbx)) {
            Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction SilentlyContinue
            return
        }

        $maxBatches = Read-Prompt (T 'prompt.maxBatches')
        if (-not ($maxBatches -as [int])) { $maxBatches = 20 } else { $maxBatches = [int]$maxBatches }

        Write-AuditLog "SOFT-DELETE inicio  =>  $Mbx  query='$kql'  estimado=$itemCount  maxLotes=$maxBatches"
        Write-Host (Ink "   [~] $(T 'del3.purging' @($maxBatches))" 'Cyan')

        $batch = 0
        do {
            $batch++
            try {
                New-ComplianceSearchAction -SearchName $searchName -Purge -PurgeType SoftDelete -Confirm:$false -ErrorAction Stop | Out-Null
            } catch {
                Write-Host (Ink "   [!] $(T 'del3.batchError' @($batch, $_.Exception.Message))" 'Yellow'); break
            }
            $actionName = "$searchName" + "_Purge"
            $atries = 0
            do {
                Start-Sleep -Seconds 5
                $a = Get-ComplianceSearchAction -Identity $actionName -ErrorAction SilentlyContinue
                $atries++
            } while ($a -and $a.Status -notin 'Completed','Failed' -and $atries -lt 40)
            Write-Host (Ink "   $(T 'del3.batchStatus' @($batch, $a.Status))" 'Dark')
            # Limpiar la accion para poder relanzar sobre la misma busqueda
            Remove-ComplianceSearchAction -Identity $actionName -Confirm:$false -ErrorAction SilentlyContinue
            # Re-estimar cuantos quedan
            Start-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue
            $rtries = 0
            do { Start-Sleep -Seconds 5; $s = Get-ComplianceSearch -Identity $searchName -ErrorAction SilentlyContinue; $rtries++ }
            while ($s -and $s.Status -notin 'Completed','Failed' -and $rtries -lt 40)
            $remaining = if ($s) { [int]$s.Items } else { 0 }
            Write-Host (Ink "       $(T 'del3.remainingEstimated' @($remaining))" 'Gray')
        } while ($remaining -gt 0 -and $batch -lt $maxBatches)

        Write-AuditLog "SOFT-DELETE fin  =>  $Mbx  lotesEjecutados=$batch  restantes=$remaining"
        Write-Host (Ink "   [+] $(T 'del3.finished' @($batch, $remaining))" 'Green')
        if ($remaining -gt 0) {
            Write-Host (Ink "   [i] $(T 'del3.stillRemaining')" 'Cyan')
        }
        Write-Host (Ink "   [i] $(T 'del3.recoverable')" 'Cyan')
        Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction SilentlyContinue
    } catch {
        Write-Host (Ink "   [x] $(T 'del3.error' @($_.Exception.Message))" 'Red')
        Remove-ComplianceSearch -Identity $searchName -Confirm:$false -ErrorAction SilentlyContinue
    }
}

# --- ACCION 4: Runbook de reparacion del cliente (guia, no ejecuta nada) -------------
function Show-ClientRepairRunbook {
    param([string]$Mbx)
    Write-Section (T 'section.remediate4') 'Green'
    Write-Host (Ink "   $(T 'runbook4.instructions')" 'Gray')
    Write-Host ''
    $steps = @(
        @((T 'runbook4.step1'), 'outlook.exe /safe'),
        @((T 'runbook4.step2'), 'outlook.exe /cleanviews'),
        @((T 'runbook4.step3'), 'outlook.exe /resetnavpane'),
        @((T 'runbook4.step4'), 'Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Outlook\*.ost" | Rename-Item -NewName { $_.Name + ".old" }'),
        @((T 'runbook4.step5'), (T 'runbook4.step5sub')),
        @((T 'runbook4.step6'), (T 'runbook4.step6sub'))
    )
    foreach ($st in $steps) {
        Write-Host (Ink "   ▶ $($st[0])" 'Cyan')
        Write-Host (Ink "     $($st[1])" 'Green')
        Write-Host ''
    }
    Write-Host (Ink "   $(T 'runbook4.reminder1')" 'Yellow')
    Write-Host (Ink "   $(T 'runbook4.reminder2')" 'Yellow')
    Write-AuditLog "Mostrado runbook de reparacion de cliente (buzon contexto: $Mbx)"
}

# --- Submenu de remediacion ---------------------------------------------------------
function Invoke-RemediationMenu {
    param([string]$Mbx)
    $back = $false
    while (-not $back) {
        Write-Section (T 'section.remediation' @($(if($Mbx){$Mbx}else{T 'menu.undefined'}))) 'Pink'
        $ropts = @(
            @('1',(T 'remMenu.opt1'),'Blue'),
            @('2',(T 'remMenu.opt2'),'Purple'),
            @('3',(T 'remMenu.opt3'),'Red'),
            @('4',(T 'remMenu.opt4'),'Green'),
            @('0',(T 'remMenu.opt0'),'Gray')
        )
        foreach ($o in $ropts) {
            Write-Host (Ink "   [$($o[0])] " $o[2] -Bold) -NoNewline
            Write-Host (Ink $o[1] 'White')
        }
        Write-Host (Ink "   $(T 'remMenu.auditLabel' @($Script:AuditPath))" 'Dark')
        $c = Read-Prompt (T 'prompt.remediationAction')
        switch ($c) {
            '1' { Enable-MailboxArchive -Mbx $Mbx }
            '2' { Move-OldItemsToArchive -Mbx $Mbx }
            '3' { Remove-OldMailboxItems -Mbx $Mbx }
            '4' { Show-ClientRepairRunbook -Mbx $Mbx }
            '0' { $back = $true }
            default { Write-Host (Ink "   $(T 'remMenu.invalidOption')" 'Red') }
        }
        if (-not $back) {
            Write-Host ''
            Write-Host (Ink "   $(T 'remMenu.pressEnter')" 'Gray') -NoNewline
            [void](Read-Host)
        }
    }
}

# =====================================================================================
#  MENU
# =====================================================================================
function Show-Menu {
    Write-Rule 'Purple'
    $opts = @(
        @('1',(T 'menu.opt1'),'Cyan'),
        @('2',(T 'menu.opt2'),'Pink'),
        @('3',(T 'menu.opt3'),'Green'),
        @('4',(T 'menu.opt4'),'Blue'),
        @('5',(T 'menu.opt5'),'Yellow'),
        @('6',(T 'menu.opt6'),'Purple'),
        @('7',(T 'menu.opt7'),'Gray'),
        @('8',(T 'menu.opt8'),'Orange'),
        @('9',(T 'menu.opt9'),'Purple'),
        @('0',(T 'menu.opt0'),'Gray')
    )
    foreach ($o in $opts) {
        Write-Host (Ink "   [$($o[0])] " $o[2] -Bold) -NoNewline
        Write-Host (Ink $o[1] 'White')
    }
    $mbxTxt = if ($Script:SessionInfo.Mailbox) { $Script:SessionInfo.Mailbox } else { T 'menu.undefined' }
    $connTxt = if (Test-EXOConnected) { T 'menu.connected' } else { T 'menu.disconnected' }
    Write-Host ''
    Write-Host (Ink "   $(T 'menu.mailboxLabel') " 'Dark') -NoNewline; Write-Host (Ink $mbxTxt 'Cyan') -NoNewline
    Write-Host (Ink "   ·   EXO: " 'Dark') -NoNewline
    Write-Host (Ink $connTxt $(if(Test-EXOConnected){'Green'}else{'Red'}))
    Write-Rule 'Purple'
}

# =====================================================================================
#  MAIN
# =====================================================================================
if (-not $Language -and -not $Auto) { Show-LanguagePicker }
Show-Banner

if ($Auto) {
    Invoke-FullDiagnostic -Mbx $Mailbox -SkipRemote:($LocalOnly) -SkipLocal:($RemoteOnly)
    if ($ReportPath -or -not $LocalOnly) { Export-Report -Path $ReportPath }
    return
}

if ($LocalOnly) { Invoke-LocalDiagnostics; Show-Summary; return }
if ($RemoteOnly) { Invoke-FullDiagnostic -Mbx $Mailbox -SkipLocal; return }

# Bucle interactivo
$run = $true
while ($run) {
    Show-Menu
    $choice = Read-Prompt (T 'prompt.selectOption')
    switch ($choice) {
        '1' { Invoke-FullDiagnostic -Mbx $Script:SessionInfo.Mailbox }
        '2' {
            $m = if ($Script:SessionInfo.Mailbox) { $Script:SessionInfo.Mailbox } else { Read-Prompt (T 'prompt.mailbox') }
            $Script:Findings.Clear(); if ($m) { Invoke-RemoteDiagnostics -Mbx $m }; Show-Summary
        }
        '3' { $Script:Findings.Clear(); Invoke-LocalDiagnostics; Show-Summary }
        '4' { Connect-EXO }
        '5' { Show-Summary }
        '6' { Export-Report }
        '7' { $Script:SessionInfo.Mailbox = Read-Prompt (T 'prompt.newMailbox') }
        '8' { Invoke-RemediationMenu -Mbx $Script:SessionInfo.Mailbox }
        '9' { Show-LanguagePicker; Show-Banner }
        '0' { $run = $false }
        default { Write-Host (Ink "   $(T 'main.invalidOption')" 'Red') }
    }
    if ($run -and $choice -in '1','2','3','5','8') {
        if ($choice -ne '8') {
            Write-Host ''
            Write-Host (Ink "   $(T 'main.pressEnterReturn')" 'Gray') -NoNewline
            [void](Read-Host)
        }
        Show-Banner
    }
}
Write-Host (Ink "  >> $(T 'main.sessionEnded')" 'Pink' -Bold)
