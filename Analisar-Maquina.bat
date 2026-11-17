<# :
@echo off
setlocal
title Analise da Maquina
rem ==================================================================
rem  ANALISE DA MAQUINA - baseado no artigo "Como analisar uma maquina"
rem  Este arquivo e .bat + PowerShell no mesmo arquivo. Basta dar dois
rem  cliques. Ele pede permissao de administrador sozinho.
rem ==================================================================
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Solicitando permissao de administrador...
  powershell -NoProfile -Command "try { Start-Process -FilePath '%~f0' -Verb RunAs -ErrorAction Stop; exit 0 } catch { exit 1 }"
  if not errorlevel 1 exit /b
  echo Permissao negada - executando com recursos limitados.
)
set "BAT_DIR=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ([IO.File]::ReadAllText('%~f0', [Text.Encoding]::UTF8))"
echo.
pause
exit /b
#>

# ====================================================================
#  PARTE POWERSHELL
# ====================================================================
$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference    = 'SilentlyContinue'
try { $Host.UI.RawUI.WindowTitle = 'Analise da Maquina' } catch {}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
$batDir  = $env:BAT_DIR
$auto    = ($env:AM_AUTO -eq '1')   # modo sem perguntas (uso em testes)
$dias    = 30
$desde   = (Get-Date).AddDays(-$dias)

# Referencias do artigo
$LIM_CPU = 65; $LIM_HDD = 55; $LIM_SSD = 46
$PING_GIGA = 2; $PING_FAST = 12

$L       = New-Object System.Collections.Generic.List[string]   # relatorio em texto
$Rel       = New-Object System.Collections.Generic.List[object]   # mesmo conteudo, estruturado (para o PDF)
$Alertas = New-Object System.Collections.Generic.List[object]

function Txt([string]$t = '', [string]$cor = 'Gray') { $L.Add($t); Write-Host $t -ForegroundColor $cor }
function W([string]$t = '', [string]$cor = 'Gray') { Txt $t $cor; $Rel.Add(@{ T = 'txt'; X = $t }) }
function Secao([string]$t) { Txt ''; Txt ('=' * 74) 'Cyan'; Txt ('  ' + $t.ToUpper()) 'Cyan'; Txt ('=' * 74) 'Cyan'; $Rel.Add(@{ T = 'sec'; X = $t }) }
function Ok([string]$t)    { Txt "  [OK]      $t" 'Green'; $Rel.Add(@{ T = 'ok'; X = $t }) }
function Info([string]$t)  { Txt "  [INFO]    $t" 'Gray';  $Rel.Add(@{ T = 'info'; X = $t }) }
function Alerta([string]$t, [string]$tipo) { Txt "  [ALERTA]  $t" 'Yellow'; $Rel.Add(@{ T = 'alerta'; X = $t }); $Alertas.Add([pscustomobject]@{ Tipo = $tipo; Texto = $t; Chave = (ChaveAlerta $t) }) }
function Item([string]$k, $v) { Txt ('  {0,-30}: {1}' -f $k, $v); $Rel.Add(@{ T = 'kv'; K = $k.Trim(); V = "$v"; Sub = $k.StartsWith(' ') }) }

$GRUPOS = [ordered]@{
    hw   = 'HARDWARE - indícios de falha física'
    sw   = 'WINDOWS / SOFTWARE'
    perf = 'DESEMPENHO / CONFIGURAÇÃO'
    net  = 'REDE'
}
$CONCL = [ordered]@{
    hw   = 'Há indícios de FALHA FÍSICA: orientar o cliente a procurar um técnico de hardware.'
    sw   = 'Problema de SOFTWARE: reinstalar/reconfigurar o programa, atualizar o Windows ou reparar o sistema (DISM /Online /Cleanup-Image /RestoreHealth + sfc /scannow).'
    perf = 'DESEMPENHO: avaliar se a configuração atende o uso diário (SSD, mais RAM), fechar programas pesados e reduzir os programas de inicialização.'
    net  = 'REDE: trocar cabo/porta do switch, rodar este .bat em outro PC para comparar e verificar o servidor.'
}
$CONCL_OK = 'Máquina sem problemas aparentes. Se o sintoma persistir, acompanhar em tempo real com o Gerenciador de Tarefas + HWMonitor no momento em que o problema ocorrer.'

# --------------------------------------------------------------------
#  EMPRESA (aparece nos relatorios em PDF) - a logo e o arquivo logo.png na pasta do .bat
# --------------------------------------------------------------------
$EMPRESA = 'GL Tech'
$SLOGAN  = 'Técnico de informática e suporte de TI'

# --------------------------------------------------------------------
#  INTERPRETACAO: areas, catalogo de problemas (linguagem simples) e nota
# --------------------------------------------------------------------
$AREAS = [ordered]@{
    estab   = @{ Nome = 'Estabilidade e memória';     Ok = 'Sem telas azuis, desligamentos inesperados ou erros de hardware.' }
    disco   = @{ Nome = 'Disco (HD / SSD)';           Ok = 'Disco saudável, sem erros e com espaço livre suficiente.' }
    temp    = @{ Nome = 'Processador e temperatura';  Ok = 'Processador sem superaquecimento e sem redução de velocidade.' }
    windows = @{ Nome = 'Windows e atualizações';     Ok = 'Windows com suporte, sem arquivos corrompidos e atualizado.' }
    desemp  = @{ Nome = 'Desempenho e configuração';  Ok = 'Processador, memória e disco dão conta do uso atual.' }
    rede    = @{ Nome = 'Rede e internet';            Ok = 'Rede e internet respondendo normalmente.' }
}
$NOTAS = @{}   # observacoes por area (ex.: item nao verificado sem administrador)
function Nota([string]$area, [string]$txt) { if (-not $NOTAS.ContainsKey($area)) { $NOTAS[$area] = @() }; $NOTAS[$area] += $txt }

$PRIO = @{
    1 = @{ Nome = 'Urgente';     Cls = 'p1'; Grupo = 'Fazer agora';               Desc = 'risco de parar de funcionar ou de perder dados' }
    2 = @{ Nome = 'Importante';  Cls = 'p2'; Grupo = 'Fazer nos próximos dias';   Desc = 'afeta o desempenho ou a segurança' }
    3 = @{ Nome = 'Recomendado'; Cls = 'p3'; Grupo = 'Quando for possível';       Desc = 'ajustes simples que deixam o computador melhor' }
}

function Mk($area, $prio, $titulo, $explica, $acao, $custo, $quem) { @{ Area = $area; Prio = $prio; Titulo = $titulo; Explica = $explica; Acao = $acao; Custo = $custo; Quem = $quem } }
$CAT = @{
    bsod           = Mk 'estab' 1 'O Windows travou com tela azul' 'A tela azul acontece quando o Windows encontra um erro grave e precisa reiniciar. As causas mais comuns são memória RAM com defeito, driver com problema ou arquivos do Windows corrompidos.' 'Testar a memória RAM (Win+R > mdsched.exe), atualizar os drivers e reparar o Windows (DISM + SFC). Se continuar, levar a um técnico de hardware.' 'Sem custo a médio' 'Técnico'
    kp41           = Mk 'estab' 2 'O computador desligou sozinho' 'O Windows registrou desligamentos que não foram feitos pelo menu Iniciar. Pode ser queda de energia, fonte com defeito, superaquecimento ou o computador travado e desligado no botão.' 'Ligar o computador em nobreak ou estabilizador, verificar a fonte e a temperatura e fazer limpeza interna. Orientar a desligar sempre pelo menu Iniciar.' 'Baixo a médio' 'Técnico / usuário'
    whea           = Mk 'estab' 1 'Erros de hardware registrados' 'O processador ou a placa-mãe avisaram o Windows sobre falhas de hardware. Isso costuma vir antes de travamentos e telas azuis.' 'Levar a um técnico de hardware para testar memória, processador e placa-mãe.' 'Médio' 'Técnico'
    memerro        = Mk 'estab' 1 'Memória RAM com defeito' 'O teste de memória do Windows encontrou erros. Memória com defeito causa travamentos, telas azuis e arquivos corrompidos.' 'Trocar o(s) pente(s) de memória com defeito.' 'Médio' 'Técnico'
    crash          = Mk 'windows' 3 'Programas fechando sozinhos' 'Alguns programas fecharam inesperadamente várias vezes no último mês.' 'Atualizar ou reinstalar os programas que mais travaram (lista nos detalhes técnicos).' 'Sem custo' 'Técnico'
    disco_falha    = Mk 'disco' 1 'Disco com sinais de defeito' 'O disco (HD ou SSD) é onde ficam o Windows e todos os arquivos. Os sinais encontrados indicam risco real de perder dados e de o computador parar de ligar.' 'Fazer BACKUP dos arquivos imediatamente e trocar o disco, de preferência por um SSD.' 'Médio' 'Técnico'
    disco_temp     = Mk 'disco' 2 'Disco esquentando acima do ideal' 'Calor em excesso diminui a vida útil do disco e pode causar falhas.' 'Melhorar a ventilação, limpar a poeira e verificar se as ventoinhas estão funcionando.' 'Baixo' 'Técnico'
    ssd_desgaste   = Mk 'disco' 2 'SSD perto do fim da vida útil' 'Todo SSD tem um limite de gravações, e este já usou quase todo esse limite.' 'Planejar a troca do SSD e manter o backup em dia.' 'Médio' 'Técnico'
    hdd_sistema    = Mk 'desemp' 2 'Windows instalado em HD (disco lento)' 'HDs mecânicos são muito mais lentos que SSDs. É a causa mais comum de lentidão para ligar o computador e abrir programas.' 'Trocar o HD por um SSD, copiando o sistema (clonagem). É a melhoria com o melhor custo-benefício.' 'Médio' 'Técnico'
    pouco_espaco   = Mk 'desemp' 2 'Pouco espaço livre no disco' 'Com o disco quase cheio o Windows fica lento e as atualizações falham.' 'Liberar espaço: Limpeza de Disco, apagar arquivos temporários e downloads antigos, mover fotos e vídeos para HD externo ou nuvem.' 'Sem custo' 'Usuário / técnico'
    ram_pouca      = Mk 'desemp' 2 'Pouca memória RAM' 'A memória RAM é onde os programas abertos trabalham. Com pouca memória, o computador fica lento ao usar vários programas ao mesmo tempo.' 'Instalar mais memória: no mínimo 8 GB, ou 16 GB para uso pesado. O técnico confirma o tipo de pente compatível.' 'Médio' 'Técnico'
    uptime         = Mk 'desemp' 3 'Computador sem reiniciar há muitos dias' 'Ficar muitos dias sem reiniciar acumula programas presos na memória e deixa atualizações pela metade.' 'Reiniciar o computador pelo menos uma vez por semana, ou desligar no fim do expediente.' 'Sem custo' 'Usuário'
    win10          = Mk 'windows' 1 'Windows sem suporte da Microsoft' 'O Windows 10 não recebe correções de segurança desde 14/10/2025. O computador fica mais exposto a vírus e invasões.' 'Atualizar para o Windows 11 (se o computador for compatível) ou planejar a troca do equipamento.' 'Médio a alto' 'Técnico'
    win11old       = Mk 'windows' 2 'Versão do Windows desatualizada' 'Esta versão do Windows 11 não recebe mais correções de segurança.' 'Atualizar o Windows 11 para a versão mais recente pelo Windows Update.' 'Sem custo' 'Técnico'
    cpu_temp       = Mk 'temp' 1 'Processador superaquecendo' 'Acima de 65 °C o processador reduz a velocidade para se proteger e pode até desligar o computador.' 'Fazer limpeza interna, trocar a pasta térmica e verificar as ventoinhas.' 'Baixo' 'Técnico'
    throttle       = Mk 'temp' 3 'Processador reduzindo a velocidade' 'O processador diminuiu a velocidade algumas vezes, normalmente por calor ou por economia de energia.' 'Conferir a temperatura no HWMonitor e o plano de energia; fazer limpeza interna se estiver quente.' 'Sem custo a baixo' 'Técnico'
    plano_energia  = Mk 'desemp' 3 'Modo de economia de energia ligado' 'Nesse modo o Windows limita o desempenho do computador para gastar menos energia.' 'Mudar o plano de energia para "Equilibrado" ou "Alto desempenho".' 'Sem custo' 'Técnico'
    win_corrompido = Mk 'windows' 2 'Arquivos do Windows corrompidos' 'Arquivos do próprio Windows estão danificados, o que causa erros, lentidão e travamentos.' 'Reparar o Windows (DISM /RestoreHealth e depois sfc /scannow). Se não resolver, reinstalar o Windows.' 'Sem custo' 'Técnico'
    reboot_pend    = Mk 'windows' 3 'Atualização esperando reinício' 'O Windows instalou atualizações que só terminam depois de reiniciar.' 'Reiniciar o computador.' 'Sem custo' 'Usuário'
    sem_update     = Mk 'windows' 2 'Windows sem atualizar há muito tempo' 'As atualizações trazem correções de segurança e de problemas. Ficar sem atualizar deixa o computador vulnerável.' 'Rodar o Windows Update e deixar as atualizações automáticas ligadas.' 'Sem custo' 'Técnico'
    update_pend    = Mk 'windows' 3 'Atualizações pendentes' 'Há atualizações disponíveis que ainda não foram instaladas. Algumas corrigem problemas e falhas de segurança.' 'Instalar as atualizações pelo Windows Update, fora do horário de uso.' 'Sem custo' 'Técnico'
    cpu_alta       = Mk 'desemp' 2 'Processador sobrecarregado' 'Durante a análise o processador trabalhou quase no limite, o que deixa tudo lento.' 'Ver nos detalhes técnicos quais programas mais usam o processador e fechar ou desinstalar o que não for necessário. Se for o uso normal, avaliar um computador mais potente.' 'Sem custo a alto' 'Técnico'
    proc_pesado    = Mk 'desemp' 3 'Programa usando muito o processador' 'Um único programa estava usando boa parte do processador.' 'Verificar se esse programa é necessário; se não for, fechar ou desinstalar. Se for, procurar atualização dele.' 'Sem custo' 'Técnico'
    mem_cheia      = Mk 'desemp' 2 'Memória quase cheia' 'Quase toda a memória RAM estava em uso. Quando ela acaba, o Windows passa a usar o disco, que é muito mais lento.' 'Fechar programas e abas que não estão em uso. Se acontecer sempre, aumentar a memória RAM.' 'Sem custo a médio' 'Usuário / técnico'
    disco_ocupado  = Mk 'desemp' 2 'Disco trabalhando no limite' 'O disco ficou ocupado quase o tempo todo, o que faz o computador travar.' 'Trocar HD por SSD e verificar a saúde do disco; ver quais programas usam muito o disco.' 'Médio' 'Técnico'
    inicializacao  = Mk 'desemp' 3 'Muitos programas abrindo junto com o Windows' 'Cada programa que abre sozinho deixa o computador mais lento para ligar e ocupa memória.' 'Desativar os desnecessários em Gerenciador de Tarefas > Aplicativos de inicialização.' 'Sem custo' 'Técnico'
    rede_sem_placa = Mk 'rede' 1 'Computador sem conexão de rede' 'Nenhuma placa de rede está conectada.' 'Verificar o cabo de rede, o Wi-Fi e o driver da placa de rede.' 'Sem custo a baixo' 'Técnico'
    rede_link      = Mk 'rede' 2 'Rede conectada em velocidade muito baixa' 'A placa de rede conectou em velocidade baixa, geralmente por cabo ruim ou porta do switch com defeito.' 'Trocar o cabo de rede e testar outra porta do switch.' 'Baixo' 'Técnico'
    rede_nome      = Mk 'rede' 2 'Servidor não encontrado pelo nome' 'O computador não conseguiu achar o endereço do equipamento pelo nome.' 'Verificar o nome do servidor e a configuração de DNS da rede.' 'Sem custo' 'Técnico'
    rede_instavel  = Mk 'rede' 1 'Comunicação com o servidor instável' 'Parte das informações enviadas ao servidor se perdeu. Isso faz o sistema travar, ficar lento ou perder a conexão.' 'Trocar o cabo, testar outra porta do switch e verificar o switch e a placa de rede. Evitar Wi-Fi para o sistema.' 'Baixo a médio' 'Técnico'
    rede_lenta     = Mk 'rede' 2 'Rede lenta até o servidor' 'O tempo de resposta até o servidor está acima do esperado para essa rede.' 'Verificar cabos e switch, usar cabo em vez de Wi-Fi e comparar com outro computador da rede.' 'Baixo' 'Técnico'
    rede_sem_resposta = Mk 'rede' 2 'Equipamento da rede não respondeu' 'Um equipamento testado não respondeu: pode estar desligado, fora da rede ou com firewall bloqueando o teste.' 'Confirmar se o equipamento está ligado e conectado e verificar o firewall dele.' 'Sem custo' 'Técnico'
    internet_off   = Mk 'rede' 1 'Sem acesso à internet' 'O computador não conseguiu acessar a internet.' 'Verificar o roteador/modem e os cabos; se continuar, acionar o provedor.' 'Sem custo' 'Técnico'
    internet_perda = Mk 'rede' 2 'Internet instável' 'Parte da comunicação com a internet se perdeu durante o teste.' 'Reiniciar o roteador/modem, verificar os cabos e, se continuar, acionar o provedor.' 'Sem custo' 'Técnico / provedor'
    dns            = Mk 'rede' 2 'Problema para encontrar sites (DNS)' 'O computador não consegue traduzir o nome dos sites em endereços.' 'Configurar um DNS confiável (ex.: 8.8.8.8 e 1.1.1.1) ou verificar o roteador.' 'Sem custo' 'Técnico'
}
# Qual item do catalogo corresponde a cada alerta (pelo texto do alerta, na ordem)
$PADROES = @(
    @('tela\(s\) azul', 'bsod'), @('Kernel-Power 41', 'kp41'), @('WHEA', 'whea'), @('Diagnóstico de Memória.*ERROS', 'memerro'),
    @('travamentos de programas', 'crash'), @('°C - acima do limite|°C \(limite', 'disco_temp'), @('vida útil consumida', 'ssd_desgaste'),
    @("com status '|leitura/escrita não corrigidos|PREVÊ FALHA|erro de disco no log|com saúde '", 'disco_falha'),
    @('HD mecânico', 'hdd_sistema'), @('pouco espaço livre', 'pouco_espaco'), @('GB de RAM', 'ram_pouca'), @('sem reiniciar', 'uptime'),
    @('Windows 10: sem suporte', 'win10'), @('Windows 11 versão .* fora do suporte', 'win11old'), @('^Temperatura de .* acima de', 'cpu_temp'),
    @('limitada pelo firmware', 'throttle'), @('Economia de energia', 'plano_energia'), @('Imagem do Windows com corrupção|SFC encontrou', 'win_corrompido'),
    @('aguardando REINICIALIZAÇÃO', 'reboot_pend'), @('Nenhuma atualização instalada há', 'sem_update'), @('pendente\(s\)', 'update_pend'),
    @('Processador com uso alto', 'cpu_alta'), @("^Processo '", 'proc_pesado'), @('Memória quase cheia', 'mem_cheia'), @('^Disco ocupado', 'disco_ocupado'),
    @('iniciando com o Windows', 'inicializacao'), @('Nenhuma placa de rede', 'rede_sem_placa'), @('conectada a apenas', 'rede_link'),
    @('resolver o nome', 'rede_nome'), @('de pacotes para a internet', 'internet_perda'), @('Sem acesso à internet', 'internet_off'),
    @('DNS não está resolvendo', 'dns'), @('conexões com .* falharam|dos pacotes para', 'rede_instavel'),
    @('lenta \(média|Ping mínimo para|Ping instável', 'rede_lenta'), @('^Sem resposta de', 'rede_sem_resposta')
)
function ChaveAlerta([string]$t) { foreach ($p in $PADROES) { if ($t -match $p[0]) { return $p[1] } }; 'outro' }

function CalcularMelhorias {
    $areaTipo = @{ hw = 'estab'; sw = 'windows'; perf = 'desemp'; net = 'rede' }
    $lista = foreach ($g in ($Alertas | Group-Object Chave)) {
        $c = $CAT[$g.Name]
        if (-not $c) { $c = Mk $areaTipo[$g.Group[0].Tipo] 2 'Outro ponto de atenção' $g.Group[0].Texto 'Avaliar com o técnico.' 'A definir' 'Técnico' }
        [pscustomobject]@{ Chave = $g.Name; Area = $c.Area; Prio = [int]$c.Prio; Titulo = $c.Titulo; Explica = $c.Explica; Acao = $c.Acao; Custo = $c.Custo; Quem = $c.Quem; Evidencias = @($g.Group | ForEach-Object { $_.Texto }) }
    }
    @($lista | Sort-Object Prio, Titulo)
}
function Pontuacao($mel) {
    $p = 100
    foreach ($m in $mel) { $p -= @{ 1 = 25; 2 = 10; 3 = 4 }[[int]$m.Prio] }
    if (@($mel | Where-Object { $_.Prio -eq 1 }).Count) { $p = [math]::Min($p, 64) }   # problema urgente: no maximo "Regular"
    [math]::Max(0, $p)
}
function Classe([int]$p) {
    if ($p -ge 85)     { @{ Nome = 'Ótimo';   Cls = 'otimo';   Frase = 'O computador está em ótimo estado. Nenhuma ação urgente.' } }
    elseif ($p -ge 70) { @{ Nome = 'Bom';     Cls = 'bom';     Frase = 'O computador está bem, com alguns pontos simples para melhorar.' } }
    elseif ($p -ge 50) { @{ Nome = 'Regular'; Cls = 'regular'; Frase = 'O computador precisa de atenção. Veja abaixo o que fazer.' } }
    else               { @{ Nome = 'Crítico'; Cls = 'critico'; Frase = 'O computador tem problemas sérios. Recomendamos agir o quanto antes.' } }
}
function StatusArea([string]$k) {
    $ps = @($MELHORIAS | Where-Object { $_.Area -eq $k } | ForEach-Object { [int]$_.Prio })
    if (-not $ps.Count) { 'ok' } elseif (($ps | Measure-Object -Minimum).Minimum -eq 1) { 'critico' } else { 'atencao' }
}
function StatusTexto([string]$s) { @{ ok = 'Tudo certo'; atencao = 'Atenção'; critico = 'Problema' }[$s] }
$Config = [ordered]@{}   # configuracao da maquina (para o relatorio de melhorias)
function Ev($log, $ids, $provRegex) {
    $h = @{ LogName = $log; StartTime = $desde }
    if ($ids) { $h.Id = $ids }
    $r = @(Get-WinEvent -FilterHashtable $h -ErrorAction SilentlyContinue)
    if ($provRegex) { $r = @($r | Where-Object { $_.ProviderName -match $provRegex }) }
    ,$r
}
function Curto([string]$s, [int]$n = 110) {
    $s = ($s -replace '\s+', ' ').Trim()
    if ($s.Length -gt $n) { $s.Substring(0, $n) + '...' } else { $s }
}
function Achar([string[]]$caminhos, [string[]]$nomes) {
    foreach ($c in $caminhos) { if ($c -and (Test-Path -LiteralPath $c)) { return $c } }
    foreach ($raiz in @($batDir, $pastaFerr)) {
        if (-not $raiz -or -not (Test-Path -LiteralPath $raiz)) { continue }
        foreach ($n in $nomes) {
            $f = Get-ChildItem -LiteralPath $raiz -Filter $n -Recurse -Depth 3 -File -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($f) { return $f.FullName }
        }
    }
    return $null
}

# --------------------------------------------------------------------
#  FERRAMENTAS (CrystalDiskInfo / HWMonitor) - download automatico
# --------------------------------------------------------------------
$cdiFixos = @("$env:ProgramFiles\CrystalDiskInfo\DiskInfo64.exe", "${env:ProgramFiles(x86)}\CrystalDiskInfo\DiskInfo64.exe", "${env:ProgramFiles(x86)}\CrystalDiskInfo\DiskInfo32.exe")
$cdiNomes = @('DiskInfo64.exe', 'DiskInfo32.exe')
$hwmFixos = @("$env:ProgramFiles\CPUID\HWMonitor\HWMonitor.exe", "${env:ProgramFiles(x86)}\CPUID\HWMonitor\HWMonitor.exe")
$hwmNomes = @('HWMonitor_x64.exe', 'HWMonitor.exe', 'HWMonitor_x32.exe')

# Pasta onde as ferramentas portateis ficam (ao lado do .bat; se nao der, ProgramData)
$pastaFerr = $null
foreach ($p in @($(if ($batDir) { Join-Path $batDir 'Ferramentas' }), (Join-Path $env:ProgramData 'AnaliseMaquina\Ferramentas'))) {
    if (-not $p) { continue }
    try {
        if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p -Force -ErrorAction Stop | Out-Null }
        $teste = Join-Path $p '.teste'; [IO.File]::WriteAllText($teste, 'x'); Remove-Item -LiteralPath $teste -Force
        $pastaFerr = $p; break
    } catch {}
}

$curlExe = "$env:SystemRoot\System32\curl.exe"
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
function BaixarTexto([string]$url, [string]$ua = 'Mozilla/5.0') {
    if (Test-Path $curlExe) { return ((& $curlExe -sL --max-time 30 -A $ua $url) -join "`n") }
    try { (Invoke-WebRequest $url -UseBasicParsing -TimeoutSec 30 -UserAgent $ua).Content } catch { '' }
}
function BaixarArquivo([string]$url, [string]$destino) {
    Remove-Item -LiteralPath $destino -Force -ErrorAction SilentlyContinue
    if (Test-Path $curlExe) { & $curlExe -sSL --max-time 300 -A 'Wget/1.21' -o $destino $url 2>$null }
    else { try { Invoke-WebRequest $url -OutFile $destino -UseBasicParsing -UserAgent 'Wget/1.21' -TimeoutSec 300 } catch {} }
    if (-not (Test-Path -LiteralPath $destino) -or (Get-Item -LiteralPath $destino).Length -lt 100000) { return $false }
    $fs = [IO.File]::OpenRead($destino); $b1 = $fs.ReadByte(); $b2 = $fs.ReadByte(); $fs.Close()
    return ($b1 -eq 0x50 -and $b2 -eq 0x4B)   # "PK" = arquivo zip valido
}
function InstalarWinget([string]$id) {
    if (-not $isAdmin) { return }   # sem admin o instalador abriria janela de permissao (UAC) no meio da analise
    $wg =(Get-Command winget.exe -ErrorAction SilentlyContinue).Source
    if (-not $wg) { $wg = Get-ChildItem "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe" -ErrorAction SilentlyContinue | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName }
    if (-not $wg) { return }
    & $wg install --id $id -e --silent --accept-source-agreements --accept-package-agreements --disable-interactivity *> $null
}
function ObterCrystalDiskInfo {
    if (Achar $cdiFixos $cdiNomes) { return }
    Write-Host '  Baixando CrystalDiskInfo (portátil, sourceforge.net)...' -ForegroundColor Cyan
    $rss = BaixarTexto 'https://sourceforge.net/projects/crystaldiskinfo/rss?path=/' 'Wget/1.21'   # SourceForge recusa UA de navegador
    $m = [regex]::Match($rss, 'https://sourceforge\.net/projects/crystaldiskinfo/files/[\d.]+/CrystalDiskInfo[\d_]+\.zip/download')
    if ($m.Success -and $pastaFerr) {
        $zip = Join-Path $pastaFerr 'CrystalDiskInfo.zip'
        if (BaixarArquivo $m.Value $zip) {
            Expand-Archive -LiteralPath $zip -DestinationPath (Join-Path $pastaFerr 'CrystalDiskInfo') -Force
            Remove-Item -LiteralPath $zip -Force
        }
    }
    if (-not (Achar $cdiFixos $cdiNomes)) {
        Write-Host '  Download direto falhou - tentando instalar pelo winget...' -ForegroundColor DarkGray
        InstalarWinget 'CrystalDewWorld.CrystalDiskInfo'
    }
    if (Achar $cdiFixos $cdiNomes) { Write-Host '  CrystalDiskInfo pronto.' -ForegroundColor Green } else { Write-Host '  Não foi possível obter o CrystalDiskInfo.' -ForegroundColor Red }
}
function ObterHWMonitor {
    if (Achar $hwmFixos $hwmNomes) { return }
    Write-Host '  Baixando HWMonitor (portátil, cpuid.com)...' -ForegroundColor Cyan
    $html = BaixarTexto 'https://www.cpuid.com/softwares/hwmonitor.html'
    $vers = @([regex]::Matches($html, 'hwmonitor_(\d+\.\d+)\.zip') | ForEach-Object { [version]$_.Groups[1].Value } | Sort-Object -Descending)
    if ($vers.Count -and $pastaFerr) {
        $zip = Join-Path $pastaFerr 'HWMonitor.zip'
        if (BaixarArquivo "https://download.cpuid.com/hwmonitor/hwmonitor_$($vers[0]).zip" $zip) {
            Expand-Archive -LiteralPath $zip -DestinationPath (Join-Path $pastaFerr 'HWMonitor') -Force
            Remove-Item -LiteralPath $zip -Force
        }
    }
    if (-not (Achar $hwmFixos $hwmNomes)) {
        Write-Host '  Download direto falhou - tentando instalar pelo winget...' -ForegroundColor DarkGray
        InstalarWinget 'CPUID.HWMonitor'
    }
    if (Achar $hwmFixos $hwmNomes) { Write-Host '  HWMonitor pronto.' -ForegroundColor Green } else { Write-Host '  Não foi possível obter o HWMonitor.' -ForegroundColor Red }
}

# --------------------------------------------------------------------
#  DESCOBERTA DA REDE - acha o servidor e os computadores sozinho
# --------------------------------------------------------------------
$PORTAS_DB = @{ 1433 = 'SQL Server'; 3050 = 'Firebird'; 5432 = 'PostgreSQL'; 3306 = 'MySQL'; 1521 = 'Oracle' }
$PRIVADO   = '^(10\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.)'

function NomesNetBIOS([string[]]$ips) {
    # Pergunta o nome do computador direto pela porta 137 (igual ao "nbtstat -A"), todos de uma vez
    $nomes = @{}
    if (-not $ips.Count) { return $nomes }
    $req = [byte[]](0x12, 0x34, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0x20, 0x43, 0x4B) + [byte[]](@(0x41) * 30) + [byte[]](0, 0, 0x21, 0, 1)
    $udp = New-Object Net.Sockets.UdpClient(0)
    $udp.Client.ReceiveTimeout = 300
    foreach ($ip in $ips) { try { [void]$udp.Send($req, $req.Length, $ip, 137) } catch {} }
    $fim = (Get-Date).AddSeconds(2)
    while ((Get-Date) -lt $fim) {
        try {
            $ep = New-Object Net.IPEndPoint([Net.IPAddress]::Any, 0)
            $d = $udp.Receive([ref]$ep)
        } catch { continue }
        if ($d.Length -lt 57) { continue }
        for ($k = 0; $k -lt $d[56]; $k++) {
            $o = 57 + $k * 18
            if ($o + 18 -gt $d.Length) { break }
            if ($d[$o + 15] -eq 0 -and ($d[$o + 16] -band 0x80) -eq 0) {
                $nomes[$ep.Address.ToString()] = [Text.Encoding]::ASCII.GetString($d, $o, 15).Trim(); break
            }
        }
    }
    $udp.Close()
    $nomes
}

function DescobrirRede {
    $r = [pscustomobject]@{ MeuIp = $null; Gw = $null; Base = $null; Hosts = @(); SouServidor = $false; Obs = @(); Sugestao = @() }
    $cfg = @(Get-NetIPConfiguration | Where-Object { $_.NetAdapter.Status -eq 'Up' -and $_.IPv4Address -and $_.IPv4DefaultGateway })[0]
    if (-not $cfg) { $cfg = @(Get-NetIPConfiguration | Where-Object { $_.NetAdapter.Status -eq 'Up' -and $_.IPv4Address })[0] }
    if (-not $cfg) { return $r }
    $r.MeuIp = @($cfg.IPv4Address)[0].IPAddress
    $r.Gw    = @($cfg.IPv4DefaultGateway.NextHop)[0]
    $r.Base  = ($r.MeuIp -split '\.')[0..2] -join '.'
    $Tab = @{}
    $novo = { param($ip) if (-not $Tab.ContainsKey($ip)) { $Tab[$ip] = [pscustomobject]@{ IP = $ip; Nome = ''; NB = $false; Ping = $false; Portas = @(); Conexoes = @(); Cliente = $false; Mapeado = ''; Pontos = 0; Tipo = '' } }; $Tab[$ip] }

    # 1) Varredura por ping em toda a faixa (x.x.x.1 a 254) - em paralelo
    $tarefas = @{}
    foreach ($i in 1..254) { $ip = "$($r.Base).$i"; if ($ip -ne $r.MeuIp) { $tarefas[$ip] = (New-Object Net.NetworkInformation.Ping).SendPingAsync($ip, 1000) } }
    try { [void][Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]@($tarefas.Values), 5000) } catch {}
    foreach ($ip in @($tarefas.Keys)) { $t = $tarefas[$ip]; if ($t.Status -eq 'RanToCompletion' -and "$($t.Result.Status)" -eq 'Success') { (& $novo $ip).Ping = $true } }

    # 2) Tabela ARP: pega quem bloqueia ping mas respondeu na rede
    Get-NetNeighbor -AddressFamily IPv4 -InterfaceIndex $cfg.InterfaceIndex | Where-Object {
        $_.IPAddress -like "$($r.Base).*" -and $_.IPAddress -ne $r.MeuIp -and $_.IPAddress -notmatch '\.(0|255)$' -and
        "$($_.State)" -in 'Reachable', 'Stale', 'Delay', 'Probe' -and $_.LinkLayerAddress -and $_.LinkLayerAddress -notmatch '^(00-?){5}00$|^(FF-?){5}FF$'
    } | ForEach-Object { $null = & $novo $_.IPAddress }

    # 3) Conexoes abertas desta maquina (ex.: sistema conectado no banco do servidor)
    $proc = @{}; Get-Process | ForEach-Object { $proc[$_.Id] = $_.ProcessName }
    $escuta = @(Get-NetTCPConnection -State Listen | ForEach-Object { [int]$_.LocalPort } | Select-Object -Unique)
    $bancoLocal = @($PORTAS_DB.Keys | Where-Object { $escuta -contains $_ } | ForEach-Object { $PORTAS_DB[$_] })
    if (-not $bancoLocal.Count -and (Get-Service 'MSSQL*' | Where-Object { $_.Status -eq 'Running' -and $_.Name -notmatch 'FDLauncher|Launchpad' })) { $bancoLocal = @('SQL Server') }
    Get-NetTCPConnection -State Established | Where-Object { $_.RemoteAddress -match $PRIVADO -and $_.RemoteAddress -ne $r.MeuIp } | ForEach-Object {
        $h = & $novo $_.RemoteAddress
        if ($escuta -contains [int]$_.LocalPort) {
            # conexao recebida: o outro equipamento e cliente desta maquina
            $h.Cliente = $true
            if ($PORTAS_DB.ContainsKey([int]$_.LocalPort)) { $r.SouServidor = $true }
        }
        else {
            $desc = '{0}:{1}' -f $proc[[int]$_.OwningProcess], $_.RemotePort
            if ($h.Conexoes -notcontains $desc) { $h.Conexoes += $desc }
        }
    }

    if ($r.SouServidor) { $r.Obs += "Esta máquina é o SERVIDOR ($($bancoLocal -join '/') com outros computadores conectados) - o teste será feito contra esses computadores." }
    elseif ($bancoLocal.Count) { $r.Obs += "Esta máquina também tem $($bancoLocal -join '/') instalado, mas nenhum outro computador está conectado nele agora." }

    # 4) Unidades de rede mapeadas (\\SERVIDOR\pasta)
    $maps = @(Get-SmbMapping | ForEach-Object { $_.RemotePath }) + @(Get-CimInstance Win32_MappedLogicalDisk | ForEach-Object { $_.ProviderName })
    foreach ($m in ($maps | Where-Object { $_ -match '^\\\\([^\\]+)\\' } | Select-Object -Unique)) {
        $srv = ($m -split '\\')[2]
        $ip = $srv
        if ($srv -notmatch '^\d+\.\d+\.\d+\.\d+$') { try { $ip = ([Net.Dns]::GetHostAddresses($srv) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).IPAddressToString } catch { $ip = $null } }
        if ($ip -and $ip -match $PRIVADO) { $h = & $novo $ip; $h.Mapeado = $m; if (-not $h.Nome) { $h.Nome = $srv } }
    }

    $ips = @($Tab.Keys)
    # 5) Portas abertas (banco de dados, compartilhamento, area de trabalho remota, impressora)
    $portas = @(1433, 3050, 5432, 3306, 445, 3389, 9100)
    $tcp = New-Object System.Collections.Generic.List[object]
    foreach ($ip in $ips) { foreach ($p in $portas) { $c = New-Object Net.Sockets.TcpClient; $tcp.Add(@($ip, $p, $c, $c.ConnectAsync($ip, $p))) } }
    if ($tcp.Count) { try { [void][Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]@($tcp | ForEach-Object { $_[3] }), 2500) } catch {} }
    foreach ($x in $tcp) { if ($x[3].Status -eq 'RanToCompletion' -and $x[2].Connected) { $Tab[$x[0]].Portas += $x[1] }; try { $x[2].Close() } catch {} }

    # 6) Nomes: NetBIOS (computadores Windows) e DNS reverso (demais)
    $nb = NomesNetBIOS $ips
    foreach ($ip in $nb.Keys) { if ($Tab.ContainsKey($ip)) { $Tab[$ip].NB = $true; if (-not $Tab[$ip].Nome) { $Tab[$ip].Nome = $nb[$ip] } } }
    $dns = @{}
    foreach ($ip in ($ips | Where-Object { -not $Tab[$_].Nome })) { $dns[$ip] = [Net.Dns]::GetHostEntryAsync($ip) }
    if ($dns.Count) { try { [void][Threading.Tasks.Task]::WaitAll([Threading.Tasks.Task[]]@($dns.Values), 3000) } catch {} }
    foreach ($ip in $dns.Keys) { $t = $dns[$ip]; if ($t.Status -eq 'RanToCompletion' -and $t.Result.HostName -and $t.Result.HostName -ne $ip) { $Tab[$ip].Nome = $t.Result.HostName } }

    # 7) Classificar cada equipamento
    foreach ($h in $Tab.Values) {
        $db    = @($h.Portas | Where-Object { $PORTAS_DB.ContainsKey($_) })
        $conDb = @($h.Conexoes | ForEach-Object { [int](($_ -split ':')[-1]) } | Where-Object { $PORTAS_DB.ContainsKey($_) })
        if ($conDb.Count) { $h.Pontos += 10 } elseif ($h.Conexoes.Count) { $h.Pontos += 3 }
        if ($db.Count) { $h.Pontos += 5 }
        if ($h.Mapeado) { $h.Pontos += 4 }
        $bancos = (@($db + $conDb) | Select-Object -Unique | ForEach-Object { $PORTAS_DB[$_] }) -join '/'
        $h.Tipo = if ($h.IP -eq $r.Gw) { 'Roteador (gateway)' }
                  elseif ($bancos) { "SERVIDOR ($bancos)" }
                  elseif ($h.Mapeado) { 'SERVIDOR (arquivos)' }
                  elseif ($h.Portas -contains 9100) { 'Impressora' }
                  elseif ($h.Cliente) { 'Computador (usa este servidor)' }
                  elseif ($h.NB -or $h.Portas -contains 445 -or $h.Portas -contains 3389) { 'Computador' }
                  else { 'Outro dispositivo' }
        if ($h.IP -eq $r.Gw) { $h.Pontos = 0 }
    }
    $r.Hosts = @($Tab.Values | Sort-Object @{ e = { $_.Pontos }; Descending = $true }, @{ e = { $_.Tipo -notlike 'Computador*' } }, @{ e = { [int](($_.IP -split '\.')[3]) } })

    # 8) Sugestao de alvos para o teste
    if ($r.SouServidor) {
        $r.Sugestao = @($r.Hosts | Where-Object { $_.Cliente -and $_.Ping } | Select-Object -First 3 | ForEach-Object { $_.IP })
        if (-not $r.Sugestao.Count) { $r.Sugestao = @($r.Hosts | Where-Object { $_.Tipo -like 'Computador*' -and $_.Ping } | Select-Object -First 3 | ForEach-Object { $_.IP }) }
    } else {
        $r.Sugestao = @($r.Hosts | Where-Object { $_.Pontos -ge 3 } | Select-Object -First 2 | ForEach-Object { $_.IP })
    }
    if (-not $r.Sugestao.Count -and $r.Gw) { $r.Sugestao = @($r.Gw) }
    $r
}

function LinhasRede($r) {
    $l = New-Object System.Collections.Generic.List[string]
    $l.Add(('  Esta máquina: {0}  IP {1}  gateway {2}' -f $env:COMPUTERNAME, $r.MeuIp, $r.Gw))
    foreach ($o in $r.Obs) { $l.Add("  >> $o") }
    if (-not $r.Hosts.Count) { $l.Add('  Nenhum outro equipamento encontrado na rede.'); return $l }
    $l.Add('')
    $l.Add(('  {0,3}  {1,-15}  {2,-22}  {3,-30}  {4}' -f 'Nº', 'IP', 'Nome', 'Tipo', 'Detalhes'))
    $l.Add(('  {0,3}  {1,-15}  {2,-22}  {3,-30}  {4}' -f '--', '--', '----', '----', '--------'))
    $n = 0
    foreach ($h in $r.Hosts) {
        $n++
        $det = @()
        if ($h.Conexoes.Count) { $det += 'conectado: ' + (($h.Conexoes | Select-Object -First 3) -join ', ') }
        if ($h.Cliente)        { $det += 'conectado nesta máquina' }
        if ($h.Mapeado)        { $det += "unidade $($h.Mapeado)" }
        if ($h.Portas.Count)   { $det += 'portas ' + ($h.Portas -join ',') }
        if (-not $h.Ping)      { $det += 'NÃO responde ping' }
        $nome = if ($h.Nome) { $h.Nome } else { '?' }
        if ($nome.Length -gt 22) { $nome = $nome.Substring(0, 22) }
        $l.Add(('  {0,3}  {1,-15}  {2,-22}  {3,-30}  {4}' -f $n, $h.IP, $nome, $h.Tipo, ($det -join ' | ')))
    }
    $l
}

function TesteTcp([string]$ip, [int]$porta, [int]$n) {
    $t = New-Object System.Collections.Generic.List[double]; $falha = 0
    for ($i = 0; $i -lt $n; $i++) {
        $c = New-Object Net.Sockets.TcpClient
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try { if ($c.ConnectAsync($ip, $porta).Wait(2000) -and $c.Connected) { $t.Add([math]::Round($sw.Elapsed.TotalMilliseconds, 1)) } else { $falha++ } } catch { $falha++ }
        $c.Close(); Start-Sleep -Milliseconds 150
    }
    $o = [pscustomobject]@{ Enviados = $n; Falhas = $falha; Min = $null; Med = $null; Max = $null }
    if ($t.Count) { $m = $t | Measure-Object -Minimum -Maximum -Average; $o.Min = $m.Minimum; $o.Max = $m.Maximum; $o.Med = [math]::Round($m.Average, 1) }
    $o
}
function Pergunta([string]$txt, [string]$envVar) {
    if ($auto) { return [Environment]::GetEnvironmentVariable($envVar) }
    Read-Host $txt
}
function TestePing([string]$alvo, [int]$n, [int]$bytes = 32) {
    $p = New-Object System.Net.NetworkInformation.Ping
    $buf = New-Object byte[] $bytes
    $t = New-Object System.Collections.Generic.List[double]; $perd = 0
    for ($i = 0; $i -lt $n; $i++) {
        try { $r = $p.Send($alvo, 2000, $buf); if ($r.Status -eq 'Success') { $t.Add($r.RoundtripTime) } else { $perd++ } } catch { $perd++ }
        Start-Sleep -Milliseconds 150
    }
    $o = [pscustomobject]@{ Enviados = $n; Perdidos = $perd; Perda = [math]::Round($perd / $n * 100, 1); Min = $null; Med = $null; Max = $null }
    if ($t.Count) { $m = $t | Measure-Object -Minimum -Maximum -Average; $o.Min = $m.Minimum; $o.Max = $m.Maximum; $o.Med = [math]::Round($m.Average, 1) }
    $o
}

# --------------------------------------------------------------------
#  PERGUNTAS INICIAIS
# --------------------------------------------------------------------
Clear-Host
Write-Host ''
Write-Host '  ANALISE DA MAQUINA - hardware, Windows, desempenho e rede' -ForegroundColor Cyan
Write-Host '  Baseado no artigo "Como analisar uma máquina" (base de conhecimento SkyTech)' -ForegroundColor DarkGray
Write-Host ''
if (-not $isAdmin) {
    Write-Host '  ATENÇÃO: rodando SEM administrador. Saúde/temperatura do disco, temperatura' -ForegroundColor Red
    Write-Host '  do processador e integridade do Windows ficarão incompletos.' -ForegroundColor Red
    Write-Host ''
}
$sintoma  = Pergunta '  Sintoma relatado pelo cliente (lentidão, travamento, erro...) [Enter p/ pular]' 'AM_SINTOMA'
$sfcTxt   = Pergunta '  Rodar verificação de arquivos do Windows (SFC - leva de 5 a 15 min)? (S/N) [N]' 'AM_SFC'
$rodarSfc = ($sfcTxt -match '^[sS]')

# Baixa CrystalDiskInfo / HWMonitor se nao estiverem na maquina nem na pasta do .bat
$faltaCdi = -not (Achar $cdiFixos $cdiNomes)
$faltaHwm = -not (Achar $hwmFixos $hwmNomes)
if (($faltaCdi -or $faltaHwm) -and $env:AM_BAIXAR -ne 'N') {
    Write-Host ''
    if ($faltaCdi) { ObterCrystalDiskInfo }
    if ($faltaHwm) { ObterHWMonitor }
}

# Descobre o servidor e os computadores da rede
Write-Host ''
Write-Host '  Procurando o servidor e os computadores da rede (10 a 20 segundos)...' -ForegroundColor Cyan
$rede = DescobrirRede
$gw = $rede.Gw
$linhasRede = LinhasRede $rede
Write-Host ''
$linhasRede | ForEach-Object { Write-Host $_ }
$sugNums = @($rede.Sugestao | ForEach-Object { $s = $_; $i = [array]::IndexOf(@($rede.Hosts | ForEach-Object { $_.IP }), $s); if ($i -ge 0) { $i + 1 } else { $s } }) -join ' '
Write-Host ''
$alvosTxt = Pergunta "  Números dos equipamentos para o teste de rede (ex: 1 3) ou IP/nome [Enter = $sugNums]" 'AM_ALVOS'
$alvos = @()
foreach ($tk in @($alvosTxt -split '[\s,;]+' | Where-Object { $_ })) {
    if ($tk -match '^\d{1,3}$' -and [int]$tk -ge 1 -and [int]$tk -le $rede.Hosts.Count) { $alvos += $rede.Hosts[[int]$tk - 1].IP }
    else { $alvos += $tk }
}
if (-not $alvos.Count) { $alvos = @($rede.Sugestao) }
$alvos = @($alvos | Select-Object -Unique)

Write-Host ''
Write-Host '  Analisando... não feche esta janela.' -ForegroundColor Cyan

# Consulta ao Windows Update roda em paralelo (pode demorar)
$jobWU = Start-Job -ScriptBlock {
    try {
        $s = New-Object -ComObject Microsoft.Update.Session
        $b = $s.CreateUpdateSearcher()
        $n = $b.GetTotalHistoryCount()
        if ($n -gt 0) {
            $h = $b.QueryHistory(0, [Math]::Min($n, 100)) | Where-Object { $_.ResultCode -eq 2 -and $_.Operation -eq 1 } | Sort-Object Date -Descending | Select-Object -First 1
            if ($h) { 'ULTIMA|' + $h.Date.ToLocalTime().ToString('dd/MM/yyyy') + '|' + [int]((Get-Date) - $h.Date).TotalDays + '|' + $h.Title }
        }
        $r = $b.Search("IsInstalled=0 and IsHidden=0")
        foreach ($u in $r.Updates) { 'PEND|' + $u.Title }
        'FIM'
    } catch { 'ERRO|' + $_.Exception.Message }
}

# --------------------------------------------------------------------
#  IDENTIFICACAO
# --------------------------------------------------------------------
Secao 'Identificação'
Item 'Computador'           $env:COMPUTERNAME
Item 'Usuário'              "$env:USERDOMAIN\$env:USERNAME"
Item 'Data da análise'      (Get-Date -Format 'dd/MM/yyyy HH:mm')
Item 'Executado como admin' $(if ($isAdmin) { 'Sim' } else { 'Não' })
if ($sintoma) { Item 'Sintoma relatado' $sintoma }

# --------------------------------------------------------------------
#  1. SISTEMA
# --------------------------------------------------------------------
Secao '1. Sistema e configuração da máquina'
$os  = Get-CimInstance Win32_OperatingSystem
$cs  = Get-CimInstance Win32_ComputerSystem
$cpu = @(Get-CimInstance Win32_Processor)[0]
$bb  = Get-CimInstance Win32_BaseBoard
$cv  = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$build = [int]$os.BuildNumber
$ramGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
$up = (Get-Date) - $os.LastBootUpTime
$pentes = @(Get-CimInstance Win32_PhysicalMemory)
$descPentes = ($pentes | ForEach-Object { '{0} GB {1} MHz' -f [math]::Round($_.Capacity / 1GB), $_.ConfiguredClockSpeed }) -join ' + '

Item 'Windows'            ("{0} {1} - versão {2} (build {3}.{4})" -f $os.Caption, $os.OSArchitecture, $cv.DisplayVersion, $os.BuildNumber, $cv.UBR)
Item 'Fabricante / modelo' "$($cs.Manufacturer) $($cs.Model)"
Item 'Placa-mãe'          "$($bb.Manufacturer) $($bb.Product)"
Item 'Processador'        ("{0} ({1} núcleos / {2} threads)" -f $cpu.Name.Trim(), $cpu.NumberOfCores, $cpu.NumberOfLogicalProcessors)
Item 'Memória RAM'        ("{0} GB  [{1} pente(s): {2}]" -f $ramGB, $pentes.Count, $descPentes)
Item 'Ligada há'          ("{0} dia(s), {1} h {2} min (desde {3})" -f $up.Days, $up.Hours, $up.Minutes, $os.LastBootUpTime.ToString('dd/MM/yyyy HH:mm'))
$Config['Computador']  = "$env:COMPUTERNAME - $($cs.Manufacturer) $($cs.Model)"
$Config['Windows']     = "$($os.Caption) - versão $($cv.DisplayVersion)"
$Config['Processador'] = "$($cpu.Name.Trim()) ($($cpu.NumberOfCores) núcleos)"
$Config['Memória RAM'] = "$ramGB GB ($descPentes)"

if ($ramGB -lt 8) { Alerta "Apenas $ramGB GB de RAM - pode ser insuficiente para o uso diário (recomendado 8 GB ou mais)." 'perf' }
else { Ok "Quantidade de RAM adequada ($ramGB GB)." }
if ($up.Days -ge 7) { Alerta "Máquina ligada há $($up.Days) dias sem reiniciar - reiniciar pode resolver lentidão e concluir atualizações." 'sw' }
if ($build -lt 22000) { Alerta 'Windows 10: sem suporte da Microsoft desde 14/10/2025 (sem correções de segurança).' 'sw' }
elseif ($build -lt 26100) { Alerta "Windows 11 versão $($cv.DisplayVersion) fora do suporte - atualizar para 24H2/25H2." 'sw' }
else { Ok "Versão do Windows suportada ($($cv.DisplayVersion))." }

# --------------------------------------------------------------------
#  2. LOGS DO WINDOWS (memoria / Windows corrompido)
# --------------------------------------------------------------------
Secao "2. Logs do Windows - memória, telas azuis e erros (últimos $dias dias)"
$bsod  = Ev 'System' 1001 'WER-SystemErrorReporting'
$kp41  = Ev 'System' 41   'Kernel-Power'
$whea  = Ev 'System' $null 'WHEA-Logger'
$memd  = Ev 'System' @(1101, 1102, 1201, 1202) 'MemoryDiagnostics-Results'
$dumps = @(Get-ChildItem "$env:SystemRoot\Minidump" -Filter *.dmp -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -gt $desde })

if ($bsod.Count -or $dumps.Count) {
    $qtd = [math]::Max($bsod.Count, $dumps.Count)
    $codigos = @($bsod | ForEach-Object { if ($_.Message -match '0x[0-9a-fA-F]{8}') { $Matches[0] } } | Select-Object -Unique) -join ', '
    $ult = if ($bsod.Count) { $bsod[0].TimeCreated.ToString('dd/MM/yyyy HH:mm') } else { $dumps[0].LastWriteTime.ToString('dd/MM/yyyy HH:mm') }
    Alerta "$qtd tela(s) azul(is) registrada(s), última em $ult. Códigos: $codigos - possível problema de MEMÓRIA ou WINDOWS CORROMPIDO." 'hw'
} else { Ok 'Nenhuma tela azul registrada.' }

if ($kp41.Count) { Alerta "$($kp41.Count) desligamento(s) inesperado(s) (Kernel-Power 41) - queda de energia, fonte, superaquecimento ou travamento total." 'hw' }
else { Ok 'Nenhum desligamento inesperado.' }

$wheaErr = @($whea | Where-Object { $_.Level -le 3 })
if ($wheaErr.Count) { Alerta "$($wheaErr.Count) erro(s) de hardware WHEA (processador / memória / barramento PCIe). Último: $(Curto $wheaErr[0].Message 90)" 'hw' }
else { Ok 'Nenhum erro de hardware (WHEA).' }

if ($memd.Count) {
    $m = $memd[0]
    if ($m.Id -in 1102, 1202) { Alerta "Diagnóstico de Memória do Windows ENCONTROU ERROS na RAM em $($m.TimeCreated.ToString('dd/MM/yyyy')) - trocar/testar os pentes." 'hw' }
    else { Ok "Diagnóstico de Memória do Windows sem erros em $($m.TimeCreated.ToString('dd/MM/yyyy'))." }
} else { Info 'Nenhum teste de memória recente. Para testar a RAM: Win+R > mdsched.exe (reinicia a máquina).' }

$crash = Ev 'Application' 1000 'Application Error'
if ($crash.Count) {
    W "  Programas que travaram/fecharam sozinhos ($($crash.Count) ocorrências):"
    $crash | Group-Object { $_.Properties[0].Value } | Sort-Object Count -Descending | Select-Object -First 6 | ForEach-Object { W ('      {0,4}x  {1}' -f $_.Count, $_.Name) }
    if ($crash.Count -ge 10) { Alerta "$($crash.Count) travamentos de programas no período - verificar/reinstalar os programas mais citados acima." 'sw' }
} else { Ok 'Nenhum programa travou (Application Error 1000).' }

$err7 = @(Get-WinEvent -FilterHashtable @{ LogName = 'System'; Level = 1, 2; StartTime = (Get-Date).AddDays(-7) } -ErrorAction SilentlyContinue)
W ''
Item 'Erros/críticos no log Sistema (7 dias)' $err7.Count
if ($err7.Count) {
    W '    Origens mais frequentes:'
    $err7 | Group-Object ProviderName | Sort-Object Count -Descending | Select-Object -First 8 | ForEach-Object { W ('      {0,4}x  {1}' -f $_.Count, $_.Name) }
    W '    Erros mais recentes:'
    $err7 | Select-Object -First 5 | ForEach-Object { W ('      {0}  [{1} {2}] {3}' -f $_.TimeCreated.ToString('dd/MM HH:mm'), $_.ProviderName, $_.Id, (Curto $_.Message 80)) }
}

# --------------------------------------------------------------------
#  3. DISCO
# --------------------------------------------------------------------
Secao '3. Disco (HD / SSD) - saúde, temperatura e espaço'
$discoSistema = (Get-Partition -DriveLetter C | Get-Disk).Number
$pds = @(Get-PhysicalDisk | Sort-Object { [int]$_.DeviceId })
if (-not $pds.Count) { Info 'Não foi possível listar os discos (requer administrador).' }
if (-not $isAdmin) { Nota 'disco' 'Saúde detalhada do disco não verificada (rode como administrador).' }
$discosResumo = @()
foreach ($d in $pds) {
    $mt = "$($d.MediaType)"; $bt = "$($d.BusType)"
    $tipo = if ($mt -in 'SSD', '4' -or $bt -in 'NVMe', '17') { 'SSD' } elseif ($mt -in 'HDD', '3') { 'HDD' } else { 'Desconhecido' }
    $discosResumo += ('{0} - {1} {2}, {3} GB{4}' -f $d.FriendlyName, $tipo, $bt, [math]::Round($d.Size / 1GB), $(if ([int]$d.DeviceId -eq $discoSistema) { ' (Windows)' } else { '' }))
    $hs = "$($d.HealthStatus)"
    $hsPt = switch ($hs) { 'Healthy' { 'Saudável' } '0' { 'Saudável' } 'Warning' { 'AVISO' } '1' { 'AVISO' } 'Unhealthy' { 'COM FALHA' } '2' { 'COM FALHA' } default { $hs } }
    W ''
    W ("  Disco {0}: {1}  -  {2} {3}  -  {4} GB{5}" -f $d.DeviceId, $d.FriendlyName, $tipo, $bt, [math]::Round($d.Size / 1GB), $(if ([int]$d.DeviceId -eq $discoSistema) { '  (disco do Windows)' } else { '' }))
    Item '  Saúde (Windows)' $hsPt
    if ($hsPt -ne 'Saudável') { Alerta "Disco $($d.FriendlyName) com status '$hsPt' - possível FALHA FÍSICA no HD/SSD." 'hw' }

    $rc = $d | Get-StorageReliabilityCounter
    if ($rc) {
        $lim = if ($tipo -eq 'SSD') { $LIM_SSD } else { $LIM_HDD }
        if ($rc.Temperature) {
            Item '  Temperatura' ("{0} °C (máx. registrada {1} °C, limite {2} °C)" -f $rc.Temperature, $rc.TemperatureMax, $lim)
            if ($rc.Temperature -gt $lim) { Alerta "Disco $($d.FriendlyName) a $($rc.Temperature) °C - acima do limite de $lim °C para $tipo." 'hw' }
        }
        if ($rc.Wear -ne $null -and $tipo -eq 'SSD') {
            Item '  Desgaste (vida útil usada)' "$($rc.Wear) %"
            if ($rc.Wear -ge 90) { Alerta "SSD $($d.FriendlyName) com $($rc.Wear)% da vida útil consumida - planejar troca." 'hw' }
        }
        if ($rc.PowerOnHours) { Item '  Horas ligado' $rc.PowerOnHours }
        $errs = [int]$rc.ReadErrorsUncorrected + [int]$rc.WriteErrorsUncorrected
        if ($rc.ReadErrorsUncorrected -ne $null) {
            Item '  Erros leitura/escrita não corrigidos' ("{0} / {1}" -f [int]$rc.ReadErrorsUncorrected, [int]$rc.WriteErrorsUncorrected)
            if ($errs -gt 0) { Alerta "Disco $($d.FriendlyName) com $errs erro(s) de leitura/escrita não corrigidos - possível FALHA FÍSICA." 'hw' }
        }
    }
    if ([int]$d.DeviceId -eq $discoSistema -and $tipo -eq 'HDD') { Alerta 'Windows instalado em HD mecânico - causa comum de lentidão. Recomendado trocar por SSD.' 'perf' }
}

if ($discosResumo.Count) { $Config['Discos'] = $discosResumo -join ' | ' }
$fp = @(Get-CimInstance -Namespace root/wmi -ClassName MSStorageDriver_FailurePredictStatus | Where-Object { $_.PredictFailure })
if ($fp.Count) { Alerta "O S.M.A.R.T. de $($fp.Count) disco(s) PREVÊ FALHA - fazer backup imediatamente e trocar o disco." 'hw' }

$evDisk = Ev 'System' @(7, 11, 51, 55, 129, 153) '^(disk|Ntfs|Microsoft-Windows-Ntfs|storahci|stornvme|iaStor.*)$'
if ($evDisk.Count) {
    $res = ($evDisk | Group-Object { "$($_.ProviderName) $($_.Id)" } | ForEach-Object { "$($_.Name) ($($_.Count)x)" }) -join ', '
    Alerta "$($evDisk.Count) evento(s) de erro de disco no log: $res - possível problema no HD/SSD ou cabo." 'hw'
} else { Ok 'Nenhum erro de disco no log do Windows.' }

W ''
W '  Espaço livre:'
foreach ($v in @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3')) {
    if (-not $v.Size) { continue }
    $livreGB = [math]::Round($v.FreeSpace / 1GB, 1); $pct = [math]::Round($v.FreeSpace / $v.Size * 100)
    W ('      {0}  {1,8} GB livres de {2,6} GB  ({3}% livre)' -f $v.DeviceID, $livreGB, [math]::Round($v.Size / 1GB), $pct)
    if ($v.DeviceID -eq 'C:') { $Config['Espaço livre no C:'] = '{0} GB livres de {1} GB ({2}%)' -f $livreGB, [math]::Round($v.Size / 1GB), $pct }
    if ($pct -lt 10 -or ($v.DeviceID -eq 'C:' -and $livreGB -lt 15)) { Alerta "Unidade $($v.DeviceID) com pouco espaço livre ($livreGB GB / $pct%) - causa lentidão e falha em atualizações." 'perf' }
}

# CrystalDiskInfo (se existir instalado ou na pasta do .bat)
$cdi = Achar $cdiFixos $cdiNomes
$cdiTexto = $null
W ''
if ($cdi -and $isAdmin) {
    W '  CrystalDiskInfo encontrado - gerando relatório S.M.A.R.T....'
    $txt = Join-Path (Split-Path $cdi) 'DiskInfo.txt'
    Remove-Item -LiteralPath $txt -Force -ErrorAction SilentlyContinue
    $pr = Start-Process -FilePath $cdi -ArgumentList '/CopyExit' -PassThru -WindowStyle Hidden
    if ($pr -and -not $pr.WaitForExit(90000)) { $pr.Kill() }
    if (Test-Path -LiteralPath $txt) {
        $cdiTexto = Get-Content -LiteralPath $txt
        $atual = $null; $discos = @()
        foreach ($ln in $cdiTexto) {
            if ($ln -match '^\s*(Model|Modelo)\s*:\s*(.+)$') { $atual = [ordered]@{ Modelo = $Matches[2].Trim(); Saude = ''; Temp = $null; SSD = $false }; $discos += $atual }
            elseif ($atual -and $ln -match '^\s*(Health Status|Status de Sa\S*|Estado de Sa\S*)\s*:\s*(.+)$') { $atual.Saude = $Matches[2].Trim() }
            elseif ($atual -and $ln -match '^\s*(Temperature|Temperatura)\s*:\s*(\d+)') { $atual.Temp = [int]$Matches[2] }
            elseif ($atual -and $ln -match '^\s*(Rotation Rate|Rota\S* )\S*.*:.*SSD') { $atual.SSD = $true }
            elseif ($atual -and $ln -match '^\s*Interface\s*:.*NVM') { $atual.SSD = $true }
        }
        foreach ($c in $discos) {
            $lim = if ($c.SSD) { $LIM_SSD } else { $LIM_HDD }
            W ("      {0}  -  Saúde: {1}  -  Temperatura: {2} °C" -f $c.Modelo, $c.Saude, $(if ($c.Temp -ne $null) { $c.Temp } else { '?' }))
            if ($c.Saude -match 'Caution|Bad|Cuidado|Ruim|Aten|Alerta|Cr.tico') { Alerta "CrystalDiskInfo: disco $($c.Modelo) com saúde '$($c.Saude)' - possível FALHA FÍSICA." 'hw' }
            if ($c.Temp -ne $null -and $c.Temp -gt $lim) { Alerta "CrystalDiskInfo: disco $($c.Modelo) a $($c.Temp) °C (limite $lim °C)." 'hw' }
        }
        Info 'Relatório completo do CrystalDiskInfo anexado no final deste arquivo.'
    } else { Info 'CrystalDiskInfo não gerou o relatório.' }
} elseif ($cdi) {
    Info 'CrystalDiskInfo encontrado, mas precisa de administrador para gerar o relatório.'
} else {
    Info 'CrystalDiskInfo não disponível (download recusado ou sem internet). Para análise S.M.A.R.T. completa, extraia a versão portátil (crystalmark.info) na pasta deste .bat.'
}

# --------------------------------------------------------------------
#  4. PROCESSADOR E TEMPERATURA
# --------------------------------------------------------------------
Secao '4. Processador e temperatura'
Item 'Clock atual / máximo' ("{0} / {1} MHz" -f $cpu.CurrentClockSpeed, $cpu.MaxClockSpeed)
$tz = @(Get-CimInstance -Namespace root/wmi -ClassName MSAcpi_ThermalZoneTemperature)
$temps = @($tz | ForEach-Object { [math]::Round($_.CurrentTemperature / 10 - 273.15, 1) } | Where-Object { $_ -gt 5 -and $_ -lt 120 })
if ($temps.Count) {
    $tmax = ($temps | Measure-Object -Maximum).Maximum
    Item 'Temperatura (sensor ACPI)' (($temps | ForEach-Object { "$_ °C" }) -join ' | ')
    if ($tmax -gt $LIM_CPU) { Alerta "Temperatura de $tmax °C - acima de $LIM_CPU °C. Verificar ventoinha, poeira e pasta térmica." 'hw' }
    else { Ok "Temperatura dentro do limite de $LIM_CPU °C (sensor ACPI é aproximado - confirmar no HWMonitor)." }
} else {
    Info 'O Windows não expõe a temperatura do processador nesta máquina - confira no HWMonitor (oferecido no final).'
    Nota 'temp' 'Temperatura não informada pelo Windows - conferir no HWMonitor.'
}
$thr = Ev 'System' 37 'Kernel-Processor-Power'
if ($thr.Count) { Alerta "Velocidade do processador foi limitada pelo firmware $($thr.Count) vez(es) - comum em superaquecimento ou modo de economia de energia." 'perf' }
else { Ok 'Nenhuma limitação de velocidade do processador registrada.' }
$plano = (powercfg /getactivescheme) -replace '^.*\((.+)\)\s*\*?\s*$', '$1'
if ($plano) {
    Item 'Plano de energia' $plano
    if ($plano -match 'Economia|Power saver') { Alerta 'Plano de energia em "Economia de energia" - deixa a máquina lenta. Mudar para Equilibrado/Alto desempenho.' 'perf' }
}

# --------------------------------------------------------------------
#  5. INTEGRIDADE DO WINDOWS
# --------------------------------------------------------------------
Secao '5. Integridade do sistema operacional'
if ($isAdmin) {
    W '  Verificando a imagem do Windows (DISM CheckHealth)...'
    $img = Repair-WindowsImage -Online -CheckHealth
    $st = "$($img.ImageHealthState)"
    if ($st -eq 'Healthy') { Ok 'Imagem do Windows sem corrupção registrada (DISM).' }
    elseif ($st) { Alerta "Imagem do Windows com corrupção ($st). Executar: DISM /Online /Cleanup-Image /RestoreHealth e depois sfc /scannow." 'sw' }
    else { Info 'Não foi possível consultar o DISM.' }

    if ($rodarSfc) {
        W '  Executando SFC /verifyonly (pode levar de 5 a 15 minutos, aguarde)...'
        $enc = [Console]::OutputEncoding
        [Console]::OutputEncoding = [Text.Encoding]::Unicode
        $sfcOut = ((& "$env:SystemRoot\System32\sfc.exe" /verifyonly) -join "`n") -replace "`0", ''
        [Console]::OutputEncoding = $enc
        if ($sfcOut -match 'não encontrou nenhuma viola|nao encontrou nenhuma viola|did not find any integrity') { Ok 'SFC: nenhum arquivo do Windows corrompido.' }
        elseif ($sfcOut -match 'viola|violation') { Alerta 'SFC encontrou arquivos do Windows CORROMPIDOS. Executar: DISM /Online /Cleanup-Image /RestoreHealth e depois sfc /scannow.' 'sw' }
        else { Info ('Resultado do SFC: ' + (Curto (($sfcOut -split "`n" | Where-Object { $_.Trim() }) | Select-Object -Last 1) 150)) }
    } else { Info 'SFC não executado (opção N). Para verificar manualmente: sfc /scannow' }
} else {
    Info 'Requer administrador - integridade do Windows não verificada.'
    Nota 'windows' 'Integridade dos arquivos do Windows não verificada (rode como administrador).'
}

# --------------------------------------------------------------------
#  6. ATUALIZACOES
# --------------------------------------------------------------------
Secao '6. Atualizações do Windows'
$reboot = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or
          (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')
if ($reboot) { Alerta 'Há atualizações aguardando REINICIALIZAÇÃO para serem concluídas.' 'sw' }
W '  Consultando o Windows Update (até 3 minutos)...'
$null = Wait-Job $jobWU -Timeout 180
if ($jobWU.State -eq 'Completed') {
    $wu = @(Receive-Job $jobWU)
    $ultima = $wu | Where-Object { $_ -like 'ULTIMA|*' } | Select-Object -First 1
    $pend   = @($wu | Where-Object { $_ -like 'PEND|*' } | ForEach-Object { $_.Substring(5) })
    $erro   = $wu | Where-Object { $_ -like 'ERRO|*' } | Select-Object -First 1
    if ($ultima) {
        $u = $ultima -split '\|', 4
        Item 'Última atualização instalada' "$($u[1]) ($($u[2]) dias atrás) - $(Curto $u[3] 70)"
        if ([int]$u[2] -gt 45) { Alerta "Nenhuma atualização instalada há $($u[2]) dias." 'sw' }
    }
    if ($erro) { Info ('Windows Update retornou erro: ' + $erro.Substring(5)) }
    elseif ($pend.Count) {
        Alerta "$($pend.Count) atualização(ões) pendente(s) - alguns problemas são corrigidos em versões mais recentes:" 'sw'
        $pend | Select-Object -First 15 | ForEach-Object { W "      - $(Curto $_ 100)" }
    } else { Ok 'Nenhuma atualização pendente.' }
} else {
    Info 'O Windows Update não respondeu a tempo. Verifique manualmente em Configurações > Windows Update.'
    Nota 'windows' 'Windows Update não respondeu a tempo - atualizações não verificadas.'
    Stop-Job $jobWU
}
Remove-Job $jobWU -Force

# --------------------------------------------------------------------
#  7. CONSUMO DE RECURSOS (Gerenciador de Tarefas)
# --------------------------------------------------------------------
Secao '7. Consumo de recursos (amostra de ~6 segundos)'
$nCpu = [Environment]::ProcessorCount
$p1 = @{}; Get-Process | ForEach-Object { if ($_.CPU -ne $null) { $p1[$_.Id] = $_.CPU } }
$sw = [Diagnostics.Stopwatch]::StartNew()
$cpuS = @(); $dskS = @()
for ($i = 0; $i -lt 6; $i++) {
    $cpuS += [double](Get-CimInstance Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'").PercentProcessorTime
    $idle = (Get-CimInstance Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'").PercentIdleTime
    if ($idle -ne $null) { $dskS += [math]::Max(0, 100 - [double]$idle) }
    Start-Sleep -Milliseconds 800
}
$seg = $sw.Elapsed.TotalSeconds
$procs = @(Get-Process)
$cpuMed = [math]::Round(($cpuS | Measure-Object -Average).Average)
$dskMed = if ($dskS.Count) { [math]::Round(($dskS | Measure-Object -Average).Average) } else { $null }
$os = Get-CimInstance Win32_OperatingSystem
$memUso = [math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize * 100)
$memUsoGB = [math]::Round(($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / 1MB, 1)

Item 'CPU (média)'          "$cpuMed %"
Item 'Memória em uso'       "$memUso %  ($memUsoGB de $ramGB GB)"
if ($dskMed -ne $null) { Item 'Disco ocupado (média)' "$dskMed %" }
Item 'Processos em execução' $procs.Count

if ($cpuMed -ge 80) { Alerta "Processador com uso alto ($cpuMed%) - veja os processos abaixo." 'perf' }
if ($memUso -ge 85) { Alerta "Memória quase cheia ($memUso%) - fechar programas ou aumentar a RAM." 'perf' }
if ($dskMed -ne $null -and $dskMed -ge 80) { Alerta "Disco ocupado $dskMed% do tempo - gargalo de disco (típico de HD mecânico ou disco com defeito)." 'perf' }
if ($cpuMed -lt 80 -and $memUso -lt 85 -and ($dskMed -eq $null -or $dskMed -lt 80)) { Ok 'Consumo de CPU, memória e disco normal no momento.' }

$topCpu = foreach ($p in $procs) {
    if ($p.Id -ne 0 -and $p.CPU -ne $null -and $p1.ContainsKey($p.Id)) {
        [pscustomobject]@{ Nome = $p.ProcessName; PID = $p.Id; CPU = [math]::Round(($p.CPU - $p1[$p.Id]) / $seg / $nCpu * 100, 1) }
    }
}
W ''
W '  Processos que mais usam CPU agora:'
$topCpu | Sort-Object CPU -Descending | Select-Object -First 8 | ForEach-Object {
    W ('      {0,6} %   {1} (PID {2})' -f $_.CPU, $_.Nome, $_.PID)
    if ($_.CPU -ge 50) { Alerta "Processo '$($_.Nome)' usando $($_.CPU)% da CPU." 'perf' }
}
W ''
W '  Programas que mais usam memória (somando todas as janelas/processos):'
$procs | Group-Object ProcessName | ForEach-Object {
    [pscustomobject]@{ Nome = $_.Name; Qtd = $_.Count; MB = [math]::Round((($_.Group | Measure-Object WorkingSet64 -Sum).Sum) / 1MB) }
} | Sort-Object MB -Descending | Select-Object -First 8 | ForEach-Object { W ('      {0,7} MB   {1}{2}' -f $_.MB, $_.Nome, $(if ($_.Qtd -gt 1) { " ($($_.Qtd) processos)" } else { '' })) }

$ini = @(Get-CimInstance Win32_StartupCommand)
W ''
Item 'Programas na inicialização' $ini.Count
$ini | Select-Object -First 20 | ForEach-Object { W "      - $($_.Name)" }
if ($ini.Count -gt 15) { Alerta "$($ini.Count) programas iniciando com o Windows - desativar os desnecessários (Gerenciador de Tarefas > Aplicativos de inicialização)." 'perf' }

# --------------------------------------------------------------------
#  8. REDE
# --------------------------------------------------------------------
Secao '8. Testes de rede (ping e tracert)'
$ads = @(Get-NetAdapter -Physical | Where-Object { $_.Status -eq 'Up' })
if (-not $ads.Count) { Alerta 'Nenhuma placa de rede física conectada.' 'net' }
foreach ($a in $ads) {
    $wifi = ($a.PhysicalMediaType -match '802\.11|Wireless') -or ($a.InterfaceDescription -match 'Wi-?Fi|Wireless|WLAN')
    Item "Placa: $($a.Name)" ("{0} - {1}{2}" -f $a.InterfaceDescription, $a.LinkSpeed, $(if ($wifi) { ' (Wi-Fi)' } else { '' }))
    if (-not $wifi -and $a.Speed -lt 100000000) { Alerta "Placa '$($a.Name)' conectada a apenas $($a.LinkSpeed) - verificar cabo/porta do switch." 'net' }
}
$cfgs = @(Get-NetIPConfiguration | Where-Object { $_.NetAdapter.Status -eq 'Up' -and $_.IPv4Address })
foreach ($c in $cfgs) {
    Item "IP ($($c.InterfaceAlias))" ("{0}  gateway {1}  DNS {2}" -f ($c.IPv4Address.IPAddress -join ', '), ($c.IPv4DefaultGateway.NextHop -join ', '), (($c.DNSServer | Where-Object { $_.AddressFamily -eq 2 }).ServerAddresses -join ', '))
}

W ''
W '  Equipamentos encontrados na rede (descoberta automática):'
foreach ($ln in $linhasRede) { W $ln }
$servDetect = @($rede.Hosts | Where-Object { $_.Tipo -like 'SERVIDOR*' })
if (-not $rede.SouServidor -and -not $servDetect.Count) { Info 'Nenhum servidor identificado automaticamente (sistema fechado ou banco em outra rede). Se souber o IP, informe-o na pergunta do início.' }
if ($gw -and $alvos -contains $gw) { Info "Testando o gateway ($gw): roteador/modem responde mais devagar que um servidor, os tempos servem só como referência." }

foreach ($alvoOrig in $alvos) {
    W ''
    $ip = $null
    try { $ip = ([Net.Dns]::GetHostAddresses($alvoOrig) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).IPAddressToString } catch {}
    if (-not $ip) { Alerta "Não foi possível resolver o nome '$alvoOrig' (DNS)." 'net'; continue }
    $hInfo = @($rede.Hosts | Where-Object { $_.IP -eq $ip })[0]
    $alvo = if ($hInfo -and $hInfo.Nome) { "$($hInfo.Nome) [$ip]" } elseif ($alvoOrig -ne $ip) { "$alvoOrig [$ip]" } else { $ip }
    if ($hInfo -and $hInfo.Tipo) { $alvo += " - $($hInfo.Tipo)" }
    $rota = Find-NetRoute -RemoteIPAddress $ip | Select-Object -First 1
    $ad = if ($rota) { Get-NetAdapter -InterfaceIndex $rota.InterfaceIndex } else { $null }
    $wifi = $ad -and (($ad.PhysicalMediaType -match '802\.11|Wireless') -or ($ad.InterfaceDescription -match 'Wi-?Fi|Wireless|WLAN'))
    $ref = $null; $refTxt = ''
    if ($ad -and -not $wifi) {
        if ($ad.Speed -ge 1000000000) { $ref = $PING_GIGA; $refTxt = "Gigabit: esperado até ~$PING_GIGA ms" }
        elseif ($ad.Speed -ge 100000000) { $ref = $PING_FAST; $refTxt = "FastEthernet: esperado até ~$PING_FAST ms" }
    } elseif ($wifi) { $refTxt = 'Wi-Fi: latência maior é normal; para teste conclusivo use cabo' }
    if ($gw -and $ip -eq $gw) { $ref = $null }   # gateway: só perda de pacotes gera alerta

    # Pacotes de 65500 bytes, como no artigo (ping -l 65500): é com esse tamanho que valem as referências 1-2 ms / 11-12 ms
    W "  Ping para $alvo via $(if ($ad) { "$($ad.Name) $($ad.LinkSpeed)" } else { '?' }) - 50 pacotes de 65500 bytes..."
    $r = TestePing $ip 50 65500
    if ($r.Med -ne $null) {
        Item '  Tempo mín / méd / máx' ("{0} / {1} / {2} ms   [{3}]" -f $r.Min, $r.Med, $r.Max, $refTxt)
    }
    Item '  Pacotes perdidos' ("{0} de {1} ({2}%)" -f $r.Perdidos, $r.Enviados, $r.Perda)
    if ($r.Perdidos -eq $r.Enviados) {
        # Ping bloqueado? Mede pelo tempo de conexao em uma porta aberta (ex.: 1433 do SQL Server)
        $pTcp = @($hInfo.Portas | Sort-Object { if ($PORTAS_DB.ContainsKey($_)) { 0 } else { 1 } })[0]
        if (-not $pTcp) { foreach ($pp in 1433, 445, 3389, 3050, 5432, 80) { $tt = TesteTcp $ip $pp 1; if (-not $tt.Falhas) { $pTcp = $pp; break } } }
        if ($pTcp) {
            $tt = TesteTcp $ip $pTcp 30
            Info "O ping é bloqueado pelo firewall de $alvo - medindo pela conexão na porta $pTcp."
            if ($tt.Med -ne $null) { Item '  Conexão TCP mín / méd / máx' ("{0} / {1} / {2} ms" -f $tt.Min, $tt.Med, $tt.Max) }
            Item '  Conexões que falharam' ("{0} de {1}" -f $tt.Falhas, $tt.Enviados)
            if ($tt.Falhas) { Alerta "$($tt.Falhas) de $($tt.Enviados) conexões com $alvo falharam - instabilidade na rede." 'net' }
            elseif ($tt.Med -gt 10) { Alerta "Conexão com $alvo lenta (média $($tt.Med) ms numa rede local)." 'net' }
            else { Ok "Comunicação com $alvo normal (via TCP)." }
            Info 'Para o teste de ping completo, liberar o ICMP (ping) no firewall do equipamento.'
        } else { Alerta "Sem resposta de $alvo - equipamento desligado, fora da rede ou firewall bloqueando ping." 'net' }
        continue
    }
    elseif ($r.Perdidos -gt 0) { Alerta "Perda de $($r.Perda)% dos pacotes para $alvo - não pode haver perda. Verificar cabo, switch, placa de rede ou Wi-Fi." 'net' }
    if ($ref -and $r.Min -ne $null -and $r.Min -gt $ref) { Alerta "Ping mínimo para $alvo de $($r.Min) ms (média $($r.Med) ms) - acima do esperado ($refTxt). Possível gargalo na rede ou no servidor." 'net' }
    elseif ($ref -and $r.Med -ne $null -and $r.Med -gt ($ref * 3)) { Alerta "Ping instável para ${alvo}: mínimo $($r.Min) ms mas média $($r.Med) ms e máximo $($r.Max) ms - oscilação na rede." 'net' }
    if ($r.Perdidos -eq 0 -and ($ref -eq $null -or ($r.Min -le $ref -and $r.Med -le $ref * 3))) { Ok "Comunicação com $alvo normal." }
}

W ''
W '  Internet (8.8.8.8) - 10 pacotes...'
$ri = TestePing '8.8.8.8' 10
if ($ri.Med -ne $null) { Item '  Tempo médio / perda' ("{0} ms / {1}%" -f $ri.Med, $ri.Perda) }
if ($ri.Perdidos -eq $ri.Enviados) { Alerta 'Sem acesso à internet (8.8.8.8 não responde).' 'net' }
elseif ($ri.Perdidos -gt 0) { Alerta "Perda de $($ri.Perda)% de pacotes para a internet." 'net' }
$dnsOk = $false; try { $null = [Net.Dns]::GetHostAddresses('www.google.com'); $dnsOk = $true } catch {}
if ($dnsOk) { Ok 'Resolução de nomes (DNS) funcionando.' } else { Alerta 'DNS não está resolvendo nomes (www.google.com) - verificar DNS da placa de rede.' 'net' }

$encAnt = [Console]::OutputEncoding
try { [Console]::OutputEncoding = [Text.Encoding]::GetEncoding([Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage) } catch {}
foreach ($tAlvo in @(@($alvos | Select-Object -First 1) + 'google.com.br')) {
    if (-not $tAlvo) { continue }
    W ''
    W "  Tracert para $tAlvo (caminho que os pacotes percorrem até o destino):"
    & tracert.exe -4 -d -h 20 -w 1000 $tAlvo | Where-Object { $_.Trim() } | ForEach-Object { W "      $_" }
}
[Console]::OutputEncoding = $encAnt

# --------------------------------------------------------------------
#  RESUMO, NOTA E MELHORIAS
# --------------------------------------------------------------------
$MELHORIAS = @(CalcularMelhorias)
$NOTA      = Pontuacao $MELHORIAS
$CLASSE    = Classe $NOTA

Secao 'Resumo e interpretação dos resultados'
Item 'Nota da máquina' ("{0}/100 - {1}" -f $NOTA, $CLASSE.Nome)
foreach ($k in $AREAS.Keys) { Item $AREAS[$k].Nome (StatusTexto (StatusArea $k)) }
if (-not $Alertas.Count) { Ok 'Nenhum problema evidente encontrado nos testes automáticos.' }
foreach ($k in $GRUPOS.Keys) {
    $itens = @($Alertas | Where-Object { $_.Tipo -eq $k })
    if (-not $itens.Count) { continue }
    W ''
    W "  >> $($GRUPOS[$k]) ($($itens.Count))" 'Yellow'
    $itens | ForEach-Object { W "     - $($_.Texto)" }
}
W ''
W '  CONCLUSÃO:' 'Cyan'
$tipos = @($Alertas | ForEach-Object { $_.Tipo } | Select-Object -Unique)
foreach ($k in $CONCL.Keys) { if ($k -in $tipos) { W "   * $($CONCL[$k])" 'Yellow' } }
if (-not $tipos.Count) { W "   * $CONCL_OK" 'Green' }
$REFS = "Referências: CPU até $LIM_CPU °C | HD até $LIM_HDD °C | SSD até $LIM_SSD °C | ping até ~$PING_GIGA ms (Gigabit) ou ~$PING_FAST ms (FastEthernet), sem perda de pacotes. Teste de rede: comparar entre 2 ou mais PCs."
W ''
W "  $REFS" 'DarkGray'

Secao 'Melhorias recomendadas'
if (-not $MELHORIAS.Count) { Ok 'Nenhuma correção necessária. Seguir as boas práticas de manutenção preventiva.' }
foreach ($m in $MELHORIAS) {
    W ("  [{0}] {1}" -f $PRIO[$m.Prio].Nome.ToUpper(), $m.Titulo) 'Yellow'
    W "      O que fazer: $($m.Acao)"
    W "      Investimento: $($m.Custo)  |  Quem faz: $($m.Quem)"
}

if ($cdiTexto) {
    $L.Add(''); $L.Add('=' * 74); $L.Add('  ANEXO - RELATÓRIO COMPLETO DO CRYSTALDISKINFO'); $L.Add('=' * 74)
    foreach ($ln in $cdiTexto) { $L.Add($ln) }
    $Rel.Add(@{ T = 'anexo'; X = 'Anexo - relatório completo do CrystalDiskInfo'; Linhas = @($cdiTexto) })
}

# --------------------------------------------------------------------
#  RELATORIOS FORMATADOS (HTML -> PDF pelo Microsoft Edge)
# --------------------------------------------------------------------
function Hx([string]$s) { [Net.WebUtility]::HtmlEncode($s) }

# Logo: logo.png / logo.jpg / logo.svg na pasta do .bat (embutida no PDF)
$logoUri = $null
if ($batDir) {
    foreach ($n in 'logo.png', 'logo.jpg', 'logo.jpeg', 'logo.svg') {
        $lp = Join-Path $batDir $n
        if (Test-Path -LiteralPath $lp) {
            $mime = @{ '.png' = 'image/png'; '.jpg' = 'image/jpeg'; '.jpeg' = 'image/jpeg'; '.svg' = 'image/svg+xml' }[[IO.Path]::GetExtension($lp).ToLower()]
            $logoUri = "data:$mime;base64," + [Convert]::ToBase64String([IO.File]::ReadAllBytes($lp))
            break
        }
    }
}

$CSS = @'
@page { size: A4; margin: 11mm 12mm 13mm; }
:root { --escuro:#0b1220; --azul:#1f4e79; --ciano:#0ea5c6; --verde:#1e7b34; --amarelo:#a15c07; --vermelho:#b42318; --cinza:#5f6b7a; --borda:#d9dee5; }
* { -webkit-print-color-adjust: exact; print-color-adjust: exact; box-sizing: border-box; }
body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 9.5pt; color: #1d2733; background: #fff; margin: 0; line-height: 1.4; }
.topo { display: flex; align-items: center; gap: 20px; background: linear-gradient(120deg, #0b1220 0%, #12233d 55%, #0c3b52 100%); color: #fff; border-radius: 10px; padding: 10px 22px; margin-bottom: 14px; }
.topo img { height: 100px; }
.topo .marca { font-size: 22pt; font-weight: 800; letter-spacing: 2px; color: #67e8f9; }
.topo .t1 { font-size: 18pt; font-weight: 700; }
.topo .t2 { color: #b6c6dc; font-size: 9.5pt; margin-top: 4px; }
.topo .t3 { color: #67e8f9; font-size: 9.5pt; margin-top: 2px; }
h2 { font-size: 12pt; color: var(--escuro); border-bottom: 2px solid var(--ciano); padding-bottom: 3px; margin: 16px 0 8px; break-after: avoid; }
h2 .leg { font-size: 9pt; font-weight: 400; color: var(--cinza); }
p { margin: 4px 0; }
.legenda { color: var(--cinza); font-size: 8.3pt; }
.nota { display: flex; align-items: center; gap: 20px; border: 1px solid var(--borda); border-radius: 10px; padding: 12px 20px; break-inside: avoid; }
.nota .num { font-size: 38pt; font-weight: 800; line-height: 1; white-space: nowrap; }
.nota .num span { font-size: 13pt; color: var(--cinza); font-weight: 600; }
.nota .info { flex: 1; }
.nota .cl { font-size: 14pt; font-weight: 700; }
.nota .frase { color: #3b4756; margin: 2px 0 8px; font-size: 10pt; }
.barra { height: 11px; background: #e9edf2; border-radius: 6px; overflow: hidden; }
.barra i { display: block; height: 100%; border-radius: 6px; }
.c-otimo .num, .c-otimo .cl { color: var(--verde); } .c-otimo .barra i { background: #22a447; }
.c-bom .num, .c-bom .cl { color: #0e7490; } .c-bom .barra i { background: #14a3c7; }
.c-regular .num, .c-regular .cl { color: var(--amarelo); } .c-regular .barra i { background: #f59e0b; }
.c-critico .num, .c-critico .cl { color: var(--vermelho); } .c-critico .barra i { background: #dc2626; }
.cards { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }
.card { border: 1px solid var(--borda); border-left: 6px solid; border-radius: 8px; padding: 7px 11px; break-inside: avoid; }
.card .ch { display: flex; align-items: center; gap: 8px; }
.card .ch b { flex: 1; font-size: 10pt; }
.dot { width: 12px; height: 12px; border-radius: 50%; flex: none; }
.card .lbl { font-size: 8pt; font-weight: 700; padding: 1px 9px; border-radius: 10px; }
.card .cb { color: #3b4756; font-size: 8.8pt; margin-top: 3px; }
.card ul { margin: 2px 0 0; padding-left: 16px; }
.card .nt { color: var(--cinza); font-style: italic; font-size: 8pt; margin-top: 3px; }
.s-ok { border-left-color: #22a447; } .s-ok .dot { background: #22a447; } .s-ok .lbl { background: #e6f4ea; color: var(--verde); }
.s-atencao { border-left-color: #f59e0b; } .s-atencao .dot { background: #f59e0b; } .s-atencao .lbl { background: #fff4e0; color: var(--amarelo); }
.s-critico { border-left-color: #dc2626; } .s-critico .dot { background: #dc2626; } .s-critico .lbl { background: #fdecea; color: var(--vermelho); }
.achado { border: 1px solid var(--borda); border-left: 5px solid; border-radius: 8px; padding: 8px 12px; margin: 7px 0; break-inside: avoid; }
.achado .ah { display: flex; align-items: center; gap: 9px; margin-bottom: 3px; }
.achado .ah b { font-size: 10.5pt; }
.pr { font-size: 7.5pt; font-weight: 700; padding: 2px 9px; border-radius: 10px; text-transform: uppercase; letter-spacing: .5px; flex: none; }
.p1 { border-left-color: #dc2626; } .p1 .pr { background: #fdecea; color: var(--vermelho); }
.p2 { border-left-color: #f59e0b; } .p2 .pr { background: #fff4e0; color: var(--amarelo); }
.p3 { border-left-color: #38bdf8; } .p3 .pr { background: #e0f2fe; color: #0369a1; }
.faz { background: #f1f7fb; border-radius: 6px; padding: 5px 9px; margin-top: 5px; }
.chips { display: flex; gap: 6px; flex-wrap: wrap; margin-top: 6px; }
.chip { font-size: 8pt; background: #eef1f4; color: #334155; border-radius: 10px; padding: 1px 10px; }
.evid { color: var(--cinza); font-size: 7.8pt; margin-top: 5px; }
.ok-geral { background: #e6f4ea; color: var(--verde); border-radius: 8px; padding: 10px 14px; font-weight: 600; }
table.t { width: 100%; border-collapse: collapse; margin: 4px 0; break-inside: avoid; }
table.t th { background: var(--escuro); color: #fff; font-weight: 600; text-align: left; padding: 5px 8px; font-size: 8.8pt; }
table.t td { padding: 5px 8px; border-bottom: 1px solid #e6eaef; vertical-align: top; }
table.cfg td:first-child { color: var(--cinza); width: 165px; }
ul.boas { margin: 4px 0; padding-left: 18px; } ul.boas li { margin: 3px 0; }
.quebra { break-before: page; }
.tec h1 { font-size: 15pt; color: var(--escuro); margin: 0 0 2px; }
.tec h2 { font-size: 10.5pt; color: #fff; background: var(--azul); border: 0; padding: 4px 9px; border-radius: 4px; margin: 14px 0 6px; }
.tec h4 { font-size: 9.5pt; margin: 9px 0 3px; color: var(--azul); break-after: avoid; }
.kv { display: grid; grid-template-columns: 215px 1fr; gap: 8px; padding: 2px 4px; border-bottom: 1px solid #eef1f4; break-inside: avoid; }
.kv span:first-child { color: var(--cinza); } .kv.sub span:first-child { padding-left: 14px; }
.st { display: flex; gap: 8px; align-items: flex-start; padding: 3px 4px; break-inside: avoid; }
.st b { font-size: 7.5pt; padding: 1px 0; border-radius: 3px; width: 52px; text-align: center; flex: none; margin-top: 1px; }
.st.ok b { background: #e6f4ea; color: var(--verde); } .st.al b { background: #fff1d6; color: var(--amarelo); } .st.in b { background: #eef1f4; color: var(--cinza); }
.st.al span { font-weight: 600; }
pre { font-family: Consolas, 'Courier New', monospace; font-size: 7.6pt; background: #f6f8fa; border: 1px solid var(--borda); border-radius: 4px; padding: 6px 8px; margin: 4px 0; white-space: pre-wrap; word-break: break-word; }
footer { color: var(--cinza); font-size: 7.5pt; text-align: center; margin-top: 16px; border-top: 1px solid var(--borda); padding-top: 5px; }
'@

function InicioHtml([string]$titulo) {
    $marca = if ($logoUri) { "<img src='$logoUri' alt='$(Hx $EMPRESA)'>" } else { "<div class='marca'>$(Hx $EMPRESA)</div>" }
    $s = "<!DOCTYPE html><html lang='pt-BR'><head><meta charset='utf-8'><title>$(Hx $titulo) - $(Hx $env:COMPUTERNAME)</title><style>$CSS</style></head><body>"
    $s += "<div class='topo'>$marca<div><div class='t1'>$(Hx $titulo)</div><div class='t2'>Computador: $(Hx $env:COMPUTERNAME) &nbsp;·&nbsp; $(Get-Date -Format 'dd/MM/yyyy HH:mm')</div>"
    if ($sintoma) { $s += "<div class='t3'>Problema relatado: $(Hx $sintoma)</div>" }
    $s + '</div></div>'
}
function RodapeHtml { "<footer>$(Hx $EMPRESA) &nbsp;·&nbsp; $(Hx $SLOGAN) &nbsp;·&nbsp; relatório gerado automaticamente em $(Get-Date -Format 'dd/MM/yyyy HH:mm')</footer></body></html>" }
function NotaHtml {
    "<div class='nota c-$($CLASSE.Cls)'><div class='num'>$NOTA<span>/100</span></div><div class='info'><div class='cl'>Nota da máquina: $(Hx $CLASSE.Nome)</div><div class='frase'>$(Hx $CLASSE.Frase)</div><div class='barra'><i style='width:$([math]::Max(3, $NOTA))%'></i></div></div></div>"
}
function AchadoHtml($m, [switch]$Completo) {
    $pr = $PRIO[[int]$m.Prio]
    $s = "<div class='achado $($pr.Cls)'><div class='ah'><span class='pr'>$($pr.Nome)</span><b>$(Hx $m.Titulo)</b></div><p>$(Hx $m.Explica)</p><div class='faz'><b>O que fazer:</b> $(Hx $m.Acao)</div>"
    if ($Completo) {
        $s += "<div class='chips'><span class='chip'>Investimento: $(Hx $m.Custo)</span><span class='chip'>Quem faz: $(Hx $m.Quem)</span><span class='chip'>Área: $(Hx $AREAS[$m.Area].Nome)</span></div>"
        $s += "<div class='evid'><b>O que a análise encontrou:</b> " + (($m.Evidencias | ForEach-Object { Hx $_ }) -join ' &nbsp;·&nbsp; ') + '</div>'
    }
    $s + '</div>'
}

# ---------- Relatorio de analise ----------
function GerarHtml {
    $sb = New-Object Text.StringBuilder
    $add = { param($s) [void]$sb.Append($s) }
    & $add (InicioHtml 'Relatório de análise da máquina')
    & $add (NotaHtml)

    & $add "<h2>Painel de saúde <span class='leg'>&nbsp;verde = tudo certo · amarelo = atenção · vermelho = resolver logo</span></h2><div class='cards'>"
    foreach ($k in $AREAS.Keys) {
        $st = StatusArea $k
        $ms = @($MELHORIAS | Where-Object { $_.Area -eq $k })
        $corpo = if ($ms.Count) { '<ul>' + (($ms | ForEach-Object { "<li>$(Hx $_.Titulo)</li>" }) -join '') + '</ul>' } else { Hx $AREAS[$k].Ok }
        $nt = if ($NOTAS.ContainsKey($k)) { ($NOTAS[$k] | ForEach-Object { "<div class='nt'>$(Hx $_)</div>" }) -join '' } else { '' }
        & $add "<div class='card s-$st'><div class='ch'><span class='dot'></span><b>$(Hx $AREAS[$k].Nome)</b><span class='lbl'>$(StatusTexto $st)</span></div><div class='cb'>$corpo</div>$nt</div>"
    }
    & $add '</div>'

    & $add '<h2>O que encontramos, em palavras simples</h2>'
    if (-not $MELHORIAS.Count) { & $add "<div class='ok-geral'>Nenhum problema foi encontrado. $(Hx $CONCL_OK)</div>" }
    foreach ($m in $MELHORIAS) { & $add (AchadoHtml $m) }
    if ($MELHORIAS.Count) { & $add "<p class='legenda'>O plano completo, com investimento estimado e quem faz cada item, está no <b>Relatório de melhorias</b>.</p>" }

    # Detalhes tecnicos (tudo o que foi coletado), a partir de uma nova pagina
    & $add "<div class='tec quebra'><h1>Detalhes técnicos</h1><p class='legenda'>Dados completos coletados na análise, para uso do técnico.</p>"
    $pre = New-Object System.Collections.Generic.List[string]
    $fechaPre = { if ($pre.Count) { & $add ('<pre>' + (Hx ($pre -join "`n")) + '</pre>'); $pre.Clear() } }
    $pular = $false
    foreach ($e in $Rel) {
        if ($e.T -eq 'sec') { & $fechaPre; $pular = ($e.X -like 'Resumo*' -or $e.X -like 'Melhorias*'); if (-not $pular) { & $add "<h2>$(Hx $e.X)</h2>" }; continue }
        if ($pular) { continue }
        if ($e.T -eq 'txt') {
            $x = $e.X -replace '^  ', ''
            $tr = $x.Trim()
            if (-not $tr) { continue }
            if ($tr -match '^(Consultando|Verificando|Executando) ') { continue }   # mensagens de progresso
            if ($x -notmatch '^\s' -and ($tr -match ':$' -or $tr -match '\.\.\.$' -or $tr -match '^Disco \d+:')) { & $fechaPre; & $add "<h4>$(Hx ($tr -replace '\.\.\.$', ''))</h4>"; continue }
            $pre.Add($x); continue
        }
        & $fechaPre
        switch ($e.T) {
            'ok'     { & $add "<div class='st ok'><b>OK</b><span>$(Hx $e.X)</span></div>" }
            'alerta' { & $add "<div class='st al'><b>ALERTA</b><span>$(Hx $e.X)</span></div>" }
            'info'   { & $add "<div class='st in'><b>INFO</b><span>$(Hx $e.X)</span></div>" }
            'kv'     { & $add "<div class='kv$(if ($e.Sub) { ' sub' })'><span>$(Hx $e.K)</span><span>$(Hx $e.V)</span></div>" }
            'anexo'  { & $add "<h2>$(Hx $e.X)</h2><pre>$(Hx ($e.Linhas -join "`n"))</pre>" }
        }
    }
    & $fechaPre
    & $add "<p class='legenda'>$(Hx $REFS)</p></div>"
    & $add (RodapeHtml)
    $sb.ToString()
}

# ---------- Relatorio de melhorias ----------
$BOAS = @(
    'Fazer backup dos arquivos importantes pelo menos uma vez por semana (HD externo, pen drive ou nuvem).',
    'Reiniciar o computador pelo menos uma vez por semana.',
    'Manter o Windows e o antivírus sempre atualizados.',
    'Fazer limpeza interna (poeira e ventoinhas) a cada 6 a 12 meses.',
    'Usar nobreak ou estabilizador para proteger contra quedas de energia.',
    'Evitar instalar programas desconhecidos e deixar só o necessário abrindo com o Windows.',
    'Repetir esta análise a cada 6 meses ou sempre que notar lentidão.'
)
function GerarHtmlMelhorias {
    $sb = New-Object Text.StringBuilder
    $add = { param($s) [void]$sb.Append($s) }
    & $add (InicioHtml 'Relatório de melhorias')
    & $add (NotaHtml)
    & $add "<p style='margin-top:10px'>Com base na análise feita neste computador, listamos o que pode ser feito para deixá-lo <b>mais rápido, estável e seguro</b>, em ordem de prioridade.</p>"

    if ($MELHORIAS.Count) {
        & $add "<h2>Resumo do plano</h2><table class='t'><tr><th>#</th><th>Melhoria</th><th>Prioridade</th><th>Investimento</th><th>Quem faz</th></tr>"
        $i = 0
        foreach ($m in $MELHORIAS) { $i++; & $add "<tr><td>$i</td><td>$(Hx $m.Titulo)</td><td class='$($PRIO[$m.Prio].Cls)'><span class='pr'>$($PRIO[$m.Prio].Nome)</span></td><td>$(Hx $m.Custo)</td><td>$(Hx $m.Quem)</td></tr>" }
        & $add '</table>'
        foreach ($p in 1, 2, 3) {
            $ms = @($MELHORIAS | Where-Object { $_.Prio -eq $p })
            if (-not $ms.Count) { continue }
            & $add "<h2>$($PRIO[$p].Grupo) <span class='leg'>&nbsp;$($PRIO[$p].Desc)</span></h2>"
            foreach ($m in $ms) { & $add (AchadoHtml $m -Completo) }
        }
    } else {
        & $add "<div class='ok-geral' style='margin-top:10px'>Nenhuma correção é necessária no momento. Para manter o computador assim, siga as boas práticas abaixo.</div>"
    }

    & $add "<h2>Configuração atual do computador</h2><table class='t cfg'>"
    foreach ($k in $Config.Keys) { & $add "<tr><td>$(Hx $k)</td><td>$(Hx $Config[$k])</td></tr>" }
    & $add '</table>'

    & $add "<h2>Manutenção preventiva (boas práticas)</h2><ul class='boas'>"
    foreach ($b in $BOAS) { & $add "<li>$(Hx $b)</li>" }
    & $add '</ul>'
    & $add "<p class='legenda'><b>Investimento estimado:</b> Sem custo = só configuração ou limpeza de arquivos · Baixo = até cerca de R$ 150 · Médio = cerca de R$ 150 a R$ 600 · Alto = acima de R$ 600. Valores aproximados; confirme sempre em orçamento.</p>"
    & $add (RodapeHtml)
    $sb.ToString()
}

function HtmlParaPdf([string]$html, [string]$pdf) {
    $nav = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
             "$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
             "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe") | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
    if (-not $nav) { return $false }
    $perfil = Join-Path $env:TEMP ('am_pdf_' + [guid]::NewGuid().ToString('N'))
    $uri = ([Uri]$html).AbsoluteUri
    $arg = @('--headless', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-extensions',
             "--user-data-dir=`"$perfil`"", '--no-pdf-header-footer', '--print-to-pdf-no-header', "--print-to-pdf=`"$pdf`"", "`"$uri`"")
    Remove-Item -LiteralPath $pdf -Force -ErrorAction SilentlyContinue
    $p = Start-Process -FilePath $nav -ArgumentList $arg -PassThru -WindowStyle Hidden
    if ($p -and -not $p.WaitForExit(60000)) { try { $p.Kill() } catch {} }
    for ($i = 0; $i -lt 20 -and -not (Test-Path -LiteralPath $pdf); $i++) { Start-Sleep -Milliseconds 500 }
    Remove-Item -LiteralPath $perfil -Recurse -Force -ErrorAction SilentlyContinue
    (Test-Path -LiteralPath $pdf) -and ((Get-Item -LiteralPath $pdf).Length -gt 1000)
}
function SalvarRelatorio([string]$conteudo, [string]$base) {
    $html = "$base.html"; $pdf = "$base.pdf"
    [IO.File]::WriteAllText($html, $conteudo, (New-Object Text.UTF8Encoding $true))
    if (HtmlParaPdf $html $pdf) { Remove-Item -LiteralPath $html -Force; return $pdf }
    return $html
}

# --------------------------------------------------------------------
#  SALVAR RELATORIOS
# --------------------------------------------------------------------
$nome = 'Analise_{0}_{1}' -f $env:COMPUTERNAME, (Get-Date -Format 'yyyy-MM-dd_HH-mm')
$destinos = @()
if ($batDir) { $destinos += (Join-Path $batDir 'Relatorios') }
$destinos += [Environment]::GetFolderPath('Desktop')
$arq = $null
foreach ($dst in $destinos) {
    try {
        if (-not (Test-Path -LiteralPath $dst)) { New-Item -ItemType Directory -Path $dst -Force -ErrorAction Stop | Out-Null }
        $tmp = Join-Path $dst "$nome.txt"
        [IO.File]::WriteAllLines($tmp, $L, (New-Object Text.UTF8Encoding $true))
        $arq = $tmp; break
    } catch {}
}
Write-Host ''
if ($arq) {
    $pastaRel = Split-Path -Parent $arq
    $abrir = @()
    try {
        Write-Host '  Gerando os relatórios em PDF...' -ForegroundColor Cyan
        $abrir += SalvarRelatorio (GerarHtml) (Join-Path $pastaRel $nome)
        $abrir += SalvarRelatorio (GerarHtmlMelhorias) (Join-Path $pastaRel ($nome -replace '^Analise_', 'Melhorias_'))
    } catch { Write-Host "  Erro ao gerar os relatórios formatados: $($_.Exception.Message)" -ForegroundColor Yellow }
    if (@($abrir | Where-Object { $_ -like '*.html' }).Count) { Write-Host '  Edge/Chrome não encontrado - relatórios salvos em HTML (abrem no navegador).' -ForegroundColor Yellow }
    if (-not $abrir.Count) { $abrir = @($arq) }
    Write-Host ("  Nota da máquina: {0}/100 - {1}" -f $NOTA, $CLASSE.Nome) -ForegroundColor Cyan
    foreach ($a in $abrir) { Write-Host "  Salvo: $a" -ForegroundColor Green }
    Write-Host "  (versão em texto: $arq)" -ForegroundColor DarkGray
    if (-not $auto) { foreach ($a in $abrir) { Start-Process -FilePath $a } }
} else { Write-Host '  Não foi possível salvar o relatório.' -ForegroundColor Red }

# --------------------------------------------------------------------
#  MONITORAMENTO EM TEMPO REAL (Gerenciador de Tarefas + HWMonitor)
# --------------------------------------------------------------------
$hwm = Achar $hwmFixos $hwmNomes
if (-not $auto) {
    Write-Host ''
    if (-not $hwm) { Write-Host '  HWMonitor não disponível (download recusado ou sem internet).' -ForegroundColor DarkGray }
    $resp = Read-Host ('  Abrir o Gerenciador de Tarefas' + $(if ($hwm) { ' e o HWMonitor' } else { '' }) + ' para acompanhar em tempo real? (S/N) [N]')
    if ($resp -match '^[sS]') {
        Start-Process taskmgr.exe
        if ($hwm) { Start-Process -FilePath $hwm }
    }
}
