# Analisar-Maquina

Script `.bat` que faz sozinho o checklist do artigo **"Como analisar uma máquina"** da base de conhecimento SkyTech e gera um relatório em `.txt`.

## Como usar

1. Copie a pasta para a máquina (pendrive, AnyDesk ou rede).
2. Dê dois cliques em `Analisar-Maquina.bat` e clique em **Sim** na permissão de administrador.
3. Responda às perguntas: sintoma, se quer rodar o SFC e quais equipamentos testar na rede (Enter aceita a sugestão).
4. No final, o relatório abre no Bloco de Notas e fica salvo em `Relatorios\`.

> Se aparecer "O Windows protegeu o computador", clique em **Mais informações > Executar assim mesmo**.

## O que é analisado

| Área | Verificações |
|---|---|
| Sistema | Windows e versão, modelo, processador, RAM, tempo ligado |
| Logs do Windows | Tela azul (BugCheck), desligamento inesperado (Kernel-Power 41), erros de hardware (WHEA), teste de memória, programas que travaram |
| Disco | Saúde, temperatura, desgaste, erros de leitura/escrita, eventos de disco (ex.: 153), espaço livre e relatório S.M.A.R.T. do CrystalDiskInfo |
| Processador | Temperatura, limitação de velocidade, plano de energia |
| Integridade | DISM CheckHealth e SFC (opcional) |
| Atualizações | Última instalada, pendentes, reinício pendente |
| Consumo | CPU, memória, disco, processos que mais consomem, programas na inicialização |
| Rede | Encontra sozinho o servidor e os PCs da rede, ping com pacotes de 65500 bytes, tracert |

No final, os alertas são separados em **hardware**, **software**, **desempenho** e **rede**, com a conclusão: falha física (procurar técnico) ou software (reinstalar/reconfigurar).

## Referências usadas (do artigo)

- Temperaturas: CPU até 65 °C, HD até 55 °C, SSD até 46 °C
- Ping (pacote de 65500 bytes): até ~2 ms em rede Gigabit, ~11–12 ms em FastEthernet, sem perda de pacotes

## Ferramentas

O CrystalDiskInfo e o HWMonitor são baixados automaticamente dos sites oficiais (versão portátil) para a pasta `Ferramentas\` na primeira execução. Leve essa pasta junto para usar em máquinas sem internet.
