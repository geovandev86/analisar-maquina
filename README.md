# Analisar-Maquina

Script `.bat` que faz sozinho o checklist do artigo **"Como analisar uma máquina"** da base de conhecimento SkyTech e gera um relatório em **PDF** (e também em `.txt`).

## Como usar

1. Copie a pasta para a máquina (pendrive, AnyDesk ou rede).
2. Dê dois cliques em `Analisar-Maquina.bat` e clique em **Sim** na permissão de administrador.
3. Responda às perguntas: sintoma, se quer rodar o SFC e quais equipamentos testar na rede (Enter aceita a sugestão).
4. No final, dois PDFs abrem sozinhos e ficam salvos em `Relatorios\` (junto com uma versão em `.txt` para colar no chamado):
   - **Relatório de análise**: nota da máquina de 0 a 100, painel de saúde por área (verde / amarelo / vermelho), os problemas explicados em palavras simples e, no final, os detalhes técnicos.
   - **Relatório de melhorias**: plano de ação por prioridade (fazer agora / próximos dias / quando possível), com o porquê, como fazer, investimento estimado e quem faz, a configuração atual da máquina e boas práticas de manutenção.

   Os PDFs são gerados pelo Microsoft Edge, que já vem no Windows.

## Sua logo nos relatórios

Coloque a logo da sua empresa como `logo.png` (ou `.jpg` / `.svg`) na mesma pasta do `.bat`. O nome da empresa e a frase do rodapé ficam nas variáveis `$EMPRESA` e `$SLOGAN`, no começo da parte PowerShell do `.bat`. A logo não vai para o repositório (está no `.gitignore`).

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
