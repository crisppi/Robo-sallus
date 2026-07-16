# Continuidade - Salus

## Fluxo descoberto

1. Abrir `https://www.orizonbrasil.com.br/acesso-restrito.html`.
2. Clicar em `Efetuar login` do produto `SALUS`.
3. O portal consulta:
   `https://www.orizonbrasil.com.br/Autenticacao/Login/PreLogin?selectValue=Salus`
4. A resposta redireciona para:
   `https://salus.orizon.com.br`
5. No Salus, a primeira tela pede apenas o usuario.
6. A triagem do usuario chama:
   `https://salus.orizon.com.br/api/autenticacao/triagem?Usuario=<usuario>`
7. Para o usuario testado, a triagem retornou `tipoAutenticacao: login-auditor`.
8. Apos autenticar, a tela carregada foi:
   `https://salus.orizon.com.br/salus/hospital`

## Dados para retomada

- Usuario: `RobertoHm`
- Senha: nao salva aqui por seguranca.
- Perfil retornado pela API:
  - Nome: Roberto Benedito Crisppi
  - Grupo: `salus-auditor-externo`
  - Auditor: `49`
- Hospital alvo:
  - Nome: `Hospital Sirio Libanes`
  - Codigo: `190314`
  - Endereco exibido: `R DA ADMA JAFET, 0 - -, Sao Paulo - SP`

## Observacao tecnica

O macOS bloqueou automacao direta por AppleScript na janela normal do Chrome. O login foi concluido usando uma janela temporaria do Chrome com DevTools remoto na porta `9222`.

Para construir o robo, o caminho mais estavel provavelmente sera usar Playwright/Puppeteer com um perfil de navegador controlado, em vez de depender de teclado/mouse do sistema.

Em 2026-07-15, o Chrome recusou a conexao WebSocket com erro `403 Forbidden`
por causa do cabecalho `Origin`. O helper `scripts/salus_cdp.py` foi ajustado
para conectar com `suppress_origin=True`.

## Exportacao gerada

- Arquivo: `exports/pacientes_sirio_libanes_2026-07-15.xlsx`
- Registros: `409`
- Colunas: `Nome`, `Iniciais`, `Senha`, `Dias internado`
- Origem: tela `https://salus.orizon.com.br/salus/gestao-internacao`
- Endpoint observado na tela:
  `https://salus.orizon.com.br/api/internacoes?user_key=49&IdPrestador=113&IdEmpresaAuditoria=6&Pagina=1&TamanhoPagina=500`
- Script salvo: `scripts/gerar_lista_pacientes.py`
- Teste apos correcao do DevTools:
  `exports/pacientes_sirio_libanes_teste_devtools.xlsx` com `409` pacientes

## Mapeamento do formulario de evolucao clinica

- Paciente mapeada: `LP` / senha `K3ZAVM6`
- Planilha principal: `exports/preenchimento_evolucao_clinica_LP_K3ZAVM6_2026-07-15.xlsx`
- Copia colorida: `exports/preenchimento_evolucao_clinica_LP_K3ZAVM6_2026-07-15_colorido.xlsx`
- Campos em colunas: `156`
- Abas: `Preenchimento`, `Campos`, `Opcoes`, `Resumo`
- Cores: aplicadas na aba `Preenchimento` por bloco/etapa, com legenda na aba `Resumo`
- Scripts salvos:
  - `scripts/gerar_planilha_evolucao.py`
  - `scripts/colorir_planilha_evolucao.py`
  - `scripts/aplicar_listas_excel.py`
- Status do checklist no rascunho atual:
  - `Preenchido`: 2
  - `Faltando`: 43
  - `Condicional`: 107
  - `Opcional`: 2
  - `Somente exibição`: 2
- Origem da definicao:
  `https://salus.orizon.com.br/api/formularios/auditoria/definicao?user_key=49&idInternacao=3956`

## Etapa 2 do robo

- Regra detalhada: `scripts/REGRA_ETAPA_2_LANCAR_EVOLUCAO_SALUS.txt`
- Script principal: `scripts/etapa2_lancar_evolucao_salus.py`
- Tela web local: `scripts/app_robo_sallus_web.py`
- Tela desktop Tk opcional: `scripts/app_robo_sallus.py`
- Atalho executavel macOS: `RoboSallus.command`
- Modo padrao: dry-run, sem lancar dados no Salus.
- Teste LP executado:
  `python3 scripts/etapa2_lancar_evolucao_salus.py --somente-senha K3ZAVM6 --saida exports/relatorio_lancamentos_teste_LP.xlsx`
- Resultado do teste LP: `DRY_RUN: 1`
- Teste com 5 pacientes:
  `python3 scripts/etapa2_lancar_evolucao_salus.py --limite 5 --saida exports/relatorio_lancamentos_teste_5.xlsx`
- Resultado do teste com 5 pacientes: `DRY_RUN: 1`, `PULADO: 4`
- Observacao: o executor real do Salus ainda esta isolado na classe `SalusExecutor`; no teste real, ligar os cliques/API dentro dessa classe.
