# Robo Sallus

Projeto para automatizar o acesso e preenchimento de elementos no portal Salus/Orizon.

## Estado atual

- Portal inicial: https://www.orizonbrasil.com.br/acesso-restrito.html
- Produto correto no portal Orizon: `SALUS`
- Destino do produto: https://salus.orizon.com.br
- Login testado com sucesso em 2026-07-15.
- Tela alcançada apos login: `https://salus.orizon.com.br/salus/hospital`
- Hospital alvo para continuar: `Hospital Sirio Libanes - cod: 190314`

## Arquivos gerados

- Lista de pacientes: `exports/pacientes_sirio_libanes_2026-07-15.xlsx`
- Planilha de preenchimento da evolucao clinica: `exports/preenchimento_evolucao_clinica_LP_K3ZAVM6_2026-07-15.xlsx`
- Copia colorida da planilha: `exports/preenchimento_evolucao_clinica_LP_K3ZAVM6_2026-07-15_colorido.xlsx`

## Scripts

- `scripts/REGRA_FLUXO_ROBO_SALLUS.txt`: regra principal do fluxo em 2 etapas automatizadas, com pausa humana entre elas.
- `scripts/REGRA_ETAPA_2_LANCAR_EVOLUCAO_SALUS.txt`: regra detalhada do lancamento senha por senha no Salus.
- `scripts/app_robo_sallus_web.py`: tela web local com botoes Etapa 1 e Etapa 2, cards e log.
- `scripts/app_robo_sallus.py`: versao desktop Tk, depende de Python com tkinter.
- `scripts/salus_cdp.py`: chama APIs do Salus usando o Chrome ja logado.
- `scripts/gerar_lista_pacientes.py`: baixa a lista de pacientes para Excel.
- `scripts/gerar_planilha_evolucao.py`: gera a planilha de evolucao clinica com campos em colunas.
- `scripts/colorir_planilha_evolucao.py`: colore a aba `Preenchimento` por bloco/etapa.
- `scripts/aplicar_listas_excel.py`: aplica listas suspensas para campos de escolha unica e orienta campos de multipla escolha.
- `scripts/etapa2_lancar_evolucao_salus.py`: executa a regra da etapa 2 em modo simulacao ou lancamento real.

Exemplos:

```bash
python3 scripts/gerar_lista_pacientes.py
python3 scripts/colorir_planilha_evolucao.py
python3 scripts/aplicar_listas_excel.py
python3 scripts/etapa2_lancar_evolucao_salus.py --limite 5
```

Para abrir a tela do robo:

```bash
python3 scripts/app_robo_sallus_web.py
```

Ou, no macOS, abrir o arquivo:

```bash
./RoboSallus.command
```

## Observacoes

Credenciais nao devem ser salvas em texto puro neste repositorio. Para continuar, informe a senha em tempo de execucao ou use uma variavel de ambiente/local seguro.
