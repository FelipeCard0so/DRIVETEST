# 🗺 DT 3.0 — Automação Drive Test

> Ferramenta profissional de automação para engenheiros de RF em campo.  
> Organiza atividades, otimiza rotas de deslocamento, gera relatórios técnicos, atualiza planilhas e publica mapa interativo online — tudo com um único clique.

---

## 📋 Visão Geral

O **DT 3.0** foi desenvolvido para eliminar o trabalho manual repetitivo no planejamento de campanhas de Drive Test em redes móveis 2G/3G/4G/5G. O engenheiro recebe as atividades, coloca na pasta correta e executa o programa — o restante é automático.

### O que o programa faz

- Lê as novas atividades da planilha `Felipe_DATA.xlsx`
- Normaliza e organiza as frequências de rádio (2G/3G/4G/5G)
- Busca PCI, Azimuth e UF na base de dados 4G (`Base_4G.xlsx`)
- **Otimiza a rota de deslocamento** usando algoritmos matemáticos avançados (Vizinho Mais Próximo + 2-OPT + Or-Opt + 3-OPT com múltiplos pontos de partida)
- Gera o `RELATORIO_SSV_COMPLETO.txt` com todas as informações técnicas por site
- Gera o `MAPA_ROTAS.html` — mapa interativo com rota, hotéis, status por cor
- Publica o mapa automaticamente no **GitHub Pages** (acesso público via link)
- Atualiza a planilha de desempenho no **Google Sheets**

---

## 🚀 Modos de Execução

Ao iniciar o programa, você escolhe entre dois modos:

```
[1] Execução completa
    Processa novas atividades, otimiza rota,
    gera relatório, atualiza Sheets e mapa.

[2] Atualizar mapa
    Lê status atuais do Sheets e gera novo
    mapa sem reprocessar nada mais.
```

O **Modo 2** é ideal para atualizar o mapa durante o dia conforme as atividades vão sendo concluídas, sem precisar rodar o processo inteiro.

---

## 🗂 Estrutura de Pastas

```
DT 3.0/
│
├── dt30_main.py              ← Código principal (ou dt30_main.exe)
├── Base_4G.xlsx              ← Base de dados RF (PCI, Azimuth, UF, coordenadas)
├── credentials.json          ← Credenciais Google Sheets (não versionar!)
├── github_token.txt          ← Token de acesso GitHub Pages (não versionar!)
│
├── novas atividades/         ← Coloque aqui o arquivo Felipe_DATA.xlsx
│   └── processados/          ← Arquivos já processados são movidos aqui
│
└── out/
    └── RELATORIO_SSV_COMPLETO.txt  ← Relatório gerado automaticamente
```

---

## 🗺 Mapa Interativo

O mapa gerado exibe todas as atividades com **cores por status**:

| Cor | Significado |
|-----|-------------|
| 🟠 Laranja (estrela) | Ponto de partida (localização atual) |
| 🔵 Azul | Atividade pendente — faz parte da rota |
| 🟢 Verde | Atividade concluída |
| 🔴 Vermelho | Atividade improdutiva |
| 🟣 Roxo | Aguardando para deslocar (fora da rota) |
| 🟠 Laranja (torre) | Hotel / Pousada |

### Funcionalidades do mapa

- **Linha de rota** conectando todos os pontos pendentes na ordem otimizada
- **Popup** em cada marcador com site, cidade, frequências, hotel e observações
- **Botão "Marcar como Concluída"** em cada atividade pendente (atualiza a rota visualmente)
- **Checkboxes** para mostrar/ocultar camadas: Concluídas, Improdutivas, Hotéis, Aguardando
- **Contador** de pendentes e concluídas em tempo real

---

## ⚙️ Algoritmo de Otimização de Rota

O DT 3.0 utiliza um pipeline de 4 camadas para encontrar a rota mais curta:

1. **Múltiplos pontos de partida** — testa o Vizinho Mais Próximo saindo de cada cidade do pool, eliminando a dependência de um único ponto inicial
2. **2-OPT** — inverte segmentos da rota para remover cruzamentos
3. **Or-Opt** — move blocos de 1, 2 ou 3 cidades para posições melhores (resolve o clássico problema de "passar por uma cidade e voltar depois")
4. **3-OPT** — testa 7 variantes de reconexão para cruzamentos complexos que o 2-OPT não resolve
5. **Or-Opt final** — polish de refinamento após o 3-OPT

---

## 📊 Status das Atividades

O programa respeita os seguintes status na planilha do Google Sheets:

| Status | Comportamento |
|--------|---------------|
| `✓ Atividade concluída` | Fixo — não entra na otimização, aparece no mapa em verde |
| `IMPRODUTIVO` | Fixo — não entra na otimização, aparece no mapa em vermelho |
| `Aguardando para deslocar` | Separado — vai ao final da planilha e aparece em roxo no mapa |
| `>> EM DESLOCAMENTO` | Entra normalmente na otimização |
| `Nova Atividade` | Marcador de atividade recém-inserida |
| *(vazio)* | Entra normalmente na otimização |

---

## 🔧 Pré-requisitos

### Python 3.10+

```bash
pip install pandas openpyxl requests gspread google-auth
```

### Arquivos necessários

- `Base_4G.xlsx` — base de dados RF com colunas: `SITE`, `LATITUDE`, `LONGITUDE`, `CIDADE`, `[P]UF`, `PCI`, `AZIMUTH`
- `credentials.json` — conta de serviço Google Cloud com acesso à API Sheets e Drive
- `github_token.txt` — token de acesso GitHub (3 linhas: token / usuário / repositório)

---

## 🔑 Configuração das APIs

### Google Sheets

1. Acesse [console.cloud.google.com](https://console.cloud.google.com)
2. Crie um projeto → ative **Google Sheets API** e **Google Drive API**
3. Crie uma **Service Account** com papel Editor
4. Baixe o JSON e renomeie para `credentials.json`
5. Compartilhe sua planilha com o e-mail da Service Account

### GitHub Pages

Crie o arquivo `github_token.txt` com exatamente 3 linhas:

```
ghp_seu_token_aqui
seu-usuario-github
nome-do-repositorio
```

O token precisa do escopo `repo`. Gere em: GitHub → Settings → Developer settings → Personal access tokens.

---

## 📦 Compilar como executável (.exe)

```bash
pyinstaller --onefile --console dt30_main.py
```

Mova o `dt30_main.exe` da pasta `dist/` para a pasta `DT 3.0/` junto com os demais arquivos.

> ⚠️ Use `--console` (não `--noconsole`) pois o programa requer interação com o usuário durante a execução.

---

## 📁 Formato do arquivo de entrada

O arquivo `Felipe_DATA.xlsx` deve conter as colunas:

| Coluna | Descrição |
|--------|-----------|
| `Site` | Código do site (ex: MGVJG) |
| `Cidade` | Nome da cidade |
| `Latitude` | Coordenada latitude |
| `Longitude` | Coordenada longitude |
| `Frequencia` | Frequências do site (ex: `4G:700/1800/2100\|5G:3500`) |
| `Tecnologia` | Tecnologia (4G, 5G, etc.) |
| `Regional` | UF do site |
| `Vendor` | Integrador de campo |
| `Integradora` | Empresa demandante |

---

## 🗓 Ciclo Mensal

O programa detecta automaticamente o mês vigente e trabalha na aba correspondente do Google Sheets (`JANEIRO/2026`, `FEVEREIRO/2026`, etc.). Na virada do mês, a nova aba é criada automaticamente duplicando a estrutura da anterior.

---

## 🔒 Segurança

Os arquivos `credentials.json` e `github_token.txt` **nunca devem ser versionados**. Adicione ao `.gitignore`:

```
credentials.json
github_token.txt
```

---

## 👤 Autor

Desenvolvido por **Felipe** — Engenheiro de RF  
Automação de campo para campanhas de Drive Test em redes móveis brasileiras.
