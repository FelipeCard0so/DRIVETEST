# DT 3.0 — Automação de Drive Test

Ferramenta de automação para técnicos de campo que executam **Drive Tests** (SSV) em redes móveis. Organiza atividades, otimiza rotas, atualiza planilhas em tempo real e gera um mapa interativo publicado automaticamente no GitHub Pages.

> Desenvolvido para uso em campo real, atendendo operadoras como Vivo, através de equipes como NGSolution, TELEQUIPE, EOLEN e STEIN.

---

## 📌 O que o projeto faz

O DT 3.0 nasceu da necessidade real de organizar o trabalho diário de um técnico de Drive Test: múltiplas atividades espalhadas por diferentes cidades e estados, troca constante de planilhas, cálculo manual de rotas e falta de visibilidade do progresso do dia.

O sistema automatiza todo esse fluxo:

- **Lê e processa** novas atividades vindas de planilhas Excel
- **Otimiza a rota** entre os sites pendentes, priorizando o vizinho mais próximo
- **Calcula rotas reais** (com trânsito) via Mapbox Directions API
- **Atualiza o Google Sheets** automaticamente com o status de cada atividade
- **Gera um mapa interativo** e publica no GitHub Pages a cada execução
- **Normaliza dados de frequência e tecnologia** (2G/3G/4G/5G) vindos de fontes inconsistentes

---

## 🚀 Funcionalidades

- 🗺️ Mapa interativo em **Mapbox GL JS**, com rota real calculada via API de Directions
- 📊 Integração direta com **Google Sheets** (leitura e escrita via `gspread`)
- 🧭 Otimização de rota pelo algoritmo do vizinho mais próximo
- 🔄 Reconhecimento automático do ponto de partida com base no último status de atividade finalizada (`✓ Atividade concluída`, `IMPRODUTIVO`, `CANCELADO`, `ÁREA DE RISCO`)
- 🧹 Normalização robusta de tecnologia (`LTE → 4G`, `NR → 5G`) e de frequências, preservando portadoras duplicadas (ex.: `3G:850/850` = dois testes distintos)
- 🌐 Publicação automática do mapa atualizado no **GitHub Pages**
- 📦 Empacotado como executável standalone via **PyInstaller**

---

## 🛠️ Tecnologias utilizadas

| Camada | Tecnologia |
|---|---|
| Linguagem | Python |
| Planilhas | Google Sheets API (`gspread`) |
| Mapa | Mapbox GL JS v2.15.0 |
| Publicação | GitHub Pages |
| Empacotamento | PyInstaller (`--onefile --console`) |

---

## ⚙️ Regras de negócio

Algumas decisões de design seguem regras específicas do processo real de Drive Test:

- Frequências duplicadas **nunca são removidas** — cada repetição representa um teste distinto
- Portadoras com prefixo `RS` são tratadas como distintas da mesma frequência sem prefixo (ex.: `RS2600 ≠ 2600`)
- `LTE` é sempre normalizado para `4G`, `NR` para `5G`
- O vizinho mais próximo tem prioridade máxima na otimização de rota
- Nenhuma atividade é removida da planilha — apenas atualizada
- A primeira linha da planilha contém fórmulas e nunca é sobrescrita

---

## 📁 Estrutura esperada

```
DT 3.0/
├── dt30_main.py
├── credentials.json       # Credenciais da Google Service Account
├── config.json             # Configurações (IDs de planilhas, tokens)
└── dist/
    └── dt30_main.exe       # Executável gerado via PyInstaller
```

---

## ▶️ Como executar

```bash
# Clonar o repositório
git clone https://github.com/FelipeCard0so/DRIVETEST.git

# Instalar dependências
pip install gspread google-auth requests

# Executar
python dt30_main.py
```

Ou, para gerar o executável:

```bash
pyinstaller --onefile --console dt30_main.py
```

---

## 🧩 Status do projeto

Em evolução ativa. Melhorias recentes incluem rotas reais com trânsito via Mapbox, cálculo de ETA, aviso de limite de quota da API e correção do reconhecimento do ponto de partida considerando os quatro status de atividade finalizada.

**Limitação conhecida:** na virada de mês, quando uma nova aba é criada na planilha, o sistema ainda pode perder a referência do último local visitado, exigindo inserção manual de coordenadas.

---

## 👤 Autor

**Felipe Cardoso**
Técnico de campo (Drive Test / SSV) e desenvolvedor de ferramentas de automação para o próprio fluxo de trabalho.
GitHub: [@FelipeCard0so](https://github.com/FelipeCard0so)
