# Gerador de Atas — Secretário Executivo de IA

Aplicativo Streamlit que utiliza IA Generativa (Google Gemini) para automatizar a elaboração de atas executivas dos fóruns de **S&OE** e **S&OP** do Grupo Aço Cearense.

## O que faz

- Recebe a transcrição bruta de uma reunião (texto ou áudio transcrito)
- Corrige automaticamente erros de reconhecimento de voz para termos técnicos do setor (CA-50, CA-60, Fio-máquina, etc.)
- Gera a ata formatada com: participantes, decisões, encaminhamentos e status de tarefas
- Exporta o documento final em `.docx` com formatação corporativa pronta para envio

## Tipos de reunião suportados

| Tipo | Descrição |
|------|-----------|
| Supply Aço | Reunião de Alinhamento S&OE Aço |
| Supply SIN | Reunião de Alinhamento S&OE SIN |
| Demanda | Reunião de Alinhamento S&OE (Demanda) |

## Stack

- **Python 3.10+**
- **Streamlit** — interface web
- **Google Gemini** (`gemini-2.5-pro`) — geração e estruturação da ata
- **python-docx** — exportação em Word com formatação corporativa

## Como executar

### Opção 1 — Windows (recomendado)
```
Executar App.bat
```

### Opção 2 — Terminal
```bash
pip install -r requirements.txt
streamlit run app.py
```

### Variável de ambiente necessária
```
GEMINI_API_KEY=sua_chave_aqui
```

## Estrutura
```
app.py               # Aplicação principal
requirements.txt     # Dependências
Executar App.bat     # Atalho de execução Windows
.devcontainer/       # Configuração para dev em container
```

## Contexto

Projeto desenvolvido para o time de **Planejamento Integrado (S&OP/S&OE)**, eliminando o trabalho manual de elaboração de atas e padronizando o formato dos registros executivos.