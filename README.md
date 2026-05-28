# Conferência Razão × SPED

Ferramenta de auditoria fiscal do HUB EFCT — cruzamento Triple Check entre Razão Contábil e blocos SPED (C100-EFD, A100-CONTRI, F100-CONTRI), com extração automática de NF do "Complemento Histórico", enriquecimento de CNPJ via mapa SPED, busca de Razão Social oficial + CNAEs na BrasilAPI, filtro CST 50–67 nos blocos de contribuições, coluna de análise de divergências e aba consolidada de Pendências por CNPJ.

**Arquitetura:** Opção B do manual EFCT (Flask + Render) — escolhida para processar arquivos SPED grandes (10–100 MB) sem limitações de memória do navegador.

## Stack

| Componente | Tecnologia |
|---|---|
| Frontend | HTML/CSS/JS vanilla (`index.html`) — UI no padrão visual EFCT |
| Backend | Python 3.11 + Flask 3 |
| Processamento de Excel | pandas + openpyxl |
| Enriquecimento CNPJ → CNAE | BrasilAPI (primária) + CNPJa (fallback), com cache em memória |
| Servidor de aplicação | Gunicorn (`--bind 0.0.0.0:10000 --timeout 300 --workers 1 --threads 4`) |
| Container | Docker (`python:3.11-slim`) |
| Hospedagem | Render.com |

## Como usar

1. Acesse a URL de produção (ver `HUB_METADATA.json`)
2. Faça upload dos **4 arquivos** Excel:
   - **Razão Contábil** (planilha matriz, aba "RAZÃO CONTABIL")
   - **C100-EFD FISCAL** (bloco fiscal ICMS/IPI)
   - **A100-CONTRI** (bloco serviços EFD-Contribuições)
   - **F100-CONTRI** (bloco demais operações EFD-Contribuições)
3. (Opcional) Marque "Pular busca de CNAE" se quiser auditoria mais rápida
4. Clique em **"Processar cruzamento"**
5. Aguarde 10–60s (o servidor processa, busca CNAEs em paralelo e devolve)
6. Clique em **"⬇ Baixar planilha cruzada"**

A planilha gerada tem 2 abas:
- **RAZÃO CONTABIL** — original + 9 novas colunas (NF Extraída · Razão Social · CNPJ · 1º CNAE · 2º CNAE · C100-EFD · A100-CONTRI · F100-CONTRI · Análise do Cruzamento)
- **PENDÊNCIAS DE CRUZAMENTO** — consolidação por CNPJ dos lançamentos não localizados, ordenada DESC pelo valor

## Regras de cruzamento

- **Chave de match para C100-EFD e A100**: NF Extraída + CNPJ — quando há múltiplos itens da mesma NF, **soma todos os itens** e compara o total com a Vlr Partida. Diferença &gt; R$ 100 é sinalizada como divergência na coluna Análise.
- **Chave de match para F100**: CNPJ + Período + Vlr Operação ≈ Vlr Partida (F100 não tem NF — é design do SPED para operações sem nota)
- **Filtro CST 50–67** aplicado em A100 (col 42 `CST Cofins`) e F100 (col 23 `CST Cofins`) — só lançamentos com direito a crédito entram na soma
- **Tolerância R$ 100,00** para considerar "valor batido"
- **Coluna ANÁLISE DO CRUZAMENTO**: texto enxuto — quando bate, mostra `"Encontrado em C100-EFD"`; quando diverge, mostra só o valor `R$ 380,00` (com sinal); quando não localiza, `"Não localizado nos blocos de crédito"`

## Padrão visual EFCT aplicado

- Fonte: **Exo 2** em toda a planilha
- Paleta oficial: navy `#0C2B38` + verde-limão `#B3BC2B`
- Cabeçalho originais 1-15: navy + texto lime
- Cabeçalho colunas adicionadas 16-24: lime + texto navy
- Dados em fundo branco com texto navy (destaque por cor do texto: âmbar p/ divergência, vermelho p/ alerta)
- Sem gridlines (visual limpo, profissional)

## Endpoints

| Método | Rota | Descrição |
|---|---|---|
| `GET` | `/` | Serve `index.html` |
| `GET` | `/health` | Health check para Render |
| `POST` | `/api/processar` | Recebe 4 arquivos via multipart, devolve xlsx cruzado |

Formato da resposta de erro:
```json
{ "erro": "mensagem humana" }
```

Headers especiais na resposta de sucesso:
- `Content-Disposition: attachment; filename="<nome> - CRUZADO.xlsx"`
- `X-Stat-linhas_razao`, `X-Stat-match_efd`, `X-Stat-pendencias`, etc. — estatísticas para a UI

## Como rodar localmente (dev)

```bash
pip install -r requirements.txt
python app.py
# http://localhost:10000
```

## Como rodar via Docker

```bash
docker build -t conferencia-razao-sped .
docker run -p 10000:10000 conferencia-razao-sped
# http://localhost:10000
```

## Deploy no Render

1. Push do repositório para GitHub
2. https://render.com → New Web Service → conectar o repo
3. Render detecta o `Dockerfile` automaticamente
4. Aguardar build (~3–5 min)
5. Copiar a URL gerada (ex: `https://conferencia-razao-sped.onrender.com`)

**Notas sobre o Render Free tier:**
- Cold start ~50s após 15 min de inatividade
- 512 MB RAM (suficiente pra arquivos C100-EFD de até ~30 MB)
- Para arquivos maiores ou volume alto, considerar plano Starter ($7/mês com 2 GB RAM)
