# PDF Livre

Converte PDF, PPTX travado/achatado e imagens de slide em apresentações
`.pptx` totalmente editáveis (texto, formas, tabelas e imagens como objetos
nativos do PowerPoint, não como imagem raster).

Projeto da SEAD-GO / SCTP-TransformaLAB.

## Como funciona

O backend decide automaticamente a melhor rota para cada arquivo:

1. **Trilha A — liberar**: PPTX nativo travado/agrupado. Só destrava e
   desagrupa; nunca perde fidelidade (checado por SSIM ≥ 0.99).
2. **Trilha C — vetorial**: PDF ou PPTX com texto vivo (extraído
   diretamente, sem OCR). Reconstrução fiel, alta precisão.
3. **Trilha D — visão**: imagem achatada ou PDF escaneado, sem texto vivo.
   Duas variantes:
   - **Convencional** (padrão, sem custo): OpenCV + Tesseract em
     português.
   - **IA** (opcional): modelo de visão Claude, ativado somente quando a
     variável `ANTHROPIC_API_KEY` está definida. Tem teto de gasto diário
     configurável e cai de volta ao modo convencional automaticamente se a
     chave falhar ou o teto for atingido — a conversão nunca trava por
     causa da IA.

Ver `backend/Requisitos_funcionais.docx` (RF01–RF27, RNF01–RNF25) para a
especificação completa que orientou o desenho.

## Estrutura

```
backend/          API FastAPI (Python) — todo o processamento
  app/
    main.py        rotas HTTP e orquestração dos jobs
    unlock.py       Trilha A
    pdf2pptx.py      Trilha C
    vision_cv.py      Trilha D convencional (OpenCV/Tesseract)
    vision_ai.py       Trilha D com IA (opcional, com teto de gasto)
    rebuild.py       monta o .pptx final a partir das regiões detectadas
    audit.py          calcula o Índice de Editabilidade
    normalize.py      normaliza qualquer formato de entrada em imagens/PDF
  testes/          smoke tests (SSIM, roteamento)
  Dockerfile       imagem com LibreOffice + Tesseract-por + Poppler
supabase/
  migrations/      histórico de conversões + buckets de Storage (RLS por usuário)
```

## Rodando localmente

```bash
cd backend
python -m venv .venv && source .venv/bin/activate   # ou .venv\Scripts\activate no Windows
pip install -r requirements.txt
cp .env.example .env    # preencha o que for usar (tudo opcional exceto o básico)
uvicorn app.main:app --reload
```

Abra `http://localhost:8000`.

Sem preencher `ANTHROPIC_API_KEY`, o app funciona 100% no modo convencional
— é o padrão recomendado até a validação com o TI/segurança.

### Ativando a IA de visão (opcional)

1. Gere uma chave em <https://console.anthropic.com/settings/keys>.
2. Coloque-a em `backend/.env` como `ANTHROPIC_API_KEY=sk-ant-...`
   (nunca em um chat, commit ou issue — se uma chave real vazar em algum
   lugar, revogue-a imediatamente no console e gere outra).
3. Ajuste `PPTLIVRE_TETO_DIARIO_USD` conforme o orçamento aceitável.

## Segurança e dados (RNF16/RNF18/RNF19)

- O arquivo original é apagado do servidor assim que o processamento
  termina.
- O resultado fica em uma pasta isolada por identificador aleatório,
  não listável, e expira em 20 minutos.
- Limite de 8 requisições de conversão por minuto por IP.
- **O deploy atual ainda não tem autenticação** — o link, quando
  publicado, é de acesso restrito por enquanto (uso combinado), com
  camada de login (Supabase Auth) planejada como próximo passo antes de
  uma divulgação mais ampla.

## Próximos passos

- Integrar Supabase (histórico de conversões + Storage) no backend Python.
- Login via Supabase Auth.
- Deploy (Render/Fly.io/Railway, já que não há acesso à infraestrutura da
  SEAD para hospedagem interna).
