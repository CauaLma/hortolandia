# CAD 2026 Dashboard — pronto para Render

Dashboard web com backend Python e leitura da planilha `data/cad2026.xlsx`.

## Publicar no Render

1. Crie um repositório no GitHub e envie todos os arquivos desta pasta.
2. No Render, clique em **New +** > **Blueprint**.
3. Conecte o repositório.
4. O Render vai ler o `render.yaml` automaticamente.
5. Clique em **Apply** / **Deploy**.

Também funciona criando como **Web Service** manual:

- Environment: `Python`
- Build Command: `pip install -r requirements.txt`
- Start Command: `python app.py`

## Rodar local

```cmd
py -m pip install -r requirements.txt
py app.py
```

Depois abra:

```txt
http://localhost:5052
```
