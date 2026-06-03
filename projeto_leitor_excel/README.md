# Motor de Escoragem — Visualizador e Exportador

Aplicação web para visualizar e exportar configurações de motor de escoragem em Excel.

## Funcionalidades

- Upload de arquivo JSON de escoragem
- Visualização das variáveis, regras e pontuações diretamente no navegador
- Visualização da classificação de risco com faixas de pontuação
- Exportação para Excel (.xlsx) com layout profissional

## Deploy — Streamlit Community Cloud (gratuito)

### 1. Subir o projeto no GitHub

```bash
git init
git add .
git commit -m "primeiro commit"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
git push -u origin main
```

### 2. Fazer o deploy

1. Acesse **[share.streamlit.io](https://share.streamlit.io)**
2. Faça login com sua conta GitHub
3. Clique em **New app**
4. Selecione o repositório e a branch `main`
5. Em **Main file path**, coloque `app.py`
6. Clique em **Deploy**

A aplicação ficará disponível em uma URL pública em poucos minutos.

## Rodar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Estrutura

```
├── app.py                  # Interface Streamlit
├── exportar_escoragem.py   # Logica de parsing e geracao do Excel
├── requirements.txt        # Dependencias
└── .streamlit/
    └── config.toml         # Configuracao do tema
```
