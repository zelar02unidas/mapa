# Mapa Zelar — Ativos por Unidade

Painel interativo para visualizar os ativos instalados por unidade:
mapa com agrupamento por região, filtros por cliente, unidade, segmento,
estado, consultor, quantidade de ativos, distância e tempo desde a última
visita, além de tabela ordenável com exportação para CSV.

## Como funciona a atualização dos dados

O site tenta baixar a planilha **ao vivo** a cada visita, nesta ordem:

1. `https://raw.githubusercontent.com/SEU-USUARIO/SEU-REPOSITORIO/main/Mapa - Zelar.xlsx`
2. `https://cdn.jsdelivr.net/gh/SEU-USUARIO/SEU-REPOSITORIO@main/Mapa - Zelar.xlsx`

Basta atualizar o arquivo `Mapa - Zelar.xlsx` no repositório que o site
mostra os dados novos automaticamente (não precisa publicar nada, nem
atualizar arquivos gerados).

Se nenhuma das duas URLs responder, o site usa uma **cópia embutida**
(`js/data.js`) dos dados e avisa na tela.

## Como publicar (GitHub Pages)

1. Crie um repositório **público** no GitHub (ex.: `mapa-zelar`).
2. Envie todos os arquivos desta pasta para a raiz do repositório,
   **incluindo a planilha** `Mapa - Zelar.xlsx`.
   - Sem Git? Na página do repositório use *Add file → Upload files*.
3. Habilite o site: *Settings → Pages → Source: Deploy from a branch →
   Branch: `main` → Save*.
4. Edite o arquivo `js/config.js` e preencha as duas variáveis:

   ```js
   GITHUB_USER: "seu-usuario",      // dono do repositório
   GITHUB_REPO: "mapa-zelar"        // nome do repositório
   ```

5. Aguarde alguns minutos e acesse
   `https://SEU-USUARIO.github.io/SEU-REPOSITORIO/`.

Pronto. Na próxima atualização da planilha no GitHub, o site já carrega
os dados novos sozinho.

## Estrutura do projeto

| Arquivo                 | Função                                              |
| ----------------------- | --------------------------------------------------- |
| `index.html`            | Página principal                                    |
| `css/style.css`         | Estilos (inclui tema claro/escuro)                  |
| `js/config.js`          | Suas credenciais do repositório (edite aqui)        |
| `js/app.js`             | Lógica do site (mapa, filtros, tabela, exportação)  |
| `js/geo.js`             | Conversão cidade + localização em latitude/longitude|
| `js/data.js`            | Cópia embutida dos dados (plano B automático)       |
| `lib/`                  | Bibliotecas de terceiros (Leaflet, SheetJS, etc.)   |
| `Mapa - Zelar.xlsx`     | A planilha oficial com os dados                     |

## Como regenerar a cópia embutida (opcional)

A cópia em `js/data.js` só precisa ser regenerada se você quiser que o
site funcione mesmo sem internet/GitHub. Requer Node.js e a planilha na
raiz do projeto:

```bash
npm install --no-save xlsx open-location-code
node tools/build.mjs
```

## Observações sobre a planilha

- A aba usada é **Mapa** (ou a primeira aba encontrada).
- A coluna **LOCALIZAÇÃO** (código de área aberta) dá a precisão máxima
  do ponto no mapa. Sem ela, o ponto é aproximado pelo centro do
  município (os cartões no mapa mostram quando é aproximado).
- Datas são escritas como mês/ano (ex.: `JUL/24`).

## Teste local

```bash
python3 -m http.server 8899
# acesse http://localhost:8899/index.html
```

> Sem abrir pelo endereço `localhost` (por ex. dando duplo clique no
> arquivo), as bibliotecas de mapa não funcionam.

## Mapa dos arquivos

- `Captura_de_tela_*.png` — exemplos de tela (podem ser apagados).
