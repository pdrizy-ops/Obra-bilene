# Casa Lodge · Bilene — Controlo de Obra (PWA)

App instalável (PWA) para controlo de fases, materiais e custos da obra Casa Lodge, em
Bilene. Não precisa de build — corre directamente no browser a partir destes ficheiros.

## Publicar no GitHub Pages (obter um link)

**Opção A — pela interface do GitHub (sem usar o terminal)**

1. Vai a [github.com/new](https://github.com/new) e cria um repositório novo, por exemplo
   `casa-lodge-obra` (pode ser público ou privado — Pages funciona em ambos nos planos
   pagos; nos gratuitos só em repositórios públicos).
2. Abre o repositório criado → **Add file → Upload files**.
3. Arrasta **todos** os ficheiros desta pasta (`index.html`, `app.jsx`, `manifest.json`,
   `service-worker.js`, `README.md` e a pasta `icons/` completa) e faz **Commit changes**.
4. Vai a **Settings → Pages** (barra lateral esquerda).
5. Em **Source**, escolhe **Deploy from a branch** → branch **main** → pasta **/ (root)**
   → **Save**.
6. Espera 1-2 minutos. O GitHub mostra o link no topo da mesma página, algo como:
   `https://o-teu-utilizador.github.io/casa-lodge-obra/`

**Opção B — pelo terminal (git)**

```bash
cd casa-lodge-pwa
git init
git add .
git commit -m "Casa Lodge — app de controlo de obra"
git branch -M main
git remote add origin https://github.com/O-TEU-UTILIZADOR/casa-lodge-obra.git
git push -u origin main
```

Depois activa o Pages tal como nos passos 4-5 acima.

## Instalar como app

Abre o link em Chrome ou Safari no telemóvel/computador. Vai aparecer a opção
**"Instalar app"** (ou **"Adicionar ao ecrã principal"** no iPhone, através do botão de
partilha). A app fica com ícone próprio e abre em ecrã inteiro, sem barra do browser.

## Onde ficam os dados

Todos os dados (fases, materiais, custos, progresso) ficam guardados no `localStorage`
do próprio dispositivo/browser onde a app é usada — não há servidor nem base de dados.
Isto significa:

- Os dados **não são partilhados automaticamente** entre o teu telemóvel e o computador —
  cada dispositivo tem a sua própria cópia.
- Limpar os dados do browser (ou desinstalar a app) apaga o progresso registado nesse
  dispositivo.
- Usa o botão **"Exportar Planilha"** regularmente para guardar uma cópia em `.xlsx` —
  é também a forma mais simples de levar os dados de um dispositivo para outro.

## Actualizar a app depois de publicada

Sempre que quiseres alterar algo (novo componente por defeito, cor, texto), edita os
ficheiros e volta a fazer commit/upload no mesmo repositório — o GitHub Pages actualiza
o link automaticamente em 1-2 minutos.

## Estrutura dos ficheiros

```
casa-lodge-pwa/
├── index.html          → página principal, carrega React/Babel/XLSX via CDN
├── app.jsx             → toda a aplicação (React, sem passo de build)
├── manifest.json        → identidade da PWA (nome, ícone, cor)
├── service-worker.js    → cache para funcionamento offline do essencial da app
├── icons/                → ícones da app (192px, 512px, apple-touch, favicon)
└── README.md
```
