# Guia Completo — CodeCraft Dashboard no GitHub Pages

## SEU LINK FINAL SERÁ:
https://adrianac3gestao-lgtm.github.io/codecraft-dashboard

---

## PASSO 1 — Criar o repositório no GitHub (fazer só uma vez)

1. Acesse https://github.com e faça login
2. Clique no botão verde **"New"** (canto superior esquerdo)
3. Preencha:
   - Repository name: `codecraft-dashboard`
   - Marque: **Public**
   - NÃO marque nenhuma outra opção
4. Clique em **"Create repository"**

---

## PASSO 2 — Instalar o Git no computador (fazer só uma vez)

1. Acesse: https://git-scm.com/download/win
2. Baixe e instale com todas as opções PADRÃO (só clicar Next)
3. Após instalar, abra o **Git Bash** (aparece no menu Iniciar)
4. Execute estes 2 comandos (substitua pelo seu e-mail do Gmail):

   git config --global user.name "Adriana C3 Gestao"
   git config --global user.email "seuemail@gmail.com"

---

## PASSO 3 — Conectar a pasta do dashboard ao GitHub (fazer só uma vez)

Abra o Git Bash e execute os comandos abaixo UM POR VEZ:

   cd "C:/Users/adria/OneDrive/C3 Gestao/EquipeC3/CODECRAFT/7. Dashboard"

   git init

   git add .

   git commit -m "primeiro envio"

   git branch -M main

   git remote add origin https://github.com/adrianac3gestao-lgtm/codecraft-dashboard.git

   git push -u origin main

Quando pedir login, use seu usuário e senha do GitHub.

---

## PASSO 4 — Ativar o GitHub Pages (fazer só uma vez)

1. Acesse: https://github.com/adrianac3gestao-lgtm/codecraft-dashboard
2. Clique em **Settings** (engrenagem)
3. No menu esquerdo clique em **Pages**
4. Em "Source" selecione: **Deploy from a branch**
5. Em "Branch" selecione: **main** e pasta **/ (root)**
6. Clique em **Save**
7. Aguarde 2 minutos

SEU LINK ESTARÁ ATIVO:
https://adrianac3gestao-lgtm.github.io/codecraft-dashboard

---

## PASSO 5 — Renomear o arquivo HTML (fazer só uma vez)

O GitHub Pages abre automaticamente um arquivo chamado index.html
Renomeie o arquivo na sua pasta:
   codecraft_dashboard.html  →  index.html

Depois execute no Git Bash:
   cd "C:/Users/adria/OneDrive/C3 Gestao/EquipeC3/CODECRAFT/7. Dashboard"
   git add .
   git commit -m "renomear para index"
   git push

---

## ROTINA DIÁRIA — Atualizar o dashboard

Basta dar DUPLO CLIQUE no arquivo:
   ATUALIZAR_DASHBOARD.bat

Ele vai fazer tudo automaticamente:
1. Rodar o Python e atualizar os dados do Excel
2. Enviar para o GitHub
3. Em 1 minuto o link do cliente já mostra os dados novos

