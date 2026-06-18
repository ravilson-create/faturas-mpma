# Consulta de Faturas de Energia — MPMA (versão hospedada)

Este app fica acessível por um link de internet, para qualquer pessoa que
você compartilhar o endereço. Os dados são salvos num banco compartilhado
(Supabase) — todo mundo que acessar o link vê o mesmo histórico.

Vamos usar três serviços, todos gratuitos no volume que você vai usar:
- **GitHub** — guarda o código do projeto (obrigatório para o próximo passo)
- **Supabase** — banco de dados na nuvem, onde as faturas ficam salvas
- **Vercel** — hospeda o site e gera o link público

Vai levar entre 20 e 40 minutos na primeira vez. Depois de configurado,
qualquer atualização futura no app é só enviar o código de novo.

---

## Parte 1 — Criar o banco de dados (Supabase)

1. Acesse https://supabase.com e clique em **Start your project**
2. Crie uma conta (pode entrar com GitHub, Google, ou e-mail)
3. Clique em **New project**
   - Dê um nome, ex: `faturas-mpma`
   - Crie uma senha de banco de dados (anote em algum lugar seguro, mas você
     não vai precisar dela diretamente — é gerenciada pelo Supabase)
   - Escolha a região mais próxima (ex: South America - São Paulo)
   - Clique em **Create new project** e espere ~2 minutos enquanto provisiona

4. No menu lateral esquerdo, clique em **SQL Editor**
5. Clique em **New query**
6. Abra o arquivo `supabase_setup.sql` (incluído neste pacote), copie todo o
   conteúdo, cole no editor, e clique em **Run** (ou Ctrl+Enter)
   - Isso cria a tabela `faturas` onde os dados ficam guardados
   - Se aparecer "Success. No rows returned", funcionou

7. No menu lateral, clique em **Settings** (ícone de engrenagem) → **API**
8. Você vai precisar de dois valores desta tela, copie e guarde:
   - **Project URL** (algo como `https://abcdefgh.supabase.co`)
   - **anon public** key (uma string longa começando com `eyJ...`)

---

## Parte 2 — Criar conta no GitHub e subir o código

1. Acesse https://github.com e clique em **Sign up** (se ainda não tem conta)
2. Depois de criar a conta, clique no `+` no canto superior direito → **New repository**
3. Dê um nome (ex: `faturas-mpma`), deixe como **Public** ou **Private**
   (tanto faz, a Vercel funciona com os dois), e clique em **Create repository**
4. Na tela que aparece, clique em **uploading an existing file**
5. Arraste TODOS os arquivos da pasta `project/` (deste pacote) para a área
   de upload — exceto o arquivo `.env.example`, esse não precisa subir
6. Role para baixo e clique em **Commit changes**

Pronto, seu código está no GitHub.

---

## Parte 3 — Hospedar o site (Vercel)

1. Acesse https://vercel.com e clique em **Sign Up**
2. Escolha **Continue with GitHub** (mais simples — conecta direto com a
   conta que você acabou de criar)
3. Autorize a Vercel a acessar seus repositórios do GitHub
4. No painel da Vercel, clique em **Add New** → **Project**
5. Encontre o repositório `faturas-mpma` que você criou e clique em **Import**
6. Na tela de configuração, antes de clicar em Deploy, abra a seção
   **Environment Variables** e adicione duas variáveis:

   | Name | Value |
   |------|-------|
   | `VITE_SUPABASE_URL` | a Project URL que você copiou do Supabase |
   | `VITE_SUPABASE_ANON_KEY` | a anon public key que você copiou do Supabase |

7. Clique em **Deploy**
8. Espere 1 a 2 minutos. Quando terminar, a Vercel mostra um link, algo como
   `https://faturas-mpma.vercel.app` — esse é o endereço público do seu app

Pronto. Esse link já funciona para qualquer pessoa, em qualquer navegador,
sem precisar instalar nada.

---

## Como usar depois de hospedado

Abra o link da Vercel, envie a planilha Excel da Equatorial no campo de
upload, e os dados ficam salvos no Supabase — visíveis para qualquer pessoa
que acessar o mesmo link, no mês atual e nos seguintes.

---

## Atualizando o app no futuro

Se você (ou eu, numa conversa futura) precisar alterar alguma coisa no
código:
1. Edite os arquivos
2. Suba a versão nova no mesmo repositório do GitHub (substituindo os
   arquivos antigos, do mesmo jeito que no passo de upload)
3. A Vercel detecta a mudança automaticamente e atualiza o site sozinha,
   em menos de um minuto — sem precisar fazer nada na Vercel

---

## Sobre segurança e acesso

Este app não tem tela de login. Qualquer pessoa com o link consegue ver os
dados e enviar novas planilhas. Para o seu caso (uso interno, link não
divulgado publicamente), isso costuma ser aceitável, mas vale ter em mente:

- Não publique o link em lugares públicos (redes sociais, etc.)
- Se no futuro precisar de controle de acesso (senha, login por e-mail),
  o Supabase tem esse recurso embutido (Authentication) — é um passo
  adicional que podemos implementar quando for necessário
- Os dados trafegam e ficam salvos com a proteção padrão do Supabase
  (conexão criptografada), mas a política de acesso atual ("Permitir
  leitura pública") significa que não há autenticação de quem lê ou grava

## Problemas comuns

**Tela em branco ou erro no console do navegador (F12)** — confira se as
duas variáveis de ambiente (`VITE_SUPABASE_URL` e `VITE_SUPABASE_ANON_KEY`)
foram configuradas corretamente na Vercel, sem espaços extras.

**"Não foi possível carregar os dados do banco"** — geralmente significa que
o script SQL (`supabase_setup.sql`) não foi executado, ou a tabela `faturas`
não existe ainda. Volte à Parte 1, passo 6.

**Upload da planilha não aparece para outras pessoas** — confirme que todos
estão acessando o mesmo link da Vercel (não um link de preview antigo).
