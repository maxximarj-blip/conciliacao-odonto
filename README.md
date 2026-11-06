# Conciliação de Produção Odontológica

Sistema web para conciliar a produção do sistema da clínica com as planilhas informadas por cada dentista e calcular as comissões. Tudo fica gravado com histórico.

- **Site:** GitHub Pages (arquivos estáticos deste repositório)
- **Banco, login e arquivos:** Supabase (Postgres + Auth + Storage)

---

## Publicação (cerca de 15 minutos, uma vez só)

### 1. Supabase: criar o projeto e o banco
1. Acesse [supabase.com](https://supabase.com), clique em **New project**, dê um nome (ex.: `conciliacao-odonto`) e escolha a região **South America (São Paulo)**.
2. No menu lateral, abra **SQL Editor** → **New query**. Cole o conteúdo de `supabase/schema.sql` e clique em **Run**. Deve aparecer *Success*.
3. Em **Authentication** → **Sign In / Providers**, desative **Allow new users to sign up**. Assim só entra quem você cadastrar.
4. Em **Authentication** → **Users** → **Add user** → **Create new user**, crie o seu usuário (e-mail + senha) marcando **Auto Confirm User**.
   - **O primeiro usuário criado vira administrador automaticamente.** Os próximos entram como operador.
5. Em **Project Settings** → **API** (ou **Data API**), copie a **Project URL** e a chave **anon public**.

### 2. Configurar o site
Abra `js/config.js` e cole os dois valores:

```js
window.APP_CONFIG = {
  supabaseUrl: 'https://xxxxxxxx.supabase.co',
  supabaseAnonKey: 'eyJhbGciOi...',
  clinica: 'Odonto Premier',
};
```

> A chave **anon public** pode ficar no site: os dados só abrem com login, e as regras do banco (RLS) protegem o resto. **Nunca** coloque a chave `service_role` aqui.

### 3. GitHub: publicar
1. Crie um repositório novo no GitHub (ex.: `conciliacao-odonto`).
2. Envie todos os arquivos desta pasta, mantendo a estrutura (`index.html`, `js/`, `supabase/`, `.nojekyll`). Pela web dá para fazer em **Add file** → **Upload files**, arrastando a pasta inteira.
3. Em **Settings** → **Pages**, configure **Source: Deploy from a branch**, **Branch: `main` / (root)** e clique em **Save**.
4. Em 1 ou 2 minutos o site fica disponível em `https://SEU-USUARIO.github.io/conciliacao-odonto/`.

> GitHub Pages em repositório **privado** exige plano pago (Pro/Team). Em repositório público o código fica visível, mas os dados não: eles ficam no Supabase, atrás do login.

### 4. (Recomendado) liberar o domínio no Supabase
Em **Authentication** → **URL Configuration**, coloque o endereço do GitHub Pages em **Site URL**.

---

## Usuários e permissões

| Papel | Pode |
|---|---|
| **admin** | tudo, inclusive **reabrir** uma competência fechada (com motivo obrigatório) |
| **operador** | cadastrar, subir planilhas, decidir divergências, salvar versões e fechar a competência |

### Solicitar acesso e Esqueci minha senha
Na tela de login há duas opções:
- **Solicitar acesso:** a pessoa informa nome, e-mail e função.
- **Esqueci minha senha:** a pessoa informa o e-mail.

Cada pedido fica gravado na tabela `solicitacoes`. O aviso por e-mail é enviado pelo próprio Supabase (Edge Function + Resend). Veja **Aviso por e-mail** abaixo. Limite: 3 pedidos por e-mail a cada 24h.

O admin vê os pedidos no botão **Solicitações** (no topo do sistema):
- **Acesso:** o painel mostra os dados e uma senha gerada. Crie o usuário no Supabase (**Authentication → Users → Add user**) e use **Copiar mensagem p/ enviar** para avisar a pessoa.
- **Senha:** o painel já traz o SQL de redefinição pronto. Rode no **SQL Editor** e envie a mensagem.
- Depois, marque o pedido como **atendido** ou **recusado**. Isso fica na auditoria.

### Aviso por e-mail (Resend + Supabase), configurar uma vez
1. Crie uma conta grátis em **resend.com** usando o e-mail que vai **receber** os avisos (ex.: maxximarj@gmail.com). No plano grátis, sem domínio próprio, o Resend só envia para o e-mail da própria conta.
2. No Resend: **API Keys → Create API Key** (permissão *Sending access*) e copie a chave `re_...`.
3. No Supabase: **Edge Functions → Deploy a new function → Via Editor**, nome `notificar-solicitacao`. Cole o conteúdo de `supabase/functions/notificar-solicitacao/index.ts` e clique em **Deploy**.
4. No Supabase: **Edge Functions → Secrets**, adicione:
   - `RESEND_API_KEY` = a chave `re_...`
   - `EMAIL_DESTINO` = maxximarj@gmail.com
   - `SITE_URL` = endereço do site no GitHub Pages (opcional)
5. No Supabase: **Database → Webhooks → Create a new hook** (se pedir, clique em *Enable webhooks* antes):
   - Name: `aviso_solicitacao`
   - Table: `solicitacoes` · Events: **Insert**
   - Type: **Supabase Edge Functions** → `notificar-solicitacao` · Method **POST**
   - Em HTTP Headers, aceite o *Authorization* que o Supabase sugere (service role)
   - **Create webhook**
6. Teste com um "Esqueci minha senha" na tela de login. Se o e-mail não chegar, veja os logs em **Edge Functions → notificar-solicitacao → Logs**.

Para tornar alguém admin, rode no SQL Editor:
```sql
update public.perfis set papel = 'admin' where email = 'pessoa@clinica.com';
```

---

## Como o histórico funciona

- **Uma conciliação por competência (mês).** Tudo é salvo automaticamente: planilhas originais (no Storage), decisões, regras e totais.
- **Salvar versão** grava uma foto do momento, com observação.
- **Fechar competência** trava tudo e grava as comissões em `comissoes_fechadas`, junto com o detalhamento das linhas.
- **Reabrir** (só admin) exige motivo. O fechamento anterior continua guardado, e o novo fechamento vira a versão vigente.
- **Nada é apagado.** Dentistas e especialidades são *inativados*. Planilhas substituídas ou removidas continuam no Storage. Versões, comissões fechadas e auditoria não aceitam alteração nem exclusão.
- **Auditoria:** toda alteração em cadastros, conciliações e arquivos fica registrada com usuário, data, valor anterior e valor novo.

### Tabelas principais (Supabase → Table Editor)
| Tabela | Conteúdo |
|---|---|
| `dentistas`, `especialidades` | cadastros |
| `conciliacoes` | uma linha por competência (status, decisões, totais) |
| `conciliacao_arquivos` | planilhas enviadas (ativas e substituídas) |
| `conciliacao_versoes` | versões imutáveis (salvamento, fechamento, reabertura) |
| `comissoes_fechadas` / `comissoes_vigentes` | comissões por dentista e mês (todas / só o último fechamento) |
| `auditoria` | log de todas as alterações |

---

## Regras da conciliação

1. Guia informada por um dentista → **ele é o executor**, mesmo que o sistema aponte outro.
2. Guia informada por mais de um dentista → **conflito**, o usuário escolhe.
3. Lançamento sem guia válida, só com o ID do atendimento → atribui se o valor bater (tolerância). Se não bater, vira **divergência**.
4. Sem planilha de dentista → usa o dentista do sistema, se estiver cadastrado.
5. Divergências de valor e lançamentos não localizados são decididos pelo usuário e entram na comissão como ajustes.

Comissão: **percentual** sobre a base (situações marcadas), **diária fixa** × dias com atendimento, ou **valor fixo** mensal. O padrão vem da especialidade, e o cadastro do dentista pode sobrescrever.

---

## Atualizações futuras
Para publicar uma nova versão do site, basta substituir os arquivos no GitHub. O banco não muda. Se uma atualização trouxer mudanças no `schema.sql`, rode o arquivo de novo no SQL Editor: ele pode ser executado várias vezes sem perder dados.
