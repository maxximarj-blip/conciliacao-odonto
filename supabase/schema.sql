-- =====================================================================
--  Conciliação de Produção Odontológica — schema Supabase
--  Rode este arquivo inteiro no Supabase: SQL Editor → New query → Run
--  Pode rodar de novo sem perder dados (idempotente).
--
--  Histórico:
--   • nada é apagado: dentistas/especialidades são inativados; arquivos
--     substituídos ficam guardados; versões e comissões fechadas são imutáveis
--   • toda alteração em cadastros e conciliações vai para a tabela auditoria
-- =====================================================================

create extension if not exists pgcrypto;

-- ---------------------------------------------------------------------
-- Perfis de usuário (criado automaticamente a cada novo usuário do Auth)
-- ---------------------------------------------------------------------
create table if not exists public.perfis (
  id          uuid primary key references auth.users(id) on delete cascade,
  nome        text,
  email       text,
  papel       text not null default 'operador' check (papel in ('admin','operador')),
  criado_em   timestamptz not null default now()
);

create or replace function public.tg_novo_usuario() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  insert into public.perfis (id, email, nome, papel)
  values (new.id, new.email, coalesce(new.raw_user_meta_data->>'nome', split_part(new.email,'@',1)),
          case when not exists (select 1 from public.perfis) then 'admin' else 'operador' end)
  on conflict (id) do nothing;
  return new;
end $$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created after insert on auth.users
  for each row execute function public.tg_novo_usuario();

create or replace function public.eh_admin() returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.perfis where id = auth.uid() and papel = 'admin');
$$;

-- ---------------------------------------------------------------------
-- Cadastros
-- ---------------------------------------------------------------------
create table if not exists public.especialidades (
  id            uuid primary key default gen_random_uuid(),
  nome          text not null,
  forma         text not null default 'percentual' check (forma in ('percentual','diaria','fixo')),
  valor         numeric(12,2) not null default 0 check (valor >= 0),
  ativo         boolean not null default true,
  criado_em     timestamptz not null default now(),
  atualizado_em timestamptz not null default now(),
  atualizado_por uuid default auth.uid()
);

create table if not exists public.dentistas (
  id               uuid primary key default gen_random_uuid(),
  nome             text not null,
  aliases          text,
  especialidade_id uuid references public.especialidades(id),
  forma            text check (forma in ('percentual','diaria','fixo')),   -- null = padrão da especialidade
  valor            numeric(12,2) check (valor >= 0),                         -- null = padrão da especialidade
  ativo            boolean not null default true,
  criado_em        timestamptz not null default now(),
  atualizado_em    timestamptz not null default now(),
  atualizado_por   uuid default auth.uid()
);
create unique index if not exists dentistas_nome_ativo_uk on public.dentistas (lower(nome)) where ativo;

-- ---------------------------------------------------------------------
-- Conciliações (uma por competência)
-- ---------------------------------------------------------------------
create table if not exists public.conciliacoes (
  id              uuid primary key default gen_random_uuid(),
  competencia     date not null unique check (extract(day from competencia) = 1),
  status          text not null default 'rascunho' check (status in ('rascunho','fechada')),
  arquivo_sistema text,
  regras          jsonb not null default '{}'::jsonb,   -- situações da base, tolerância
  decisoes        jsonb not null default '{}'::jsonb,   -- overrides, resoluções, lançamentos, diárias
  totais          jsonb not null default '{}'::jsonb,   -- último cálculo (para a lista do histórico)
  versao          int  not null default 0,
  criado_por      uuid default auth.uid(),
  criado_em       timestamptz not null default now(),
  atualizado_por  uuid default auth.uid(),
  atualizado_em   timestamptz not null default now(),
  fechada_por     uuid,
  fechada_em      timestamptz
);

create table if not exists public.conciliacao_arquivos (
  id              uuid primary key default gen_random_uuid(),
  conciliacao_id  uuid not null references public.conciliacoes(id),
  tipo            text not null check (tipo in ('sistema','dentista')),
  nome            text not null,
  dentista_id     uuid references public.dentistas(id),
  storage_path    text not null,
  tamanho         bigint,
  ativo           boolean not null default true,        -- false = substituído/removido (fica guardado)
  enviado_por     uuid default auth.uid(),
  enviado_em      timestamptz not null default now()
);
create index if not exists conc_arq_conc_idx on public.conciliacao_arquivos (conciliacao_id);

-- versões imutáveis (salvamento manual, fechamento, reabertura)
create table if not exists public.conciliacao_versoes (
  id              bigint generated always as identity primary key,
  conciliacao_id  uuid not null references public.conciliacoes(id),
  versao          int not null,
  tipo            text not null check (tipo in ('salvamento','fechamento','reabertura')),
  regras          jsonb,
  decisoes        jsonb,
  totais          jsonb,
  resumo          jsonb,       -- comissão por dentista
  detalhe         jsonb,       -- linhas conciliadas (no fechamento)
  arquivos        jsonb,       -- arquivos ativos naquele momento
  observacao      text,
  salvo_por       uuid default auth.uid(),
  salvo_por_email text,
  salvo_em        timestamptz not null default now()
);
create index if not exists conc_ver_conc_idx on public.conciliacao_versoes (conciliacao_id, versao);

-- comissões de cada fechamento (para relatórios mês a mês)
create table if not exists public.comissoes_fechadas (
  id              bigint generated always as identity primary key,
  conciliacao_id  uuid not null references public.conciliacoes(id),
  competencia     date not null,
  versao          int not null,
  dentista_id     uuid references public.dentistas(id),
  dentista_nome   text not null,
  especialidade   text,
  forma           text,
  valor_ref       numeric(12,2),
  guias           int,
  procedimentos   int,
  producao        numeric(14,2),
  ajustes         numeric(14,2),
  base            numeric(14,2),
  diarias         int,
  comissao        numeric(14,2),
  fechada_em      timestamptz not null default now()
);
create index if not exists com_fech_comp_idx on public.comissoes_fechadas (competencia, dentista_id);

-- ---------------------------------------------------------------------
-- Auditoria: toda alteração em cadastros e conciliações
-- ---------------------------------------------------------------------
create table if not exists public.auditoria (
  id            bigint generated always as identity primary key,
  tabela        text not null,
  registro_id   uuid,
  acao          text not null,
  antes         jsonb,
  depois        jsonb,
  usuario       uuid default auth.uid(),
  usuario_email text,
  em            timestamptz not null default now()
);
create index if not exists auditoria_reg_idx on public.auditoria (tabela, registro_id, em desc);

create or replace function public.tg_auditoria() returns trigger
language plpgsql security definer set search_path = public as $$
declare v_email text;
begin
  select email into v_email from public.perfis where id = auth.uid();
  if tg_op = 'UPDATE' then
    if to_jsonb(old) - 'atualizado_em' = to_jsonb(new) - 'atualizado_em' then return new; end if;
    new.atualizado_em := now();
    new.atualizado_por := auth.uid();
  end if;
  insert into public.auditoria (tabela, registro_id, acao, antes, depois, usuario_email)
  values (tg_table_name,
          case when tg_op = 'DELETE' then old.id else new.id end,
          lower(tg_op),
          case when tg_op = 'INSERT' then null else to_jsonb(old) end,
          case when tg_op = 'DELETE' then null else to_jsonb(new) end,
          v_email);
  return case when tg_op = 'DELETE' then old else new end;
end $$;

create or replace function public.tg_auditoria_arquivos() returns trigger
language plpgsql security definer set search_path = public as $$
declare v_email text;
begin
  select email into v_email from public.perfis where id = auth.uid();
  insert into public.auditoria (tabela, registro_id, acao, antes, depois, usuario_email)
  values ('conciliacao_arquivos', new.id, lower(tg_op), case when tg_op='UPDATE' then to_jsonb(old) end, to_jsonb(new), v_email);
  return new;
end $$;

drop trigger if exists aud_especialidades on public.especialidades;
create trigger aud_especialidades before insert or update on public.especialidades for each row execute function public.tg_auditoria();
drop trigger if exists aud_dentistas on public.dentistas;
create trigger aud_dentistas before insert or update on public.dentistas for each row execute function public.tg_auditoria();
drop trigger if exists aud_conciliacoes on public.conciliacoes;
create trigger aud_conciliacoes before insert or update on public.conciliacoes for each row execute function public.tg_auditoria();
drop trigger if exists aud_arquivos on public.conciliacao_arquivos;
create trigger aud_arquivos after insert or update on public.conciliacao_arquivos for each row execute function public.tg_auditoria_arquivos();

-- conciliação fechada não pode ser alterada (só reabrindo pela função)
create or replace function public.tg_trava_fechada() returns trigger
language plpgsql as $$
begin
  if old.status = 'fechada' and coalesce(current_setting('app.liberar_fechada', true), '') <> 'sim' then
    raise exception 'Competência fechada: reabra para alterar.';
  end if;
  return new;
end $$;
drop trigger if exists trava_fechada on public.conciliacoes;
create trigger trava_fechada before update on public.conciliacoes for each row execute function public.tg_trava_fechada();

create or replace function public.tg_trava_arquivo_fechada() returns trigger
language plpgsql as $$
begin
  if exists (select 1 from public.conciliacoes c where c.id = new.conciliacao_id and c.status = 'fechada') then
    raise exception 'Competência fechada: reabra para alterar os arquivos.';
  end if;
  return new;
end $$;
drop trigger if exists trava_arquivo_fechada on public.conciliacao_arquivos;
create trigger trava_arquivo_fechada before insert or update on public.conciliacao_arquivos for each row execute function public.tg_trava_arquivo_fechada();

-- ---------------------------------------------------------------------
-- Funções de negócio (chamadas pelo app via rpc)
-- ---------------------------------------------------------------------
create or replace function public._arquivos_ativos(p_id uuid) returns jsonb
language sql stable security definer set search_path = public as $$
  select coalesce(jsonb_agg(jsonb_build_object('id', id, 'tipo', tipo, 'nome', nome, 'dentista_id', dentista_id, 'storage_path', storage_path) order by enviado_em), '[]'::jsonb)
  from public.conciliacao_arquivos where conciliacao_id = p_id and ativo;
$$;

-- salva uma versão (foto do momento), sem fechar
create or replace function public.salvar_versao(p_id uuid, p_totais jsonb, p_resumo jsonb, p_observacao text default null)
returns int language plpgsql security definer set search_path = public as $$
declare v int; c public.conciliacoes;
begin
  if auth.uid() is null then raise exception 'Não autenticado'; end if;
  select * into c from public.conciliacoes where id = p_id for update;
  if not found then raise exception 'Conciliação não encontrada'; end if;
  if c.status = 'fechada' then raise exception 'Competência fechada.'; end if;
  update public.conciliacoes set versao = versao + 1, totais = coalesce(p_totais, totais) where id = p_id returning versao into v;
  insert into public.conciliacao_versoes (conciliacao_id, versao, tipo, regras, decisoes, totais, resumo, arquivos, observacao, salvo_por_email)
  values (p_id, v, 'salvamento', c.regras, c.decisoes, p_totais, p_resumo, public._arquivos_ativos(p_id), p_observacao,
          (select email from public.perfis where id = auth.uid()));
  return v;
end $$;

-- fecha a competência: trava, grava versão com detalhe e as comissões
create or replace function public.fechar_conciliacao(p_id uuid, p_totais jsonb, p_resumo jsonb, p_detalhe jsonb, p_observacao text default null)
returns int language plpgsql security definer set search_path = public as $$
declare v int; c public.conciliacoes; r jsonb;
begin
  if auth.uid() is null then raise exception 'Não autenticado'; end if;
  select * into c from public.conciliacoes where id = p_id for update;
  if not found then raise exception 'Conciliação não encontrada'; end if;
  if c.status = 'fechada' then raise exception 'Competência já está fechada.'; end if;
  update public.conciliacoes
     set versao = versao + 1, status = 'fechada', totais = p_totais, fechada_por = auth.uid(), fechada_em = now()
   where id = p_id returning versao into v;
  insert into public.conciliacao_versoes (conciliacao_id, versao, tipo, regras, decisoes, totais, resumo, detalhe, arquivos, observacao, salvo_por_email)
  values (p_id, v, 'fechamento', c.regras, c.decisoes, p_totais, p_resumo, p_detalhe, public._arquivos_ativos(p_id), p_observacao,
          (select email from public.perfis where id = auth.uid()));
  for r in select * from jsonb_array_elements(coalesce(p_resumo, '[]'::jsonb)) loop
    insert into public.comissoes_fechadas (conciliacao_id, competencia, versao, dentista_id, dentista_nome, especialidade, forma, valor_ref,
                                           guias, procedimentos, producao, ajustes, base, diarias, comissao)
    values (p_id, c.competencia, v, nullif(r->>'dentistaId','')::uuid, r->>'nome', r->>'especialidade', r->>'forma', (r->>'valorRef')::numeric,
            (r->>'guias')::int, (r->>'procs')::int, round((r->>'producao')::numeric, 2), round(coalesce((r->>'ajustes')::numeric, 0), 2),
            round((r->>'base')::numeric, 2), (r->>'diarias')::int, round((r->>'comissao')::numeric, 2));
  end loop;
  return v;
end $$;

-- reabre (somente admin); fica registrado com motivo
create or replace function public.reabrir_conciliacao(p_id uuid, p_motivo text)
returns int language plpgsql security definer set search_path = public as $$
declare v int; c public.conciliacoes;
begin
  if not public.eh_admin() then raise exception 'Apenas administradores podem reabrir uma competência.'; end if;
  if coalesce(trim(p_motivo), '') = '' then raise exception 'Informe o motivo da reabertura.'; end if;
  select * into c from public.conciliacoes where id = p_id for update;
  if c.status <> 'fechada' then raise exception 'A competência não está fechada.'; end if;
  perform set_config('app.liberar_fechada', 'sim', true);
  update public.conciliacoes set versao = versao + 1, status = 'rascunho', fechada_por = null, fechada_em = null
   where id = p_id returning versao into v;
  insert into public.conciliacao_versoes (conciliacao_id, versao, tipo, regras, decisoes, totais, arquivos, observacao, salvo_por_email)
  values (p_id, v, 'reabertura', c.regras, c.decisoes, c.totais, public._arquivos_ativos(p_id), p_motivo,
          (select email from public.perfis where id = auth.uid()));
  return v;
end $$;

-- ---------------------------------------------------------------------
-- RLS: usuários logados leem tudo; ninguém apaga nada
-- ---------------------------------------------------------------------
alter table public.perfis              enable row level security;
alter table public.especialidades      enable row level security;
alter table public.dentistas           enable row level security;
alter table public.conciliacoes        enable row level security;
alter table public.conciliacao_arquivos enable row level security;
alter table public.conciliacao_versoes enable row level security;
alter table public.comissoes_fechadas  enable row level security;
alter table public.auditoria           enable row level security;

do $$
declare t text;
begin
  foreach t in array array['perfis','especialidades','dentistas','conciliacoes','conciliacao_arquivos','conciliacao_versoes','comissoes_fechadas','auditoria'] loop
    execute format('drop policy if exists ler_%1$s on public.%1$I', t);
    execute format('create policy ler_%1$s on public.%1$I for select to authenticated using (true)', t);
  end loop;
  foreach t in array array['especialidades','dentistas','conciliacoes','conciliacao_arquivos'] loop
    execute format('drop policy if exists inserir_%1$s on public.%1$I', t);
    execute format('create policy inserir_%1$s on public.%1$I for insert to authenticated with check (true)', t);
    execute format('drop policy if exists alterar_%1$s on public.%1$I', t);
    execute format('create policy alterar_%1$s on public.%1$I for update to authenticated using (true) with check (true)', t);
  end loop;
end $$;

-- perfis: cada um edita o próprio nome; papel só via SQL (admin)
drop policy if exists alterar_proprio_perfil on public.perfis;
create policy alterar_proprio_perfil on public.perfis for update to authenticated using (id = auth.uid()) with check (id = auth.uid() and papel = (select p.papel from public.perfis p where p.id = auth.uid()));

revoke delete on public.perfis, public.especialidades, public.dentistas, public.conciliacoes, public.conciliacao_arquivos,
                 public.conciliacao_versoes, public.comissoes_fechadas, public.auditoria from anon, authenticated;
revoke insert, update on public.conciliacao_versoes, public.comissoes_fechadas, public.auditoria from anon, authenticated;
revoke all on public.perfis, public.especialidades, public.dentistas, public.conciliacoes, public.conciliacao_arquivos,
              public.conciliacao_versoes, public.comissoes_fechadas, public.auditoria from anon;

grant execute on function public.salvar_versao(uuid, jsonb, jsonb, text) to authenticated;
grant execute on function public.fechar_conciliacao(uuid, jsonb, jsonb, jsonb, text) to authenticated;
grant execute on function public.reabrir_conciliacao(uuid, text) to authenticated;
revoke execute on function public.salvar_versao(uuid, jsonb, jsonb, text) from anon;
revoke execute on function public.fechar_conciliacao(uuid, jsonb, jsonb, jsonb, text) from anon;
revoke execute on function public.reabrir_conciliacao(uuid, text) from anon;

-- comissões vigentes: último fechamento de cada competência fechada
create or replace view public.comissoes_vigentes with (security_invoker = true) as
  select cf.* from public.comissoes_fechadas cf
  join public.conciliacoes c on c.id = cf.conciliacao_id and c.status = 'fechada'
  where cf.versao = (select max(x.versao) from public.comissoes_fechadas x where x.conciliacao_id = cf.conciliacao_id);
revoke all on public.comissoes_vigentes from anon;
grant select on public.comissoes_vigentes to authenticated;

-- ---------------------------------------------------------------------
-- Storage: bucket privado com as planilhas originais (sem apagar)
-- ---------------------------------------------------------------------
insert into storage.buckets (id, name, public) values ('planilhas', 'planilhas', false)
on conflict (id) do nothing;

drop policy if exists planilhas_ler on storage.objects;
create policy planilhas_ler on storage.objects for select to authenticated using (bucket_id = 'planilhas');
drop policy if exists planilhas_enviar on storage.objects;
create policy planilhas_enviar on storage.objects for insert to authenticated with check (bucket_id = 'planilhas');

-- ---------------------------------------------------------------------
-- Especialidades padrão (só na primeira vez)
-- ---------------------------------------------------------------------
insert into public.especialidades (nome, forma, valor)
select * from (values ('Clínico','percentual',30::numeric), ('Ortodontia','percentual',30), ('Endodontia','percentual',40), ('Cirurgia','diaria',0)) v(nome, forma, valor)
where not exists (select 1 from public.especialidades);
