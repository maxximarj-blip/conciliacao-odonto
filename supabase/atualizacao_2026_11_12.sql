-- =====================================================================
--  ATUALIZAÇÃO 2026.11.12: índice de guias (busca no histórico) + glosas
--  Supabase → SQL Editor → cole e rode UMA vez.
-- =====================================================================

-- 1) arquivos de glosa podem ser guardados na competência
alter table public.conciliacao_arquivos drop constraint if exists conciliacao_arquivos_tipo_check;
alter table public.conciliacao_arquivos add constraint conciliacao_arquivos_tipo_check check (tipo in ('sistema','dentista','glosa'));

-- 2) índice das guias de cada competência (sistema e produção informada)
--    usado para achar guias de outros meses ("não encontrado")
create table if not exists public.guias_index (
  id              bigint generated always as identity primary key,
  conciliacao_id  uuid not null references public.conciliacoes(id) on delete cascade,
  competencia     date not null,
  tipo            text not null check (tipo in ('sistema','informado')),
  guia            text,
  atendimento     text,
  dentista_id     uuid,
  dentista_nome   text,
  paciente        text,
  procedimento    text,
  valor           numeric(14,2),
  data            date,
  situacao        text,
  criado_em       timestamptz not null default now()
);
create index if not exists guias_index_guia_idx on public.guias_index (guia);
create index if not exists guias_index_atend_idx on public.guias_index (atendimento);
create index if not exists guias_index_conc_idx on public.guias_index (conciliacao_id);
alter table public.guias_index enable row level security;
drop policy if exists ler_guias_index on public.guias_index;
create policy ler_guias_index on public.guias_index for select to authenticated using (true);
drop policy if exists inserir_guias_index on public.guias_index;
create policy inserir_guias_index on public.guias_index for insert to authenticated with check (true);
drop policy if exists apagar_guias_index on public.guias_index;
create policy apagar_guias_index on public.guias_index for delete to authenticated using (true);
revoke all on public.guias_index from anon;
grant select, insert, delete on public.guias_index to authenticated;
