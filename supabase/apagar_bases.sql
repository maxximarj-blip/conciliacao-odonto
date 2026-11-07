-- =====================================================================
--  ATUALIZAÇÃO 2026.11.07: botão "Apagar bases" no Histórico (só admin)
--  Supabase → SQL Editor → cole e rode UMA vez.
-- =====================================================================
create or replace function public.apagar_bases(p_dentistas boolean, p_auditoria boolean, p_confirmacao text)
returns jsonb language plpgsql security definer set search_path = public as $$
declare v_paths text[]; v_conc int; v_dent int := 0;
begin
  if not public.eh_admin() then raise exception 'Apenas administradores podem apagar as bases.'; end if;
  if coalesce(upper(trim(p_confirmacao)), '') <> 'APAGAR' then raise exception 'Digite APAGAR para confirmar.'; end if;
  select coalesce(array_agg(storage_path), '{}') into v_paths from public.conciliacao_arquivos;
  select count(*) into v_conc from public.conciliacoes;
  truncate table public.comissoes_fechadas, public.conciliacao_versoes, public.conciliacao_arquivos, public.conciliacoes restart identity cascade;
  if p_dentistas then
    delete from public.regras_procedimento r using public.dentistas d where d.id = r.dentista_id and upper(d.nome) <> 'CLINICA';
    delete from public.dentistas where upper(nome) <> 'CLINICA';
    get diagnostics v_dent = row_count;
  end if;
  if p_auditoria then truncate table public.auditoria restart identity; end if;
  return jsonb_build_object('competencias', v_conc, 'dentistas', v_dent, 'arquivos', v_paths);
end $$;
revoke all on function public.apagar_bases(boolean, boolean, text) from public, anon;
grant execute on function public.apagar_bases(boolean, boolean, text) to authenticated;

-- admin pode apagar as planilhas guardadas no Storage
drop policy if exists planilhas_apagar on storage.objects;
create policy planilhas_apagar on storage.objects for delete to authenticated using (bucket_id = 'planilhas' and public.eh_admin());
