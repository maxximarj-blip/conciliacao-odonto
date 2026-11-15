-- =====================================================================
--  ATUALIZAÇÃO 2026.11.12b: apagar por competência ou por arquivo (só admin)
--  Supabase → SQL Editor → cole e rode UMA vez.
-- =====================================================================
create or replace function public.apagar_competencia(p_id uuid, p_confirmacao text)
returns jsonb language plpgsql security definer set search_path = public as $$
declare v_paths text[]; v_comp date;
begin
  if not public.eh_admin() then raise exception 'Apenas administradores podem apagar.'; end if;
  if coalesce(upper(trim(p_confirmacao)), '') <> 'APAGAR' then raise exception 'Digite APAGAR para confirmar.'; end if;
  select competencia into v_comp from public.conciliacoes where id = p_id;
  if v_comp is null then raise exception 'Competência não encontrada.'; end if;
  select coalesce(array_agg(storage_path), '{}') into v_paths from public.conciliacao_arquivos where conciliacao_id = p_id;
  delete from public.comissoes_fechadas where conciliacao_id = p_id;
  delete from public.conciliacao_versoes where conciliacao_id = p_id;
  delete from public.conciliacao_arquivos where conciliacao_id = p_id;
  delete from public.guias_index where conciliacao_id = p_id;
  delete from public.conciliacoes where id = p_id;
  return jsonb_build_object('competencia', v_comp, 'arquivos', v_paths);
end $$;
revoke all on function public.apagar_competencia(uuid, text) from public, anon;
grant execute on function public.apagar_competencia(uuid, text) to authenticated;

create or replace function public.apagar_arquivos(p_ids uuid[])
returns jsonb language plpgsql security definer set search_path = public as $$
declare v_paths text[]; v_conc uuid[]; v_sis boolean;
begin
  if not public.eh_admin() then raise exception 'Apenas administradores podem apagar.'; end if;
  if coalesce(array_length(p_ids, 1), 0) = 0 then raise exception 'Nenhum arquivo selecionado.'; end if;
  if exists (select 1 from public.conciliacao_arquivos a join public.conciliacoes c on c.id = a.conciliacao_id where a.id = any(p_ids) and c.status = 'fechada') then
    raise exception 'A competência está fechada: reabra antes de apagar arquivos.';
  end if;
  select coalesce(array_agg(storage_path), '{}'), coalesce(array_agg(distinct conciliacao_id), '{}'), coalesce(bool_or(tipo = 'sistema' and ativo), false)
    into v_paths, v_conc, v_sis from public.conciliacao_arquivos where id = any(p_ids);
  delete from public.conciliacao_arquivos where id = any(p_ids);
  delete from public.guias_index where conciliacao_id = any(v_conc);           -- refeito ao abrir o mês
  perform set_config('app.liberar_fechada', 'sim', true);
  update public.conciliacoes set decisoes = decisoes - 'conciliado',
         arquivo_sistema = case when v_sis then null else arquivo_sistema end
   where id = any(v_conc);
  return jsonb_build_object('arquivos', v_paths, 'competencias', v_conc);
end $$;
revoke all on function public.apagar_arquivos(uuid[]) from public, anon;
grant execute on function public.apagar_arquivos(uuid[]) to authenticated;

-- recarrega o cache da API para as funções novas aparecerem na hora
notify pgrst, 'reload schema';
