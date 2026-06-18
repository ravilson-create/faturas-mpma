-- Cole este script inteiro no SQL Editor do Supabase (Supabase -> SQL Editor -> New query)
-- e clique em "Run". Isso cria a tabela onde o app guarda os dados das faturas.

create table if not exists faturas (
  fatura text primary key,
  instalacao text,
  fatura_agrupadora text,
  municipio text,
  complemento text,
  endereco text,
  categoria text,
  tipo_tarifa integer,
  mes_referencia date,
  vencimento date,
  data_leitura_anterior date,
  data_leitura_atual date,
  dic_apurado numeric,
  fic_apurado numeric,
  uc text,
  kwh numeric,
  val_consumo numeric,
  consumo_ponta_kwh numeric,
  val_consumo_ponta numeric,
  consumo_fora_ponta_kwh numeric,
  val_consumo_fora_ponta numeric,
  demanda_ativa_kw numeric,
  val_demanda_ativa numeric,
  demanda_ponta_kw numeric,
  val_demanda_ponta numeric,
  demanda_fora_ponta_kw numeric,
  val_demanda_fora_ponta numeric,
  demanda_ultrapassagem_kw numeric,
  val_demanda_ultrapassagem numeric,
  val_demanda_geracao numeric,
  reativo_excedente_kwh numeric,
  val_reativo_excedente numeric,
  reativo_excedente_fp_kwh numeric,
  val_reativo_excedente_fp numeric,
  reativo_excedente_np_kwh numeric,
  val_reativo_excedente_np numeric,
  val_bandeira numeric,
  cip numeric,
  irpj numeric,
  csll numeric,
  pis_retido numeric,
  cofins_retido numeric,
  dev_geracao numeric,
  dev_geracao_np numeric,
  dev_geracao_fp numeric,
  multa numeric,
  juros numeric,
  correcao_monetaria numeric,
  credito_dic_fic_dmic numeric,
  credito_dicri numeric,
  credito_prazo_atendimento numeric,
  credito_pagamento_conta numeric,
  taxa_religacao numeric,
  ligacao_provisoria numeric,
  parcelamento numeric,
  icms_base numeric,
  icms_val numeric,
  cofins_base numeric,
  cofins_val numeric,
  pis_base numeric,
  pis_val numeric,
  valor_total numeric,
  nao_identificado numeric,
  demanda_contratada_kw numeric,
  demanda_ponta_col_kw numeric,
  demanda_fora_ponta_col_kw numeric,
  consumo_total_kwh numeric,
  dev_geracao_total numeric,
  tem_gd boolean
);

-- Permite que o app (usando a chave pública "anon") leia e grave dados.
-- Como o app não tem login de usuário, qualquer pessoa com o link do site
-- consegue ver e enviar planilhas. Isso é aceitável para uso interno restrito
-- por não-divulgação do link, mas NÃO é controle de acesso real.
alter table faturas enable row level security;

create policy "Permitir leitura publica" on faturas
  for select using (true);

create policy "Permitir insercao publica" on faturas
  for insert with check (true);

create policy "Permitir atualizacao publica" on faturas
  for update using (true);
