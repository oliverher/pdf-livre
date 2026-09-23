-- PDF Livre — esquema inicial
-- Guarda o histórico de conversões e os metadados dos arquivos.
-- Os arquivos em si vivem no Storage (buckets 'entradas' e 'saidas'),
-- não no banco — aqui só ficam referência, status e métricas.

create table if not exists public.conversoes (
  id uuid primary key default gen_random_uuid(),
  usuario_id uuid references auth.users(id) on delete cascade not null,
  nome_arquivo_original text not null,
  formato_origem text not null check (formato_origem in ('pdf','pptx','potx','png','jpg','jpeg','webp')),
  rota text check (rota in ('vetorial','visao')),
  status text not null default 'pendente' check (status in ('pendente','processando','concluida','erro')),
  mensagem_erro text,
  caminho_storage_entrada text,
  caminho_storage_saida text,
  slides integer,
  objetos integer,
  indice_editabilidade numeric(4,3),
  criado_em timestamptz not null default now(),
  concluido_em timestamptz,
  expira_em timestamptz not null default (now() + interval '20 minutes')
);

create index if not exists conversoes_usuario_idx on public.conversoes(usuario_id, criado_em desc);
create index if not exists conversoes_expira_idx on public.conversoes(expira_em) where status = 'concluida';

-- RLS: cada pessoa só vê e mexe nas próprias conversões (RNF18/RNF19 do
-- documento de requisitos — nenhum usuário público enxerga arquivo de outro).
alter table public.conversoes enable row level security;

create policy "usuario_le_suas_conversoes"
  on public.conversoes for select
  using (auth.uid() = usuario_id);

create policy "usuario_cria_suas_conversoes"
  on public.conversoes for insert
  with check (auth.uid() = usuario_id);

create policy "usuario_atualiza_suas_conversoes"
  on public.conversoes for update
  using (auth.uid() = usuario_id);

create policy "usuario_apaga_suas_conversoes"
  on public.conversoes for delete
  using (auth.uid() = usuario_id);

-- Storage: dois buckets privados. O backend Python grava/lê com a service
-- role (contorna RLS de propósito); o navegador nunca acessa o bucket
-- diretamente, só via URL assinada de curta duração.
insert into storage.buckets (id, name, public)
values ('entradas', 'entradas', false)
on conflict (id) do nothing;

insert into storage.buckets (id, name, public)
values ('saidas', 'saidas', false)
on conflict (id) do nothing;

create policy "usuario_le_suas_entradas"
  on storage.objects for select
  using (bucket_id = 'entradas' and auth.uid()::text = (storage.foldername(name))[1]);

create policy "usuario_le_suas_saidas"
  on storage.objects for select
  using (bucket_id = 'saidas' and auth.uid()::text = (storage.foldername(name))[1]);
