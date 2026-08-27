# Configuração do banco central (Supabase)

O dashboard já está pronto para usar o Supabase. Faltam **3 passos** — leva uns 10 minutos.

---

## Passo 1 — Criar o projeto

1. Acesse [supabase.com](https://supabase.com) e crie uma conta (grátis)
2. **New project** → dê um nome (ex.: `painel-seminovos`) → escolha uma senha para o banco → **Create**
3. Espere ~2 minutos enquanto o projeto é provisionado

---

## Passo 2 — Criar a tabela

No painel do Supabase, vá em **SQL Editor** → **New query**, cole o SQL abaixo e clique em **Run**:

```sql
-- Tabela que guarda o estado do dashboard (uma linha por "sala")
create table public.dashboards (
  id          text primary key,
  pnl         jsonb not null default '{}',
  frentes     jsonb not null default '{}',
  meta        jsonb not null default '{}',
  updated_at  timestamptz not null default now()
);

-- Cria a linha da sala padrão
insert into public.dashboards (id) values ('seminovos');

-- Habilita Row Level Security
alter table public.dashboards enable row level security;

-- Permite que qualquer visitante leia e escreva NESTA tabela.
-- A proteção de acesso é feita pela senha na tela de entrada do dashboard.
create policy "leitura publica"
  on public.dashboards for select
  using (true);

create policy "escrita publica"
  on public.dashboards for insert
  with check (true);

create policy "atualizacao publica"
  on public.dashboards for update
  using (true) with check (true);

-- Habilita replicação em tempo real para esta tabela
alter publication supabase_realtime add table public.dashboards;
```

---

## Passo 3 — Colar as credenciais no HTML

No Supabase, vá em **Settings** → **API** e copie dois valores:

| Campo no Supabase | Onde colar no HTML |
|---|---|
| **Project URL** | `window.SUPABASE_URL` |
| **anon / publishable key** | `window.SUPABASE_KEY` |

No arquivo `dashboard_aceleracao_seminovos.html`, procure estas duas linhas (perto da linha 610):

```js
window.SUPABASE_URL = 'COLE_AQUI_SUA_PROJECT_URL';
window.SUPABASE_KEY = 'COLE_AQUI_SUA_PUBLISHABLE_KEY';
```

Substitua pelos seus valores, por exemplo:

```js
window.SUPABASE_URL = 'https://abcdefgh.supabase.co';
window.SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6...';
```

Faça commit e push. Em 1 minuto o GitHub Pages atualiza.

> **É seguro colocar a `anon key` no HTML?** Sim. Ela é feita para isso — fica visível no navegador de qualquer usuário por design. A proteção real vem das políticas de RLS acima. O que **nunca** pode ir no HTML é a `service_role key`.

---

## Como saber se funcionou

Abra o dashboard e olhe o cabeçalho:

| Indicador | Significado |
|---|---|
| Ponto **verde** + "Só você" | Conectado ao banco. Funcionando. |
| Ponto **cinza** + "Modo local" | Credenciais não preenchidas — ainda no modo local. |
| Ponto **amarelo** piscando | Conectou mas deu erro. Passe o mouse em cima para ver a mensagem. |

No rodapé, ao editar qualquer campo, deve aparecer *"Alterações salvas e compartilhadas com o time."*

**Teste definitivo:** preencha um número, feche o navegador **completamente**, abra em outro navegador (ou celular) com o mesmo link. O número deve estar lá.

---

## Salas separadas (opcional)

Se você quiser painéis independentes — por exemplo um por trimestre — basta mudar o `?room=` na URL e criar a linha correspondente:

```sql
insert into public.dashboards (id) values ('seminovos-q4');
```

E compartilhar: `...dashboard_aceleracao_seminovos.html?room=seminovos-q4`

---

## Upgrade futuro: login por e-mail

Hoje o acesso é protegido pela senha na tela de entrada (`Lead2sales`), que roda no navegador. Isso segura acesso casual, mas alguém tecnicamente habilidoso consegue contornar lendo o código-fonte.

Se em algum momento você precisar de proteção real, o caminho é o **Supabase Auth** (login por e-mail/magic link) trocando as políticas de RLS por versões que exigem usuário autenticado:

```sql
-- Exemplo: só usuários logados podem editar
drop policy "escrita publica" on public.dashboards;
drop policy "atualizacao publica" on public.dashboards;

create policy "escrita autenticada"
  on public.dashboards for insert
  to authenticated with check (true);

create policy "atualizacao autenticada"
  on public.dashboards for update
  to authenticated using (true) with check (true);
```

Me avise se quiser que eu implemente essa parte.
