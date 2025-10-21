set role postgres;

create unique index if not exists planning_state_id_key
    on public.planning_state (id);

insert into public.planning_state (id, label)
values ('planning_gardes_state_v080', 'Planning gardes v0.80')
on conflict (id) do nothing;

-- After ensuring the planning state exists, rerun any inserts into planning_passwords
-- that reference 'planning_gardes_state_v080'.

reset role;
