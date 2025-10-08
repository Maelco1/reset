set role postgres;

insert into public.planning_state (id, label)
values ('planning_gardes_state_v080', 'Planning gardes v0.80')
on conflict (id) do nothing;

-- After ensuring the planning state exists, rerun any inserts into planning_passwords
-- that reference 'planning_gardes_state_v080'.

reset role;
