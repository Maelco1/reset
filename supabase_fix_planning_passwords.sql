set role postgres;

insert into public.planning_state (id, label)
select 'planning_gardes_state_v080', 'Planning gardes v0.80'
where not exists (
    select 1
    from public.planning_state
    where id = 'planning_gardes_state_v080'
);

-- After ensuring the planning state exists, rerun any inserts into planning_passwords
-- that reference 'planning_gardes_state_v080'.

reset role;
