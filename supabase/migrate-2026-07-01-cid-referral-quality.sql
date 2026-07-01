alter table public.cid_referral_exports
  add column if not exists readiness_score integer not null default 0 check (readiness_score between 0 and 100),
  add column if not exists quality_assessment jsonb not null default '{}'::jsonb,
  add column if not exists executive_summary jsonb not null default '{}'::jsonb,
  add column if not exists redaction_manifest jsonb not null default '[]'::jsonb,
  add column if not exists version_comparison jsonb not null default '{}'::jsonb,
  add column if not exists supersedes_referral_id text references public.cid_referral_exports(referral_id) on delete set null;

create index if not exists idx_cid_referral_supersedes
  on public.cid_referral_exports(supersedes_referral_id);
