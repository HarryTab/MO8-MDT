-- Source order: supplied Acts for definitions and FPN tariffs, then the public
-- Trello for current short-form guidance and in-game arrest durations.
-- PACE entries are recorded as related powers, not as chargeable offences.
alter table public.operational_offences
  add column if not exists suggested_outcome text not null default 'FPN',
  add column if not exists allowed_outcomes text[] not null default array['FPN', 'Warning', 'Arrested', 'Reported for Summons', 'No Further Action']::text[],
  add column if not exists escalation_guidance text not null default '',
  add column if not exists repeat_fine integer not null default 0,
  add column if not exists legislation_source text not null default '',
  add column if not exists section_reference text not null default '',
  add column if not exists source_url text not null default '',
  add column if not exists guidance_notes text not null default '',
  add column if not exists related_powers text[] not null default array[]::text[];

with catalogue (
  code, title, category, description, base_fine, repeat_fine, arrest_minutes,
  disposal_class, legislation_source, section_reference, guidance_notes,
  related_powers
) as (
  values
    ('MO8-RD-001', 'Driving without due care or attention', 'Road Traffic', 'Driving below the standard expected of a competent and careful driver.', 1500, 2000, 4, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(1)', 'Examples include failing to follow signs or indicate. Uploaded FPN Act tariff takes precedence over the older Trello amount.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-RD-002', 'Dangerous driving', 'Road Traffic', 'Driving far below the expected standard in a way that poses significant danger to the public or other road users.', 3000, 4500, 4, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(2)', 'The higher tariff applies to repeat offending or intentional vehicle ramming that immobilises a vehicle.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-RD-003', 'Motor racing on public roads', 'Road Traffic', 'Two or more vehicles taking part in an unauthorised race, competition or trial on a public road.', 1500, 1500, 3, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(3)', 'Home Office-authorised events are excluded.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-RD-004', 'Leaving a vehicle in a dangerous position', 'Road Traffic', 'Leaving a vehicle stationary in a position or condition posing significant danger to road users or the public.', 750, 2000, 3, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(4)', 'The source requires the dangerous position to have been witnessed by a constable.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-RD-005', 'Failure to wear protective headgear', 'Road Traffic', 'A motorcycle, moped, EAPC or quad-bike rider or passenger failing to wear required protective headgear.', 200, 500, 0, 'FPN', 'Roads Classification and Offences Act 2024', 'Section 4(5)', 'Pedal-cycle riders are excluded unless electrically assisted as specified by the Act.', array[]::text[]),
    ('MO8-RD-006', 'Driving or parking in a bus lane', 'Road Traffic', 'Driving or parking wholly or partly in a bus lane without an applicable exemption.', 350, 600, 0, 'FPN', 'Roads Classification and Offences Act 2024', 'Section 4(6)', 'Apply public-service and authorised recovery exemptions.', array[]::text[]),
    ('MO8-RD-007', 'Failure to comply with traffic directions', 'Road Traffic', 'A driver or pedestrian refusing a lawful traffic direction given by a constable.', 400, 800, 3, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(7)', 'Includes directions to stop, divert or remain within a specified route.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-RD-008', 'Failure to comply with traffic signs', 'Road Traffic', 'Failing to follow a traffic sign, signal or road marking.', 200, 400, 0, 'FPN', 'Roads Classification and Offences Act 2024', 'Section 4(8)', '', array[]::text[]),
    ('MO8-RD-009', 'Driving without proper licensing', 'Road Traffic', 'Operating a vehicle with absent, obscured, incomplete or otherwise improper registration plates.', 750, 1500, 0, 'FPNSeizure', 'Roads Classification and Offences Act 2024', 'Section 4(9)', 'The Act permits an FPN, arrest or vehicle seizure according to circumstances.', array['ROCA 2024 - vehicle seizure powers']),
    ('MO8-RD-010', 'Exceeding the speed limit', 'Road Traffic', 'Operating a vehicle above the speed limit applying to the road or displayed by signage.', 2000, 3500, 0, 'FPNSeizure', 'Roads Classification and Offences Act 2024', 'Section 4(10)', 'The uploaded FPN Act tariff takes precedence over the older Trello amount.', array['ROCA 2024 - vehicle seizure powers']),
    ('MO8-RD-011', 'Driving a vehicle with an illegal modification', 'Road Traffic', 'Driving a vehicle with a prohibited modification such as obscuring tint, unauthorised blue lights or sirens.', 300, 600, 0, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(11)', 'Consider prohibition or seizure where the vehicle remains unsafe or unlawfully equipped.', array['ROCA 2024 - vehicle seizure powers']),
    ('MO8-RD-012', 'Failure to stop for a police constable', 'Road Traffic', 'Failing to stop a vehicle when lawfully required to do so by a constable.', 2000, 2000, 4, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(12)', 'Record the stop requirement and evidence that it would have been apparent to the driver.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-RD-013', 'Obstruction of the highway', 'Road Traffic', 'Wilfully obstructing free passage along a highway without lawful authority or sufficient excuse.', 400, 1750, 3, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(13)', '', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-RD-014', 'Tampering with a motor vehicle', 'Road Traffic', 'Interfering with, entering or climbing onto a vehicle to impede, damage or manipulate it.', 1500, 1500, 3, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(14)', 'Roleplay interference with brakes or mechanisms is included where evidenced.', array['PACE 2025 Section 2 - stop and search']),
    ('MO8-RD-015', 'Illegal parking', 'Road Traffic', 'Parking in a prohibited location or in a manner meeting the illegal-parking conditions in the Act.', 300, 600, 0, 'FPNSeizure', 'Roads Classification and Offences Act 2024', 'Section 4(15) and Section 8', 'Check signs, line restrictions and public-service exemptions before enforcement.', array['ROCA 2024 - vehicle seizure powers']),
    ('MO8-RD-016', 'Using a vehicle in a dangerous condition', 'Road Traffic', 'Using a vehicle whose condition, load, passengers or equipment creates danger to another person.', 1500, 1500, 3, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(16)', 'Record the specific defect, load or passenger condition creating the danger.', array['ROCA 2024 - vehicle seizure powers']),
    ('MO8-RD-017', 'Failure to stop after an accident', 'Road Traffic', 'Leaving a reportable road collision without satisfying an applicable exception or police direction.', 500, 1500, 3, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(17)', 'The uploaded FPN Act tariff takes precedence over the older Trello amount.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-RD-018', 'Failure to identify', 'Road Traffic', 'A driver failing to identify themselves when lawfully required, or supplying false or misleading identity details.', 500, 1000, 4, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(18)', 'The uploaded FPN Act tariff takes precedence over the older Trello amount.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-RD-019', 'Misuse of warning equipment', 'Road Traffic', 'Unnecessary or unauthorised use of amber warning equipment or similar vehicle warning devices.', 250, 1000, 3, 'Mixed', 'Roads Classification and Offences Act 2024', 'Section 4(19)', 'Consider warning and vehicle seizure before arrest where proportionate.', array['ROCA 2024 - vehicle seizure powers']),
    ('MO8-RD-020', 'Failure to use headlights', 'Road Traffic', 'Failing to use headlights during darkness or materially reduced visibility.', 150, 500, 0, 'FPN', 'Roads Classification and Offences Act 2024', 'Section 4(20)', '', array[]::text[]),
    ('MO8-RD-021', 'Driving or parking on cycle tracks', 'Road Traffic', 'Driving or parking a motor vehicle wholly or partly on a cycle track.', 300, 950, 0, 'FPN', 'Roads Classification and Offences Act 2024', 'Section 4(21)', 'Tariff taken from the current Trello guidance because it is not listed in the supplied FPN Act.', array[]::text[]),
    ('MO8-RD-022', 'Careless or inconsiderate cycling', 'Road Traffic', 'Cycling below the standard expected of a competent and careful cyclist or without due consideration for others.', 650, 900, 0, 'FPN', 'Roads Classification and Offences Act 2024', 'Section 4(22)', '', array[]::text[]),
    ('MO8-RD-023', 'Improper use of a yellow box junction', 'Road Traffic', 'Entering and stopping in a yellow box junction without a clear exit or applicable exception.', 200, 400, 0, 'FPN', 'Roads Classification and Offences Act 2024', 'Section 4(23)', '', array[]::text[]),
    ('MO8-RD-024', 'Wrongful recovery of a vehicle', 'Road Traffic', 'A recovery operator unlawfully recovering or retaining a vehicle contrary to the Act.', 1000, 2000, 0, 'FPN', 'Roads Classification and Offences Act 2024', 'Section 7a', 'Confirm the operator was not acting under a lawful recovery authority.', array[]::text[]),
    ('MO8-RD-025', 'Misuse of public transport distress alarms', 'Road Traffic', 'Improper activation or misuse of public-transport emergency or distress equipment.', 500, 1000, 0, 'FPN', 'Roads Classification and Offences Act 2024', 'Public transport provisions', '', array[]::text[]),

    ('MO8-PR-001', 'Failure to comply with a direction to leave land', 'Property', 'A trespasser refusing a lawful direction to leave, or returning within thirty minutes of the direction.', 0, 0, 0, 'Mixed', 'Proprietary Offences Act 2025', 'Section 1', 'The source permits an FPN or arrest but does not supply a tariff.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PR-002', 'Aggravated trespass', 'Property', 'Trespassing with intent to intimidate, obstruct or disrupt people carrying out a lawful activity.', 0, 0, 0, 'Arrest', 'Proprietary Offences Act 2025', 'Section 2', 'Public roads and highways are excluded.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PR-003', 'Theft', 'Property', 'Dishonestly appropriating property belonging to another with intent to permanently deprive them of it.', 0, 0, 0, 'Mixed', 'Proprietary Offences Act 2025', 'Section 3', 'The source permits an FPN or arrest but does not supply a tariff.', array['PACE 2025 Section 2 - stop and search', 'PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PR-004', 'Equipped to steal', 'Property', 'Possessing an article made, adapted or intended for use in theft without lawful authority.', 0, 0, 0, 'Mixed', 'Proprietary Offences Act 2025', 'Section 3a', 'No completed theft or identified target is required. The source permits an FPN or arrest but gives no tariff.', array['PACE 2025 Section 2 - stop and search', 'PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PR-005', 'Trespass of a protected site', 'Property', 'Unauthorised entry onto a signed or formally designated government, Crown, Home Office or defence protected site.', 0, 0, 0, 'Mixed', 'Proprietary Offences Act 2025', 'Section 3a in supplied PDF', 'Check designation, signage, authorised-person exceptions and the public-highway exclusion.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PR-006', 'Taking a motor vehicle without authority', 'Property', 'Taking a motor vehicle without authority and driving it or allowing oneself to be carried in or on it.', 0, 0, 0, 'Mixed', 'Proprietary Offences Act 2025', 'Section 4', 'The source permits an FPN or arrest but does not supply a tariff.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PR-007', 'Robbery', 'Property', 'Committing theft while using force or putting a person in fear of immediate force in order to steal.', 0, 0, 4, 'Arrest', 'Proprietary Offences Act 2025', 'Section 5', 'Custody time is an MO8 operational suggestion because the supplied Act gives no numeric tariff.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PR-008', 'Criminal damage', 'Property', 'Unlawfully damaging tangible property belonging to another while aware of an unreasonable risk of damage.', 0, 0, 0, 'Mixed', 'Proprietary Offences Act 2025', 'Section 6', 'The source permits an FPN or arrest but does not supply a tariff.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PR-009', 'Arson', 'Property', 'Criminal damage by fire with intent to destroy property or recklessness as to whether it would be damaged.', 0, 0, 5, 'Arrest', 'Proprietary Offences Act 2025', 'Section 7', 'On-duty authorised fire-service use of the fire tool is excluded. Custody time is an MO8 operational suggestion.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PR-010', 'Making off without payment', 'Property', 'Dishonestly leaving without paying for goods or services where immediate payment is required or expected.', 0, 0, 0, 'Mixed', 'Proprietary Offences Act 2025', 'Section 8', 'The source gives no numeric tariff.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PR-011', 'Animal cruelty', 'Property', 'Intentionally causing unnecessary harm or death to an animal.', 0, 0, 4, 'Arrest', 'Proprietary Offences Act 2025', 'Section 8a', 'Emergency-service animals may also engage emergency-worker assault offences. Custody time is an MO8 operational suggestion.', array['PACE 2025 Section 5 - arrest necessity']),

    ('MO8-WP-001', 'Possession of an offensive weapon', 'Weapons', 'Possessing an offensive weapon without lawful authority or a recognised seizure exception.', 0, 0, 4, 'Arrest', 'Offensive Weapons Act 2024', 'Section 2', 'Examples include knives, machetes, swords and compound crossbows.', array['PACE 2025 Section 2 - stop and search', 'PACE 2025 Section 5 - arrest necessity']),
    ('MO8-WP-002', 'Possession of a higher-grade weapon', 'Weapons', 'Possessing a higher-grade weapon without being a lawfully authorised person.', 0, 0, 4, 'Arrest', 'Offensive Weapons Act 2024', 'Section 3(1)', '', array['PACE 2025 Section 2 - stop and search', 'PACE 2025 Section 5 - arrest necessity']),
    ('MO8-WP-003', 'Use of a higher-grade weapon', 'Weapons', 'Using a higher-grade weapon without lawful authority.', 0, 0, 5, 'Arrest', 'Offensive Weapons Act 2024', 'Section 3(2)', '', array['PACE 2025 Section 2 - stop and search', 'PACE 2025 Section 5 - arrest necessity']),
    ('MO8-WP-004', 'Brandishing a controlled weapon', 'Weapons', 'Equipping, waving or flourishing a controlled weapon in public without a work, sport or authority exception.', 0, 0, 3, 'Arrest', 'Offensive Weapons Act 2024', 'Section 4(1)', '', array['PACE 2025 Section 2 - stop and search', 'PACE 2025 Section 5 - arrest necessity']),
    ('MO8-WP-005', 'Unlawful discharge of a controlled weapon', 'Weapons', 'Discharging a controlled weapon in public without an applicable work, sport or private-land permission exception.', 0, 0, 4, 'Arrest', 'Offensive Weapons Act 2024', 'Section 4(2)', '', array['PACE 2025 Section 2 - stop and search', 'PACE 2025 Section 5 - arrest necessity']),
    ('MO8-WP-006', 'Possession of an illegal firearm', 'Weapons', 'Possessing an illegal firearm without a lawful seizure exception.', 0, 0, 4, 'Arrest', 'Offensive Weapons Act 2024', 'Section 5(1)', '', array['PACE 2025 Section 2 - stop and search', 'PACE 2025 Section 5 - arrest necessity']),
    ('MO8-WP-007', 'Discharge of an illegal firearm', 'Weapons', 'Using or discharging an illegal firearm.', 0, 0, 4, 'Arrest', 'Offensive Weapons Act 2024', 'Section 5(2)', '', array['PACE 2025 Section 2 - stop and search', 'PACE 2025 Section 5 - arrest necessity']),
    ('MO8-WP-008', 'Unlawful distribution of weapons', 'Weapons', 'Supplying or distributing offensive, higher-grade or illegal weapons without a recognised authority or vendor exception.', 0, 0, 5, 'Arrest', 'Offensive Weapons Act 2024', 'Section 6', '', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-WP-009', 'Unlicensed possession of a controlled firearm', 'Weapons', 'Possessing a controlled firearm without the required in-game firearms licence.', 0, 0, 4, 'Arrest', 'Offensive Weapons Act 2024', 'Section 7', 'The Act directs this conduct to possession of an offensive weapon where applicable.', array['PACE 2025 Section 2 - stop and search', 'PACE 2025 Section 5 - arrest necessity']),

    ('MO8-PO-001', 'Riot', 'Public Order', 'Six or more people using or threatening unlawful violence for a common purpose in a way causing reasonable fear for safety.', 0, 0, 5, 'Arrest', 'Public Order Offences Act 2024', 'Section 1', '', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PO-002', 'Violent disorder', 'Public Order', 'Three or more people using or threatening unlawful violence for a common purpose in a way causing reasonable fear for safety.', 0, 0, 4, 'Arrest', 'Public Order Offences Act 2024', 'Section 2', '', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PO-003', 'Affray', 'Public Order', 'Using or threatening unlawful violence towards another so that a reasonable person would fear for their safety.', 0, 0, 3, 'Arrest', 'Public Order Offences Act 2024', 'Section 3', 'Words alone are insufficient under the Trello guidance.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PO-004', 'Provocation of violence', 'Public Order', 'Threatening, abusive or insulting conduct intended or likely to provoke immediate unlawful violence.', 0, 0, 3, 'Arrest', 'Public Order Offences Act 2024', 'Section 4', '', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PO-005', 'Threats to kill', 'Public Order', 'Making an imminent and plausible threat to seriously harm or kill another person.', 0, 0, 3, 'Arrest', 'Public Order / Offences Against the Person guidance', 'Threats provisions', 'Use provocation of violence where the threat is not imminent or plausibly capable of being carried out.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PO-006', 'Harassment, alarm or distress', 'Public Order', 'Continuing threatening, abusive or insulting behaviour after a specific police warning.', 0, 0, 3, 'Arrest', 'Public Order Offences Act 2024', 'Section 5', 'Record the initial warning and the subsequent repeated conduct.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PO-007', 'Wasting emergency services time', 'Public Order', 'Knowingly making a false report that wastefully deploys emergency services.', 1000, 1500, 4, 'Mixed', 'Public Order Offences Act 2024', 'Section 6', 'Arrest is suggested after repeat offending.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PO-008', 'Impersonation of a public servant', 'Public Order', 'Intentionally presenting oneself as a police, fire, ambulance, armed-forces or security public servant.', 1000, 1500, 3, 'Mixed', 'Public Order Offences Act 2024', 'Section 7', '', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-PO-009', 'Obstructing an emergency worker', 'Public Order', 'Wilfully obstructing a constable or hindering an emergency worker responding to or dealing with an emergency.', 0, 0, 4, 'Arrest', 'Public Order Offences Act 2024', 'Section 8', '', array['PACE 2025 Section 5 - arrest necessity']),

    ('MO8-AP-001', 'Battery', 'Offences Against the Person', 'Applying unlawful force to another person, including minor roleplayed physical contact.', 0, 0, 3, 'Arrest', 'Offences Against the Person Act 2025', 'Section 1', 'A warning is advised for genuinely minor conduct where proportionate.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-AP-002', 'Assault occasioning actual bodily harm', 'Offences Against the Person', 'Unlawful force causing injury more significant than a push or shove and interfering with health or comfort.', 0, 0, 4, 'Arrest', 'Offences Against the Person Act 2025', 'Section 2', '', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-AP-003', 'Unlawful wounding or grievous bodily harm', 'Offences Against the Person', 'Unlawfully wounding another or inflicting serious physical injury.', 0, 0, 4, 'Arrest', 'Offences Against the Person Act 2025', 'Section 3', '', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-AP-004', 'Assault with intent to resist arrest', 'Offences Against the Person', 'Using unlawful force in order to resist or prevent a lawful arrest.', 0, 0, 4, 'Arrest', 'Offences Against the Person Act 2025', 'Section 4', '', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-AP-005', 'Threats to kill', 'Offences Against the Person', 'Making a plausible and imminent threat to harm or kill another person.', 0, 0, 3, 'Arrest', 'Offences Against the Person Act 2025', 'Section 5', 'Do not duplicate the public-order threats charge for the same act.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-AP-006', 'False imprisonment', 'Offences Against the Person', 'Intentionally and unlawfully restraining another person''s freedom of movement by force, threat or coercion.', 0, 0, 3, 'Arrest', 'Offences Against the Person Act 2025', 'Section 6', '', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-AP-007', 'Murder', 'Offences Against the Person', 'Unlawfully killing another person with malicious intent.', 0, 0, 5, 'Arrest', 'Offences Against the Person Act 2025', 'Section 7', 'The Trello states four or more minutes; five minutes is the MO8 catalogue default and may be overridden.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-AP-008', 'Assault on an emergency worker', 'Offences Against the Person', 'Committing battery against an on-duty police, fire or ambulance worker.', 0, 0, 4, 'Arrest', 'Offences Against the Person Act 2025', 'Section 8', 'Use for battery-level conduct; more serious injury may justify the corresponding injury offence.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-AP-009', 'Accessory to an offence', 'Offences Against the Person', 'Assisting, encouraging or enabling another person to commit an offence.', 0, 0, 3, 'Arrest', 'Offences Against the Person Act 2025', 'Section 9', 'Record the principal offence and the assistance provided.', array['PACE 2025 Section 5 - arrest necessity']),
    ('MO8-AP-010', 'Causing the death of an emergency worker', 'Offences Against the Person', 'Intentionally killing or intentionally assisting in killing an on-duty emergency worker.', 0, 0, 5, 'Arrest', 'Offences Against the Person Act 2025', 'Section 9a', 'Do not also record murder for the same act.', array['PACE 2025 Section 5 - arrest necessity'])
)
insert into public.operational_offences (
  code, title, category, description, default_fine, repeat_fine,
  default_prison_minutes, default_points, suggested_outcome,
  allowed_outcomes, escalation_guidance, legislation_source,
  section_reference, source_url, guidance_notes, related_powers,
  version, effective_from, active, updated_at
)
select
  code, title, category, description, base_fine, repeat_fine,
  arrest_minutes, 0,
  case
    when disposal_class = 'FPN' then 'FPN'
    when disposal_class = 'FPNSeizure' then 'FPN'
    when disposal_class = 'Mixed' and base_fine > 0 then 'FPN'
    else 'Arrested'
  end,
  case
    when disposal_class = 'FPN' then array['Warning', 'FPN', 'No Further Action']::text[]
    when disposal_class = 'FPNSeizure' then array['Warning', 'FPN', 'Vehicle Seizure', 'No Further Action']::text[]
    when disposal_class = 'Mixed' then array['Warning', 'FPN', 'Arrested', 'Reported for Summons', 'No Further Action']::text[]
    else array['Arrested', 'Reported for Summons', 'No Further Action']::text[]
  end,
  case
    when repeat_fine > base_fine and base_fine > 0 then 'Repeat offence: use the repeat tariff of up to GBP ' || repeat_fine || ' or consider the next permitted disposal where justified.'
    when disposal_class = 'Mixed' then 'Review earlier CAD contacts before deciding between an FPN and arrest.'
    else 'Review previous CAD contacts and apply the documented disposal proportionately.'
  end,
  legislation_source, section_reference,
  'https://trello.com/b/LeCJIeWP/ps-new-simplified-in-game-legislation-trello',
  guidance_notes, related_powers, 1, date '2025-01-01', true, now()
from catalogue
on conflict (code, version) do update set
  title = excluded.title,
  category = excluded.category,
  description = excluded.description,
  default_fine = excluded.default_fine,
  repeat_fine = excluded.repeat_fine,
  default_prison_minutes = excluded.default_prison_minutes,
  default_points = excluded.default_points,
  suggested_outcome = excluded.suggested_outcome,
  allowed_outcomes = excluded.allowed_outcomes,
  escalation_guidance = excluded.escalation_guidance,
  legislation_source = excluded.legislation_source,
  section_reference = excluded.section_reference,
  source_url = excluded.source_url,
  guidance_notes = excluded.guidance_notes,
  related_powers = excluded.related_powers,
  active = true,
  updated_at = now();
