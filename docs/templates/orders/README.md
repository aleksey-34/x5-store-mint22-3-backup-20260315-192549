# Order Templates (P11-P21)

This folder contains order templates used by ARM employee checklist generation.

## Template Set

- P11: `order_11_ps_responsible_template.md`
- P12: `order_12_permit_issuer_template.md`
- P13: `order_13_height_works_template.md`
- P14: `order_14_fire_safety_template.md`
- P15: `order_15_loading_unloading_template.md`
- P16: `order_16_pressure_vessels_template.md`
- P17: `order_17_close_shift_template.md`
- P18: `order_18_internship_template.md`
- P19: `order_19_independent_work_admission_template.md`
- P20: `order_20_concrete_heating_template.md`
- P21: `letter_admission_workers_equipment_template.md`

## Placeholders

Templates support placeholders in `{{KEY}}` format.

Most used keys:

- `{{ORG_FULL_NAME}}`, `{{ORG_SHORT_NAME}}`, `{{ORG_INN}}`, `{{ORG_OGRNIP}}`, `{{ORG_ADDRESS}}`
- `{{PROJECT_OBJECT_NAME}}`, `{{PROJECT_CODE}}`
- `{{ORDER_NUMBER}}`, `{{ORDER_DATE}}`, `{{ISSUE_DATE}}`, `{{ISSUE_CITY}}`
- `{{RESPONSIBLE_PERSON}}`, `{{RESPONSIBLE_POSITION}}`, `{{RESPONSIBLE_SIGNATURE}}`
- `{{LEADER_NAME}}`, `{{LEADER_SIGNATURE}}`
- `{{WORKERS_TABLE}}`, `{{WORKERS_BULLETS}}`
- `{{PERMIT_START_DATE}}`, `{{PERMIT_END_DATE}}`
- `{{GUIDANCE}}`

Organization fields are auto-loaded from `docs/templates/organization/company_requisites_card.md` when keys are present there.
