# Object Folder Structure Standard

Use this tree for every construction object.

```text
<object_code>_<short_name>/
  00_incoming_requests/
  01_orders_and_appointments/
  02_personnel/
    00_registries/
    employees/
      <employee_id>_<last_name>/
        01_identity_and_contract/
        02_admission_orders/
        03_briefings_and_training/
        04_attestation_and_certificates/
        05_ppe_issue/
        06_permits_and_work_admission/
        07_medical_and_first_aid/
  03_hse_and_fire_safety/
    instructions/
    briefings/
    incidents_and_microtrauma/
    ppe_and_equipment_checks/
    permits/
  04_journals/
    production/
    labor_safety/
  05_execution_docs/
    ppr/
    pprv_work_at_height/
    admission_acts/
    hidden_work_acts/
    work_reports/
  06_normative_base/
  07_monthly_control/
  08_outgoing_submissions/
  09_archive/
    scan_bundles/
  10_scan_inbox/
    manual_review/
```

## Naming convention

Use this format for all controlled files:

```text
YYYYMMDD_<doc_type_code>_<short_subject>_vNN.<ext>
```

Examples:

- `20260309_ORDER_APPOINT_HSE_RESPONSIBLE_v01.pdf`
- `20260309_JOURNAL_FIRE_BRIEFING_v02.xlsx`
- `20260309_CERT_WELDER_IVANOV_II_v01.pdf`

## Scan inbox naming convention

For automated scan routing, put files into `10_scan_inbox/` with this name pattern:

```text
YYYYMMDD__DOC_TYPE__SUBJECT__[EMPLOYEE_ID].ext
```

Supported DOC_TYPE values:

- `AWR` for hidden work acts
- `PASSPORT` for employee passports (EMPLOYEE_ID is required)
- `ORDER` for orders

Examples:

- `20260310__AWR__foundation_grid_a1_a7.pdf`
- `20260310__PASSPORT__ivanov_ii__001.pdf`
- `20260310__ORDER__appoint_hse_responsible.pdf`

## Mandatory metadata fields

Each controlled file should have these attributes in a registry row:

- object code
- document type and code
- responsible person
- employee (if applicable)
- issue date
- expiry date (if applicable)
- revision
- storage folder path
- status: draft, active, superseded, archived

## Multi-employee dossier rules

- one dossier folder per employee
- one current active document per requirement (latest revision)
- expired and replaced files move to `09_archive/` with unchanged filename
- certificates and permits must include expiry reminders in the registry
