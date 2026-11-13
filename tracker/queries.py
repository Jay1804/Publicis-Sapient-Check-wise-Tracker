"""SQL for the Publicis Sapient Check-wise tracker.

The original Grafana query joined ~15 tables with correlated sub-selects and
@user-variables, scanning every row. It is replaced by small index-driven
queries (all keyed on the case/check id list of the requested window) that are
merged in pandas - see pipeline.py. Placeholders `{ids}` are filled with
integer id lists only (never user text).
"""

CLIENT_NAME = "Publicis Sapient (A division of TLG India Private Limited)"

COMPANY_IDS = "SELECT company_id FROM ec_master_company WHERE company_name = %(client)s"

# 1. case set (company join only - everything else is fetched by id list)
CASES = """
SELECT ecm.case_id, ecm.candidate_id, ecm.process_id, ecm.client_office_id, ecm.case_ars_no,
       ecm.received_date, ecm.case_accepted_date,
       ecm.created_date AS case_created_date, ecm.CASE_EXPECTED_CLOSURE_DATE AS case_due_date,
       ecm.case_status, ecm.reopen_date AS case_reopen_date, ecm.DATA_SOURCE, emc.company_name
FROM ec_case_master ecm
JOIN ec_master_company emc ON emc.company_id = ecm.client_id
WHERE ecm.client_id IN ({ids})
  AND ecm.case_status NOT IN (8, 14)
  AND ecm.received_date >= '{d_from}' AND ecm.received_date <= '{d_to}'
"""
CANDIDATES = """
SELECT candidate_id, first_name, middle_name, last_name, email_address, mobile_no
FROM ec_case_candidates WHERE candidate_id IN ({ids})
"""
PROCESSES = "SELECT process_id, process_name FROM ec_client_process WHERE process_id IN ({ids})"
LOCATIONS = "SELECT office_id, office_name FROM ec_master_company_locations WHERE office_id IN ({ids})"
FIELDS = """
SELECT case_id, case_flex_field1, case_flex_field2, case_flex_field3, case_flex_field4,
       case_flex_field5, case_flex_field6, case_flex_field7
FROM ec_case_fields WHERE case_id IN ({ids})
"""

# 2. checks of those cases
CHECKS = """
SELECT k.case_check_id, k.case_id, k.check_status, k.check_created_on, k.CHECK_EXPECTED_CLOSURE,
       k.check_name, ec.check_name AS check_unique_name, k.go_ahead_date, k.closure_comments,
       emd.disposition_name, k.VER_OVERSEAS, k.VER_SUMMARY, k.VER_TYPE, k.family_id,
       k.check_severity, k.check_closure_date, k.reopen_date AS check_reopen_date, k.insuff_remarks,
       REPLACE(fam.family_name, ' Family', '') AS department
FROM ec_case_checks k
LEFT JOIN ec_checks ec ON ec.check_id = k.check_id
LEFT JOIN ec_master_disposition emd ON emd.disposition_id = k.check_disposition_id
LEFT JOIN ec_check_families fam ON fam.family_id = k.family_id
WHERE k.case_id IN ({ids})
"""

INSUFF = """
SELECT case_check_id, insuff_id, Insuff_raised_date, Insuff_fulfil_date
FROM ec_case_insuff_details WHERE case_check_id IN ({ids})
"""

SOURCE = "SELECT case_check_id, org_id FROM ec_case_check_verification_source WHERE case_check_id IN ({ids})"

REPORTS = """
SELECT case_id, case_report_id, report_type, report_severity, report_sent_on
FROM ec_case_reports WHERE report_status = 5 AND case_id IN ({ids})
"""

# the only slow table (no index on action name) -> run in parallel chunks
HISTORY = """
SELECT case_id, action_id, action_taken_on
FROM ec_case_history WHERE action_taken = 'New Check Added' AND case_id IN ({ids})
"""

ORG_COMPANY = "SELECT company_id AS id, company_name AS name FROM ec_master_company WHERE company_id IN ({ids})"
ORG_INSTITUTE = "SELECT institute_id AS id, institute_name AS name FROM ec_master_educational_institute WHERE institute_id IN ({ids})"
ORG_CITY = "SELECT city_id AS id, city_name AS name FROM ec_master_city WHERE city_id IN ({ids})"

# Ops client TAT reference used by the Ageing column
TAT = """
SELECT ec.client_external_id, emc.company_name AS client_name, ecp.process_name,
       CONCAT(ec.client_external_id, ecp.process_name) AS formula,
       ecp.tat AS process_tat, ecp.TAT_DAYS_TYPE AS cal,
       ectc.TAT_catg AS new_tat_type, ectc.actual_tat AS new_tat_day
FROM ec_client ec
LEFT JOIN ec_client_process ecp ON ec.client_id = ecp.client_id
LEFT JOIN ec_master_company emc ON ec.client_id = emc.company_id
LEFT JOIN ec_case_tat_config ectc ON ec.client_id = ectc.client_id AND ecp.process_id = ectc.process_id
WHERE emc.company_name = %(client)s
"""
