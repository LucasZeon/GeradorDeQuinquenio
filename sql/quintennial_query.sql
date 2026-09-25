-- Extraction query used by src/main.py (SQLite dialect, fictional schema).
-- The original production query targeted a corporate Oracle schema and used internal tables.
-- DECODE/TO_DATE; table, column and code names were replaced for publication.

SELECT EMP.badge_id AS "MATRÍCULA",
       EMP.full_name AS "NOME_FUNCIONÁRIO",
       ATTR_TAX.value AS "CPF",
       EMP.pis AS "PIS",
       ATTR_JOB.value AS "FUNÇÃO",
       DEP.description AS "LOTAÇÃO",
       EMP.daily_minutes / 60.0 * 5 AS "JORNADA",
       EMP.hire_date AS "ADMISSÃO",
       CASE
           WHEN EMP.termination_date = '1899-12-30' THEN 'ATIVO'
           ELSE strftime('%d/%m/%Y', EMP.termination_date)
       END AS "RESCISÃO",
       EMP.company_code AS "EMPRESA",
       CASE EMP.employer_code
           WHEN 'ENT-A'  THEN 'ENTITY A'
           WHEN 'ENT-B'  THEN 'ENTITY B'
           WHEN 'ENT-C1' THEN 'ENTITY C (CLT)'
           WHEN 'ENT-C2' THEN 'ENTITY C (COMMISSIONED)'
           WHEN 'ENT-X'  THEN 'APPRENTICE/OTHER'
           ELSE 'OTHER'
       END AS "VÍNCULO"
FROM employee EMP
INNER JOIN department DEP
        ON (EMP.department_id = DEP.department_id
            AND EMP.company_code = DEP.company_code)
INNER JOIN employee_attribute ATTR_TAX
        ON (ATTR_TAX.badge_id = EMP.badge_id
            AND ATTR_TAX.company_code = EMP.company_code
            AND ATTR_TAX.attribute_code = 'TAX_ID')
INNER JOIN employee_attribute ATTR_JOB
        ON (ATTR_JOB.badge_id = EMP.badge_id
            AND ATTR_JOB.company_code = EMP.company_code
            AND ATTR_JOB.attribute_code = 'JOB_TITLE')
WHERE EMP.company_code IN ('CO1', 'CO2')
  AND EMP.badge_id NOT IN ('99999901', '99999902')
ORDER BY EMP.full_name
