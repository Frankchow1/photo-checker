
SET @project_kw = '建设银行杭州分行商户维护项目';
SET @date_from  = '2026-08-01';
SET @date_to    = '2026-08-31';

SELECT
  t.TASK_NO           AS task_no,
  t.ID                AS task_id,
  t.RELATED_ITEM_NAME AS project,
  t.CUST_NAME         AS customer,
  t.FLOW_NAME         AS flow_name,
  t.MERCHANTS_NO      AS merchant_no,
  t.MERCHANTS_NAME    AS merchant,
  t.TASK_UNIQUE       AS task_unique,
  t.INSPECTION_CYCLE  AS inspect_cycle,
  t.TASK_STATE        AS task_state,
  t.AUDITOR           AS auditor_id,
  t.AUDITOR_NAME      AS auditor_name,
  t.AUDITOR_PHONE     AS auditor_phone,
  t.TASK_SUBMIT_TIME  AS submit_time,
  t.ADDRESS           AS import_address,
  t.LONGITUDE         AS import_lng,
  t.LATITUDE          AS import_lat,
  i.taskconf_item_id  AS item_id,
  i.doc_name          AS item_name,
  i.doc_url           AS doc_url,
  i.create_by         AS upload_by,
  i.create_time       AS upload_time,
  i.address           AS live_address,
  i.longitude         AS live_lng,
  i.latitude          AS live_lat
FROM collectc.t_pos_product_task202609071635 t
JOIN collectc.t_collectc_item i
  ON i.task_id = t.ID
WHERE t.TASK_SUBMIT_TIME >= @date_from
  AND t.TASK_SUBMIT_TIME <  DATE_ADD(@date_to, INTERVAL 1 DAY)
  AND ( @project_id IS NULL OR t.RELATED_ITEM_ID = @project_id )
  AND i.item_type = 14
  AND i.doc_url IS NOT NULL
  AND i.doc_url <> ''
  AND ( i.doc_name COLLATE utf8mb4_general_ci LIKE '%门头%'
     OR i.doc_name COLLATE utf8mb4_general_ci LIKE '%门面%' )
  AND i.doc_name COLLATE utf8mb4_general_ci NOT LIKE '%办理机构%'
ORDER BY t.TASK_SUBMIT_TIME, i.create_time;

