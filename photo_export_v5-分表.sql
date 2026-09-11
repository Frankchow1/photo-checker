-- 1-1  任务主表  右键导出为 task.csv
SET NAMES utf8mb4;
SET SESSION group_concat_max_len = 10000000;

SET @project_kw = '建设银行杭州分行商户维护项目';
SET @date_from  = '2026-08-01';
SET @date_to    = '2026-08-31';



SELECT
t.ID                AS task_id,
t.TASK_NO           AS task_no,
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
TRIM(TRAILING '.' FROM TRIM(TRAILING '0' FROM t.LONGITUDE)) AS import_lng,
TRIM(TRAILING '.' FROM TRIM(TRAILING '0' FROM t.LATITUDE))  AS import_lat
FROM collectc.t_pos_product_task202609071635 t
WHERE t.TASK_SUBMIT_TIME >= @date_from
AND t.TASK_SUBMIT_TIME <  DATE_ADD(@date_to, INTERVAL 1 DAY)
AND t.RELATED_ITEM_NAME COLLATE utf8mb4_general_ci LIKE CONCAT('%', @project_kw, '%')
AND IFNULL(t.DELETE_FLAG, 0) = 0
ORDER BY t.TASK_SUBMIT_TIME;
-- 1-2  生成 task_id 清单  双击 id_list 那一格全选复制  粘进脚本二
SELECT
COUNT(*)                          AS task_cnt,
GROUP_CONCAT(t.ID SEPARATOR ',')  AS id_list
FROM collectc.t_pos_product_task202609071635 t
WHERE t.TASK_SUBMIT_TIME >= @date_from
AND t.TASK_SUBMIT_TIME <  DATE_ADD(@date_to, INTERVAL 1 DAY)
AND t.RELATED_ITEM_NAME COLLATE utf8mb4_general_ci LIKE CONCAT('%', @project_kw, '%')
AND IFNULL(t.DELETE_FLAG, 0) = 0;





*脚本二 · 在有 `t_collectc_item` 的库上跑**

```sql
SET NAMES utf8mb4;

-- 把脚本一 1-2 出的 id_list 整串粘到下面括号里
-- 必须一行写完  不要加引号  不要换行  末尾不要留逗号
SELECT
  i.task_id            AS task_id,
  i.id                 AS item_row_id,
  i.taskconf_item_id   AS item_id,
  i.doc_name           AS item_name,
  i.item_type          AS item_type,
  i.doc_url            AS doc_url,
  i.address            AS live_address,
  i.longitude          AS live_lng,
  i.latitude           AS live_lat,
  i.create_by          AS upload_by,
  i.create_time        AS upload_time
FROM collectc.t_collectc_item i
WHERE i.task_id IN (
56308111367798738,56281424856946754
)
  AND IFNULL(i.delete_flag, 0) = 0
  AND i.item_type = 14
  AND i.doc_url IS NOT NULL
  AND i.doc_url <> ''
  AND ( i.doc_name COLLATE utf8mb4_general_ci LIKE '%门头%'
     OR i.doc_name COLLATE utf8mb4_general_ci LIKE '%门面%' )
  AND i.doc_name COLLATE utf8mb4_general_ci NOT LIKE '%办理机构%'
ORDER BY i.task_id, i.create_time;