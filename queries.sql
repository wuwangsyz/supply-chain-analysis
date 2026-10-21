-- 供应链与配送履约分析：SQL 查询
-- 数据库：data\orders.db（表 orders，视图 orders_v）
-- 口径见 metrics.md。订单数用 COUNT(DISTINCT "Order Id")，订单级延迟 = 同一订单任一行 Late_delivery_risk = 1。

-- ============================================================
-- Q1 地区配送延迟
-- 业务问题：哪些地区延迟率高，且订单量大，应优先改善？
-- 口径：按地区统计订单数和订单级延迟率。同一订单地区唯一（已验证，0 个例外）。
-- ============================================================
WITH o AS (
    SELECT "Order Id",
           "Order Region",
           MAX(Late_delivery_risk) AS late
    FROM orders
    GROUP BY "Order Id", "Order Region"
)
SELECT "Order Region",
       COUNT(*)                                AS orders,
       SUM(late)                               AS late_orders,
       100.0 * SUM(late) / COUNT(*)            AS late_rate_pct
FROM o
GROUP BY "Order Region"
ORDER BY orders DESC;

-- ============================================================
-- Q2 运输方式承诺时效
-- 业务问题：各运输方式的承诺天数与实际天数是否匹配？
-- 口径：每单取一行计算天数（同一订单的运输方式和两个天数已验证唯一，0 个例外）。
-- ============================================================
WITH o AS (
    SELECT "Order Id",
           "Shipping Mode"                       AS mode,
           MAX("Days for shipment (scheduled)")  AS sched_days,
           MAX("Days for shipping (real)")       AS real_days,
           MAX(Late_delivery_risk)               AS late
    FROM orders
    GROUP BY "Order Id"
)
SELECT mode,
       COUNT(*)                               AS orders,
       AVG(sched_days)                        AS avg_sched_days,
       AVG(real_days)                         AS avg_real_days,
       AVG(real_days) - AVG(sched_days)       AS avg_gap_days,
       100.0 * SUM(late) / COUNT(*)           AS late_rate_pct
FROM o
GROUP BY mode
ORDER BY orders DESC;

-- Q2 敏感性检查：排除已取消订单（Shipping canceled）。
-- 延迟标志只对"Late delivery"成立；取消订单即使实际天数超过计划，也不标记为延迟。
WITH o AS (
    SELECT "Order Id",
           "Shipping Mode"                       AS mode,
           MAX("Delivery Status")                AS status,
           MAX(Late_delivery_risk)               AS late
    FROM orders
    GROUP BY "Order Id"
)
SELECT mode,
       COUNT(*)                               AS orders,
       100.0 * SUM(late) / COUNT(*)           AS late_rate_pct
FROM o
WHERE status <> 'Shipping canceled'
GROUP BY mode
ORDER BY orders DESC;

-- ============================================================
-- Q3 月度需求节奏（按 12 个完整季度汇总：2015Q1–2017Q4）
-- 业务问题：订单量在各时期如何变化？哪些时期需要提前备料或排产？
-- 口径：订单按下单月份归入季度。同一订单下单月唯一（已验证，0 个例外）。
-- 注意：2017-10 起每单行数由约 3 降为 1，2017Q4 的订单数不能直接与前期比较（见 metrics.md 第 4 节）。
-- ============================================================
WITH o AS (
    SELECT "Order Id",
           MIN(year_month)         AS ym,
           MAX(Late_delivery_risk) AS late
    FROM orders_v
    GROUP BY "Order Id"
),
q AS (
    SELECT *,
           substr(ym, 1, 4) || 'Q' || ((CAST(substr(ym, 6, 2) AS INTEGER) - 1) / 3 + 1) AS quarter
    FROM o
    WHERE ym BETWEEN '2015-01' AND '2017-12'
)
SELECT quarter,
       COUNT(*)                           AS orders,
       SUM(late)                          AS late_orders,
       100.0 * SUM(late) / COUNT(*)       AS late_rate_pct
FROM q
GROUP BY quarter
ORDER BY quarter;

-- Q3 记录结构检查：每月行数、订单数、每单行数。
-- 用于判断 2017-10 起的记录结构变化（每单行数由约 3 降为 1）。
SELECT year_month,
       COUNT(*)                                             AS item_rows,
       COUNT(DISTINCT "Order Id")                           AS orders,
       ROUND(1.0 * COUNT(*) / COUNT(DISTINCT "Order Id"), 2) AS rows_per_order
FROM orders_v
GROUP BY year_month
ORDER BY year_month;

-- ============================================================
-- Q4 品类需求与履约
-- 业务问题：哪些品类订单量大、销售额高，但履约表现差？
-- 口径：主窗口为断点前（2015-01 至 2017-09）。品类名去掉尾随空格。
--       订单数 = 含该品类的订单数（一单可含多个品类，各品类加总会大于 65,752）。
--       销售额 = SUM(Sales)，按订单项求和，每行只属于一个品类，不重复。
--       延迟率 = 含该品类的订单中，订单级延迟的比例。
-- ============================================================
WITH pair AS (
    SELECT "Order Id",
           TRIM("Category Name")                 AS category,
           MAX(Late_delivery_risk)               AS late
    FROM orders_v
    WHERE year_month BETWEEN '2015-01' AND '2017-09'
    GROUP BY "Order Id", TRIM("Category Name")
),
sales AS (
    SELECT TRIM("Category Name") AS category,
           SUM(Sales)            AS sales
    FROM orders_v
    WHERE year_month BETWEEN '2015-01' AND '2017-09'
    GROUP BY TRIM("Category Name")
)
SELECT p.category,
       COUNT(*)                            AS orders_with_category,
       s.sales                             AS sales,
       SUM(p.late)                         AS late_orders,
       100.0 * SUM(p.late) / COUNT(*)      AS late_rate_pct
FROM pair p
JOIN sales s ON s.category = p.category
GROUP BY p.category, s.sales
ORDER BY sales DESC;

-- Q4 断点检查：断点前后品类组成是否相同（共有品类数）。
WITH pre AS (SELECT DISTINCT TRIM("Category Name") AS c FROM orders_v WHERE year_month < '2017-10'),
     post AS (SELECT DISTINCT TRIM("Category Name") AS c FROM orders_v WHERE year_month >= '2017-10')
SELECT (SELECT COUNT(*) FROM pre)                                 AS pre_categories,
       (SELECT COUNT(*) FROM post)                                AS post_categories,
       (SELECT COUNT(*) FROM pre WHERE c IN (SELECT c FROM post)) AS shared_categories;

-- Q4 断点后窗口（记为 Q4P）：把上面 Q4 的 BETWEEN '2015-01' AND '2017-09' 换成
-- BETWEEN '2017-10' AND '2018-01' 即可，其余口径不变。假设口径（Q4S）在
-- scripts\03_validate.py 的 q4_post_scenario() 中计算，SQL 只负责记录值。

-- Q4P 三方核对用的单项检查（Fishing，断点后）：SQL / pandas / Excel 三方比对此结果。
WITH pair AS (
    SELECT "Order Id", MAX(Late_delivery_risk) AS late
    FROM orders_v
    WHERE year_month >= '2017-10' AND TRIM("Category Name") = 'Fishing'
    GROUP BY "Order Id"
)
SELECT COUNT(*)                          AS orders,
       SUM(late)                         AS late_orders,
       100.0 * SUM(late) / COUNT(*)      AS late_rate_pct,
       (SELECT SUM(Sales) FROM orders_v
        WHERE year_month >= '2017-10' AND TRIM("Category Name") = 'Fishing') AS sales
FROM pair;

-- ============================================================
-- 断点分析（结果与判定见 break_analysis.md）
-- 分割点：year_month < '2017-10' 为断点前，>= '2017-10' 为断点后。
-- ============================================================

-- 断点分析 1：两段的规模、ID 范围、结构量、单价。break_analysis.md 第 1、4 节的主要依据。
WITH s AS (
    SELECT *, CASE WHEN year_month < '2017-10' THEN 'pre' ELSE 'post' END AS seg
    FROM orders_v
)
SELECT seg,
       COUNT(*)                                              AS item_rows,
       COUNT(DISTINCT "Order Id")                            AS orders,
       ROUND(1.0 * COUNT(*) / COUNT(DISTINCT "Order Id"), 3) AS rows_per_order,
       CAST(MIN("Order Item Id") AS INT)                     AS min_item_id,
       CAST(MAX("Order Item Id") AS INT)                     AS max_item_id,
       CAST(MIN("Order Id") AS INT)                          AS min_order_id,
       CAST(MAX("Order Id") AS INT)                          AS max_order_id,
       COUNT(DISTINCT Market)                                AS markets,
       COUNT(DISTINCT TRIM("Category Name"))                 AS categories,
       COUNT(DISTINCT "Order Country")                       AS order_countries,
       COUNT(DISTINCT "Order Customer Id")                   AS customers,
       ROUND(AVG("Order Item Quantity"), 4)                  AS avg_qty,
       ROUND(AVG("Order Item Product Price"), 2)             AS avg_unit_price,
       ROUND(AVG(Sales), 2)                                  AS avg_sales_per_row,
       ROUND(SUM(Sales) / COUNT(DISTINCT "Order Id"), 2)     AS sales_per_order,
       ROUND(AVG("Order Item Discount Rate"), 4)             AS avg_disc_rate
FROM s
GROUP BY seg
ORDER BY seg;

-- 断点分析 2：每月出现多少个市场、国家、品类。
-- 这一条是用来否决"市场 5→2、国家 164→40 属于断点证据"的说法：并集大小受月份数量控制。
WITH m AS (
    SELECT year_month,
           CASE WHEN year_month < '2017-10' THEN 'pre' ELSE 'post' END AS seg,
           COUNT(DISTINCT Market)                AS mkts,
           COUNT(DISTINCT "Order Country")       AS ctry,
           COUNT(DISTINCT TRIM("Category Name")) AS cats,
           COUNT(DISTINCT "Order Id")            AS orders,
           COUNT(DISTINCT "Order Customer Id")   AS custs
    FROM orders_v
    GROUP BY year_month
)
SELECT seg, COUNT(*) AS months,
       MIN(mkts) AS min_mkts, MAX(mkts) AS max_mkts,
       MIN(ctry) AS min_ctry, MAX(ctry) AS max_ctry, ROUND(AVG(ctry), 1) AS avg_ctry,
       MIN(cats) AS min_cats, MAX(cats) AS max_cats, ROUND(AVG(cats), 1) AS avg_cats,
       ROUND(AVG(1.0 * orders / custs), 3) AS avg_orders_per_customer
FROM m
GROUP BY seg;

-- 断点分析 3：每月的行数结构、数量、单价、每单金额。用于区分"断点特有"与"逐月本来就在波动"。
SELECT year_month,
       COUNT(*)                                          AS item_rows,
       COUNT(DISTINCT "Order Id")                        AS orders,
       COUNT(DISTINCT "Order Customer Id")               AS customers,
       ROUND(1.0 * COUNT(*) / COUNT(DISTINCT "Order Id"), 2)      AS rows_per_order,
       ROUND(AVG("Order Item Quantity"), 3)              AS avg_qty,
       ROUND(AVG("Order Item Product Price"), 2)         AS avg_unit_price,
       ROUND(SUM(Sales) / COUNT(DISTINCT "Order Id"), 2) AS sales_per_order,
       ROUND(AVG("Order Item Discount Rate"), 4)         AS avg_disc_rate
FROM orders_v
GROUP BY year_month
ORDER BY year_month;

-- 断点分析 4：两段的履约结果是否可比（延迟率、平均延迟天数差），以及行内金额关系是否仍然成立。
WITH o AS (
    SELECT CASE WHEN year_month < '2017-10' THEN 'pre' ELSE 'post' END AS seg,
           "Order Id"                          AS oid,
           MAX(Late_delivery_risk)             AS late,
           MAX("Days for shipping (real)")     AS real_d,
           MAX("Days for shipment (scheduled)") AS sched_d
    FROM orders_v
    GROUP BY seg, oid
)
SELECT seg,
       COUNT(*)                              AS orders,
       CAST(SUM(late) AS INT)                AS late_orders,
       ROUND(100.0 * SUM(late) / COUNT(*), 4) AS late_rate_pct,
       ROUND(AVG(real_d) - AVG(sched_d), 3)  AS avg_gap_days
FROM o
GROUP BY seg;

-- 断点分析 4b：行内计算与取值范围检查，两段都应为 0。
SELECT CASE WHEN year_month < '2017-10' THEN 'pre' ELSE 'post' END AS seg,
       COUNT(*) AS rows_cnt,
       SUM(CASE WHEN ABS(Sales - "Order Item Product Price" * "Order Item Quantity") > 0.01 THEN 1 ELSE 0 END)
           AS sales_ne_price_x_qty,
       SUM(CASE WHEN "Order Item Quantity" <> CAST("Order Item Quantity" AS INT) THEN 1 ELSE 0 END)
           AS noninteger_qty,
       SUM(CASE WHEN "Order Item Discount Rate" NOT BETWEEN 0 AND 1 THEN 1 ELSE 0 END)
           AS bad_disc_rate
FROM orders_v
GROUP BY seg;

-- 断点分析 5：订单号缺口。断点前缺 11,430 个（16.6%），断点后缺 22 个（0.26%）。
WITH s AS (
    SELECT CASE WHEN year_month < '2017-10' THEN 'pre' ELSE 'post' END AS seg, "Order Id" AS oid
    FROM orders_v
)
SELECT seg,
       COUNT(DISTINCT oid)                                     AS distinct_orders,
       CAST(MIN(oid) AS INT)                                   AS min_oid,
       CAST(MAX(oid) AS INT)                                   AS max_oid,
       CAST(MAX(oid) AS INT) - CAST(MIN(oid) AS INT) + 1 - COUNT(DISTINCT oid) AS id_gaps
FROM s
GROUP BY seg;

-- 断点分析 6：两段共用的订单号与客户号（订单号 0 个共用，客户号 81 个共用）。
SELECT
  (SELECT COUNT(*) FROM (
     SELECT "Order Id" FROM orders_v WHERE year_month < '2017-10' GROUP BY "Order Id"
     INTERSECT
     SELECT "Order Id" FROM orders_v WHERE year_month >= '2017-10' GROUP BY "Order Id")) AS order_ids_in_both,
  (SELECT COUNT(*) FROM (
     SELECT "Order Customer Id" FROM orders_v WHERE year_month < '2017-10' GROUP BY "Order Customer Id"
     INTERSECT
     SELECT "Order Customer Id" FROM orders_v WHERE year_month >= '2017-10' GROUP BY "Order Customer Id")) AS customers_in_both;

-- 断点分析 7：断点后的每单行数分布（8,340 单为 1 行，63 单为 2–5 行）。
WITH o AS (
    SELECT "Order Id" AS oid, COUNT(*) AS n FROM orders_v
    WHERE year_month >= '2017-10' GROUP BY "Order Id"
)
SELECT n AS rows_per_order, COUNT(*) AS orders FROM o GROUP BY n ORDER BY n;

-- 断点分析 8：Q1 的稳健性检查 —— 两段都出现的地区，延迟率是否一致。
WITH o AS (
    SELECT CASE WHEN year_month < '2017-10' THEN 'pre' ELSE 'post' END AS seg,
           "Order Id" AS oid, MAX("Order Region") AS reg, MAX(Late_delivery_risk) AS late
    FROM orders_v
    GROUP BY seg, oid
)
SELECT TRIM(reg) AS region,
       SUM(CASE WHEN seg = 'pre'  THEN 1 ELSE 0 END) AS pre_orders,
       ROUND(100.0 * SUM(CASE WHEN seg = 'pre'  THEN late ELSE 0 END)
             / NULLIF(SUM(CASE WHEN seg = 'pre'  THEN 1 ELSE 0 END), 0), 2) AS pre_late_pct,
       SUM(CASE WHEN seg = 'post' THEN 1 ELSE 0 END) AS post_orders,
       ROUND(100.0 * SUM(CASE WHEN seg = 'post' THEN late ELSE 0 END)
             / NULLIF(SUM(CASE WHEN seg = 'post' THEN 1 ELSE 0 END), 0), 2) AS post_late_pct
FROM o
GROUP BY TRIM(reg)
HAVING post_orders > 0
ORDER BY pre_orders DESC;
