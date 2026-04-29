"""
사내 노트북에서 실행 → online_sales.xlsx 저장 → GitHub에 push
"""
import prestodb
import pandas as pd

# ✏️ 비밀번호만 입력
PASSWORD = "여기에_비밀번호"

conn = prestodb.dbapi.connect(
    host        = "kakaoent-presto-adhoc.kakaoent.io",
    port        = 8443,
    user        = "journi-y222",
    catalog     = "hadoop_kent",
    schema      = "data_analysis",
    http_scheme = "https",
    auth        = prestodb.auth.BasicAuthentication("journi-y222", PASSWORD),
)

SQL = """
SELECT
    completed_dt,
    SUM(COALESCE(line_krw_amount, 0))
    + SUM(IF(delivery_rk = 1, COALESCE(krw_delivery_fee, 0), 0))
    - SUM(COALESCE(real_discount_amount, 0)) AS pay_amt
FROM data_analysis.v_berriz_commerce_mart_daily_order
WHERE partner_id = 6
GROUP BY 1
ORDER BY 1
"""

df = pd.read_sql(SQL, conn)
df.to_excel("online_sales.xlsx", index=False)
print(f"저장 완료: {len(df)}일치")
print(df.tail())
