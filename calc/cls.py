# общая классификация продукта подписки: x_studio_linko, иначе по составу строк
PROD_SQL = """
case when so.x_studio_linko is not null then so.x_studio_linko
     when exists (select 1 from sale_order_line l2 where l2.order_id=so.id and l2.product_id in (24,40,32,43)) then 'POS'
     when exists (select 1 from sale_order_line l2 where l2.order_id=so.id and l2.product_id in (25,26,41,42)) then 'SFA'
     else '?' end
"""
MRR_SQL = "case when so.currency_id=1 then so.recurring_monthly*12000 else so.recurring_monthly end"
AMT_SQL = "case when l.currency_id=1 then l.amount_signed*12000 else l.amount_signed end"
