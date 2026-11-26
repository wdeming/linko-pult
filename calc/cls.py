# общая классификация продукта подписки (с 26.11.2026 — по полю «Тип подписки» x_studio_sub_type):
#   SFA — основная SFA, POS — основная POS, AI — AI ассистент, PS — Perfect store.
# Допродукты AI и PS — отдельные подписки на сервере SFA. Если тип пуст — по x_studio_linko, иначе по составу строк.
PRODS = ('SFA', 'POS', 'AI', 'PS')          # все продукты MRR
MAIN = ('SFA', 'POS')                        # основные: количество подписок, отток, когорты, юнит-экономика
ADDON = ('AI', 'PS')                         # допродукты SFA
PNAME = {'SFA': 'SFA', 'POS': 'POS', 'AI': 'AI ассистент', 'PS': 'Perfect store'}
FAM = {'SFA': 'SFA', 'AI': 'SFA', 'PS': 'SFA', 'POS': 'POS'}   # семейство продукта (для сегментов клиентов, затрат)
PROD_SQL = """
case so.x_studio_sub_type when 'main_sfa' then 'SFA' when 'main_pos' then 'POS' when 'ai' then 'AI' when 'perfect_store' then 'PS'
  else case when so.x_studio_linko is not null then so.x_studio_linko
     when exists (select 1 from sale_order_line l2 where l2.order_id=so.id and l2.product_id in (24,40,32,43)) then 'POS'
     when exists (select 1 from sale_order_line l2 where l2.order_id=so.id and l2.product_id in (25,26,41,42)) then 'SFA'
     else '?' end end
"""
MRR_SQL = "case when so.currency_id=1 then so.recurring_monthly*12000 else so.recurring_monthly end"
AMT_SQL = "case when l.currency_id=1 then l.amount_signed*12000 else l.amount_signed end"
