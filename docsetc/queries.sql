-- fetch not filled on sheet
SELECT * FROM wasted WHERE root_order_id IS NULL AND NOT filled AND sheet_name="БпЛА"