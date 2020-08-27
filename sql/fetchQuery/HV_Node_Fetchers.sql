select * from nodes_high_voltage_data
where start_date IN
(select start_date as s from nodes_high_voltage_data where rownum<2)
