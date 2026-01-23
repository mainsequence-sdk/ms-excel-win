'''
Docstring for test.01_MScall

This script tests the MS Excel wrappers and client functionality
by logging in and retrieving an asset.

'''


from ms_excel_win import ms_excel_wrappers
import mainsequence.client as msc
import datetime
import ms_excel_win.auth as msauth

msauth._load_tokens()

print("Test Script")
ms_excel_wrappers.login_dialog()

print(ms_excel_wrappers.get_asset("90_CBPF_48"))
print("-----------------------------------")
print(ms_excel_wrappers.get_asset("BBG000BS7KS3"))
print("-----------------------------------")
print(ms_excel_wrappers.get_asset_field("BBG000BS7KS3","current_snapshot.ticker"))
print("-----------------------------------")
print(ms_excel_wrappers.get_asset("BBG000BY29C7"))
print("-----------------------------------")
print(ms_excel_wrappers.get_asset_field("BBG000BY29C7","current_snapshot.ticker"))

print("-----------------------------------")
'''
curve_table_id	discount_curves
date_start	29-Jun-25
date_end	26-Dec-25
'''

start_date = datetime.datetime(2025, 6, 29)
end_date = datetime.datetime(2025, 12, 26)
print(ms_excel_wrappers.get_data_between_dates_from_node_identifier("discount_curves",start_date,end_date))

print("-----------------------------------")
'''
node_identifier	algoseek_daily_ohlc_DEMO

unique_indetifier_list	start_date	end_date
BBG000BS7KS3	1-Jan-22	1-Jan-23
BBG000BY29C7		

'''
node_identifier="algoseek_daily_ohlc_DEMO"
unique_indetifier_list=["BBG000BS7KS3","BBG000BY29C7"]
start_date = datetime.datetime(2022, 1, 1)
end_date = datetime.datetime(2023, 1, 1)

print(ms_excel_wrappers.get_data_between_dates_from_node_identifier(node_identifier, start_date, end_date, unique_indetifier_list ))