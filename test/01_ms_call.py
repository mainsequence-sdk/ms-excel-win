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
print("----------------")
print(ms_excel_wrappers.get_asset("BBG000BS7KS3"))
print("----------------")
print(ms_excel_wrappers.get_asset("BBG000BY29C7"))
print("----------------")

'''
curve_table_id	discount_curves
date_start	29-Jun-25
date_end	26-Dec-25
'''

start_date = datetime.datetime(2025, 6, 29)
end_date = datetime.datetime(2025, 12, 26)
print(ms_excel_wrappers.get_data_between_dates_from_node_identifier("discount_curves",start_date,end_date))