'''
Docstring for test.01_MScall

This script tests the MS Excel wrappers and client functionality
by logging in and retrieving an asset.

'''


from ms_excel_win import ms_excel_wrappers
import mainsequence.client as msc

print("Test Script")
ms_excel_wrappers.login_dialog()

print(ms_excel_wrappers.get_asset("90_CBPF_48"))
print("----------------")
print(ms_excel_wrappers.get_asset("BBG000BS7KS3"))
print("----------------")
print(ms_excel_wrappers.get_asset("BBG000BY29C7"))