Most of them are script which  is used in HR function in my daily work.

1. L01_learning history.py is used to split learning history data into different files per user. The final output is the Excel file and pdf files. The test data is already uploaded into testdata folder.
The required packages are xlsxwriter, pandas.
2. L01_payslip_from_detail.py is used to generate payslip in excel and pdf. The format can be configured in the excel files, and the pay items will be removed if there is no value exist.
The required packages are xlsxwriter, pandas.


1. W10_headcount_report, input is the headcount report, and output is the headcount summary based on region, country, company, and reporting unit.
   The workspace is under temp/10headcount folder

2. t4_change_values_auto_group_v2.py is used to compare the data. The input are 2 employee data (ap version) files, and the output is a summary file.
   The configuration file is t4_auto_group.conf, and another excel file with name: t4_Fields_Group.xlsx

3. t4_read_authorization_user.py 将人员从列表转为行表。输入为权限表，输出为每个角色人员。