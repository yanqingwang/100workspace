# -*- coding: utf-8 -*-
"""
@author: Z659190
排除不需要的数据
清理失真数据 -- 离职转公司
预处理简单数据
"""
Columns_rename = {"Reporting Unit (Reporting Unit ID)": "RU",
                  'Division/Corporate Function/Region (ID)':'Division ID',
                  'Division/Corporate Function/Region (Label)':'Division Label',
                  'BU, Divisional Function/GDF (ID)':'BU ID',
                  'BU, Divisional Function/GDF (Label)':'BU Label',
                  'Regular/Limited Employment (Label)':'Regular Limited Employment',
                  'Local Payroll Area/Pay Group (Label)':'Local Payroll'
                  }

clean_columns = ['External Agency Worker', 'External Agency Worker.1', 'Contingent Worker (External Code)',
                 'ID country','National Id','City (Label)','Home Address','Zip Code',
                 'Local Payroll',
                 'Local Employee Class (External Code)','Local Employee Class (Label)','Local Employment Type (External Code)',
                 'Local Employment Type (Label)',
                 'Probation Status (Label)','Probationary Period End Date','Cell Phone Number','Contract End Date',
                 'Matrix Manager Type','Matrix Manage Global ID','Matrix Manager ID','Matrix Manager Position',
                 'Matrix Manage First Name','Matrix Manager Last Name']

clean_col_date = ['Hire Date','Hire Date.1','Termination Date.1','Termination Date']

format_columns = {'ZF Global ID': 'str', 'Admin Group (ID)': 'str'}
# format_columns = {'ZF Global ID': 'str', 'Admin Group (ID)': 'str','Termination Date': 'datetime64' }