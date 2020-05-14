import re
mail = '102@zf.com'

if re.match(r'^([a-zA-Z.-]+)@zf.com$', mail):
    print('yes')
else:
    print('no')