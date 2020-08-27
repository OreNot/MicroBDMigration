import os

import requests

path = 'C:\microservices\order_base.xlsm'

url = 'http://webdo.gk.rosatom.local/Files/База доходных договоров 2020.xlsm'


try:
    os.remove(path)
except OSError:
    pass

r = requests.get(url, allow_redirects=True)

open(path, 'wb').write(r.content)