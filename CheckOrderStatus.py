

import pandas as pd

demand = pd.read_excel('Z:\Python projects\Chris Forecast\Forecast\demand.xlsx')

print('Enter order number.')
order = input()

relaventOrders = demand[demand['GRANDPARENT'] == order].copy()
orderCheckdf = pd.DataFrame()

for part in relaventOrders['PART'].unique():
	partCheck = demand[demand['PART'] == part].copy()
	orderCheckdf = orderCheckdf.append(partCheck.copy())

highlights = orderCheckdf[(orderCheckdf['ITEM'] == 'Phantom') & (orderCheckdf['GRANDPARENT'] == order)].copy()

writer = pd.ExcelWriter('Z:\Python projects\Chris Forecast\Forecast\ByOrder.xlsx', index=False)
orderCheckdf.to_excel(writer, 'timeline')
writer.save()

print('All done!')


