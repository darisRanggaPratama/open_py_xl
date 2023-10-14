import openpyxl as pyxl
import pprint
print('Opening workbook\n')
wbook = pyxl.load_workbook('censuspopdata.xlsx')
wsheet = wbook['Population by Census Tract'] #wbook.get_sheet_by_name('Population by Census Tract')
countyData = {}

print('Reading rows')
for row in range(2, wsheet.max_row + 1):
    # Each Row has data
    state = wsheet['B' + str(row)].value
    county = wsheet['C' + str(row)].value
    pop = wsheet['D' + str(row)].value

    countyData.setdefault(state, {})
    countyData[state].setdefault(county, {'tracts': 0, 'pop': 0})
    countyData[state][county]['tracts'] += 1
    countyData[state][county]['pop'] += int(pop)

print('Writing results')
resultFile = open('census2010.py', 'w')
resultFile.write('allData = ' + pprint.pformat(countyData))
resultFile.close()
print('Done')




# readCensusExcel.py
