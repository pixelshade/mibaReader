import os
import csv
import glob
import xlsxwriter
import progressbar


PATH = r'./data/'
OUTPATH = r'./out/'
def main():
  allLinesDict = loadFiles()
  data = cropParseDict(allLinesDict)
  createExcel(data)

def loadFiles():
  print('loading files')
  res = {}
  bar = progressbar.ProgressBar()
  for fname in bar(glob.glob('{}*.txt'.format(PATH))):
    prop = fname.replace('.txt', '').split('_')[-1]
    mms = fname.replace('.txt', '').split('_')[1].split('-')[0]
    # print(mms, 'mms')

    lines = []
    with open(fname, 'rb') as f:
      reader = csv.reader(f, delimiter=';', skipinitialspace=True)
      try:
        for i, line in enumerate(reader):
          if i < 4:
            continue
          lines.append(line)
          res.setdefault((prop, mms), []).append(line)
      except Exception as e:
        print(fname, e.message)
  return res


def cropParseDict(allLinesDict):
  print('detecting borders')
  bar = progressbar.ProgressBar()
  res = {}
  for key in bar(allLinesDict):
    prop, mms = key
    lines = allLinesDict[key]
    # start, end = findBorders(lines)
    start = 21
    end = 121
    for line in lines:
      if(line == '' or len(line) == 0 or len(line[0]) == 0):
        continue
      line = line[:2] + line[start:end]

      for i, x in enumerate(line):
        try:
          if(x == '()'):
            line[i] = None
          elif ':' in x:
            line[i] = x
          elif x is None:
            line[i] = x
          elif '.' in x:
            line[i] = float(x)
          else:
            line[i] = x
        except Exception as e:
          print(key,e,i, line)
      res.setdefault((prop, mms), []).append(line)
  return res

def addAvgs(dataSht):
  lenline = 0
  maxlenline = 0
  startLine = 4
  for i, line in enumerate(dataSht):
    line.insert(0,None)
    lenline = len(line)
    maxlenline = max(maxlenline, lenline)
    endAlph = xlsxwriter.utility.xl_col_to_name(lenline-2)
    rowi = startLine + i + 5 #len(stats)
    stats = [
      '=AVERAGE(I{}:{}{})'.format(rowi, endAlph, rowi),
      '=MIN(I{}:{}{})'.format(rowi, endAlph, rowi),
      '=MAX(I{}:{}{})'.format(rowi, endAlph, rowi),
      None,
      '{}'.format(i)
    ]
    # print(len(line))
    dataSht[i] = stats + line
    # print(len(line))
  firstLineI = len(dataSht)+3

  padding = [None] * 7
  firstLineAvgs = ['AVERAGE', 'MIN', 'MAX', None, None, None, None, 'AVERAGE']
  for i in range(len(firstLineAvgs), len(firstLineAvgs) + maxlenline - 3):
    endAlph = xlsxwriter.utility.xl_col_to_name(i)
    firstLineAvgs.append('=AVERAGE({}{}:{}{})'.format(endAlph, startLine, endAlph, firstLineI))
  dataSht.insert(0, firstLineAvgs)

  firstLineAvgs = list(padding) + ['MAX']
  for i in range(len(firstLineAvgs), len(firstLineAvgs) + maxlenline -3):
    endAlph = xlsxwriter.utility.xl_col_to_name(i)
    firstLineAvgs.append('=MAX({}{}:{}{})'.format(endAlph, startLine, endAlph, firstLineI))
  dataSht.insert(0, firstLineAvgs)

  firstLineAvgs = list(padding) + ['MIN']
  for i in range(len(firstLineAvgs), len(firstLineAvgs) + maxlenline - 3):
    endAlph = xlsxwriter.utility.xl_col_to_name(i)
    firstLineAvgs.append('=MIN({}{}:{}{})'.format(endAlph, startLine, endAlph, firstLineI))
  dataSht.insert(0, firstLineAvgs)

  return dataSht

def addConditionalFormatting(worksheet, dataSht):
  rows = len(dataSht)
  cols = len(dataSht[rows / 2])
  colName = xlsxwriter.utility.xl_col_to_name(cols)
  end = '{}{}'.format(colName, rows)
  # main raw data
  worksheet.conditional_format('I4:{}'.format(end), {'type': '3_color_scale'})
  # top stats
  worksheet.conditional_format('I1:{}1'.format(colName), {'type': '3_color_scale'})
  worksheet.conditional_format('I2:{}2'.format(colName), {'type': '3_color_scale'})
  worksheet.conditional_format('I3:{}3'.format(colName), {'type': '3_color_scale'})
  # left stats
  worksheet.conditional_format('A4:A{}'.format(rows), {'type': '3_color_scale'})
  worksheet.conditional_format('B4:B{}'.format(rows), {'type': '3_color_scale'})
  worksheet.conditional_format('C4:C{}'.format(rows), {'type': '3_color_scale'})

def findBorders(lines):
  half = len(lines) / 2
  line = lines[half]
  start = 0
  lenline = len(line)
  end = lenline

  for i, l in enumerate(line[2:]):
    if l != '()' and l!='':
      start = i + 2
      break
  for i, l in enumerate(reversed(line[2:])):
    if l != '()' and l!='':
      end = lenline - i
      break
  return start, end


def createExcel(data):
  print('creating excel')
  bar = progressbar.ProgressBar()
  for key in bar(data):
    row = 0
    col = 0
    try:
        os.makedirs(OUTPATH)
    except:
        pass
    workbook = xlsxwriter.Workbook('{}{}_{}.xlsx'.format(OUTPATH, key[0], key[1]))
    worksheet = workbook.add_worksheet('data')
    data[key] = addAvgs(data[key])
    for line in data[key]:
      worksheet.write_row(row, 0, line)
      row += 1
    addConditionalFormatting(worksheet, data[key])
    # worksheet.write(row, 0, 'Total')
    # worksheet.write(row, 1, '=SUM(B1:B4)')
  print('saving...')
  print(os.path.abspath(workbook.filename))
  workbook.close()
  print('DONE')

main()
