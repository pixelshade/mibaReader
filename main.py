import csv
import glob
import xlsxwriter
import progressbar


PATH = r'./data/'

def main():
  data = loadFiles()
  createExcel(data)

def loadFiles():
  print('loading files')
  res = {}
  bar = progressbar.ProgressBar()
  for fname in bar(glob.glob('{}*.txt'.format(PATH))):
    prop = fname.replace('.txt', '').split('_')[-1]
    mms = fname.replace('.txt', '').split('_')[1].split('-')[0]
    print(mms, 'mms')

    lines = []
    with open(fname, 'rb') as f:
      reader = csv.reader(f, delimiter=';', skipinitialspace=True)
      for i, line in enumerate(reader):
        if i < 4:
          continue
        lines.append(line)

    # print(len(line))
    start, end = findBorders(lines)
    # print('borders:',start, end)
    for line in lines:
      if(len(line) == 0 or len(line[0]) == ''):
        continue
      line =  line[:2]+line[start:end]
      res.setdefault((prop, mms), []).append(line)

  return res


def addAvgs(dataSht):
  lenline = 0
  for i, line in enumerate(dataSht):
    line.append(None)
    lenline = len(line)
    endAlph = xlsxwriter.utility.xl_col_to_name(lenline-2)

    stats = [
      '=AVERAGE(C{}:{}{})'.format(i + 1, endAlph, i + 1),
      '=MIN(C{}:{}{})'.format(i + 1, endAlph, i + 1),
      '=MAX(C{}:{}{})'.format(i + 1, endAlph, i + 1),
      None,
      '{}'.format(i)
    ]
    # print(len(line))
    line = stats + line
    # print(len(line))
  lastLineI = len(dataSht)+2
  lastLineAvgs = [None,None,None]
  startLine = 4

  for i in range(len(lastLineAvgs), lenline):
    endAlph = xlsxwriter.utility.xl_col_to_name(i)
    lastLineAvgs.append('=AVERAGE({}{}:{}{})'.format(endAlph, startLine, endAlph, lastLineI))
  dataSht.insert(0, lastLineAvgs)

  lastLineAvgs = [None,None,None]
  for i in range(len(lastLineAvgs), lenline):
    endAlph = xlsxwriter.utility.xl_col_to_name(i)
    lastLineAvgs.append('=MAX({}{}:{}{})'.format(endAlph, startLine, endAlph, lastLineI))
  dataSht.insert(0, lastLineAvgs)

  lastLineAvgs = [None, None, None]
  for i in range(len(lastLineAvgs), lenline):
    endAlph = xlsxwriter.utility.xl_col_to_name(i)
    lastLineAvgs.append('=MIN({}{}:{}{})'.format(endAlph, startLine, endAlph, lastLineI))
  dataSht.insert(0, lastLineAvgs)

  return dataSht

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
  workbook = xlsxwriter.Workbook('hello.xlsx')
  for key in data:
    bar = progressbar.ProgressBar()

    row = 0
    col = 0
    worksheet = workbook.add_worksheet(key[0])
    data[key] = addAvgs(data[key])
    for line in bar(data[key]):
      worksheet.write_row(row, 0, line)
      # worksheet.write(row, col,     item)
      # worksheet.write(row, col + 1, cost)
      row += 1

    # worksheet.write(row, 0, 'Total')
    # worksheet.write(row, 1, '=SUM(B1:B4)')
  print('saving...')
  workbook.close()

main()
