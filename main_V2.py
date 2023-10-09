import os
import imagesize
import xlsxwriter

fdir = r"C:\Users\XXX\Downloads\full_size"

# get the list of all pictures in directory
pictures = []
for (dirpath, dirnames, filenames) in os.walk(fdir):

    'SIZE MATTERS - check only sized strings'
    # pictures += [os.path.join(dirpath, file) for file in filenames if file.endswith(('.jpg','.jpeg','.png','.gif'))]
    pictures += [os.path.join(dirpath, file) for file in filenames if file.endswith('')]

# check pictures atributes
sizes = []
resolutions = []
for item in pictures:

    # get file size in kb
    sizes.append(os.stat(item).st_size)

    # get file resolution
    width, height = imagesize.get(item)
    resolutions.append(str(height) + 'x' + str(width))

print('Pictures scanned: ', len(pictures))

# prepare excel
filename = r"C:\Users\XXX\Downloads\results.xlsx"
workbook = xlsxwriter.Workbook(filename)
worksheet = workbook.add_worksheet()

# prepare headers
worksheet.write(0, 0, 'url')
worksheet.write(0, 1, 'size kb')
worksheet.write(0, 2, 'resolution')
worksheet.write(0, 3, 'mix')
worksheet.write(0, 4, 'group')
worksheet.write(0, 5, 'duplicate?')
worksheet.write(0, 6, 'uniformed pictures')
worksheet.write(0, 7, '!!! SORT MIX !!!')

# write to excel
row = 0
col = 0
for item in pictures:
    worksheet.write(row + 1, col, item)
    worksheet.write(row + 1, col + 1, sizes[row])
    worksheet.write(row + 1, col + 2, resolutions[row])
    if row == 1:
        worksheet.write(row, col + 3, '=B2&"x"&C2')
        worksheet.write(row, col + 4, 1)
        worksheet.write(row, col + 5, '=IF(OR(E2=E1,E2=E3),"YES","")')
        worksheet.write(row, col + 6, '=IF(AND(F2="YES",E2=E1),G1,A2)')
    if row == 2: worksheet.write(row, col + 4, '=IF(B3=B2,E2,E2+1)')
    row += 1

workbook.close()
os.startfile(filename)
