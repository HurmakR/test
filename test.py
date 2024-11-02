import csv


with open('repair_data.csv', newline='') as csvfile1, open('orderreturn_data.csv', newline='') as csvfile2:
    completed_repairs = csv.DictReader(csvfile1)
    return_orders = csv.DictReader(csvfile2)
    counter = 0
    if 'repair_data' in csvfile1.name:
        file1type='GSX'
    else:
        file1type='IT4'
    print(file1type)
    for repair in completed_repairs:
        found = False
        counter += 1
        print('-' * 20, counter)
        for order in return_orders:
            if order['Serial Number'] == repair['Serial Number']:
                print('bingo')
                found = True
        if not found: print('Failed')
