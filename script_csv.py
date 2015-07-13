import io
import os.path
import re
import csv

def get_size(start_path = '.'):
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(start_path):
        for f in filenames:
            fp = os.path.join(dirpath, f)
            total_size += os.path.getsize(fp)
    return total_size

def missing_required_item(school_item):
    if school_item[0]=='' or \
       school_item[3]=='' or \
       school_item[8]=='' or \
       school_item[9]=='' or \
       school_item[11]=='' or \
       school_item[13]=='':
        return True
    else:
        return False

root_path = "Y:\\"
#root_path = "C:\\fakeZ"
print("Looking at directories in '" + root_path + "'")

root_list = os.listdir(root_path)
root_list = sorted(root_list)

csvfile = open('summary.csv', 'w', newline='')
csvwriter = csv.writer(csvfile)
schools = []
items = []

for school in root_list:
    if re.match("^[0-9]", school):
        schools.append(school)

rsps_path = os.path.join(root_path, schools[21])
items_list = os.listdir(rsps_path)
cvt_map = {u'一': 10, u'二': 20, u'三': 30, u'四': 40}
items_map = {}

for item in items_list:
    value = int(item[4]) + cvt_map.get(item[1])
    items_map[value] = item

for key in sorted(items_map.keys()):
	items.append(items_map.get(key))

rows = []
first_row = ['', u'缺*成果']
first_row.extend(items)
rows.append(first_row)

for school in schools:
    school_path = os.path.join(root_path, school)
    new_row = [school, '']
    school_item = []
    for item in items:
        item_path = os.path.join(school_path, item)
        size = get_size(item_path)
        if size > 0:
            school_item.append('V')
        else:
            school_item.append('')
    missing_requirement = missing_required_item(school_item)
    new_row.extend(school_item)
    if missing_requirement:
        new_row[1] = 'V'
    rows.append(new_row)

csvwriter.writerows(rows)

#print(items)
#print (schools)
csvfile.close()
