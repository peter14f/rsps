import re
import json

tab = {}
p = re.compile("([0-9]+).+")

with open("phone.csv", "rt", encoding = "utf-8-sig") as in_file:
    for line in in_file:
        line = line[:-1]
        cells = line.split(",")
        match = p.search(cells[0])
        school = match.group(1)
        new_entry = {
            "school": cells[0],
            "number": cells[1]
            }
        tab[school] = new_entry

print(json.dumps(tab, ensure_ascii=False))

with open('phone.json', 'w') as outfile:
    json.dump(tab, outfile, ensure_ascii=False)
