import csv
import math

INPUT_FILE = "FR_194130526.csv"
BASE_NAME = "PR02_R5_130526"
NUM_FILES = 4

with open(INPUT_FILE, newline="", encoding="utf-8") as f:
    reader = csv.reader(f)
    header = next(reader)
    ids = [row for row in reader]

total = len(ids)
GROUP_SIZE = math.ceil(total / NUM_FILES)
num_files = NUM_FILES
print(f"Total IDs encontrados: {total} -> {num_files} archivos de ~{GROUP_SIZE} cada uno")

for i in range(num_files):
    start = i * GROUP_SIZE
    end = min(start + GROUP_SIZE, total)
    chunk = ids[start:end]

    filename = f"{BASE_NAME}{i + 1}.csv"
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        writer.writerow(header)
        writer.writerows(chunk)

    print(f"{filename}: {len(chunk)} IDs (filas {start + 1} a {end})")
