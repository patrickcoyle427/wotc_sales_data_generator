import csv

with open('data.csv', newline='') as csvfile:

    reader = csv.reader(csvfile, delimiter=',')

    next(reader, None)
    # Skips the header line

    data = []

    for row in reader:

        data.append((row[1], row[0], row[2], row[3], row[4], row[5]))


    for i in data:

        print(i)
