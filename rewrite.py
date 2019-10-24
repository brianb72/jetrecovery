# 2017 BrianB
# Converts the DateTime value in the csv from a Double to a 'YYYY-MM-DD HH:MM:SS' string.

from datetime import datetime, timedelta
import csv

count = 0

# This is epoch for Jet4 databases
access_epoch = datetime(1899, 12, 30)

with open("converted.csv", "w") as output_file:
    writer = csv.writer(output_file, delimiter=',')
    with open("output.csv", "r") as input_file:
        reader = csv.reader(input_file)

        # Write the header
        line = reader.next()
        writer.writerow(line)

        # Convert the rows
        for row in reader:
            d_as_double = float(row[1])
            d_integer_part = int(d_as_double)
            d_fractional_part = abs(d_as_double - d_integer_part)
            row[1] = access_epoch + timedelta(days=d_integer_part) + timedelta(days=d_fractional_part)
            writer.writerow(row)
            count += 1
            if count % 100000 == 0:
                print("Proccessed {} lines".format(count))


