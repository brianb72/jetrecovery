**Jet Recovery**

Recover data from a damaged Microsoft Access database.
Copyright
I wrote this in 2017 over a weekend to save a customer from my old IT job.

Program Description:

    A customer was operating an in-house testing rig that used a pneumatic pump and sensors to measure 
    the long term performance of valves. Over one year of data was recorded in a Microsoft Access 
    database.

    The first ~10.5 megabytes of their database file was overwritten with random data. The customer 
    attempted to use several commercial recovery tools without success. The only backup of the file 
    was months old. The customer sent me a copy of the corrupted database, as well as their old backup.

    The files were MDB Jet4 databases. The damage to the begining of the newer file destroyed the 
    database definition page and all table definitions. By looking at the old file I obtained table 
    names and definitions.

    I wrote this utility in C to scan the raw MDB file and recover rows of the customers data. All of
    the data that was needed was located in one table, 'tblResults', which I extracted to a CSV file.

    The corrupted file had an auto-incrementing ID column, and the maximum row value in the data was 
    2,149,759. The first row observed after the corrupted data was 405,378. The older file had row ID's
    of 62 to ~900,000, and I was able to copy data from the older file to replace the missing corrupted 
    data. I spot checked overlapping ID's between the old and new files and the data matched.

    'rewrite.py' is a short script that converts the datetime column in the CSV to the proper value.
