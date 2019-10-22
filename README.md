# Convert Yahoo Groups Polls to Excel

If you've recently backed up a Yahoo Group using e.g. [`yahoo-groups-archiver`](https://github.com/nsapa/yahoo-group-archiver), you will have noticed that your group's polls are archived in a series of `.json` files. This script will convert those files to an Excel workbook.

## Usage

0. Install this package's dependencies using `pip`: `pip install -r requirements.txt`

1. Use [`yahoo-groups-archiver`](https://github.com/nsapa/yahoo-group-archiver) to download/archive your Yahoo Group; it will create an output directory with the same name as your group, containing a variety of subdirectories (`email`, `photos`, etc.).

2. `python poll2xlsx.py PATH/TO/WHEREVER/YOUR/GROUP/DOWNLOADED/TO/polls my_group_polls.xlsx`

3. Enjoy!