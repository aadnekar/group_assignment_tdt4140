# Assign Groups Command Tool
A command tool for assigning groups in TDT4140 at NTNU.

## Requirements

- Python 3

You may need to run `pip3 install -r requirements.txt`

## How to use

What we aim to do is to run a questioneer in Microsoft Forms and export the results to xlsx format. We may run this script using on that file to generate the groups that we want.

* You have to provide the xlsx file exported from Forms by using the `-f` flag.
* You may also provide `-O` which stands for *output* and is gives you the option to select an appropriate file name for the result xlsx file. If this flag is not provided result file will default to `result.xlsx`.
* You have to provide one of these argumenters:
  * `-n` (--number_of_groups): assigns students to one of a given total number of groups.
  * `-Gs` (--group_size): calculates the total number of groups with regards to the number of students and a maximum total group numbers of students per group.

If both `-n` and `-Gs` are provided, the `-n` (number of groups) thrumphs the other.

An example is:

```bash
./assign_groups.py --group_size 9 -O custom_output -f exported_file.xlsx
```
