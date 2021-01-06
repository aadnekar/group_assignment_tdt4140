# Assign Groups Command Tool
A command tool for assigning groups in TDT4140 at NTNU.

## Requirements

- Python 3

You may need to run `pip3 install -r requirements.txt`

## How to use

What we aim to do is to run a questioneer in Microsoft Forms and export the results to xlsx format. We may run this script using on that file to generate the groups that we want.

You have to provide the xlsx file exported from Forms by using the `-f` flag.

You may provide `--group_size` to alter the intended group size. We want the size to be consistent since we decide the groups, we may try to get as many groups as possible with this size.

You may also provide `-O` which stands for *output* and is gives you the option to select an appropriate file name for the result xlsx file.

An example is:

```bash
./assign_groups.py --group_size 9 -O custom_output -f exported_file.xlsx
```
