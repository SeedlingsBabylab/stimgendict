# sgd (stimuli generated dictionary)

This script walks a directory tree looking for files of the form AB_CD_stimuli.xlsx.
It pulls the pair_alpha, pair_words, pair_kind, and subject_number from each file and fills them
into an output csv file.

The script depends on openpyxl. You can install this with python's package manager, pip, with:

```bash
$: pip install openpyxl
```


## usage:

```bash
$: python sgd.py directory_root output_csv_path
```

The /data directory has nested _stimuli.xlsx files that you can try the script on.

for example:

```bash
$: python sgd.py data output.csv
```
# csv output

The csv output will be of the form:

pair_alpha, pair_words, pair_kind, SubjectNumber