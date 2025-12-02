## Docs demo

### This python script is used to read a txt file containing variables names, a docx file as a template and it will generate an output directory containing output docx documents.

### Prequisites
Install python3 and python virtual-env.

### Usage
Create a directory where you want to run this script and open a terminal there and run below commands

Steps
1. [ ] Create a python virtualenv in current directory
2. [ ] Activate python virtualenv
3. [ ] Make changes in input file or template file and Run script

### Create python virtualenv using below command
```sh
python3 -m venv env
```

### Install a required python library
```sh
pip install docxtpl
```

### Activate pyton virtualenv using below command
```sh
source env/bin/activate
```

### Run script using below command
where
-k is path to variables text file
-t is path to docx document template file
-o is path to output folder name
-n is path and name to a newly generated docx document

```sh
main.py -k user-info.txt -t template.docx -o output -n "john-doe.docx"
```
