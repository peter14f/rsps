# script_xls_new.py

Goes through all the subdirectories in root_path and creates a spreadsheet that summarizes which 
level-2 directories do not have any files in its level-3 subdirectories.

Not that all level-2 directories all have the same level-3 subdirectories.

Example:

/home/upload
├── 001White
│   ├── Item1
│   │   ├── a.txt
│   │   └── b.txt
│   ├── Item2
│   │   └── a.jpg
│   └── Item3
│       └── c.pdf
├── 002Red
│   ├── Item1
│   │   ├── a.txt
│   │   └── b.txt
│   ├── Item2
│   └── Item3
│       └── c.pdf
└── 003Blue
    ├── Item1
    │   ├── a.txt
    │   └── b.txt
    ├── Item2
    │   └── a.jpg
    └── Item3
    
should generate the following excel spreadsheet
         
         Item1  Item2  Item3
001White   V      V      V 
002Red     V             V
003Blue    V      V         

script_xls.py and script_xls_new.py use Python library openpyxl to produce the .xlsx file.
