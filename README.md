#script_xls_new.py

Goes through all the subdirectories in *root_path* and creates a spreadsheet that summarizes which 
level-2 directories do not have any files in its level-3 subdirectories.

Note that the program assumes that all level-2 directories all have the same level-3 subdirectories.

Example:

<pre>
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
</pre>
should generate the following excel spreadsheet

<table>
    <tr>
        <td></td>
        <td>Item1</td>
        <td>Item2</td>
        <td>Item3</td>
    </tr>
    <tr>
        <td>001White</td>
        <td>V</td>
        <td>V</td>
        <td>V</td>
    </tr>
    <tr>
        <td>002Red</td>
        <td>V</td>
        <td></td>
        <td>V</td>
    </tr>
    <tr>
        <td>003Blue</td>
        <td>V</td>
        <td>V</td>
        <td></td>
    </tr>
</table>

script_xls_new.py uses the Python library **openpyxl** to produce the .xlsx file.
