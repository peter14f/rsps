#script_xls_new.py
#script_xls.py

Goes through all the subdirectories in root_path and creates a spreadsheet that summarizes which 
level-2 directories do not have any files in its level-3 subdirectories.

Not that all level-2 directories all have the same level-3 subdirectories.

Example:

<pre>
/home/upload<br />
├── 001White<br />
│   ├── Item1<br><br />
│   │   ├── a.txt<br />
│   │   └── b.txt<br />
│   ├── Item2<br />
│   │   └── a.jpg<br />
│   └── Item3<br />
│       └── c.pdf<br />
├── 002Red<br />
│   ├── Item1<br />
│   │   ├── a.txt<br />
│   │   └── b.txt<br />
│   ├── Item2<br />
│   └── Item3<br />
│       └── c.pdf<br />
└── 003Blue<br />
    ├── Item1<br />
    │   ├── a.txt<br />
    │   └── b.txt<br />
    ├── Item2<br />
    │   └── a.jpg<br />
    └── Item3<br />
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

script_xls.py and script_xls_new.py use Python library openpyxl to produce the .xlsx file.
