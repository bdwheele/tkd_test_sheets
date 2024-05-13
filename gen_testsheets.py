#!/bin/env python3
import argparse
from string import Template
from pathlib import Path
import csv
import sys
import subprocess

ranks = {
    'Y': ('yellow', '8th Kup / Yellow Belt'),
    'O': ('orange', '7th Kup / Orange Belt'),
    'G': ('green', '6th Kup / Green Belt'),
    'P': ('purple', '5th Kup / Purple Belt'),
    'b': ('blue', '4th Kup / Blue Belt'),
    'B': ('brown', '3rd Kup / Brown Belt'),
    'R': ('red', '2nd Kup / Red Belt'),
    'T': ('temp', '1st Kup / Temp Belt'),
    '1': ('1stdan', '1st Dan Black Belt'),
    '2': ('2nddan', '2nd Dan Black Belt'),
    '3': ('3rddan', '3rd Dan Black Belt')
}


def main():
    outdir = sys.path[0] + "/test_sheets"
    # get our test sheet template
    template = Template((Path(sys.path[0]) / "template.html").read_text())

    for rank in ranks:
        tables = {}
        subtitles = {}
        cur_table = None
        with open(Path(sys.path[0]) / "inventory.csv",newline='') as csvfile:
            rdr = csv.DictReader(csvfile)
            for row in rdr:
                if row['Label'] == '':
                    # this row has no label, so we'll skip it.
                    continue
                if row[rank] != '':
                    if row['Type'] == 'T':
                        cur_table = row['Label']
                        if cur_table not in tables:
                            tables[cur_table] = []                    
                    elif row['Type'] == 'S':
                        subtitles[cur_table] = row['Label']
                    else:
                        tables[cur_table].append([row['Type'], row['Label'], row[rank]])
                
        # we now have all of the tables in memory...generate the HTML for them.
        html_tables = []
        for tablename, tabledata in tables.items():
            if tablename in subtitles:
                subtitle = f"""<br/><span class="subtitle">{subtitles[tablename]}</span>"""
            else:
                subtitle=""
            html = f"""
        <table class="ttable">       
            <caption>{tablename}{subtitle}</caption>
            <colgroup>
                <col class="narrow"/>
                <col class="narrow"/>
                <col/>
            </colgroup>
            <tr><td>&nbsp;</td><th>Score</th><th>Comments</th></tr>"""
            collapse = False
            for row in tabledata:
                if row[0] == 'H':
                    if collapse:
                        html += "<tr><td/>&nbsp;<td/><td/></tr>"
                    html += f"""<tr><td class="theader">{row[1]}</td><td></td><td></td></tr>"""
                    collapse = row[2] == 'C'
                else:
                    if (not collapse) or (collapse and row[2] == 'N'):
                        html += f"""<tr><td>{row[1]}</td><td></td><td></td></tr>"""                
            html += "</table>"            
            html_tables.append(html)

        html_file = f"{outdir}/{ranks[rank][0]}.html"
        with open(html_file, "w") as f:
            f.write(template.safe_substitute(title=ranks[rank][1], tables='\n'.join(html_tables)))

        # generate the word doc
        subprocess.run(['pandoc', html_file, "-o", f"{outdir}/{ranks[rank][0]}.docx"])
        # generate the pdf
        subprocess.run(['weasyprint', html_file, f"{outdir}/{ranks[rank][0]}.pdf"])






if __name__ == "__main__":
    main()