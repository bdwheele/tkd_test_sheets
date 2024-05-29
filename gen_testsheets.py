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
    parser = argparse.ArgumentParser()
    parser.add_argument("--full", default=False, action="store_true", help="Don't collapse headers")
    args = parser.parse_args()


    outdir = sys.path[0] + "/test_sheets"
    # get our test sheet template
    template = Template((Path(sys.path[0]) / "template.html").read_text())

    # get the techniques sheet template
    tech_template = Template((Path(sys.path[0]) / "techniques.html").read_text())


    # walk the ranks, generating each sheet as we go.
    for rank in ranks:
        print(f"Generating {ranks[rank]}")
        data = read_inventory(Path(sys.path[0]) / "inventory.csv", rank)

        # generate the HTML tables for the data...
        html_tables = []
        tech_tables = []
        for tdata in data['tables']:
            thtml = html = """<table class="ttable">\n"""
            title = tdata['title']
            if tdata['subtitle']:
                # append the subtitle
                title += f"""</br><span class="subtitle">{tdata['subtitle']}</span>"""
            html += f"""  <caption>{title}</caption>\n"""
            thtml += f"""  <caption>{title}</caption>\n"""
            html += f"""  <colgroup><col class="narrow"/><col class="narrow"/><col/></colgroup>\n"""
            thtml += f"""  <colgroup><col class="narrow"/><col/></colgroup>\n"""
            html += f"""  <tr><td>&nbsp;</td><th>Score</th><th>Comments</th></tr>\n"""
            thtml += f"""  <tr><td>&nbsp;</td><th>Notes</th></tr>\n"""
            for header in tdata['headers']:
                if not header['techniques']:
                    # skip empty headers
                    continue
                if header['label'] != '':
                    html += f"""  <tr><td class="theader">{nbsp(header['label'])}</td><td></td><td></td></tr>\n"""
                    thtml += f"""  <tr><td class="theader">{nbsp(header['label'])}</td><td></td></tr>\n"""
                header_rows = 0
                for t in header['techniques']:
                    label = nbsp(t['label'])
                    if t['type'] == 'N':
                        # new technique
                        label = f'<span class="new">{label}</span>'
                    elif t['type'] != 'X':
                        label = f'{label} ({t["type"]})'

                    if args.full or header['type'] == 'X' or (header['type'] == 'C' and t['type'] != 'X'):
                        # either show all techniques or only show those that 
                        # are new or have a count
                        html += f"  <tr><td>{label}</td><td></td><td></td></tr>\n"
                        header_rows += 1
                        
                    thtml += f"  <tr><td>{label}</td><td></td></tr>\n"
                                        
                if header['type'] == 'C' and header_rows == 0:
                    # add a blank row for comments.
                    html += "  <tr><td>&nbsp;</td><td></td><td></td></tr>\n"
                    thtml += "  <tr><td>&nbsp;</td><td></td></tr>\n"

            html += """</table>"""
            thtml += """</table>"""
            html_tables.append(html)
            tech_tables.append(thtml)




        # generate the test sheet
        html_file = f"{outdir}/{ranks[rank][0]}.html"
        with open(html_file, "w") as f:
            f.write(template.safe_substitute(title=ranks[rank][1], 
                                             tables='\n'.join(html_tables),
                                             revision=data['revision']))
        # generate the word doc
        subprocess.run(['pandoc', html_file, "-o", f"{outdir}/{ranks[rank][0]}.docx"])
        # generate the pdf
        subprocess.run(['weasyprint', html_file, f"{outdir}/{ranks[rank][0]}.pdf"])

        # generate the technique list
        html_file = f"{outdir}/{ranks[rank][0]}-techniques.html"
        with open(html_file, "w") as f:
            f.write(tech_template.safe_substitute(title=ranks[rank][1], 
                                             tables='\n'.join(tech_tables),
                                             revision=data['revision']))
        # generate the word doc
        subprocess.run(['pandoc', html_file, "-o", f"{outdir}/{ranks[rank][0]}-techniques.docx"])
        # generate the pdf
        subprocess.run(['weasyprint', html_file, f"{outdir}/{ranks[rank][0]}-techniques.pdf"])




def nbsp(text):
    return text.replace(' ', '&nbsp;')

def read_inventory(invfile, rank):        
    # read the inventory and return the data for the given rank
    revision = "No Revision"
    tables = []                
    with open(invfile, newline='') as csvfile:
        rdr = csv.DictReader(csvfile)
        for row in rdr:
            # make sure that leading/trailing whitespace is stripped...
            row = {k: v.strip() for k, v in row.items()}
            if row['Type'] == 'R':
                # this is the test sheet revision
                revision = row['Label']
                continue

            if row['Type'] == '#' or row['Label'] == '' or row[rank] == '':
                # ignore rows that don't have a label or are a comment, or
                # are blank for the rank we're looking for.
                continue

            elif row['Type'] == 'T':
                # create a new table
                tables.append({'title': row['Label'],
                                'subtitle': None,
                                'headers': [{'label': '',
                                             'type': 'X',
                                             'techniques': []}]})
                                    
            elif row['Type'] == 'S':
                # set the subtitle of the current table
                tables[-1]['subtitle'] = row['Label']

            elif row['Type'] == 'H':
                # start a new header section
                tables[-1]['headers'].append({'label': row['Label'],
                                              'type': row[rank],
                                              'techniques': []})
            else:
                # this must just be a technique...add it where needed.
                tables[-1]['headers'][-1]['techniques'].append({'type': row[rank],
                                                                'label': row['Label']})

    return {'revision': revision,
            'tables': tables}


if __name__ == "__main__":
    main()