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
    template = Template((Path(sys.path[0]) / "test_template.html").read_text())

    # get the techniques sheet template
    tech_template = Template((Path(sys.path[0]) / "techniques_template.html").read_text())


    # walk the ranks, generating each sheet as we go.    
    for rank in ranks:
        print(f"Generating {ranks[rank]}")

        # read the csv inventory
        data = read_inventory(Path(sys.path[0]) / "inventory.csv", rank)

        for sheet in [('test_template.html', gen_test_content, '-test'),
                      ('techniques_template.html', gen_tech_content, '-techniques')]:
            content = sheet[1](data, args.full)
            filebase = ranks[rank][0] + sheet[2]
            template = Template((Path(sys.path[0], sheet[0]).read_text()))
            with open(f"{outdir}/{filebase}.html", "w") as f:
                f.write(template.safe_substitute(title=ranks[rank][1],
                                                 content=content,
                                                 revision=data['revision']))
            
            # generate the word doc & pdf
            subprocess.run(['pandoc', f'{outdir}/{filebase}.html', "-o", f"{outdir}/{filebase}.docx"])
            subprocess.run(['weasyprint', f'{outdir}/{filebase}.html', f"{outdir}/{filebase}.pdf"])


def gen_test_content(data, full=False):
    """Generate the content for a test sheet"""
    html_tables = []
    for tdata in data['tables']:
        html = """<table class="ttable">\n"""
        
        title = tdata['title']
        if tdata['subtitle']:
            # append the subtitle
            title += f"""<br/><span class="subtitle">{tdata['subtitle']}</span>"""
        html += f"""  <caption>{title}</caption>\n"""
        html += f"""  <colgroup><col class="narrow"/><col class="narrow"/><col/></colgroup>\n"""
        html += f"""  <tr><td>&nbsp;</td><th>Score</th><th>Comments</th></tr>\n"""
        for header in tdata['headers']:
            if not header['techniques']:
                # skip empty headers
                continue
            if header['label'] != '':
                html += f"""  <tr><td class="theader">{nbsp(header['label'])}</td><td></td><td></td></tr>\n"""
            header_rows = 0
            for t in header['techniques']:
                label = nbsp(t['label'])
                if t['type'] == 'N':
                    # new technique
                    label = f'<span class="new">{label}</span>'
                elif t['type'] != 'X':
                    label = f'{label} ({t["type"]})'

                if full or header['type'] == 'X' or (header['type'] == 'C' and t['type'] != 'X'):
                    # either show all techniques or only show those that 
                    # are new or have a count
                    html += f"  <tr><td>{label}</td><td></td><td></td></tr>\n"
                    header_rows += 1
            if header['type'] == 'C' and header_rows == 0:
                # add a blank row for comments.
                html += "  <tr><td>&nbsp;</td><td></td><td></td></tr>\n"

        html += """</table>"""
        html_tables.append(html)
    
    return "\n".join(html_tables)


def gen_tech_content(data, full=False):
    "generate technique sheet content"    
    tech_tables = []
    for tdata in data['tables']:        
        title = tdata['title']
        tech_html = f"""  <span class="title">{title}</span>\n"""
        if tdata['subtitle']:
            # append the subtitle
            tech_html += f"""<br/><span class="subtitle">{tdata['subtitle']}</span>"""        

        tech_html += """<ul>\n"""        
        for header in tdata['headers']:
            if not header['techniques']:
                # skip empty headers
                continue
            if header['label'] != '':
                tech_html += f"""  <li class="section">{header['label']}</li>\n"""
            for t in header['techniques']:
                label = t['label']
                if t['type'] == 'N':
                    # new technique
                    label = f'<span class="new">{label}</span>'
                elif t['type'] != 'X':
                    label = f'{label} ({t["type"]})'

                tech_html += f'  <li class="technique">{label}</li>\n'
                                    
        tech_html += """</ul>"""
        tech_tables.append(tech_html)

    return "\n".join(tech_tables)


def nbsp(text):
    "Force any spaces to non-breaking spaces"
    return text.replace(' ', '&nbsp;')

def fix_text(text):
    "Fix any html-special codes and/or things excel put in"
    # special character table
    special = {'&': '&amp;',
               '<': '&lt;',
               '>': '&gt;',
               '’': "'",
               '"': '"',
               '“': '"'}
    for q, r in special.items():
        text = text.replace(q, r)

    return text



def read_inventory(invfile, rank):        
    # read the inventory and return the data for the given rank
    revision = "No Revision"
    tables = []                
    with open(invfile, newline='') as csvfile:
        rdr = csv.DictReader(csvfile)
        for row in rdr:
            # make sure that leading/trailing whitespace is stripped...
            row = {k: fix_text(v.strip()) for k, v in row.items()}
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