#!/bin/env python3
import argparse
from string import Template
from pathlib import Path
import csv
import os
import sys
import subprocess
from datetime import datetime
import base64

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

qr_code_cache = {}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--full", default=False, action="store_true", help="Don't collapse headers")    
    parser.add_argument("--writerbin", type=str, default="oowriter", help="binary for libreoffice writer")
    parser.add_argument("--evergreen", default=False, action="store_true", help="Don't append date to filenames")
    args = parser.parse_args()

    if args.evergreen:
        outdir = sys.path[0] + "/evergreen_sheets"
    else:
        outdir = sys.path[0] + "/test_sheets"

    # walk the ranks, generating each sheet as we go.    
    docs = []
    for rank in ranks:
        print(f"Generating {ranks[rank]}")

        # read the csv inventory
        data = read_inventory(Path(sys.path[0]) / "inventory.csv", rank)


        for sheet in [('techniques_template.html', gen_tech_content, '-techniques'),
                      ('test_template.html', gen_test_content, '-test')]:
            revision = data.get('revision', datetime.strftime(datetime.now(), '%Y-%m-%d'))
            content = sheet[1](data, args.full)
            qr_codes = "<tr><td>" + "</td><td>".join([f'{k}<br><img src="data:img/png;base64, {str(v, "utf-8")}" alt="{k}">' for k,v in qr_code_cache.items()]) + "</td></tr>"
            filebase = ranks[rank][0] + sheet[2] + ("" if args.evergreen else "-" + revision)
            template = Template((Path(sys.path[0], sheet[0]).read_text()))            
            with open(f"{outdir}/{filebase}.html", "w") as f:
                f.write(template.safe_substitute(title=ranks[rank][1],
                                                 content=content,
                                                 revision=revision,
                                                 qr_codes=qr_codes))
            
            # generate the word doc & pdf
            #subprocess.run(['pandoc', f'{outdir}/{filebase}.html', "-o", f"{outdir}/{filebase}.docx"])
            subprocess.run([args.writerbin, '--convert-to', 'docx', f'{outdir}/{filebase}.html', '--outdir', f'{outdir}'])  
            subprocess.run(['weasyprint', '--pdf-variant', 'pdf/ua-1',
                            f'{outdir}/{filebase}.html', f"{outdir}/{filebase}.pdf"])
            docs.append(f"{outdir}/{filebase}.pdf")

    # Generate special sheets
    for sheet in [('Supplemental Sheet', 'supplemental_template.html', 'supplemental', gen_supplemental, 'S'),
                  ('Skills Matrix', 'matrix_template.html', 'matrix', gen_matrix, 'M')]:
        name, template_file, basename, genfunc, tag = sheet

        # create the supplemental sheet
        print(f"Generating {name}")        
        data = read_inventory(Path(sys.path[0]) / "inventory.csv", tag)
        revision = data.get('revision', datetime.strftime(datetime.now(), '%Y-%m-%d'))
        content = genfunc(data, args.full)
        filebase = basename + ("" if args.evergreen else "-" + revision)
        #filebase = basename + "-" + revision
        template = Template((Path(sys.path[0], template_file).read_text()))
        with open(f"{outdir}/{filebase}.html", "w") as f:
            f.write(template.safe_substitute(content=content,
                                                revision=revision))
        
        # generate the word doc & pdf
        #subprocess.run(['pandoc', f'{outdir}/{filebase}.html', "-o", f"{outdir}/{filebase}.docx"])        
        subprocess.run([args.writerbin, '--convert-to', 'docx', f'{outdir}/{filebase}.html', '--outdir', f'{outdir}'])            
        subprocess.run(['weasyprint', '--pdf-variant', 'pdf/ua-1',
                        f'{outdir}/{filebase}.html', f"{outdir}/{filebase}.pdf"])
        docs.append(f"{outdir}/{filebase}.pdf")

    # Create the everything pdf
    if args.evergreen:
        subprocess.run(['pdfunite', *docs, f"{outdir}/everything.pdf"])
    else:
        subprocess.run(['pdfunite', *docs, f"{outdir}/everything-{revision}.pdf"])
    
    # create everything_doublesided pdf
    # the difference between this and the everything version is that each document will
    # start on on even pages -- so when printing it comes out right
    
    # create a blank page that we can insert as necessary
    blank = f"{outdir}/blank.pdf"
    with open(f"{outdir}/blank.html", "w") as f:
        f.write("""<html lang="en"><head><style>
@page {
    size: letter portrait;
    margin: 0.25in;
}
                </style></head></html>""")
    subprocess.run(['weasyprint', f"{outdir}/blank.html", blank])
    os.unlink(f"{outdir}/blank.html")

    new_docs = []
    for pdf in docs:
        p = subprocess.run(['pdfinfo', pdf],
                           stdout=subprocess.PIPE, encoding='utf-8')
        pages = 1
        for l in p.stdout.splitlines():
            if l.startswith("Pages:"):
                _, pages = l.split(":")
                pages = int(pages)
                break
        else:
            print(f"WARNING:  cannot determine number of pages for {pdf}, assuming odd number")
            
        new_docs.append(pdf)
        if pages % 2 == 1:
            new_docs.append(blank)

    if args.evergreen:
        subprocess.run(['pdfunite', *new_docs, f"{outdir}/everything_doublesided.pdf"])
    else:
        subprocess.run(['pdfunite', *new_docs, f"{outdir}/everything_doublesided-{revision}.pdf"])

    os.unlink(blank)



def gen_test_content(data, full=False):
    """Generate the content for a test sheet"""
    html_tables = []
    for tdata in data['tables']:
        html = """<table width="100%" border="1" class="ttable">\n"""
        
        title = tdata['title']
        if tdata['subtitle']:
            # append the subtitle
            title += f"""<br/><span class="subtitle">{tdata['subtitle']}</span>"""
        html += f"""  <caption>{title}</caption>\n"""
        html += f"""  <colgroup><col class="narrow"/><col class="score"/><col/></colgroup>\n"""
        html += f"""  <thead>\n"""
        html += f"""    <tr><td>&nbsp;</td><th>Score</th><th>Comments</th></tr>\n"""
        html += f"""  </thead>\n"""
        html += f"""  <tbody>\n"""
        for header in tdata['headers']:
            if not header['techniques']:
                # skip empty headers
                continue
            if header['label'] != '':
                hlabel = header['label']
                if header['type'] == 'O':
                    hlabel += " (optional)"
                html += f"""    <tr><td class="theader">{nbsp(hlabel)}</td><td></td><td></td></tr>\n"""
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
                    html += f"    <tr><td>{label}</td><td></td><td></td></tr>\n"
                    header_rows += 1
            if header['type'] in ('C', 'O') and header_rows == 0:
                # add a blank row for comments.
                html += "    <tr><td>&nbsp;</td><td></td><td></td></tr>\n"

        html += f"""  </tbody>\n"""
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


def gen_supplemental(data, full=False):
    style_map = {'Y': 'yellow', 'O': 'orange', 'G': 'green', 'P': 'purple',
                 'b': 'blue', 'B': 'brown', 'R': 'red', 'T': 'temp',
                 '1': 'black', '2': 'black', '3': 'black'}

    def tempize_text(text):
        return ''.join([f'<span class="temp">{x}</span>' for x in text])

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
                if t['type'] == 'T':
                    label = tempize_text(label)
                if t['type'] == '2':
                    label += " (2nd Dan)"
                if t['type'] == '3':
                    label += " (3rd Dan)"

                tech_html += f'  <li class="{style_map[t["type"]]}">{label}</li>\n'
                                    
        tech_html += """</ul>"""
        tech_tables.append(tech_html)

    return "\n".join(tech_tables)


def gen_matrix(data, full=True):

    html_tables = []
    for tdata in data['tables']:
        html = """<table class="ttable">\n"""    
        title = tdata['title']
        html += f"""  <colgroup><col/>""" + ("""<col class="narrow"/>""" * 11) +   """</colgroup>\n"""
        html += f"""  <thead><th>{title}</th><th>Y</th><th>O</th><th>G</th><th>P</th><th>b</th><th>B</th><th>R</th><th>T</th><th>1</th><th>2</th><th>3</th></thead>"""        
        for header in tdata['headers']:
            if not header['techniques']:
                # skip empty headers
                continue
            if header['label'] != '':
                html += f"""  <tr><td class="theader">{nbsp(header['label'])}</td>""" + "<td colspan='11'/>"  + "</tr>\n"
            for t in header['techniques']:
                label = nbsp(t['label'])
                html += f"  <tr><td>{label}</td>" 
                found = False
                for rank in ranks.keys():
                    if rank == t['type']:
                        found = True
                    attr = 'class="graybg"' if rank in ('b', '1') else ''

                    html += f"<td {attr}>{'X' if found else '&nbsp;'}</td>"
                
                html += "</tr>\n"

        html += """</table>"""    
        html_tables.append(html)
    
    return "\n".join(html_tables)


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
               '“': '"',
               '–': '-'}
    for q, r in special.items():
        text = text.replace(q, r)

    return text


def read_inventory(invfile, rank):        
    # read the inventory and return the data for the given rank
    revision = "No Revision"
    qr_codes = {}
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
            if row['Type'] == 'Q':
                title, url = row['Label'].split('=', 1)
                if title not in qr_code_cache:
                    p = subprocess.run(["qrencode", '-t', 'PNG', '-o', '-', url],
                                       stdout=subprocess.PIPE, check=True)
                    qr_code_cache[title] = base64.b64encode(p.stdout)
                
                qr_codes[title] = qr_code_cache[title]
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
            'tables': tables,
            'qr_codes': qr_codes}


if __name__ == "__main__":
    main()