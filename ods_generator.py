import csv
import zipfile
from xml.etree.ElementTree import Element, SubElement, tostring


def load_csv(path):
    with open(path, newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        rows = [row for row in reader]
    row_count = len(rows)
    col_count = max((len(r) for r in rows), default=0)
    return rows, row_count, col_count


def display_preview(rows, row_count, col_count):
    print(f"Rows: {row_count}, Columns: {col_count}")
    try:
        show_rows = int(input("How many rows to display? (0 for none): ") or 0)
    except ValueError:
        show_rows = 0
    try:
        show_cols = int(input("How many columns to display? (0 for none): ") or 0)
    except ValueError:
        show_cols = 0
    if show_rows > 0 and show_cols > 0:
        for r in rows[:show_rows]:
            print('\t'.join(r[:show_cols]))


def write_ods(rows, output_path):
    NS = {
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
        'table': 'urn:oasis:names:tc:opendocument:xmlns:table:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
        'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'fo': 'urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0',
        'manifest': 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0',
    }

    office = Element('office:document-content', {
        'xmlns:office': NS['office'],
        'xmlns:table': NS['table'],
        'xmlns:text': NS['text'],
        'xmlns:style': NS['style'],
        'xmlns:fo': NS['fo'],
        'office:version': '1.2',
    })

    auto_styles = SubElement(office, 'office:automatic-styles')
    style_green = SubElement(auto_styles, 'style:style', {
        'style:name': 'green',
        'style:family': 'table-cell',
    })
    SubElement(style_green, 'style:table-cell-properties', {
        'fo:background-color': '#00ff00'
    })
    style_black = SubElement(auto_styles, 'style:style', {
        'style:name': 'black',
        'style:family': 'table-cell',
    })
    SubElement(style_black, 'style:table-cell-properties', {
        'fo:background-color': '#000000',
        'fo:color': '#ffffff'
    })
    style_red = SubElement(auto_styles, 'style:style', {
        'style:name': 'red',
        'style:family': 'table-cell',
    })
    SubElement(style_red, 'style:table-cell-properties', {
        'fo:background-color': '#ff0000'
    })

    body = SubElement(office, 'office:body')
    spreadsheet = SubElement(body, 'office:spreadsheet')
    table = SubElement(spreadsheet, 'table:table', {'table:name': 'Sheet1'})

    for row in rows:
        tr = SubElement(table, 'table:table-row')
        for cell in row:
            attrib = {}
            if cell == 'Y':
                attrib['table:style-name'] = 'green'
            elif cell == 'NM':
                attrib['table:style-name'] = 'black'
            else:
                try:
                    float(cell)
                    attrib['table:style-name'] = 'red'
                except ValueError:
                    pass
            tc = SubElement(tr, 'table:table-cell', attrib)
            p = SubElement(tc, 'text:p')
            p.text = cell

    content_xml = tostring(office, encoding='utf-8', xml_declaration=True)

    manifest = Element('manifest:manifest', {
        'xmlns:manifest': NS['manifest'],
        'manifest:version': '1.2',
    })
    SubElement(manifest, 'manifest:file-entry', {
        'manifest:full-path': '/',
        'manifest:version': '1.2',
        'manifest:media-type': 'application/vnd.oasis.opendocument.spreadsheet',
    })
    SubElement(manifest, 'manifest:file-entry', {
        'manifest:full-path': 'content.xml',
        'manifest:media-type': 'text/xml',
    })
    manifest_xml = tostring(manifest, encoding='utf-8', xml_declaration=True)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', compress_type=zipfile.ZIP_STORED)
        z.writestr('content.xml', content_xml)
        z.writestr('META-INF/manifest.xml', manifest_xml)


def main():
    import argparse
    parser = argparse.ArgumentParser(description='Generate ODS from CSV with color')
    parser.add_argument('input', help='Input CSV file')
    parser.add_argument('output', help='Output ODS file')
    args = parser.parse_args()

    rows, rc, cc = load_csv(args.input)
    display_preview(rows, rc, cc)
    write_ods(rows, args.output)
    print(f'Written {args.output}')


if __name__ == '__main__':
    main()
