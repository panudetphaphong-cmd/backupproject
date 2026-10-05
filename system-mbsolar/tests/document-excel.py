import zipfile
import xml.etree.ElementTree as ET
with zipfile.ZipFile('tests/document-history-test.xlsx') as book:
    assert book.testzip() is None
    for name in book.namelist():
        ET.fromstring(book.read(name))
    ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    root=ET.fromstring(book.read('xl/worksheets/sheet1.xml'))
    rows=root.findall('s:sheetData/s:row',ns)
    assert len(rows)==3
    cells={c.attrib['r']:c for row in rows for c in row}
    assert cells['G2'].find('s:is/s:t',ns).text=='0012345678901'
    assert cells['H2'].find('s:is/s:t',ns).text=='0812345678'
    assert cells['E2'].attrib['t']=='inlineStr'
    assert cells['E2'].find('s:is/s:t',ns).text=='=1+1'
    assert not root.findall('.//s:f',ns)
    assert cells['O2'].find('s:v',ns).text=='535'
    assert root.find('s:autoFilter',ns).attrib['ref']=='A1:P3'
print('PASS: XLSX ZIP CRCs, XML parts, Thai text, numeric amounts, text identifiers, formula safety and filter range')
