# Written by Mark Lindberg. Copywrite 2017.
# Mucho thanks to Ravi and Ross.
# License: Modify and use as you wish.

import zipfile
import re
from sys import argv
from docx import Document
from docx.shared import RGBColor, Inches
# import xml.etree.ElementTree as ET

# This matches all tab styles that I found.
# If I'm missing a paragraph, let me know.
# I likely need to fix this.
# pattern = r'(<w:pPr><w:ind w:firstLine="720"/></w:pPr>|<w:r><w:tab/>|<w:r><w:lastRenderedPageBreak/><w:tab/>|<w:r><w:rPr><w:rStyle w:val=\"st\"/></w:rPr><w:tab/>|<w:r><w:rPr><w:rStyle w:val=\"st\"/></w:rPr><w:lastRenderedPageBreak/><w:tab/>|<w:r><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond" w:cs="Garamond"/></w:rPr><w:|<w:r><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond" w:cs="Garamond"/><w:u w:val="single"/></w:rPr>|<w:r w:rsidRPr="008715BA"><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond" w:cs="Garamond"/></w:rPr>|<w:r w:rsidRPr="008715BA"><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond" w:cs="Garamond"/><w:color w:val="000000"/></w:rPr>|<w:r w:rsidRPr="008715BA"><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond" w:cs="Garamond"/><w:u w:val="single"/></w:rPr>|<w:r w:rsidRPr="008715BA"><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond" w:cs="Garamond"/><w:color w:val="000000"/><w:u w:val="single"/></w:rPr>|<w:r w:rsidRPr="008715BA"><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond"/></w:rPr><w:t|<w:pPr><w:tabs>|<w:rFonts w:ascii="Garamond" w:hAnsi="Garamond"/><w:caps w:val="0"/><w:smallCaps w:val="0"/><w:rtl w:val="0"/><w:lang w:val="en-US"/>|<w:rFonts w:ascii="Garamond" w:cs="Garamond" w:hAnsi="Garamond" w:eastAsia="Garamond"/><w:caps w:val="0"/><w:smallCaps w:val="0"/></w:rPr></w:pPr>)(?!</w:r></w:p>)'
pattern = r'(<w:ind w:firstLine="360"/><w:rPr><w:rFonts w:ascii="Garamond" w:cs="Garamond" w:hAnsi="Garamond" w:eastAsia="Garamond"/><w:caps w:val="0"/><w:smallCaps w:val="0"/></w:rPr></w:pPr>)(<w:r><w:rPr>)'
reg    = re.compile(pattern)
inline = '<w:tabs><w:tab w:val="left" w:pos="360"/><w:tab w:val="left" w:pos="720"/><w:tab w:val="left" w:pos="1080"/><w:tab w:val="left" w:pos="1440"/><w:tab w:val="left" w:pos="1800"/><w:tab w:val="left" w:pos="2160"/><w:tab w:val="left" w:pos="2520"/><w:tab w:val="left" w:pos="2880"/><w:tab w:val="left" w:pos="3240"/><w:tab w:val="left" w:pos="3600"/><w:tab w:val="left" w:pos="3960"/><w:tab w:val="left" w:pos="4320"/><w:tab w:val="left" w:pos="4680"/><w:tab w:val="left" w:pos="5040"/><w:tab w:val="left" w:pos="5400"/><w:tab w:val="left" w:pos="5760"/><w:tab w:val="left" w:pos="6120"/><w:tab w:val="left" w:pos="6480"/><w:tab w:val="left" w:pos="6840"/><w:tab w:val="left" w:pos="7200"/><w:tab w:val="left" w:pos="7560"/><w:tab w:val="left" w:pos="7920"/><w:tab w:val="left" w:pos="8280"/><w:tab w:val="left" w:pos="8640"/><w:tab w:val="left" w:pos="9000"/></w:tabs><w:rPr><w:rFonts w:ascii="Garamond" w:cs="Garamond" w:hAnsi="Garamond" w:eastAsia="Garamond"/><w:caps w:val="0"/><w:smallCaps w:val="0"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond"/><w:outline w:val="0"/><w:color w:val="a7a7a7"/><w:vertAlign w:val="superscript"/><w:rtl w:val="0"/><w:lang w:val="en-US"/><w14:textFill><w14:solidFill><w14:srgbClr w14:val="A7A7A7"/></w14:solidFill></w14:textFill></w:rPr><w:t>{}</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Garamond" w:cs="Garamond" w:hAnsi="Garamond" w:eastAsia="Garamond"/><w:vertAlign w:val="superscript"/><w:rtl w:val="0"/></w:rPr><w:tab/><w:t xml:space="preserve"> </w:t></w:r>'
    
nothing = '<w:tab w:val="left" w:pos="720"/><w:tab w:val="left" w:pos="1080"/><w:tab w:val="left" w:pos="1440"/><w:tab w:val="left" w:pos="1800"/><w:tab w:val="left" w:pos="2160"/><w:tab w:val="left" w:pos="2520"/><w:tab w:val="left" w:pos="2880"/><w:tab w:val="left" w:pos="3240"/><w:tab w:val="left" w:pos="3600"/><w:tab w:val="left" w:pos="3960"/><w:tab w:val="left" w:pos="4320"/><w:tab w:val="left" w:pos="4680"/><w:tab w:val="left" w:pos="5040"/><w:tab w:val="left" w:pos="5400"/><w:tab w:val="left" w:pos="5760"/><w:tab w:val="left" w:pos="6120"/><w:tab w:val="left" w:pos="6480"/><w:tab w:val="left" w:pos="6840"/><w:tab w:val="left" w:pos="7200"/><w:tab w:val="left" w:pos="7560"/><w:tab w:val="left" w:pos="7920"/><w:tab w:val="left" w:pos="8280"/><w:tab w:val="left" w:pos="8640"/><w:tab w:val="left" w:pos="9000"/></w:tabs><w:rPr><w:rFonts w:ascii="Garamond" w:cs="Garamond" w:hAnsi="Garamond" w:eastAsia="Garamond"/><w:caps w:val="0"/><w:smallCaps w:val="0"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Garamond" w:hAnsi="Garamond"/><w:outline w:val="0"/><w:color w:val="a7a7a7"/><w:vertAlign w:val="superscript"/><w:rtl w:val="0"/><w:lang w:val="en-US"/><w14:textFill><w14:solidFill><w14:srgbClr w14:val="A7A7A7"/></w14:solidFill></w14:textFill></w:rPr><w:t>{}</w:t></w:r>'

stuff =  '''
<w:r>
    <w:pict>
        <v:shape type="#_x0000_t202" style="position:absolute;margin-left:-52.5pt;margin-top:-4pt;width:50.25pt;height:19.5pt;z-index:251659776;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-width-relative:margin;v-text-anchor:top" stroked="f">
            <v:textbox>
                <w:txbxContent>
                    <w:p>
                        <w:r>
                            <w:rPr>
                                <w:color w:val="AAAAAA" />
                            </w:rPr>
                            <w:t>{}</w:t>
                        </w:r>
                    </w:p>
                </w:txbxContent>
            </v:textbox>
        </v:shape>
    </w:pict>
</w:r>
'''
margin = '''
    <w:r>
        <w:rPr>
            <w:color w:val="AAAAAA" />
            <w:vertAlign w:val="superscript"/>
        </w:rPr>
        <w:t>{} </w:t>
    </w:r>
    '''

namespace = {
    'word': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

def insertParNums(old, newMargin, newInline):
    '''
    This inserts a textbox to the left of the paragraph, containing the
    paragraph number.
    '''
    zin   = zipfile.ZipFile(old, 'r')
    zoutMargin = zipfile.ZipFile(newInline, 'w')
    zoutInline = zipfile.ZipFile(newMargin, 'w')
    # .docx files are actually stored as a zipped up set of xml files.
    for item in zin.infolist():
        buf = zin.read(item.filename)
        # This is the one that contains the text of the document.
        if item.filename == 'word/document.xml':
            text = buf.decode('utf-8')
            # root = ET.fromstring(text)
            # breakpoint()
            print(text[:5000])
            # final = addStuff(text)
            finalMargin  = ''
            finalInline  = ''
            count  = 1
            # x is the text, y is the tab and stuff.
            for line in text.split('</w:p>'):
                line += '</w:p>'
                # print(line)
                lineMargin = lineInline = line
                try:
                    result = re.split(pattern, line, maxsplit=1)
                    print(f"{result=}")
                    x, _, y, z = result
                    print(f"{x=}")
                    print(f"{y=}")
                    print(f"{z=}")
                    print()
                except Exception as e:
                    # print(e)
                    if ('Chapter' in line or
                        'Prologue' in line or
                        'EPILOGUE' in line or
                        'Scene' in line):
                        count = 1
                else:
                    if ('Chapter' in line or
                        'Prologue' in line or
                        'EPILOGUE' in line or
                        'Scene' in line):
                        count = 1
                    else:
                        if '<w:t>' in z:
                            lineMargin = x + margin.format(count) + y + z
                            lineInline = x + inline.format(count) + y + z
                        else:
                            lineMargin = lineInline = line
                        count += 1
                finalMargin += lineMargin
                print(f"{lineInline=}")
                finalInline += lineInline
            # Write out all files, including modified text.
            zoutMargin.writestr(item, finalMargin.encode('utf-8'))
            zoutInline.writestr(item, finalInline.encode('utf-8'))
        else:
            zoutMargin.writestr(item, buf)
            zoutInline.writestr(item, buf)
    zoutMargin.close()
    zoutInline.close()
    zin.close()

def newParNums(oldFile, newFile):
    doc = Document(oldFile)
    count = 1
    gray = RGBColor(0xa7, 0xa7, 0xa7)
    for paragraph in doc.paragraphs:
        if paragraph.text.strip() == '':
            continue
        lower_text = paragraph.text.lower().strip()
        if (lower_text.startswith('prologue') or 
            lower_text.startswith('chapter') or 
            lower_text.startswith('epilogue')):
            count = 1
            continue
        runs = paragraph.runs
        paragraph.clear()
        paragraph.paragraph_format.first_line_indent = 0
        paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(0.25))
        num_run = paragraph.add_run(str(count))
        num_run.font.superscript = True
        num_run.font.color.rgb = gray
        paragraph.add_run('\t') # Indent.
        for run in runs:
            new_run = paragraph.add_run(run.text)
            new_run.style = run.style
            new_run.underline = run.underline
            new_run.italic = run.italic
            new_run.bold = run.bold
            # new_run.font = run.font
        count += 1
    doc.save(newFile)

def viewText(fname):
    '''
    Unused, simply for printing the entire underlying xml file.
    '''
    zin = zipfile.ZipFile(fname, 'r')
    for item in zin.infolist():
        if item.filename == 'word/document.xml':
            text = zin.read(item.filename).decode('utf-8')
            pr = True
            count = 0
            for line in text.split('</w:p>'):
                count+=1
                if ('EPILOGUE' in line):
                    pr = True
                if pr:
                    print()
                    print(line)
                if count > 10:
                    break
    zin.close()

if __name__ == '__main__':
    try:
        oldFile = " ".join(argv[1:])
    except:
        print("Please specify a file to convert.")
        exit()
    print("Converting", oldFile)
    if not oldFile.endswith('.docx'):
        oldFile += '.docx'
    # breakpoint()
    # viewText(oldFile)
    # exit(1)
    marginFile = oldFile.split('.docx')[0] + ' with margin numbers.docx'
    inlineFile = oldFile.split('.docx')[0] + ' with inline numbers.docx'
    # insertParNums(oldFile, marginFile, inlineFile)
    newParNums(oldFile, inlineFile)
