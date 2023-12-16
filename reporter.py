""""
Reporter unit for the MICS project
(c) Valeriy Chepelev
"""

from docxtpl import DocxTemplate


def report(template):
    doc = DocxTemplate(template)
    ldt = [{'inst': 'LD0',
            'name': 'device0',
            'desc': 'description of device 0'},
           {'inst': 'LD1',
            'name': 'device1',
            'desc': 'description of device 1'},
           {'inst': 'LD2',
            'name': 'device2',
            'desc': 'very-very long poem of long superior description of device 2'}]
    context = {'name': 'БФПО-152-КСЗ-41_200',
               'ld_tbl': ldt}
    doc.render(context)
    doc.save('generated_doc.docx')
