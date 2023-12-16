""""
Reporter unit for the MICS project
(c) Valeriy Chepelev
"""

from docxtpl import DocxTemplate


def report(template):
    doc = DocxTemplate(template)
    ldt1 = [{'inst': 'LD0',
             'name': 'device0',
             'desc': 'description of device 0'},
            {'inst': 'LD1',
             'name': 'device1',
             'desc': 'description of device 1'},
            {'inst': 'LD2',
             'name': 'device2',
             'desc': 'very-very long poem of long superior description of device 2'}]
    ldt2 = [{'inst': 'LDA',
             'name': 'deviceA',
             'desc': 'description of device A'}]
    context = {'name': 'БФПО-152-КСЗ-41_200',
               'tables': [ldt1, ldt2]}
    doc.render(context)
    doc.save('generated_doc.docx')
