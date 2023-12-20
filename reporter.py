""""
Reporter unit for the MICS project
(c) Valeriy Chepelev
"""

from docxtpl import DocxTemplate
from micsengine import list_ld, list_ln, list_do, get_associations, list_class_groups, list_ln_cdc


def get_context(ied_name, icd, nsd, associations):
    ct = {'name': ied_name}
    lds = list()
    for ld in list_ld(icd):
        lds.append(ld)
        groups = list()
        for group in list_class_groups(icd, ld['inst']):
            groups.append(group)
            lns = list()
            for ln in list_ln(icd, ld['inst'], group['class_group']):
                lns.append(ln)
            groups[-1].update({'lns': lns})
        lds[-1].update({'groups': groups})
    lns = list()
    for ln in list_ln(icd):
        lns.append(ln)
        usages = list()
        for usage in list_ln_cdc(icd, ln['lnType']):
            usages.append({'name': usage})
            dobs = list()
            for dob in list_do(icd, ln['lnType'], nsd, usage):
                dobs.append(dob)
                dobs[-1].update({'signal': get_associations(associations,
                                                            "/".join([ln['ldInst'], ln['name'], dob['name']]))})
            usages[-1].update({'dobs': dobs})
        lns[-1].update({'usages': usages})

    ct.update({'ld_table': lds,
               'ln_table': lns})
    return ct

def report(template, context):
    doc = DocxTemplate(template)
    doc.render(context)
    doc.save('generated_doc.docx')
