import argparse
import lxml.etree as xml_tree
from micsengine import list_ld, list_ln, list_do, get_associations, list_class_groups, list_ln_cdc
from associations_reader import read_data
from reporter import report, get_context


def table_gener(icd, nsd, iec_data):
    """Creates all the required mics tables.
    """
    for ld in list_ld(icd):
        print(f'LOGICAL DEVICE {ld}')
        for ln_group in list_class_groups(icd, ld['inst']):
            print(f'GROUP {ln_group["class_group"]} - {ln_group["class_desc"]}')
            for ln in list_ln(icd, ld['inst'], ln_group["class_group"]):
                print(f'LN {ln}')
                for cdc in list_ln_cdc(icd, ln['lnType']):
                    print(f'- {cdc} data objects:')
                    for dob in list_do(icd, ln['lnType'], nsd, cdc):
                        full_do_name = "/".join([ln['ldInst'], ln['name'], dob['name']])
                        print(f'  {dob} -> {get_associations(iec_data, full_do_name)}')


def test_tab_gener():
    icd_tree = xml_tree.parse('ICD-152-KSZ-41_200.icd')
    icd_root = icd_tree.getroot()
    nsd_tree = xml_tree.parse('IEC_61850-7-4_2007B.nsd')
    nsd_root = nsd_tree.getroot()
    iec_data = read_data('associations-152-КСЗ-41_200.txt')
    table_gener(icd_root, nsd_root, iec_data)


def test_report():
    icd_tree = xml_tree.parse('ICD-152-KSZ-41_200.icd')
    icd_root = icd_tree.getroot()
    nsd_tree = xml_tree.parse('IEC_61850-7-4_2007B.nsd')
    nsd_root = nsd_tree.getroot()
    associations = read_data('associations-152-КСЗ-41_200.txt')
    report('MICS_template.docx',
           get_context('БФПО-152-КСЗ-41_200', icd_root, nsd_root, associations))


def test_list_ln():
    icd_tree = xml_tree.parse('ICD-152-KSZ-41_200.icd')
    icd_root = icd_tree.getroot()
    for c in list_ln(icd_root, classgroup='L'):
        print(c)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='MICS generator by VCh.')
    parser.add_argument('-d', '--debug', action='store_true', help='execute debugging functions')
    args = parser.parse_args()
    if args.debug:
        test_report()
