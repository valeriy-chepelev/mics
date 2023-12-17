import argparse
import lxml.etree as xml_tree
from micsengine import list_ld, list_ln, list_do, get_associations, list_class_groups
from associations_reader import read_data
from reporter import report


def table_gener(icd, nsd, iec_data):
    """Creates all the required mics tables.
    """
    print('=== DEVICES AND NODES ===')
    for ld in list_ld(icd):
        print(f'LOGICAL DEVICE instance "{ld[0]}"')
        for ln in list_ln(icd, ld[0]):
            print(f'{ln[1]} Logical Node "{ln[0]}":{ln[3]} ({ln[2]})')
    print('=== NODES AND DATA ===')
    for ln in list_ln(icd):
        print(f'LOGICAL NODE "{ln[0]}":{ln[3]} ({ln[2]})')
        for do in list_do(icd, ln[4], nsd):
            print(f'{do[3]} data object "{do[0]}" ({do[1]}) [{do[2]}]: "{do[4]}"', end='')
            print(f', associated to "{get_associations(iec_data, "/".join([ln[0], do[0]]))}"')


def test_tab_gener():
    icd_tree = xml_tree.parse('ICD-152-KSZ-41_200.icd')
    icd_root = icd_tree.getroot()
    nsd_tree = xml_tree.parse('IEC_61850-7-4_2007B.nsd')
    nsd_root = nsd_tree.getroot()
    iec_data = read_data('associations-152-КСЗ-41_200.txt')
    table_gener(icd_root, nsd_root, iec_data)


def test_report():
    report('MICS_template.docx')


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
        test_list_ln()
