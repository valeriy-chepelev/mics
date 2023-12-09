# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import lxml.etree as xml_tree
from micsengine import list_ld, list_ln, _get_nsd_do, list_do, get_associations
from associations_reader import read_data

def test_list_ld():
    icd_tree = xml_tree.parse('ICD-152-KSZ-41_200.icd')
    icd_root = icd_tree.getroot()
    for ld in list_ld(icd_root):
        print(ld)


def test_list_ln():
    icd_tree = xml_tree.parse('ICD-152-KSZ-41_200.icd')
    icd_root = icd_tree.getroot()
    for ln in list_ln(icd_root):
        print(ln)


def test_do_scan():
    nsd = xml_tree.parse('IEC_61850-7-4_2007B.nsd')
    nsd_root = nsd.getroot()
    print(_get_nsd_do(nsd_root, 'PDIS', 'Beh1').get('presCond'))


def test_list_do():
    icd_tree = xml_tree.parse('ICD-152-KSZ-41_200.icd')
    icd_root = icd_tree.getroot()
    nsd_tree = xml_tree.parse('IEC_61850-7-4_2007B.nsd')
    nsd_root = nsd_tree.getroot()
    for ln in list_ln(icd_root):
        print('==============================')
        print(ln)
        for d in list_do(icd_root, ln[3], ln[4], nsd_root):
            print(d)
    pass


def test_ass_reader():
    s = get_associations(read_data('foo'), 'LD0/Seq_MSQI1/SeqV')
    print(s)

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


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
        for do in list_do(icd, ln[3], ln[4], nsd):
            print(f'{do[3]} data object "{do[0]}" ({do[1]}) [{do[2]}]: {do[4]}')


def test_tab_gener():
    icd_tree = xml_tree.parse('ICD-152-KSZ-41_200.icd')
    icd_root = icd_tree.getroot()
    nsd_tree = xml_tree.parse('IEC_61850-7-4_2007B.nsd')
    nsd_root = nsd_tree.getroot()
    iec_data = read_data('associations-152-КСЗ-41_200.txt')
    table_gener(icd_root, nsd_root, iec_data)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    test_tab_gener()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
