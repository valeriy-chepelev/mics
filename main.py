# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import lxml.etree as xml_tree
from micsengine import list_ld, list_ln, _get_nsd_do


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

def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    test_do_scan()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
