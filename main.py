import argparse
import lxml.etree as xml_tree
from associations_reader import read_data
from reporter import report, get_context
from configuration import cfg, save_config


def test_report():
    icd_tree = xml_tree.parse('ICD-152-KSZ-41_200.icd')
    icd_root = icd_tree.getroot()
    nsd_tree = xml_tree.parse('IEC_61850-7-4_2007B.nsd')
    nsd_root = nsd_tree.getroot()
    associations = read_data('associations-152-КСЗ-41_200.txt')
    report('MICS_template.docx',
           get_context('БФПО-152-КСЗ-41_200', icd_root, nsd_root, associations))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='MICS generator by VCh.')
    parser.add_argument('-d', '--debug', action='store_true', help='execute debugging functions')
    args = parser.parse_args()
    for key in cfg:
        print(f'{key}:{cfg[key]}')
    cfg['str'] = 'Test multiword по-русски'
    save_config()
    if args.debug:
        # test_report()
        pass
