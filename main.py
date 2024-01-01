import argparse
import re
import subprocess, os, platform
import lxml.etree as xml_tree
from associations_reader import read_data
from reporter import report, get_context
from configuration import cfg
import logging
from micsgui import MicsApplication, ask_target_name


def process_args():
    def chk_icd(val):
        return val if val is not None and re.match('^.+?\.((ICD)|(icd))$', val) else None

    def chk_txt(val):
        return val if val is not None and re.match('^.+?\.((TXT)|(txt))$', val) else None

    def chk_ied(val):
        return val if val is not None and re.match('^[\w,\-_А-Яа-я]{5,40}$', val) else None

    names = dict(icd=None, txt=None, ied=None)
    names['icd'] = next((r for p in (args.ICD, args.TXT, args.IED) if (r := chk_icd(p)) is not None), None)
    names['txt'] = next((r for p in (args.ICD, args.TXT, args.IED) if (r := chk_txt(p)) is not None), None)
    names['ied'] = next((r for p in (args.ICD, args.TXT, args.IED) if (r := chk_ied(p)) is not None), None)
    return names


def execute_report(forceautoname=False):
    # TODO: use parameter, redesign
    icd_tree = xml_tree.parse(names['icd'])
    icd_root = icd_tree.getroot()
    nsd_tree = xml_tree.parse('IEC_61850-7-4_2007B.nsd')
    nsd_root = nsd_tree.getroot()

    # prepare associations
    associations = dict()  # Dummy associations
    try:
        if names['txt'] is not None:
            associations = read_data(names["txt"])
    except FileNotFoundError:
        logging.critical(f'File "{names["txt"]}" not found.')
        return
    except AssertionError:
        logging.critical(f'File "{names["txt"]}" is not an associations file.')
        return

    # define target file
    head, tail = os.path.split(names['icd'])
    tgt, _ = os.path.splitext(tail)
    target_name = cfg['auto_name_prefix'] + tgt + '.docx'
    target = head + target_name
    if not (forceautoname or cfg['auto_name']):
        target = ask_target_name(target_name)
        if target == '':
            return

    # call a report
    report(cfg['template'],
           get_context(names['ied'] if names['ied'] is not None else 'IED',
                       icd_root,
                       nsd_root,
                       associations),
           target)

    # open document
    if cfg['open_after_save'] == 'True':
        if platform.system() == 'Darwin':  # macOS
            subprocess.call(('open', target))
        elif platform.system() == 'Windows':  # Windows
            os.startfile(target)
        else:  # linux variants
            subprocess.call(('xdg-open', target))
    logging.info(f'Successfully created MICS "{target}".')


if __name__ == '__main__':
    # Logging configuration
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    # parser
    parser = argparse.ArgumentParser(description='MICSer - IEC-61850 MICS generator by VCh.',
                                     epilog='''Place arguments in any order. No arguments - opens GUI.
                                            See additional options in "micser.ini".''')
    parser.add_argument('ICD', nargs='?', help='ICD filename.icd')
    parser.add_argument('TXT', nargs='?', help='optional associations filename.txt')
    parser.add_argument('IED', nargs='?', help='optional IED name')
    args = parser.parse_args()
    if args.ICD is None and args.TXT is None and args.IED is None:
        # Open GUI
        app = MicsApplication()
        app.start_app()
    else:
        if (names := process_args())['icd'] is None:
            # ICD is not defined
            parser.print_help()
        else:
            # normal command execution
            execute_report(forceautoname=True)
