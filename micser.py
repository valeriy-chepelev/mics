import argparse
import re
from reporter import execute_report, auto_ied_name
import logging
from micsgui import MicsApplication


def process_args():
    def chk_icd(val):
        return val if val is not None and re.match(r'^.+?\.((ICD)|(icd))$', val) else None

    def chk_txt(val):
        return val if val is not None and re.match(r'^.+?\.((TXT)|(txt))$', val) else None

    def chk_ied(val):
        return val if val is not None and re.match(r'^[\w,\-_А-Яа-я]{5,40}$', val) else None

    names = dict(icd=None, txt=None, ied=None)
    names['icd'] = next((r for p in (args.ICD, args.TXT, args.IED) if (r := chk_icd(p)) is not None), None)
    names['txt'] = next((r for p in (args.ICD, args.TXT, args.IED) if (r := chk_txt(p)) is not None), None)
    names['ied'] = next((r for p in (args.ICD, args.TXT, args.IED) if (r := chk_ied(p)) is not None), None)
    return names


if __name__ == '__main__':
    # Logging configuration
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s')
    # parser
    parser = argparse.ArgumentParser(description='MICSer v.1.1 - IEC-61850 MICS generator by VCh.',
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
        params = process_args()
        if params['icd'] is None:
            # ICD is not defined
            parser.print_help()
        else:
            # check the IED
            ied_name = auto_ied_name(params['icd'], params['txt']) if params['ied'] is None else params['ied']
            # normal command execution
            logging.info(f'Creating MICS for "{params["icd"]}", please wait...')
            execute_report(params['icd'], params['txt'],
                           ied_name,
                           force_auto_name=True)
