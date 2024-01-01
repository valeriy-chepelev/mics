import argparse
import re
from reporter import execute_report
import logging
from micsgui import MicsApplication


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
            execute_report(names['ícd'], names['txt'], names['ied'], forceautoname=True)
