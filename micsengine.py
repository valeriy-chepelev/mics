""""
Main data processing unit for the MICS project
(c) Valeriy Chepelev
LGPL licensing
"""

from natsort import natsort_keygen

# Dictionaries definition
_lnGroup = {'A': 'Automatic control',
            'B': 'Unknown',
            'C': 'Supervisory Control',
            'D': 'Distributed energy resources',
            'E': 'Unknown',
            'F': 'Functional blocks',
            'G': 'Generic function references',
            'H': 'Hydro power',
            'I': 'Interfacing and archiving',
            'J': 'Unknown',
            'K': 'Mechanical and non-electrical primary equipment',
            'L': 'System logical nodes',
            'M': 'Metering and measurement',
            'N': 'Unknown',
            'O': 'Restricted',
            'P': 'Protection functions',
            'Q': 'Power quality events detection related',
            'R': 'Protection related functions',
            'S': 'Supervision and monitoring',
            'T': 'Instrument transformer and sensors',
            'U': 'Unknown',
            'V': 'Unknown',
            'W': 'Wind power',
            'X': 'Switchgear',
            'Y': 'Power transformer and related functions',
            'Z': 'Further (power system) equipment'}
_cdcUsage = {'Status': ["SPS", "DPS", "INS", "ENS", "ACT", "ACD", "SEC", "BCR", "HST", "VSS"],
             'Measurement': ["MV", "CMV", "SAV", "WYE", "DEL", "SEQ", "HMV", "HWYE", "HDEL"],
             'Control': ["SPC", "DPC", "INC", "ENC", "BSC", "ISC", "APC", "BAC"],
             'Setting': ["SPG", "ING", "ENG", "ORG", "TSG", "CUG", "VSG", "ASG", "CURVE", "CSG"],
             'Description': ["DPL", "LPL", "CSD"]}
_cdcOrder = {'Status': 20,
             'Measurement': 30,
             'Control': 40,
             'Setting': 50,
             'Description': 10}

# Namespaces and maps definition
ns = {'61850': 'http://www.iec.ch/61850/2003/SCL',
      'NSD': 'http://www.iec.ch/61850/2016/NSD'}

nsURI = '{%s}' % (ns['61850'])
nsMap = {None: ns['61850']}

ns_dURI = '{%s}' % (ns['NSD'])
ns_dMap = {None: ns['NSD']}


def list_ld(icd):
    """Yields LD's info from icd object.
    Return tuple (inst, ldName, desc)"""
    lds = [(ld.get('inst'), ld.get('ldName'), ld.get('desc'))
           for ld in icd.findall('.//61850:LDevice', ns)]
    for t in sorted(lds, key=natsort_keygen(key=lambda tup: tup[0].lower())):
        yield t[0], '' if t[1] is None else t[1], '' if t[2] is None else t[2]


def _ln_name(ln):
    t = ['' if (p := ln.get('prefix')) is None else p,
         ln.get('lnClass'),
         '' if (i := ln.get('inst')) is None else i]
    return ''.join(t)


def list_ln(icd, ldinst=''):
    """Yields LN's common info from icd object.
    Return tuple (prefix+class+instance, desc)
    Sorts by LN Class, leading system LN's.
    If 'ldinst' omitted, return data from all LDs, primarily sorted."""
    ld_request = './/' if ldinst == '' \
        else f'.//61850:LDevice[@inst="{ldinst}"]/'
    # find LN0 first
    lns = [(ln.getparent().get('inst'),
            _ln_name(ln),
            ln.get('desc'),
            ln.get('lnClass')[0])
           for ln in icd.findall(f'{ld_request}61850:LN0', ns)]
    # extend with LNs
    lns.extend([(ln.getparent().get('inst'),
                 _ln_name(ln),
                 ln.get('desc'),
                 ln.get('lnClass')[0])
                for ln in icd.findall(f'{ld_request}61850:LN', ns)])
    for t in lns:  # TODO: sort - sorted(list, key=lambda x: (x[0], -x[1]))
        yield t  # TODO: out formatting
