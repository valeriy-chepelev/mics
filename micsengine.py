""""
Main data processing unit for the MICS project
(c) Valeriy Chepelev
LGPL licensing
"""

from natsort import natsort_keygen
from string import digits

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
_cdcOrder = {'Status': 'B',
             'Measurement': 'C',
             'Control': 'D',
             'Setting': 'E',
             'Description': 'A'}

# Namespaces and maps definition
ns = {'61850': 'http://www.iec.ch/61850/2003/SCL',
      'NSD': 'http://www.iec.ch/61850/2016/NSD'}

nsURI = '{%s}' % (ns['61850'])
nsMap = {None: ns['61850']}

ns_dURI = '{%s}' % (ns['NSD'])
ns_dMap = {None: ns['NSD']}


def list_ld(icd):
    """Yields LD's info from icd object.
    Return dictionary {inst, ldName, desc}"""
    lds = [(ld.get('inst'), ld.get('ldName'), ld.get('desc'))
           for ld in icd.findall('.//61850:LDevice', ns)]
    for t in sorted(lds, key=natsort_keygen(key=lambda tup: tup[0].lower())):
        yield {'inst': t[0],
               'ldName': '' if t[1] is None else t[1],
               'desc': '' if t[2] is None else t[2]}


def _ln_name(ln):
    t = ['' if (p := ln.get('prefix')) is None else p,
         ln.get('lnClass'),
         '' if (i := ln.get('inst')) is None else i]
    return ''.join(t)


def list_class_groups(icd, ldinst=''):
    """Yields list of sorted ln class group.
    Return dictionary {class_group, class_desc}, i.e. A - Automatic control.
    class_group is the first letter of class names.
    If 'ldinst' omitted, return data from all LDs."""
    ld_request = './/' if ldinst == '' \
        else f'.//61850:LDevice[@inst="{ldinst}"]/'
    letters = {ln.get('lnClass')[0] for ln in icd.findall(f'{ld_request}61850:LN', ns)}
    letters.add('L')
    for letter in sorted(list(letters)):
        yield {'class_group': letter,
               'class_desc': _lnGroup[letter]}


def list_ln(icd, ldinst='', classgroup=''):
    """Yields LN's common info from icd object.
    Return dictionary {ldInst, name=prefix+class+instance, desc, class, type}
    Sorts by: LD instance, class group, full name.
    If 'ldinst' omitted, return data from all LDs."""
    ld_request = './/' if ldinst == '' \
        else f'.//61850:LDevice[@inst="{ldinst}"]/'
    # find LN0 first
    if classgroup in ['', 'L']:
        lns = [(ln.getparent().get('inst'),  # 0- LD instance
                _ln_name(ln),  # 1- Full LN name
                ln.get('desc'),  # 2- Description
                ln.get('lnClass'),  # 3- Class name
                ln.get('lnType'))  # 4- LN type
               for ln in icd.findall(f'{ld_request}61850:LN0', ns)]
    else:
        lns = list()
    # extend with LNs
    lns.extend([(ln.getparent().get('inst'),
                 _ln_name(ln),
                 ln.get('desc'),
                 ln.get('lnClass'),
                 ln.get('lnType'))
                for ln in icd.findall(f'{ld_request}61850:LN', ns)
                if classgroup == '' or ln.get('lnClass')[0] == classgroup
                ])
    for t in sorted(lns, key=natsort_keygen(key=lambda tup: tup[0] + '/' + tup[3] + '/' + tup[1])):
        yield {'ldInst': t[0],
               'name': t[1],
               'class': t[3],
               'type': t[4],
               'desc': '' if t[2] is None else t[2]}


def _get_nsd_do(nsd, ln_class: str, do_name: str):
    """ Return NSD DataObject (lxml element) referenced by ln_class and some do_name.
    Control's the 'multi' condition for the do_name trailed with digits.
    Return None if ln_class or data object not in NSD.
    nsd should be lxml root."""
    d = None
    ln = nsd.find(f'.//NSD:LNClass[@name="{ln_class}"]', ns)
    if ln is None:
        ln = nsd.find(f'.//NSD:AbstractLNClass[@name="{ln_class}"]', ns)
    # Here ln should be found otherwise return None
    if ln is not None:
        d = ln.find(f'./NSD:DataObject[@name="{do_name}"]', ns)
        if d is None and do_name[-1] in digits:
            # Try to find as 'multy'
            d = ln.find(f'./NSD:DataObject[@name="{do_name.rstrip(digits)}"]', ns)
            if d is not None and 'multi' not in d.get('presCond'):
                d = None  # This do should be 'multi'
        if d is None and 'base' in ln.attrib:
            # Lookup in parents iteratively
            d = _get_nsd_do(nsd, ln.get('base'), do_name)
    return d


# TODO: get list of cdcUsages within ln
# TODO: convert list_do to get dictionary, filtered by cdcUsages

def list_do(icd, ln_type: str, nsd):
    """ Yields data objects of ln,
    return sorted tuple (doName, cdc, conditions, cdc_usage, description).
    icd and nsd should be lxml root."""
    ln_type_obj = icd.find(f'.//61850:LNodeType[@id="{ln_type}"]', ns)
    assert ln_type_obj is not None
    dobs = [(name := dob.get('name'),
             cdc := f.get('cdc') if (f := icd.find(f'.//61850:DOType[@id="{dob.get("type")}"]',
                                                   ns)) is not None else '',
             d.get('presCond')[0] if (d := _get_nsd_do(nsd, ln_type_obj.get('lnClass'), name)) \
                                     is not None else 'E',
             next((key for key, vals in _cdcUsage.items() if cdc in vals), ''),
             desc if (desc := dob.get('desc')) is not None else '')
            for dob in ln_type_obj.findall('61850:DO', ns)]
    for d in sorted(dobs, key=natsort_keygen(key=lambda tup: _cdcOrder[tup[3]] + '/' + tup[1])):
        yield d


def get_associations(data: dict, do_name: str) -> str:
    """Return associations, comma-separated from iec_data,
    according to full data object name.
    do_name should be in format: LDinst/prefPTOC1/Str"""
    d = {val for key, val in data.items() if
         key.startswith(do_name)}
    return '; '.join(sorted(d, key=natsort_keygen()))
