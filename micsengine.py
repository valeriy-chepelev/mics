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
    lds = [{'inst': ld.get('inst'),
            'ldName': '' if ld.get('ldName') is None else ld.get('ldName'),
            'desc': '' if ld.get('desc') is None else ld.get('desc')}
           for ld in icd.findall('.//61850:LDevice', ns)]
    for t in sorted(lds, key=natsort_keygen(key=lambda x: x['inst'])):
        yield t


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
    letters.add('L')  # system LLN0 (L) is mandatory in any LD - no check
    for letter in sorted(list(letters)):
        yield {'class_group': letter,
               'class_desc': _lnGroup[letter]}


def _ln_dict(ln):
    return dict(ldInst=ln.getparent().get('inst'),
                name=_ln_name(ln),
                desc='' if ln.get('desc') is None else ln.get('desc'),
                lnClass=ln.get('lnClass'),
                lnType=ln.get('lnType'))


def list_ln(icd, ldinst='', classgroup=''):
    """Yields LN's common info from icd object.
    Return dictionary {ldInst, name=prefix+class+instance, desc, lnClass, lnType}
    Sorts by: LD instance, class group, full name.
    If 'ldinst' omitted, return data from all LDs."""
    ld_request = './/' if ldinst == '' \
        else f'.//61850:LDevice[@inst="{ldinst}"]/'
    # find LN0 first
    if classgroup in ['', 'L']:
        lns = [_ln_dict(ln) for ln in icd.findall(f'{ld_request}61850:LN0', ns)]
    else:
        lns = list()
    # extend with LNs
    lns.extend([_ln_dict(ln) for ln in icd.findall(f'{ld_request}61850:LN', ns)
                if classgroup == '' or ln.get('lnClass')[0] == classgroup])
    for t in sorted(lns, key=natsort_keygen(key=lambda x: x['ldInst'] + '/' + x['lnClass'] + '/' + x['name'])):
        yield t


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


def list_ln_cdc(icd, ln_type: str):
    """ Yields sorted list of ln's cdc usages.
    icd should be a lxml root"""
    ln_type_obj = icd.find(f'.//61850:LNodeType[@id="{ln_type}"]', ns)
    assert ln_type_obj is not None
    cdc_set = {f.get('cdc') for dob in ln_type_obj.findall('61850:DO', ns)
               if (f := icd.find(f'.//61850:DOType[@id="{dob.get("type")}"]', ns)) is not None}
    usages = {next((key for key, values in _cdcUsage.items() if cdc in values), '')
              for cdc in cdc_set}
    for d in sorted(list(usages), key=lambda x: _cdcOrder[x]):
        yield d


def list_do(icd, ln_type: str, nsd, usage='Status'):
    """ Yields data objects of ln,
    return sorted list of dict {name, cdc, presCond, desc}.
    icd and nsd should be lxml root."""
    ln_type_obj = icd.find(f'.//61850:LNodeType[@id="{ln_type}"]', ns)
    assert ln_type_obj is not None
    cdc = ''  # compiler warning suppress
    dobs = [{'name': (name := dob.get('name')),
             'cdc': cdc,
             'presCond': ('E' if (d := _get_nsd_do(nsd, ln_type_obj.get('lnClass'), name)) is None
             else d.get('presCond')[0]).replace('A', 'C'),  # DONE: Change presCond A to C
             'desc': '' if dob.get('desc') is None else dob.get('desc')}
            for dob in ln_type_obj.findall('61850:DO', ns)
            if ((f := icd.find(f'.//61850:DOType[@id="{dob.get("type")}"]', ns)) is not None)
            and (cdc := f.get('cdc')) in _cdcUsage[usage]
            ]
    for d in sorted(dobs, key=natsort_keygen(key=lambda x: x['name'])):
        yield d


def get_associations(data: dict, do_name: str) -> str:
    """Return associations, comma-separated from iec_data,
    according to full data object name.
    do_name should be in format: LDinst/prefPTOC1/Str"""
    d = {val for key, val in data.items() if
         set(do_name.split('/')).issubset(key.split('/'))}
    return '; '.join(sorted(d, key=natsort_keygen()))
