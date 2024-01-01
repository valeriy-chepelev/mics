""""
Reporter unit for the MICS project
(c) Valeriy Chepelev
"""

from docxtpl import DocxTemplate
from micsengine import list_ld, list_ln, list_do, get_associations, list_class_groups, list_ln_cdc
import logging
import os
import platform
import subprocess
from configuration import cfg, save_config
import lxml.etree as xml_tree
from associations_reader import read_data
import tkinter.filedialog


def get_context(ied_name, icd, nsd, associations):
    ct = {'name': ied_name}
    lds = list()
    for ld in list_ld(icd):
        lds.append(ld)
        groups = list()
        for group in list_class_groups(icd, ld['inst']):
            groups.append(group)
            lns = list()
            for ln in list_ln(icd, ld['inst'], group['class_group']):
                lns.append(ln)
            groups[-1].update({'lns': lns})
        lds[-1].update({'groups': groups})
    lns = list()
    for ln in list_ln(icd):
        lns.append(ln)
        usages = list()
        for usage in list_ln_cdc(icd, ln['lnType']):
            usages.append({'name': usage})
            dobs = list()
            for dob in list_do(icd, ln['lnType'], nsd, usage):
                dobs.append(dob)
                dobs[-1].update({'signal': get_associations(associations,
                                                            "/".join([ln['ldInst'], ln['name'], dob['name']]))})
            usages[-1].update({'dobs': dobs})
        lns[-1].update({'usages': usages})
    ct.update({'ld_table': lds,
               'ln_table': lns})
    return ct


def report(template, context, target):
    doc = DocxTemplate(template)
    doc.render(context)
    doc.save(target)


def execute_report(icd_filename, txt_filename, ied_name, forceautoname=False):
    # load icd
    try:
        icd_tree = xml_tree.parse(icd_filename)
        icd_root = icd_tree.getroot()
    except Exception as E:
        logging.error(E)
        return
    # load nsd
    try:
        nsd_tree = xml_tree.parse(cfg['nsd'])
        nsd_root = nsd_tree.getroot()
    except Exception as E:
        logging.error(E)
        return
    # load associations
    associations = dict()  # Dummy associations
    try:
        if txt_filename is not None: #TODO: check to empty string
            associations = read_data(txt_filename)
    except FileNotFoundError:
        logging.error(f'File "{txt_filename}" not found.')
        return
    except AssertionError:
        logging.error(f'File "{txt_filename}" is not an associations file.')
        return
    # define target file
    head, tail = os.path.split(icd_filename) #TODO: work with path not finished by slash
    tgt, _ = os.path.splitext(tail)
    target_name = cfg['auto_name_prefix'] + tgt + '.docx'
    target = head + target_name
    if not (forceautoname or (cfg['auto_name'] == 'True')):
        target = ask_target_name(target_name)
        if target == '':
            return
    # call a report
    report(cfg['template'],
           get_context(ied_name,
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


def ask_target_name(default_name):
    """Call a standard SaveAs Dialog and return path+name of new docx file.
    Return empty string in case of user cancel."""
    target = tkinter.filedialog.asksaveasfilename(defaultextension='docx',
                                                  filetypes=(("Word file", "*.docx"), ("All Files", "*.*")),
                                                  initialdir=cfg['save_path'],
                                                  initialfile=default_name)
    if target != '':
        path, _ = os.path.split(target)
        if cfg['save_path'] != path:
            cfg['save_path'] = path
            save_config()
    return target
