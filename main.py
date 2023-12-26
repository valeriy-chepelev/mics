import argparse
import re
import subprocess, os, platform
import tkinter.filedialog

import lxml.etree as xml_tree
from associations_reader import read_data
from reporter import report, get_context
from configuration import cfg, save_config
from tkinter import *
import tkinter.ttk as ttk
import tkinter.scrolledtext as scrolled_text
import logging


class TextHandler(logging.Handler):
    # This class allows you to log to a Tkinter Text or ScrolledText widget
    # Adapted from Moshe Kaplan: https://gist.github.com/moshekaplan/c425f861de7bbf28ef06

    def __init__(self, text):
        # run the regular Handler __init__
        logging.Handler.__init__(self)
        # Store a reference to the Text it will log to
        self.text = text

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text.configure(state='normal')
            self.text.insert(END, msg + '\n')
            self.text.configure(state='disabled')
            # Autoscroll to the bottom
            self.text.yview(END)

        # This is necessary because we can't modify the Text from other threads
        self.text.after(0, append)


def fixed_map(option):
    # Fix for setting text colour for Tkinter 8.6.9
    # From: https://core.tcl.tk/tk/info/509cafafae
    #
    # Returns the style map for 'option' with any styles starting with
    # ('!disabled', '!selected', ...) filtered out.

    # style.map() returns an empty list for missing options, so this
    # should be future-safe.
    return [elm for elm in style.map('Treeview', query_opt=option) if
            elm[:2] != ('!disabled', '!selected')]


style = ttk.Style()
style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))  # fix end


#TODO: Move out GUI to another unit, main is too huge

class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid(sticky=N + S + E + W)
        top = self.winfo_toplevel()
        top.rowconfigure(0, minsize=300, weight=1)
        top.columnconfigure(0, minsize=200, weight=1)
        main_pane = Frame(master=self)
        # row 0
        Label(master=main_pane, text='ICD:').grid(column=0, row=0, stick="nsew", padx=(5, 0), pady=5)
        self.icd_entry = Entry(master=main_pane, state='disabled')
        self.icd_entry.grid(column=1, row=0, stick="ew", pady=5)
        Button(master=main_pane, text='Load', relief='flat', takefocus=0, command=self.on_open_icd). \
            grid(column=2, row=0, stick="nsew", padx=5, pady=5)
        # row 1
        Label(master=main_pane, text='TXT:').grid(column=0, row=1, stick="nsew", padx=(5, 0), pady=5)
        self.txt_entry = Entry(master=main_pane, state='disabled')
        self.txt_entry.grid(column=1, row=1, stick="ew", pady=5)
        Button(master=main_pane, text='Load', relief='flat', takefocus=0, command=self.on_open_txt) \
            .grid(column=2, row=1, stick="nsew", padx=5, pady=5)
        # row 2
        Label(master=main_pane, text='Name:').grid(column=0, row=2, stick="nsew", padx=(5, 0), pady=5)
        self.ied_entry = Entry(master=main_pane)
        self.ied_entry.grid(column=1, row=2, stick="ew", pady=5)
        Button(master=main_pane, text='Auto', relief='flat', takefocus=0). \
            grid(column=2, row=2, stick="nsew", padx=5, pady=5)
        # row 3
        Button(master=main_pane, text='Go MICSer!', relief='flat', command=self.on_mics). \
            grid(column=0, row=3, stick="nsew", padx=10, pady=10, columnspan=3)
        # row 4
        st = scrolled_text.ScrolledText(master=main_pane, state='disabled')
        st.grid(column=0, row=4, stick="nsew", padx=5, pady=5, columnspan=3)

        main_pane.columnconfigure(1, weight=1)
        main_pane.rowconfigure(4, weight=1)
        main_pane.grid(column=0, row=0, sticky='nesw')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        # Create textLogger
        text_handler = TextHandler(st)
        # Add the handler to logger
        logger = logging.getLogger()
        logger.addHandler(text_handler)

    def start_app(self):
        self.master.title('MICSer tool')
        self.master.minsize(300, 200)
        self.mainloop()

    def on_mics(self):
        logging.info('Foo')

    def on_open_icd(self):
        select = tkinter.filedialog.askopenfilename(title='Open ICD file',
                                                    defaultextension='icd',
                                                    filetypes=(("ICD file", "*.icd"), ("All Files", "*.*")),
                                                    initialdir=cfg['icd_path'])
        if os.path.isfile(select):
            path, _ = os.path.split(select)
            self.icd_entry.configure(state='normal')
            self.icd_entry.delete(0, END)
            self.icd_entry.insert(0, select)
            self.icd_entry.configure(state='disabled')
            cfg['icd_path'] = path
            save_config()

    def on_open_txt(self):
        select = tkinter.filedialog.askopenfilename(title='Open associations txt file',
                                                    defaultextension='txt',
                                                    filetypes=(("TXT file", "*.txt"), ("All Files", "*.*")),
                                                    initialdir=cfg['txt_path'])
        if os.path.isfile(select):
            path, _ = os.path.split(select)
            self.txt_entry.configure(state='normal')
            self.txt_entry.delete(0, END)
            self.txt_entry.insert(0, select)
            self.txt_entry.configure(state='disabled')
            cfg['txt_path'] = path
            save_config()


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
        target = tkinter.filedialog.asksaveasfilename(defaultextension='docx',
                                                      filetypes=(("Word file", "*.docx"), ("All Files", "*.*")),
                                                      initialdir=cfg['save_path'],
                                                      initialfile=target_name)
        if target == '':
            return
        path, _ = os.path.split(target)
        cfg['save_path'] = path
        save_config()

    # call a report
    report(cfg['template'],
           get_context(names['ied'] if names['ied'] is not None else 'IED',
                       icd_root,
                       nsd_root,
                       associations),
           target)

    # open document
    if cfg['open_doc'] == 'True':
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
        app = Application()
        app.start_app()
    else:
        if (names := process_args())['icd'] is None:
            # ICD is not defined
            parser.print_help()
        else:
            # normal command execution
            execute_report(forceautoname=True)
