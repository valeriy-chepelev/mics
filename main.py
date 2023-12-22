import argparse
import lxml.etree as xml_tree
from associations_reader import read_data
from reporter import report, get_context
from configuration import cfg, save_config
from tkinter import *
import tkinter.ttk as ttk
import tkinter.scrolledtext as ScrolledText
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


class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid(sticky=N + S + E + W)
        self.create_widgets()

    def create_widgets(self):
        top = self.winfo_toplevel()
        top.rowconfigure(0, minsize=300, weight=1)
        top.columnconfigure(0, minsize=200, weight=1)
        main_pane = Frame(master=self)
        # row 0
        Label(master=main_pane, text='ICD:').grid(column=0, row=0, stick="nsew", padx=(5, 0), pady=5)
        Entry(master=main_pane).grid(column=1, row=0, stick="nsew", pady=5)
        Button(master=main_pane, text='Load', relief='flat', takefocus=0). \
            grid(column=2, row=0, stick="nsew", padx=5, pady=5)
        # row 1
        Label(master=main_pane, text='TXT:').grid(column=0, row=1, stick="nsew", padx=(5, 0), pady=5)
        Entry(master=main_pane).grid(column=1, row=1, stick="nsew", pady=5)
        Button(master=main_pane, text='Load', relief='flat', takefocus=0) \
            .grid(column=2, row=1, stick="nsew", padx=(5), pady=5)
        # row 2
        Label(master=main_pane, text='Name:').grid(column=0, row=2, stick="nsew", padx=(5, 0), pady=5)
        Entry(master=main_pane).grid(column=1, row=2, stick="nsew", pady=5)
        Button(master=main_pane, text='Auto', relief='flat', takefocus=0). \
            grid(column=2, row=2, stick="nsew", padx=(5), pady=5)
        # row 3
        Button(master=main_pane, text='Go MICSer!', relief='flat', command=self.on_mics). \
            grid(column=0, row=3, stick="nsew", padx=10, pady=10, columnspan=3)
        # row 4
        st = ScrolledText.ScrolledText(master=main_pane, state='disabled')
        st.grid(column=0, row=4, stick="nsew", padx=5, pady=5, columnspan=3)

        main_pane.columnconfigure(1, weight=1)
        main_pane.rowconfigure(4, weight=1)
        main_pane.grid(column=0, row=0, sticky='nesw')
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        # Create textLogger
        text_handler = TextHandler(st)
        # Logging configuration
        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')
        # Add the handler to logger
        logger = logging.getLogger()
        logger.addHandler(text_handler)

    def start_app(self):
        self.master.title('MICSer tool')
        self.master.minsize(300, 200)
        self.mainloop()

    def on_mics(self):
        logging.info('Foo')


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
    if args.debug:
        # test_report()
        pass
    app = Application()
    app.start_app()
