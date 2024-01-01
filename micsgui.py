""""
GUI TK-based unit for the MICS project
(c) Valeriy Chepelev
LGPL licensing
"""
from configuration import cfg, save_config
from tkinter import *
import tkinter.ttk as ttk
import tkinter.scrolledtext as scrolled_text
import tkinter.filedialog
import logging
import os
from reporter import execute_report


# Icons graphical data
img_data_xml = b'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAACXBIWXMAAAsTAAALEwEAmpwYAAADrklEQVR4nO2aTWgTQRTHp7b4Qf1AxG8sohW12LQ7s61ShNSTXuytB/UiCKIXbypeLOilIIrV2vbNphV68FCVoiJUvHjSi4ifB09SwZbSdOclrSJqu7JJk0y2u2siNNlk9w97yb55b/a3M2/ezoSQQIECBQoUKCcxjkbhLjHBOA4zTbQSfwLA5AVijgJ2EsOo8CcAnrwo4GVfA2Ag5hjHY54CQBZJTd0zm1TAJsZFF+PiZ3oUcPzBNGwm5Q5AFuPilCU5jqkDU9uIbwCAvsYmH7wJDY5XFyJ+0QEofajaJ0XxtH3IqCTlPwXwmsvK0FmIPhQNgJkMKWDcdXnk4iQpVwAM8OG/l0jxi4E4tNh9KRIA/SAD8WS+BnAbBVFFw9qyzAFSIhx2ByE+1XVPrCTlCCAlNTK1j4K4x7iYdQChkXIGkJIaEQrl+MJmFMwqWnwvKXcACRlGBQW8tDAf4A3iCwDzooARS7/eEj8BUPrwiKVf6CsAVENWsH4xDwJQND1U8gAYxy9SEuvPqy3o9SUNoLE3tivLL+DxfGuDkgbAAM9mHl7MMZjZnE/7xt54XUkDoBwfpIc/4Ie8+wTxPaULoMNYwriYlHzezNeFGont9jSA9iGj0tzczGUJU/r0o+lYgBDuMKryziFeAhAaHK9mgI+dbBmIC5nsL343d0VXZ8UCMSL/Zu8Dd3oSgBKZ3MIAX7vZMi6eSR8yLx1ivWeg1zjFUftxh+cAKMniZNTNtm7IWMoAZyQAV5xjiTFzutjFaugV2z0FgHI8TDnGstd20b3AlyZaZRsV9HB2LHHHUh/MqKC3LfADeo1nAFANT5tzWbahHJ/bJTMK4qpk9722y1gm3zfbmG2zfIH4wzhelO3MQxLPAGCW+xTE5/oeXGvrC8QraYSM2NmYbU0fbnEbeqNbSw5AS2RyVWI3N7MCnP9fAIlkW2pTQAW9zZIjqNWX4xSA2Dnr2YFnAOSaBJOnvum3HzUrQvm+UxKUC6WUQj3TGzwFQPpEdVwGKeDHzFvF+8RGlof/5rQMNvbH13sOgFshtF+b3mjZ3z9DbCQ9/Du3QqjpdmydJwGkSmHK8ZFsSzmekH04nejkWgqbidKzAFIfQxTELal2H5B8jBIHmW1yOf62+/8AWSyxQgXKQ+YI8TWAlkRN4WMAIfOT288ADlz/usLXAMJ3jeW+BlBn7iv4GUC4w6jyNYD2xKarjwEQw6goGgDm0SsAsFhiHni7wQjgRZwCgQIFChSIlJf+AlufylTENow2AAAAAElFTkSuQmCC'
img_data_txt = b'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAACXBIWXMAAAsTAAALEwEAmpwYAAAC8ElEQVR4nO2Yz2tTQRDHV/x1UFRUtCDeKmiVttlJakXBnwc9evDgTfDmQRGrR70WBKVa2sy8UkWKSEE8qMf+C6KCKP4BSqHJzEuq+Ktd2byXmsQkKpLk9b39whLeY3d297OT2XmjlJOTk5OT018JSEz7Gs8CyRPw+IhKJgAJGvKiRhlWxqxIJgAKmka5nmgAgLwIJGcjBUC1SJnR+a40SgaIR4D4y5IXkHwGTwZU3AFUCojP1wTHD+nJ3E6VGACY31gnHrzoffBxXTvm7ziAVFbS9YMiPz8zbVaq+P8F5GaTm2G4HWvoGAAbDDVKoen1SHxOxRUAoDz+8xXJ3wD5aKvX0iEA+UOA/DTMAZp5wVzKk+5YxgCrDHE/ID/SyD+aeMKbntHZ9SqOAMrqH/d3Acl9IF5oAMJTcQZQVsor7AGS6XrpcoZye1XcAZSliU8DyafqeCC3VFIAWGmUCzXreqmSBCAzOt9Vsy5p2WTQzs/f/2gOQKsEETjdaHsA8lcgyddrpZw+yOjCdzV3OQqX3gc1wEIjO+EckQUw1WisJj5m+6TRPxXYstWepc3P2HepLB+3z7ZvwzUgT0UYgMykUU7Y1nePNwUb9wfts0YeKm0O+V3PtFkz6MlmjZzTxN/TE7l9gGa1TWfDPkN2TCrrH7A2BkbmNpTt2jmiC4Aqmlc4WBqD8qqOp1wrwcH8JY18J+jHl3+3w2+bFUWWLwCSYt/43A576tZTeseK28IYkBgARiM/LNvVKBP17cQYAKBM/rLLY4kCoEn8/V5xO2Bha/eIWWt/w2suGQAA+Upgj8c08dWgn38xIQD4fXjqu4N6nnDmrr/l8A2zCkhexx5AKisnA1v8rCIe3K5MhGIBQIfpbFUdryoVrtroQkWqu1QADcfmbcxYdgCgg80BaJUgAqcbKQ9QEZEDQM4DjPsLkIsBxgVBcreAcdcguTzALP+E4x/l8gByeYBxX4PkPoeNqweQK4iYKN1OTk5OTk4q4voJgFhNqIBQuqsAAAAASUVORK5CYII='
img_data_doc = b'iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAACXBIWXMAAAsTAAALEwEAmpwYAAADr0lEQVR4nO1ZS2gTURR99QeiqAtFEVyp4A9t5qatgqDdCCKISF24U0TxixRcuNEuBKkUBUGw900/UFFpwYqICxVcqLgUNyKIooj4QZN7k9b6ATsyTSZ5ncSYSSaTOHkH3iLvd2/Ou7/3RggNDQ0NDY2iAJKt4Bp9Bsk3waTNoj4J4FRDGjeQO4VlNdQnATLVDOTTdU0AII2D5N01RYCoEJoujS6KIjeBpIsg6UfGCiSPgcnNIuwEqABJ+1zB8UO0L7ZE1A0BGJ+bJx48XTvwcVYQ8qtOQKSbo/mDIt3ZNWRNFeF3Ae4qkBk6g9ChagTYwdBAThZMj5L2iLASAMg3/p0i6RcgtVZalyoREN8ISLfTNUAhK/gaMXlZKGOAjSZJjYA0CJJ+F7CE56sufZ4twkiAg8bLieWG5N4Js89PhCnCTIAD29xTFpFjBb8jZnKlCDsBDgBpB0j+Njke8AVRLwTYMEw+4NLrmagnAlrMkYUuvbiuCNg1ZE0NTC8oUpB9gTEk33dapOfLYvccQ/IRZxyQj7rH7TXqHs0Xv84pV6+yAR4EgeQ3SnRuyzP+SLnYPM4dpzZl/Su/9CoL4EGQIWlAmds1aR+0pqvR20D6vmrImvHXSxByn196BUnAfmXuI3Us9drjyuGulx5AfuiMRUza65dewbkAJlcouXnMPnVln2M59TzGjzvjmzqsaYA8qljAUr/0KgvgRZBlNaTf9lN/0GTI7kNXcwiQdN0Zj/ZQRPH/D77qVQ7AoyBAHlas4IjTbyC/znEB5LeKnINKfLjmt14BEkDt2T9IV+y+xt7kAuXPfTKQYhlfT6dLQO5XiDnkt14lAzwKmvyeRy/tvijGtyukDBqSbynBbmdKDr1w+ppkbLXfepUM8CjIDmaZJy37UQOT80HS2awF8GED6YSaLtf107zsvZ++FPNJrGYJsGEg3cueeGIbSH7g/I72xNYAJlrUdAnIWxTzHxZFoKYJAORTihvYpz+S9v+Y6LCmTKS8bN93A+mM4iLtldIrOAJM2qyseZ/vdCfuA3nmqKnTb71KApQgaP35dzMB6Wehwsf+8ptbF3Ci2A8fNU2ADZD0xL3WLnaUdNmaSxDdEUWi5gkwJJ9zrWX1dPNaCfLJSuvlGaUKimJiK0iOZxrSYO7edFedE+lObKi0Xp4RmCCP0ARIbQGWdgGpY4Clg6DUWcDSaVDqOsD6/wsOj9B1gNR1gFWVOgBqtGkCKgWogdPVFiCr6AIaGhoaGiJc+AML+MrPmpnTOQAAAABJRU5ErkJggg=='
img_data_mixer = b'iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAYAAABw4pVUAAAACXBIWXMAAAsTAAALEwEAmpwYAAAHFklEQVR4nO2dfYhUVRTAj9EHfVP0gVH0ZWZL6s69z1UkCvsmKqR/yuqPqCgJpVCzIMqKovqjSAp3z51dQ//JLLEiypI+SAqLPqQiQejLMDPdOWdmt4819cWd3aWZtzO+md339p735v3gwDLM3bn3nPfuPed+nAuQkZGRkZExTijDVynkH7Rhv3WE9irk19te3HUMSEMhb3evIHZlmEUgDW34lxY2yFMgDW34AfeKYSeikO8BabSt9Q/Xhra5Vo52IB4WrgeJ5PJ0Q0sapJtyIBVleKNrBenxFiydBFJp7yxMV0j7nCvJjI8opL/B9yeAZLThvGtF6XET2gbSmbai7xRtmN0ri+MX5PchCWikpc6VZZzJn9rQVm1ohcbiTJDApOX+Ea03lcI13iA6oJBXiZheaVU3WNcU2qyxcLxrm7SmG2xqizL8mmt7wAxD7Wl2g70aQeGsZ7cfqZHmaMObRpRBmgOuSbMb3N5TOrluu9E/TCNtqDYIrwPXzMz3naoMF10rT0csyvBfYUHhDFOcXF2G9ogIJNPpBlNDQaFG3iFuqiWlbvAHjbRdIf9aWW5Wnk8ECaTNDVbIq8LabMeRSqdGGfrXfgZS0Ib7UmMQQ0+Ethf53EC5n0ASGul714rUUQny3aHtzfMVoue+FNJ7zhVpIpJ88ZrQ9ub5rqq3CrkbJKEM96TnDSlMDWuv3QAReEMegjTQYfacrgy97dwITXpLGmlNwCA3Q1rwVvae4doIFYrtb6TOCumzqi7LFGdBWvAkGcTQ1kbqrA3trixnZy0gPcagd9wbYvhJ541hdZ7dvfvYEW+VhGmTSqau4BNcK1NHY5CeRjZ8VJVB/hYkopBLrhWqxypIj4a1UyPNDbwhb4JE0hEc0h2h7TS0KFDueZCINvSue4XymMQevwhrpzL0QlUZLNwLEknDglV7Z6ktvJ30ViL2ACvkR1wrVI9ROpbvOa7Zrtnr7r0Q0kKHoEjdrn6GVtj3Jwzuz/q/nIitQGkMDBXyd2F11dg/sboc7YK04QkxiA1QQ+uKxdmBMpshTXiyIvV8WH2V4Vuquzl6GSST6OAQ+eGw9tnvVJehJ0Eyg5uRBSjXNC/K0G2h7UNeGTDInSAZu5TpWrF6lOIhXx7WPmXok8oyuS66DCSjDK12rVg9WsHSlLDzMcHts/YzkMyIpc0EybTVO4+u2zDfn6CQX23WTXaONrzQtWL1KEQh9dZqjzWSxsJFdXb7L4S00SEoUm9SvrFn+CGNeGICw0aFfsvleRKkFS9ZBtmU6y6cCWnFkxWpjxSkgXICHqQ1yhSvFbd23gi2f3WuSBN9UJhYRD/xpraID/BabVtpe2fxPEgr2tDjrhWsmxXccRSkFW14frKMQX9Amsl1Fa5zrmTTlHzpWmcZGRkZGRlysPuVkF8SMED7Y0nrZ6faIQ3YXeOuFaqjMUpvh6GzIcloJGXzpbtWpo7MKPxhIicRh1GGPnetRB2x5PJ8IyQRjXyla+XpeGRLIt8SbXitAOX5sQiSgiRh15RjzZeFNDB0Hnye11083244sGL/tp9ppFeGFpBiMgo9Bkki18VenHkMvR4+p6EkMMjrYnogNkCSyOWLt0ZuCKR9Gmlxs3VRSEu0of3RGoR/hiRh0xjF8FQulpPdjvYnap3EditRd1NQ6/4S5AWD+XG5355eKqe1QF5Qa1+UNry+Za6sCGK3Ukb4ZgwEx4xc9+7TFPLXdb5/QCM/GKyT3S8V8UA/D5LAJcv8Q7WhfyI0yJoauRy/qv1d/uJgc06RuuINJBIQgd0UEGXXoANP4mA3FezSaKeHdDss8w9p5pRTlA9KyyzTzjDFyZX/vzxmVHRnyvAzdY8sByLqoTglqrptgSQQtUfTFjhaPJxU0+46z+VLF9SrhzJ0qUb+uPIz+78iTaoc8kamcu9VW8AgZU/qILkPy4P30FgRPF9u36Qo6za9k84C6QSPdkXdZUGdiT2rbNt9VXtS9GmMXZaf6+KrQTp2ISfKRusw93KZf4gd0O3AHnYBZKSD+qAzcR8k4C6qKI3h24nCer9nXdyyq1vXBa7OKB08ehaBQbpAMrmuwsUxGGTAThQGf8sGf+UgsK4x+ieOHFuiXb1Uhj8CyQSTCEdnFB5xB4cNEANTJ312/LLdVK1c6/ba7ajrpZB+B8kow8/FYhBTbvwSiRcp25ySIJV4D2nSfmXo/tHFRRFPvyclL69C/jE+g/CwrG/kYKX9ThzdVGJOWJUvxorxSdTVb4sdnNeW3VgsTbHBYzkKx9IU+9mgNzU+248U8tMgEXs72/gYg6XJepCIh3yTAOX4UlOQjztp2TKqmzfIXlFXGtW9pqGVBA+eLcgJdn2gdQ1Cc0EawRSpLSVIS0EaCnm7c8UYZzIfpGFzo7eeUWivMvxG6hImZ2RkZGRAkvkP714+6WAZJTwAAAAASUVORK5CYII='
img_data_refresh = b'iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAACXBIWXMAAAsTAAALEwEAmpwYAAAAwUlEQVR4nGNgGAWUAKNZ7wuMZ334TyrGaaDxzA/pRrM/pKGJzaOKBYYzPloaz3r/w3jW+19Gs97ZkuNjY1wWmE75ImE868MTmAKjWe+f609/I00VC4xn/mc1nvnhEIZXZ74/rjLxPzsVLPgwE094LqDYApIUEACjFhAEo0GEFeBL5kaz3s+gOJka48qos94fw8io5OYDU2KLGkoymtGsjxYEC0uSy/6Z7xejWDL7Qxp6cU+ZBbPAwVFAjo9HAQMMAAAUnYhG7Mw9ZwAAAABJRU5ErkJggg=='


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


class MicsApplication(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.grid(sticky=N + S + E + W)
        top = self.winfo_toplevel()
        top.rowconfigure(0, minsize=300, weight=1)
        top.columnconfigure(0, minsize=200, weight=1)
        main_pane = Frame(master=self)
        # icons
        self.ico_code = PhotoImage(data=img_data_xml)
        self.ico_txt = PhotoImage(data=img_data_txt)
        self.ico_doc = PhotoImage(data=img_data_doc)
        self.ico_refresh = PhotoImage(data=img_data_refresh)
        self.ico_mixer = PhotoImage(data=img_data_mixer)
        # variables
        self.icd_name = StringVar()
        self.txt_name = StringVar()
        self.ied_name = StringVar()
        # row 0-1
        Label(master=main_pane, text='ICD file:').grid(column=1, row=0, stick="ws", pady=(5, 0))
        Entry(master=main_pane, state='disabled', textvariable=self.icd_name).\
            grid(column=1, row=1, stick="ewn", padx=(0, 5))
        Button(master=main_pane, relief='flat', takefocus=0, command=self.on_open_icd, image=self.ico_code). \
            grid(column=0, row=0, rowspan=2, stick="nsew", padx=(5, 0), pady=(5, 0))
        # row 2-3
        Label(master=main_pane, text='TXT associations file:').grid(column=1, row=2, stick="ws")
        Entry(master=main_pane, state='disabled', textvariable=self.txt_name).\
            grid(column=1, row=3, stick="ewn", padx=(0, 5))
        Button(master=main_pane, relief='flat', takefocus=0, command=self.on_open_txt, image=self.ico_txt).\
            grid(column=0, row=2, rowspan=2, stick="nsew", padx=(5, 0))
        # row 4-5
        Label(master=main_pane, text='IED official name:').grid(column=1, row=4, stick="wn")
        Entry(master=main_pane, textvariable=self.ied_name).\
            grid(column=1, row=5, stick="ew", padx=(0, 5))
        Button(master=main_pane, relief='flat', takefocus=0, image=self.ico_refresh, command=self.on_autoname).\
            grid(column=0, row=4, stick="nsew", rowspan=2, padx=(5, 0))
        # row 6
        self.start_btn = Button(master=main_pane, relief='flat', takefocus=0, command=self.on_mics,
               image=self.ico_doc, state='disabled')
        self.start_btn.grid(column=0, row=6, stick="nsew", padx=(5, 0))
        Label(master=main_pane, text='Run MICSer to create a perfect MICS document.').\
            grid(column=1, row=6, stick="wns")
        # row 7
        st = scrolled_text.ScrolledText(master=main_pane, state='disabled')
        st.grid(column=0, row=7, stick="nsew", padx=5, pady=5, columnspan=2)

        main_pane.columnconfigure(1, weight=1)
        main_pane.rowconfigure(7, weight=1)
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
        self.master.wm_iconphoto(False, self.ico_mixer)
        self.master.minsize(300, 300)
        self.mainloop()

    def on_mics(self):
        execute_report(self.icd_name.get(),
                       self.txt_name.get(),
                       self.ied_name.get())

    def on_open_icd(self):
        select = tkinter.filedialog.askopenfilename(title='Open ICD file',
                                                    defaultextension='icd',
                                                    filetypes=(("ICD file", "*.icd"), ("All Files", "*.*")),
                                                    initialdir=cfg['icd_path'])
        if os.path.isfile(select):
            self.icd_name.set(select)
            self.start_btn.configure(state='normal')
            path, _ = os.path.split(select)
            if path != cfg['icd_path']:
                cfg['icd_path'] = path
                save_config()

    def on_open_txt(self):
        select = tkinter.filedialog.askopenfilename(title='Open associations txt file',
                                                    defaultextension='txt',
                                                    filetypes=(("TXT file", "*.txt"), ("All Files", "*.*")),
                                                    initialdir=cfg['txt_path'])
        if os.path.isfile(select):
            self.txt_name.set(select)
            path, _ = os.path.split(select)
            if path != cfg['txt_path']:
                cfg['txt_path'] = path
                save_config()

    def on_autoname(self):
        _, filename = os.path.split(self.icd_name.get())
        name, _ = os.path.splitext(filename)
        self.ied_name.set('IED' if name == '' else name)
