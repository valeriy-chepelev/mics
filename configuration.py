import configparser

_config_name = 'micser.ini'

_default_config = dict(template='MICS_template.docx',
                       open_after_save=True,
                       auto_save=False,
                       auto_save_prefix='MICS-',
                       icd_path='',
                       txt_path='',
                       save_path='',
                       auto_ied=True,
                       auto_ied_from='TXT',
                       nsd='IEC_61850-7-4_2007B.nsd')


def save_config():
    with open(_config_name, 'w') as configfile:
        _config.write(configfile)


_config = configparser.ConfigParser()
_config.read(_config_name)
cfg = _config['DEFAULT']
_update = False
for key, value in _default_config.items():
    if key not in cfg:
        cfg[key] = str(value)
        _update = True
if _update:
    save_config()
