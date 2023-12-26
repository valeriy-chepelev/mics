import configparser

_config_name = 'micser.ini'

_default_config = '''
[DEFAULT]
template = MICS_template.docx
open_doc = True
auto_name = True
auto_name_prefix = MICS_
icd_path = 
txt_path =
save_path =
'''

_config = configparser.ConfigParser()
if not _config.read(_config_name):
    _config.read_string(_default_config)
cfg = _config['DEFAULT']
#TODO: check default keys present in loaded config

def save_config():
    with open(_config_name, 'w') as configfile:
        _config.write(configfile)
