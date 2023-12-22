import configparser

_config_name = 'mics.ini'

_default_config = '''
[DEFAULT]
open_doc = False
auto_name = True
'''

_config = configparser.ConfigParser()
if not _config.read(_config_name):
    _config.read_string(_default_config)
cfg = _config['DEFAULT']


def save_config():
    with open(_config_name, 'w') as configfile:
        _config.write(configfile)
