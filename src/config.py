import configparser


config = configparser.RawConfigParser()
config.read('config.toml')

interface_dict = dict(config.items('INTERFACE'))



