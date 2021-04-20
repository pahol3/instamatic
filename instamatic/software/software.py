from instamatic import config

__all__ = ['Software', 'get_software']

def get_software(interface: str):
    """Grab software class with the specific 'interface'."""

    if interface == 'TIA':
        from .TIA_software import TIASoftware as cls
    else:
        raise ValueError(f'No such software interface: `{interface}`')

    return cls

def Software(name: str = None, use_server: bool = False):
    """Generic class to load sofware interface/acquisition class.

    use_server: bool
        Connect to software server running on the host/port defined in the config file

    returns: software interface/acquisition class
    """

    if name is None:
        return None

    if use_server:
        from .software_client import SoftwareClient
        sw = SoftwareClient(name=name)
    else:
        cls = get_software(name)
        sw = cls()
        sw.connect()

    return sw
