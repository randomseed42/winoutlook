__version__ = '0.1.0'

import sys

if not sys.platform.startswith('win'):
    raise SystemError('This package is only available for Windows.')

from .mailer import Mailer

__all__ = [
    'Mailer',
]
