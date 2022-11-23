import sys
sys.path.insert(0, '/home/ciecrm/branch_comparator')
sys.path.insert(0, '/home/ciecrm/.local/lib/python3.6/site-packages')

import logging
logging.basicConfig(stream=sys.stderr)

from main import app as application
