"""Allow ``python -m find_duplicates``."""

import sys

from .cli import main

sys.exit(main())
