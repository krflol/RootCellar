"""RootCellar Python worker placeholder.

This script is a bootstrap target for the process-isolated scripting host.
It currently provides only metadata output for integration smoke tests.
"""

from __future__ import annotations

import json
import platform
import sys


def main() -> int:
    payload = {
        "component": "rootcellar-python-worker",
        "python_version": sys.version.split()[0],
        "platform": platform.platform(),
        "status": "placeholder",
    }
    print(json.dumps(payload, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())