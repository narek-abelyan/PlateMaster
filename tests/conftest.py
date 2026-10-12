from pathlib import Path
import importlib.util
import sys

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="session")
def app_module():
    spec = importlib.util.spec_from_file_location("platemaster_app", ROOT / "PlateMaster1.2.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module
