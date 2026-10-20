"""Run the repository's Python tool regressions as one unittest suite.

The tests live beside the tools they cover and use `<tool>.test.py` names,
which unittest discovery cannot import as modules. Load each file by path
instead. Run through `pnpm test:python`, which supplies the third-party
dependencies (python-docx, lxml, fontTools) with uv.
"""

import importlib.util
from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parent.parent
FILES = [
    *sorted((ROOT / 'packages/legacy-office/tools').glob('*.test.py')),
    ROOT / 'scripts/test_generate_font_reference_metrics.py',
]


def load(path: Path) -> unittest.TestSuite:
    name = path.name.removesuffix('.py').replace('-', '_').replace('.', '_')
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return unittest.defaultTestLoader.loadTestsFromModule(module)


if __name__ == '__main__':
    suite = unittest.TestSuite(load(path) for path in FILES)
    result = unittest.TextTestRunner(verbosity=1).run(suite)
    sys.exit(0 if result.wasSuccessful() else 1)
