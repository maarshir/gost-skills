import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "skills" / "gost-report" / "scripts"))
sys.path.insert(0, str(ROOT / "skills" / "gost-bibliography" / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent))
