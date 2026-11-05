"""Exercise the production native JSON writer without opening Outlook."""
import json
from pathlib import Path
import subprocess
import sys

executable = Path(sys.argv[1]).resolve()
result = subprocess.run([str(executable), "--self-test-json"], check=True, capture_output=True)
payload = json.loads(result.stdout)
expected = "".join(chr(i) for i in range(32)) + '"\\utf8-çãé'
assert set(payload) == {"Mutable", "Constant", "Temporary"}
assert all(value == expected for value in payload.values())
assert b"\t" not in result.stdout
print("JSON nativo: 32 caracteres de controle, aspas, barra invertida e UTF-8 aprovados.")
