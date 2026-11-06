"""Read-only comparison of retained field hashes; never approves or changes mail."""
import argparse
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
import json
from pathlib import Path
import sqlite3


def analyze(row):
    if not row.get('StageMsgPath'):
        return None
    original = Path(row['StageMsgPath']).parent

    def latest(prefix):
        dirs = [p for p in original.parent.glob(prefix + '-*') if (p / 'field-hashes.json').exists()]
        return max(dirs, key=lambda p: p.stat().st_mtime) if dirs else None

    source, target = latest('source-current'), latest('target-current')
    if not source or not target:
        return None

    def difference(left, right, name):
        a, b = left / name, right / name
        if not a.exists() or not b.exists():
            return ['missing-profile']
        x, y = json.loads(a.read_text(encoding='utf-8-sig')), json.loads(b.read_text(encoding='utf-8-sig'))
        return sorted(k for k in x.keys() | y.keys() if x.get(k) != y.get(k))

    return dict(Key=row['Key'][:12],
                RawSourceChanges=difference(original, source, 'field-hashes.json'),
                RawTargetDifferences=difference(original, target, 'field-hashes.json'),
                RelevantSourceChanges=difference(original, source, 'relevant-fields.json'),
                RelevantTargetDifferences=difference(original, target, 'relevant-fields.json'))


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('root', type=Path)
    parser.add_argument('job')
    parser.add_argument('output', type=Path)
    args = parser.parse_args()
    uri = (args.root / 'archive.sqlite').resolve().as_uri() + '?mode=ro'
    with sqlite3.connect(uri, uri=True) as connection:
        rows = [json.loads(r[0]) for r in connection.execute('SELECT payload FROM items WHERE job=? AND state=9', (args.job,))]
    with ThreadPoolExecutor(max_workers=4) as pool:
        results = [r for r in pool.map(analyze, rows) if r]
    args.output.write_text(json.dumps(results, indent=2), encoding='utf-8')
    print('Analyzed', len(results), 'of', len(rows))
    for field in ('RawSourceChanges', 'RawTargetDifferences', 'RelevantSourceChanges', 'RelevantTargetDifferences'):
        print(field, Counter(tuple(r[field]) for r in results).most_common(8))
