"""Paste ECharts into an HTML deliverable, in place of its one empty
`<script></script>` pair (the placeholder in assets/page-shell.html).

Usage:
    python scripts/inline_echarts.py report.html            # download 5.6.0
    python scripts/inline_echarts.py report.html echarts.min.js

Without a second argument it downloads the echarts 5.6.0 package from the npm
registry, which needs network once; with one it reads that local
echarts.min.js. It records the version in a comment above the tag and rewrites
the file in place. Needs only Python 3.8+.
"""

import io
import re
import sys
import tarfile
import urllib.request

VERSION = "5.6.0"
URL = f"https://registry.npmjs.org/echarts/-/echarts-{VERSION}.tgz"


def bundle(local):
    if local:
        text = open(local, encoding="utf-8").read()
        found = re.search(r'version\s*=\s*"(\d+\.\d+\.\d+)"', text)
        return text, found.group(1) if found else "unknown"
    with urllib.request.urlopen(URL, timeout=120) as response:
        archive = tarfile.open(fileobj=io.BytesIO(response.read()))
    return archive.extractfile("package/dist/echarts.min.js").read().decode("utf-8"), VERSION


def main(page, local=None):
    html = open(page, encoding="utf-8").read()
    placeholder = re.compile(r"<script>\s*</script>")
    if len(placeholder.findall(html)) != 1:
        sys.exit("expected exactly one empty <script></script> pair to replace")
    text, version = bundle(local)
    if "</script" in text.lower():
        sys.exit("the bundle contains </script and cannot be inlined as is")
    tag = f"<!-- echarts {version}, inlined -->\n<script>{text}</script>"
    html = placeholder.sub(lambda _: tag, html, count=1)
    open(page, "w", encoding="utf-8").write(html)
    print(f"inlined echarts {version} ({len(text):,} characters) into {page}")


if __name__ == "__main__":
    main(*sys.argv[1:3])
