"""List everything an HTML deliverable loads from outside itself.

Usage:
    python scripts/check_offline.py report.html

Prints one line per hit, with its line number: a host, a protocol-relative
`//host` URL, a relative file such as `echarts.min.js` or `images/logo.png`,
`@import`, `fetch(`, `XMLHttpRequest`, and relative files set from JavaScript,
such as an ECharts `image://images/icon.png` symbol or `img.src = "logo.png"`.
The `w3.org` namespace URLs inline SVG carries are not reported, nor is the
bundle inline_echarts.py writes after its `<!-- echarts X, inlined -->` comment.
Inline every hit, or confirm it is a link the reader chooses to follow. The exit code is 1
when anything is printed. Needs only Python 3.8+.
"""

import re
import sys

PATTERN = r"""(?:src|srcset|href|action|poster|data)\s*=\s*["']?\s*(?!data:|#|mailto:|[{\[])[^\s"'>]+|url\(\s*["']?\s*(?!data:|#)[^\s"')]+|@import|fetch\(|XMLHttpRequest|googleapis|https?://(?!www\.w3\.org/)[^\s"'<>)]+|["']//[^\s"'/][^\s"']*|image://(?!https?:|data:)[^\s"']+|setAttribute\(\s*["'](?:src|srcset|href|poster|data)["']\s*,\s*["'](?!data:|#|https?:|//)[^"']+|\.(?:src|href)\s*=\s*["'](?!data:|#|https?:|//)[^"']+"""


def main(path):
    html = open(path, encoding="utf-8").read()
    # Blank out the inlined ECharts bundle, keeping its newlines so line
    # numbers still match: its own code is full of src= and data= text.
    html = re.sub(
        r"(<!-- echarts [\w.]+, inlined -->\s*<script>)(.*?)(</script>)",
        lambda m: m.group(1) + re.sub(r"[^\n]", " ", m.group(2)) + m.group(3),
        html,
        flags=re.IGNORECASE | re.DOTALL,
    )
    hits = 0
    for hit in re.finditer(PATTERN, html, re.IGNORECASE):
        print(html.count("\n", 0, hit.start()) + 1, hit.group(0))
        hits += 1
    # srcset lists several candidates. As HTML parses it, a URL runs to the
    # next whitespace, then its descriptors run to the next comma.
    for attr in re.finditer(r"srcset\s*=\s*([\"'])(.*?)\1", html, re.IGNORECASE):
        value, pos = attr.group(2), 0
        while (m := re.compile(r"[,\s]*(\S+)").match(value, pos)):
            url, pos = m.group(1), m.end()
            if url.endswith(","):
                url = url.rstrip(",")
            else:
                comma = value.find(",", pos)
                pos = len(value) if comma < 0 else comma + 1
            if url and not url.lower().startswith("data:"):
                print(html.count("\n", 0, attr.start()) + 1, "srcset", url)
                hits += 1
    return 1 if hits else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
