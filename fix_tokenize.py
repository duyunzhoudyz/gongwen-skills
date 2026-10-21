"""
Definitive tokenize-based fix for build_docx.py.
Restores from git HEAD, then processes all tokens:
- STRING tokens: replace curly quotes in content with \\uXXXX escapes
- ENCODING/COMMENT/other tokens: replace curly quotes with straight quotes
- Special: triple-curly-quote docstring delimiters -> triple-straight-quote
"""
import subprocess, tokenize, io, os

fn = 'skills/gongwen-format/scripts/build_docx.py'
result = subprocess.run(['git', 'show', f'HEAD:{fn.replace(chr(92),"/")}'], capture_output=True)
original = result.stdout.decode('utf-8')

# Characters to fix and their replacements INSIDE string literals (as \\uXXXX escapes)
IN_STRING = {
    '"': '\\u201c',
    '"': '\\u201d',
    '‘': '\\u2018',
    '’': '\\u2019',
    '×': '\\u00d7',
    '—': '\\u2014',
    '、': '\\u3001',
}

# Characters to fix OUTSIDE string literals (replace with closest ASCII/standard equivalent)
OUTSIDE_STRING = {
    '"': '"',
    '"': '"',
    '‘': "'",
    '’': "'",
    '×': '\\u00d7',
    '—': '\\u2014',
    '、': '\\u3001',
}

# First: fix triple-curly-quote docstring delimiters
# These appear as """ (U+201C U+201D U+201D) or similar
content = original
content = content.replace('"""', '"""')  # U+201C U+201D U+201D -> """
content = content.replace('"""', '"""')  # U+201D U+201D U+201D -> """
content = content.replace('"""', '"""')  # U+201C U+201C U+201C -> """
content = content.replace("'''", "'''")  # single-quote variants

# Now tokenize to find string token boundaries
try:
    tokens_list = list(tokenize.generate_tokens(io.StringIO(content).readline))
    token_ok = True
except tokenize.TokenError:
    token_ok = False

if token_ok:
    # Build set of character positions inside string tokens
    lines_list = content.splitlines(keepends=True)

    def rc_offset(lines, row, col):
        off = 0
        for i in range(row-1):
            off += len(lines[i])
        return off + col

    # Collect string spans as flat offsets
    string_spans = []
    for tok in tokens_list:
        if tok.type == tokenize.STRING:
            s = rc_offset(lines_list, tok.start[0], tok.start[1])
            e = rc_offset(lines_list, tok.end[0], tok.end[1])
            string_spans.append((s, e))

    # Build position-to-inside-string lookup (sorted spans)
    def in_string_span(pos):
        for s, e in string_spans:
            if s <= pos < e:
                return True
        return False

    # Process character by character
    result_chars = []
    for i, c in enumerate(content):
        if c in IN_STRING:
            if in_string_span(i):
                result_chars.append(IN_STRING[c])
            else:
                result_chars.append(OUTSIDE_STRING.get(c, c))
        else:
            result_chars.append(c)
    fixed = ''.join(result_chars)
else:
    # Fallback: just replace all curly quotes globally
    fixed = content
    for bad, esc in OUTSIDE_STRING.items():
        fixed = fixed.replace(bad, esc)

open(fn, 'w', encoding='utf-8').write(fixed)

try:
    compile(fixed, fn, 'exec')
    print('build_docx.py: PASS')
except SyntaxError as e:
    lns = fixed.splitlines()
    n = e.lineno
    print(f'build_docx.py: BROKEN:{n}: {e.msg}')
    if n and n <= len(lns):
        print(f'  {repr(lns[n-1][:140])}')
