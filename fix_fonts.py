# fix_fonts.py — comprehensive font replacement
import re

FILE = "mkpos.py"

with open(FILE, "r", encoding="utf-8") as f:
    content = f.read()

original = content
arial_before = len(re.findall(r'Arial', content))

# 1. family="Arial" (with any spacing)
content = re.sub(r'family\s*=\s*"Arial"', 'family=MYANMAR_FONT', content)
content = re.sub(r"family\s*=\s*'Arial'", "family=MYANMAR_FONT", content)

# 2. Tuples: ("Arial", N, "bold")
content = re.sub(
    r'\(\s*"Arial"\s*,\s*(\d+)\s*,\s*"bold"\s*\)',
    r'(MYANMAR_FONT, \1, "bold")',
    content
)
content = re.sub(
    r"\(\s*'Arial'\s*,\s*(\d+)\s*,\s*'bold'\s*\)",
    r'(MYANMAR_FONT, \1, "bold")',
    content
)

# 3. Tuples: ("Arial", N, "italic" / "normal")
content = re.sub(
    r'\(\s*"Arial"\s*,\s*(\d+)\s*,\s*"(\w+)"\s*\)',
    r'(MYANMAR_FONT, \1, "\2")',
    content
)

# 4. Tuples: ("Arial", N) — 2 items
content = re.sub(
    r'\(\s*"Arial"\s*,\s*(\d+)\s*\)',
    r'(MYANMAR_FONT, \1)',
    content
)
content = re.sub(
    r"\(\s*'Arial'\s*,\s*(\d+)\s*\)",
    r'(MYANMAR_FONT, \1)',
    content
)

# 5. String font="Arial 11" or "Arial 11 bold"
content = re.sub(
    r'font\s*=\s*"Arial\s+(\d+)(\s+bold)?"',
    lambda m: 'font=(MYANMAR_FONT, ' + m.group(1) + (', "bold")' if m.group(2) else ')'),
    content
)

# 6. Any remaining string "Arial" (fallback)
content = content.replace('"Arial"', 'MYANMAR_FONT')
content = content.replace("'Arial'", 'MYANMAR_FONT')

with open(FILE, "w", encoding="utf-8") as f:
    f.write(content)

arial_after = len(re.findall(r'Arial', content))

print(f"✅ Done!")
print(f"   Arial before: {arial_before}")
print(f"   Arial after : {arial_after}")

if arial_after > 0:
    print(f"\n⚠️  Remaining Arial lines:")
    for i, line in enumerate(content.split("\n"), 1):
        if "Arial" in line:
            print(f"   Line {i}: {line.strip()[:120]}")
else:
    print(f"\n✅ All Arial replaced!")