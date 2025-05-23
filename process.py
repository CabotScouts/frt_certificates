from pathlib import Path
from pikepdf import Pdf
from openpyxl import load_workbook

def get_initials(name: str) -> str:
    names = name.split(" ")
    initials = "".join(list(map(lambda n: n[0], names)))
    return initials

def make_if_not_exists(path: Path) -> None:
    if not path.exists():
        path.mkdir()

inputPdf: Path = Path.cwd() / "certificates.pdf"
inputList: Path = Path.cwd() / "attended.xlsx"
outputDir: Path = Path.cwd() / "output"

if not inputPdf.exists():
    exit("Missing input file (mail merged certificate PDF)")

if not inputList.exists():
    exit("Missing input file (spreadsheet of attendees)")

make_if_not_exists(outputDir)

workbook = load_workbook(filename=inputList, data_only=True, read_only=True)
sheet = workbook.active.values

# get columns from headings
c = {}
for i, v in enumerate(next(sheet)):
    c[v] = i

pdf = Pdf.open(inputPdf)
for page in pdf.pages:
    row = next(sheet)
    group = row[c["Group"]]
    number = row[c["MembershipNumber"]]
    name = row[c["Name"]]

    groupDir = outputDir / group.replace(" ", "")
    make_if_not_exists(groupDir)
    file = f"{number}_{get_initials(name)}.pdf"
    dest = groupDir / file

    out = Pdf.new()
    out.pages.append(page)
    out.save(dest)