# PyInstaller spec - build with:  pyinstaller memo_app.spec --noconfirm   (or run build.bat)
# One-folder build: dist/CastellanMemo/CastellanMemo.exe plus an editable resources/ folder next to it.
from PyInstaller.utils.hooks import collect_all, collect_data_files

datas, binaries, hiddenimports = [], [], []
for pkg in ("docx", "pymupdf", "jinja2", "extract_msg"):      # extract_msg ships data files for .msg parsing
    d, b, h = collect_all(pkg)
    datas += d; binaries += b; hiddenimports += h
datas += collect_data_files("dropbox")          # trusted-certs.crt used for TLS
datas += collect_data_files("certifi")
hiddenimports += ["openpyxl.cell._writer", "tkinter", "tkinter.ttk", "tkinter.scrolledtext",
                  "msal", "memo_app.graph_mail", "memo_app.emails"]      # imported lazily

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    excludes=["pytest", "matplotlib", "numpy.tests"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="CastellanMemo",
    console=False,          # set True while debugging to see tracebacks
    icon=None,
)
coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=False, name="CastellanMemo")
