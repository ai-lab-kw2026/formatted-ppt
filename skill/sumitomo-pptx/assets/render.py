#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPTX の全ページを PNG に書き出す（目視QA用）。

環境ごとにコマンドを使い分けなくて済むように、ここで自動判別する。

  1. Windows + PowerPoint         → PowerPoint COM で直接 PNG（最も忠実）
  2. それ以外（Linux/Mac/WSL 等）  → LibreOffice で PDF 化 → ラスタライズ
                                     （pdftoppm / PyMuPDF / pdf2image のどれか）

    python render.py deck.pptx [-o render] [--dpi 80]

成功すると render/slide-01.png … を作り、絶対パスを1行ずつ表示する。
そのパスを実際に開いて目で見ること。表示しただけで確認したことにしない。

どの経路も使えない環境では、何が足りないかと導入方法を示して終了する（終了コード2）。
その場合は「目視確認が未実施」であることを利用者に必ず伝える。
"""

import argparse
import glob
import os
import shutil
import subprocess
import sys
import tempfile


def _fail(msg, code=2):
    raise SystemExit(msg if code != 2 else msg)


# ---------------------------------------------------------------- 経路1: COM

def render_with_com(pptx, out_dir, dpi_w, dpi_h):
    """Windows + PowerPoint。使えなければ None を返す（エラーにしない）。"""
    if not sys.platform.startswith("win"):
        return None
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return None

    app = None
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        pres = app.Presentations.Open(
            os.path.abspath(pptx), True, False, False)
        pres.Export(os.path.abspath(out_dir), "PNG", dpi_w, dpi_h)
        pres.Close()
    except Exception as exc:
        print("  PowerPoint COM は使えなかった (%s)。LibreOffice を試す。" % exc,
              file=sys.stderr)
        return None
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass

    # PowerPoint は スライド1.PNG / Slide1.PNG のように出すので名前を揃える
    made = []
    files = sorted(glob.glob(os.path.join(out_dir, "*.PNG")) +
                   glob.glob(os.path.join(out_dir, "*.png")))
    for i, src in enumerate(files, start=1):
        dst = os.path.join(out_dir, "slide-%02d.png" % i)
        if os.path.abspath(src) != os.path.abspath(dst):
            shutil.move(src, dst)
        made.append(dst)
    return made or None


# --------------------------------------------------- 経路2: LibreOffice + 変換

def _soffice_bin():
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    # Windows の既定インストール先
    for cand in (r"C:\Program Files\LibreOffice\program\soffice.exe",
                 r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(cand):
            return cand
    return None


def pptx_to_pdf(pptx, out_dir):
    """LibreOffice で PDF 化する。失敗したら理由を返す。"""
    soffice = _soffice_bin()
    if not soffice:
        return None, "libreoffice-notfound"

    # サンドボックスでは既定のユーザープロファイルを作れず、
    # 「User installation could not be completed」で何も変換されないことがある。
    # 使い捨てプロファイルを明示すると通る。
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    cmd = [soffice, "--headless",
           "-env:UserInstallation=file://" + profile.replace(os.sep, "/"),
           "--convert-to", "pdf", os.path.abspath(pptx),
           "--outdir", os.path.abspath(out_dir)]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    except subprocess.TimeoutExpired:
        return None, "timeout"
    finally:
        shutil.rmtree(profile, ignore_errors=True)

    pdf = os.path.join(
        out_dir, os.path.splitext(os.path.basename(pptx))[0] + ".pdf")
    if os.path.exists(pdf):
        return pdf, None
    detail = (proc.stderr or proc.stdout or "").strip().splitlines()
    return None, ("変換失敗: " + (detail[-1] if detail else "原因不明"))


def pdf_to_pngs(pdf, out_dir, dpi):
    """PDF を1ページ1枚の PNG にする。使える手段を順に試す。"""
    # (a) poppler の pdftoppm
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-png", "-r", str(dpi), pdf,
                        os.path.join(out_dir, "slide")], check=True)
        return sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))

    # (b) PyMuPDF
    try:
        import fitz  # type: ignore
        doc = fitz.open(pdf)
        made = []
        for i, page in enumerate(doc, start=1):
            path = os.path.join(out_dir, "slide-%02d.png" % i)
            page.get_pixmap(dpi=dpi).save(path)
            made.append(path)
        doc.close()
        return made
    except ImportError:
        pass

    # (c) pdf2image
    try:
        from pdf2image import convert_from_path  # type: ignore
        made = []
        for i, im in enumerate(convert_from_path(pdf, dpi=dpi), start=1):
            path = os.path.join(out_dir, "slide-%02d.png" % i)
            im.save(path)
            made.append(path)
        return made
    except ImportError:
        return None


# ---------------------------------------------------------------- 案内文

HELP_NO_SOFFICE = """エラー: PNG目視QAに必要なツールがありません。

PowerPoint も LibreOffice も見つかりませんでした。どちらかを入れてください。

  Windows : PowerPoint があれば  python -m pip install pywin32
            無ければ https://ja.libreoffice.org/ から LibreOffice を導入
  Ubuntu/Debian : sudo apt-get install -y libreoffice-impress poppler-utils
  macOS         : brew install --cask libreoffice && brew install poppler

導入できない場合は、目視QAを飛ばしたうえで
「目視確認が未実施」であることを利用者に必ず伝えてください（黙って省略しない）。"""

HELP_NO_RASTER = """エラー: PDF までは作れましたが、PNG にするツールがありません。

次のどれかを入れてください。

  Ubuntu/Debian : sudo apt-get install -y poppler-utils
  macOS         : brew install poppler
  どのOSでも     : python -m pip install pymupdf

PDF は残してあるので、PDF を直接開いて確認しても構いません: %s"""


# ---------------------------------------------------------------- CLI

def main():
    ap = argparse.ArgumentParser(
        description="PPTX の全ページを目視QA用の PNG にする")
    ap.add_argument("pptx", help="入力 .pptx")
    ap.add_argument("-o", "--out-dir", default="render",
                    help="出力先 (既定: render)")
    ap.add_argument("--dpi", type=int, default=80,
                    help="解像度 (既定: 80。1枚あたり約1100px幅)")
    args = ap.parse_args()

    if not os.path.exists(args.pptx):
        raise SystemExit("エラー: ファイルが見つかりません → %s" % args.pptx)
    os.makedirs(args.out_dir, exist_ok=True)

    # 経路1: PowerPoint COM
    made = render_with_com(args.pptx, args.out_dir,
                           int(args.dpi * 20), int(args.dpi * 11.25))
    route = "PowerPoint COM"

    # 経路2: LibreOffice
    if not made:
        pdf, err = pptx_to_pdf(args.pptx, args.out_dir)
        if not pdf:
            raise SystemExit(HELP_NO_SOFFICE if err == "libreoffice-notfound"
                             else "エラー: LibreOffice でのPDF化に失敗 (%s)\n\n%s"
                             % (err, HELP_NO_SOFFICE))
        made = pdf_to_pngs(pdf, args.out_dir, args.dpi)
        if not made:
            raise SystemExit(HELP_NO_RASTER % os.path.abspath(pdf))
        route = "LibreOffice + ラスタライズ"

    print("OK: %d枚を書き出しました (%s)" % (len(made), route))
    for path in made:
        print(os.path.abspath(path))
    print("\n※ 上のパスを実際に開いて目で見ること。"
          "はみ出し・重なり・不自然な余白を見つけたら直して再ビルドする。")


if __name__ == "__main__":
    main()
