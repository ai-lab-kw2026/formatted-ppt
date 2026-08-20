#!/bin/bash
# Claude Code on the web 用のセットアップ。
# コンテナは破棄のたびに初期化されるため、起動時に毎回ここで環境を整える。
# 何度走っても安全（導入済みなら各ステップを飛ばす）。
set -euo pipefail

# ローカル環境（利用者のPC）では何もしない。web セッションのみ対象。
if [ "${CLAUDE_CODE_REMOTE:-}" != "true" ]; then
  exit 0
fi

PROJECT_DIR="${CLAUDE_PROJECT_DIR:-$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)}"

# --- 1. Python ライブラリ ---------------------------------------------------
# python-pptx / pillow は必須、pypdf は PDF 取り込み用。
missing_py=()
python3 -c "import pptx"  2>/dev/null || missing_py+=("python-pptx")
python3 -c "import PIL"   2>/dev/null || missing_py+=("pillow")
python3 -c "import pypdf" 2>/dev/null || missing_py+=("pypdf")

if [ ${#missing_py[@]} -gt 0 ]; then
  echo "Python ライブラリを導入: ${missing_py[*]}"
  python3 -m pip install --quiet "${missing_py[@]}" \
    || echo "警告: Python ライブラリの導入に失敗しました" >&2
fi

# --- 2. 目視QA用のツール ----------------------------------------------------
# LibreOffice Impress で pptx→PDF、poppler の pdftoppm で PDF→PNG。
# これが無いと assets/render.py が動かず、目視QAを飛ばすことになる。
need_apt=()
if ! ls /usr/lib/libreoffice/program 2>/dev/null | grep -qi impress; then
  need_apt+=("libreoffice-impress")
fi
command -v pdftoppm >/dev/null 2>&1 || need_apt+=("poppler-utils")

if [ ${#need_apt[@]} -gt 0 ]; then
  echo "目視QA用ツールを導入: ${need_apt[*]}"
  # 索引が古いと 404 で取得に失敗するため、先に update する
  apt-get update -qq >/dev/null 2>&1 || true
  DEBIAN_FRONTEND=noninteractive apt-get install -y -qq "${need_apt[@]}" >/dev/null 2>&1 \
    || echo "警告: 目視QA用ツールの導入に失敗。render.py が使えない可能性があります" >&2
fi

# --- 3. スキルを配置 --------------------------------------------------------
# リポジトリを直しただけでは ~/.claude/skills/ の実体は変わらないので毎回同期する。
SRC="$PROJECT_DIR/skill/sumitomo-pptx"
DEST="$HOME/.claude/skills/sumitomo-pptx"
if [ -d "$SRC" ]; then
  mkdir -p "$HOME/.claude/skills"
  rm -rf "$DEST"
  cp -r "$SRC" "$DEST"
  # ビルド副産物は持ち込まない
  find "$DEST" -name '__pycache__' -type d -prune -exec rm -rf {} + 2>/dev/null || true
  echo "スキルを配置: $DEST"
fi

echo "セットアップ完了"
