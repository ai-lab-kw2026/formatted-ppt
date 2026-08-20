# 導入指示書（Claude Code 用）

このファイルは、**Claude Code に読ませて実行させるための手順書**です。
人が手でコマンドを打つ必要はありません。

利用者は Claude Code に次のように頼むだけです。

> このフォルダの INSTALL.md を読んで、書いてあるとおりにセットアップして。終わったら結果を報告して。

---

## Claude への指示

以下を **上から順に** 実行してください。各ステップで結果を確認し、失敗したら止まって
利用者に何が起きたかを日本語で伝えてください。勝手に別の方法を試さないこと。

環境は Windows + PowerShell を前提とします。

### STEP 1: 配置先を用意する

```powershell
New-Item -ItemType Directory -Force "$env:USERPROFILE\.claude\skills" | Out-Null
```

### STEP 2: スキルフォルダをコピーする

このファイルと同じ場所にある `skill\sumitomo-pptx` を、上のフォルダへコピーします。

```powershell
Copy-Item ".\skill\sumitomo-pptx" "$env:USERPROFILE\.claude\skills\" -Recurse -Force
```

すでに同名フォルダがある場合は**上書きしてよい**が、上書きした事実を報告に含めること。

### STEP 3: Python を確認する

```powershell
python --version
```

- 3.10 以上が表示された → STEP 4 へ
- 表示されない／3.9 以下 → **ここで止まる**。利用者に次を伝える：
  「Python 3.10 以上が必要です。python.org または Microsoft Store から導入し、
  インストーラでは *Add python.exe to PATH* にチェックを入れてください。
  導入後、もう一度このセットアップを頼んでください。」
  ※ Claude が勝手に Python をインストールしないこと。

### STEP 4: 依存ライブラリを入れる

```powershell
python -m pip install python-pptx pillow
```

`pip` ではなく `python -m pip` を使うこと（複数 Python 環境での誤インストールを避けるため）。

### STEP 5: 検証する

```powershell
python -c "import pptx, PIL; print('libs OK', pptx.__version__, PIL.__version__)"
Test-Path "$env:USERPROFILE\.claude\skills\sumitomo-pptx\assets\template.pptx"
Test-Path "$env:USERPROFILE\.claude\skills\sumitomo-pptx\SKILL.md"
Test-Path "$env:USERPROFILE\.claude\skills\sumitomo-pptx\assets\build.py"
```

**4つすべて成功（`libs OK ...` の表示 ＋ `True` が3つ）で合格**。
1つでも欠けていれば未完了として扱い、どれが欠けたかを具体的に報告する。

### STEP 6: 報告する

次の形式で、日本語で簡潔に報告してください。

- 配置先のパス
- Python のバージョン / python-pptx / Pillow のバージョン
- 検証4項目の結果
- 最後に必ずこの一文を添える：
  **「Claude Code を再起動してください。再起動後に『パワポにして』と頼めばスキルが起動します。」**
  （スキルは起動時に読み込まれるため、再起動しないと認識されません）

---

## 動作確認（任意）

再起動後、利用者が短い資料で試せるように、次を案内してください。

> 【内容】を住商パワポでつくって

品質ルール（1枚1メッセージ・言い切りタイトル・図はネイティブ図形・16pt下限・PNG目視QA・
spec.json の保存）は SKILL.md に内蔵されているため、**利用者が毎回指定する必要はありません**。
その資料固有の事情（枚数の上限、読み手、章扉の見せ方など）だけを書き足せば足ります。
