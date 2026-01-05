# Windows用：27文字チェック（ドラッグ&ドロップ）＋自動検証

## 使い方（配布用）
1. `27check.vbs` を任意の場所（例：デスクトップ）に置きます。
2. チェックしたい `.txt`（または `.md`）を `27check.vbs` にドラッグ&ドロップします。
3. 入力ファイルと同じフォルダに `*_27check_YYYYMMDD_HHMMSS.md` が生成され、自動で開きます。

## 追加：Windows環境が無くても検証する方法（GitHub Actions）
このフォルダをそのままGitHubにpushすると、Windows上で自動テストが走ります。
- `.github/workflows/test.yml` がテスト定義です。
- 成功すると、GitHub上で緑のチェックが付きます（Windows実機で動いた証拠）。

## テストファイル
- `test_ok.txt`：違反なし（判定OK）
- `test_ng.txt`：違反あり（判定NG、2件）
- `test_blankline.txt`：空行入り（挙動確認用）
