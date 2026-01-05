# 27文字チェックツール（Windows版）

ビジュアルノベル原稿の「1行27文字」ルールをチェックするツールです。

[![test-27check](https://github.com/vitalage/27check-windows/actions/workflows/test.yml/badge.svg)](https://github.com/vitalage/27check-windows/actions/workflows/test.yml)

## 使い方

1. `27check.vbs` をデスクトップなど任意の場所に配置
2. チェックしたい `.txt` または `.md` ファイルを `27check.vbs` にドラッグ&ドロップ
3. 同じフォルダに `*_27check_YYYYMMDD_HHMMSS.md` が自動生成され、開きます

## 結果ファイルの例

### OK（違反なし）
```markdown
# 27文字チェック結果

- 対象: 原稿.txt
- 判定: OK
- ルール: 1行あたり全角27文字まで（文字数 > 27 を違反）
- 総行数: 100
- 違反数: 0

問題ありません。
```

### NG（違反あり）
```markdown
# 27文字チェック結果

- 対象: 原稿.txt
- 判定: NG
- ルール: 1行あたり全角27文字まで（文字数 > 27 を違反）
- 総行数: 100
- 違反数: 3

## 違反一覧

|行|ワード|ワード内行|文字数|超過|抜粋|
|---:|---:|---:|---:|---:|---|
|26|9|2|35|8|長すぎる文章がここに表示されます…|
```

## ルール

- **1行27文字以内**（全角換算）
- 3行で1ワード
- 超過した行は違反としてカウント

## 動作環境

- Windows 10 / 11
- 追加インストール不要（VBScript使用）

## テスト

GitHub Actionsで自動テスト済み（Windows環境）

---

**作成**: Claude Code
**バージョン**: 1.0
**更新日**: 2026-01-05
