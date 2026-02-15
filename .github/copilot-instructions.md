# Copilot Instructions

このプロジェクトはExcel VBAをオブジェクト指向的にかつ日本語で記述するためのライブラリーです。

## 参照スキルガイド (Skills)

特定のタスクを実行する際は、必ず以下の対応するドキュメントを参照し、その指針に従ってください。

- **VBAコーディング標準**
  - VBA のコーディング標準（将来的にはベストプラクティス、命名規則、アンチパターンも追加予定）
  - 📄 `.github/skills/coding-standards/SKILL.md`

## プロジェクト概要

- **言語**: VBA (Visual Basic for Applications)
- **用途**: Excelの機能を日本語クラスで提供
- **ファイル形式**: `.xltm` (Excelマクロ テンプレート)

## ディレクトリ構成

- `src/` - VBAクラスファイル（日本語命名）
  - `セル.cls` - Cellオブジェクトの日本語ラッパー
  - `セル範囲.cls` - Rangeオブジェクトの日本語ラッパー
  - `ワークシート.cls` - Worksheetオブジェクトの日本語ラッパー
  - `ワークブック.cls` - Workbookオブジェクトの日本語ラッパー
  - ほか、Excelオブジェクトモデルの日本語ラッパークラス
