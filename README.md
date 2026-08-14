# Network Design Diff

SW、AP、ルータ、Firewallなど、機器種別を問わず過去版と今回版の詳細設計書Excelを比較するVBAツールです。必要に応じて、過去Configと今回Configのテキスト差分も確認できます。

通常利用では次の2ファイルだけを使用します。

- [`UniversalDesignDiff.bas`](./UniversalDesignDiff.bas)：Excelへ貼り付けるVBA本体
- [`UNIVERSAL_DIFF_README.md`](./UNIVERSAL_DIFF_README.md)：導入方法、使い方、判定、制限事項

## 主な機能

- Excelの変更、追加、削除、数式変更、移動候補を検出
- 過去版・今回版の色付きコピーを結果ブックへ生成
- 対応表が空の場合は、比較対象シートを左から順番に対応
- 必要な場合だけ`UD_シート対応表`で異なる順番や名称を明示対応
- 表紙、改訂履歴、拠点固有値などの比較ルール
- メーカー非依存のConfig行差分と秘密情報マスク
- Python、PowerShell、外部ライブラリ、外部通信不要

## FortiGate専用試作版

FortiGateの詳細設計書と`config system interface`を照合する旧試作版は、通常利用から分離して[`fortigate-specific`](./fortigate-specific/)へ移動しました。汎用Diffを使用するだけなら、このフォルダは不要です。

## 注意

差分はレビュー候補です。設定の妥当性やメーカー固有の意味までは保証しないため、本番前に匿名化したテストファイルで確認してください。
