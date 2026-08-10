# FortiGate専用試作版（通常利用から分離）

このフォルダは、FortiGate詳細設計書の「ネットワーク」シートとFortiGate Configの`config system interface`を照合する旧試作版です。

- [`FG60_SelfContained.bas`](./FG60_SelfContained.bas)
- 実行マクロ：`RunAllChecks`
- 結果シート：`FG_比較結果`

対象は主にインターフェース設定です。Firewall Policy、SD-WAN、IPsecなどは対象外です。

通常の過去版・今回版比較には、リポジトリ直下の`UniversalDesignDiff.bas`を使用してください。同じ`.xlsm`へ両方のVBAを貼り付ける必要はありません。
