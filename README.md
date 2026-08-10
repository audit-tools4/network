# FortiGate Network Design Checker

FortiGateの詳細設計書Excelにある「ネットワーク」シートと、FortiGate configの`config system interface`をExcel VBAだけで照合する試作ツールです。

Python、PowerShell、外部ライブラリ、外部通信は使用しません。

## 会社PCへの導入

1. GitHubで[`FG60_SelfContained.bas`](./FG60_SelfContained.bas)を開く。
2. コードの先頭から末尾までコピーする。
3. 会社PCで空のExcelブックを作成し、`.xlsm`形式で保存する。
4. `Alt + F11`を押す。
5. VBAエディターで「挿入」→「標準モジュール」を選ぶ。
6. コピーしたコードを貼り付けて保存する。
7. `Alt + F8`から`RunAllChecks`を実行する。
8. 詳細設計書Excel、FortiGate configの順に選択する。
9. `FG_比較結果`シートを確認する。

`.bas`ファイルをダウンロードする必要はありません。GitHubのコード表示からコピーして貼り付けられます。

## 初期対象

詳細設計書の「ネットワーク」シートから、次の項目を抽出します。

- インターフェース名
- エイリアス
- VLANタイプ、親インターフェース、VLAN ID
- VDOM
- VRF ID
- ロール
- アドレッシングモード
- IP／ネットマスク
- セカンダリIPの有効・無効
- 管理者アクセスIPv4
- LLDP受信・送信
- ステータス

config側は`config system interface`を解析します。

## 判定

- `一致`
- `一致(未設定)`
- `不一致`
- `CONFIGオブジェクトなし`
- `CONFIG未記載`
- `要確認(省略可能)`

通常の`show`ではFortiGateのデフォルト値が省略される場合があります。可能であれば`show full-configuration`またはバックアップconfigを使用してください。

## 初期版の対象外

- DHCPサーバ
- アドレスオブジェクトの自動作成
- セキュリティモード、認証ポータル
- Explicit Web／FTP Proxy
- アウトバンドシェイピング
- SD-WAN、IPsec、Firewall Policy

## 安全性

- 詳細設計書は読み取り専用で開き、保存せず閉じます。
- configは読み取りのみで、変更しません。
- 外部通信しません。
- VBAコードには会社固有のIPアドレス、装置名、パスワード、PSK、APIトークンを含めないでください。

## 注意

これは初期版です。本番利用前に匿名化した設計書とconfigで動作確認し、意図的な不一致を正しく検出できることを確認してください。
