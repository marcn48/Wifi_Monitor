# 引き継ぎメモ（HANDOFF）
作成日: 2026-03-13

---

## プロジェクト概要

Windows用 Wi-Fi接続監視・切断原因診断システム。  
Pythonで動作し、コマンドプロンプトに常時表示しながらログをJSONで保存する。  
ブラウザで開くログビューア（log_viewer.html）で時系列グラフと各イベントを確認できる。

---

## ファイル構成（C:\WiFiMonitor\）

| ファイル | 内容 |
|---------|------|
| `wifi_monitor.py` | 監視プログラム本体 |
| `log_viewer.html` | ブラウザ用ログビューア |
| `README.md` | セットアップ・操作説明 |
| `起動.bat` | 手動起動用（ユーザーが手動作成・GitHubには含めない） |
| `アンインストール.bat` | 完全削除用（ファイル名は任意に変更可） |
| `logs\wifi_log_YYYY-MM-DD.json` | 日付ごとのログ |

---

## ユーザー環境

| 項目 | 内容 |
|------|------|
| OS | Windows 10 (Version 10.0.26200.8037) |
| Python | 3.14.3（`%LocalAppData%\Programs\Python\Python314\python.exe`） |
| SSID | BE001（5GHz / 802.11ac） |
| 拠点 | 大川原事務所（AP×3台）・広野事務所（AP×8台前後）、同じSSID |
| ゲートウェイ | 192.168.1.1 |
| ネットワーク | 会社共有（約80ユーザー）・管理者権限なし |
| speedtest-cli | 2.1.3 インストール済み |

---

## 現在の主要設定値（wifi_monitor.py）

| 定数 | 値 | 意味 |
|------|-----|------|
| `CHECK_INTERVAL` | 5秒 | 監視間隔 |
| `STATUS_LOG_EVERY` | 12回 | 60秒ごとに詳細表示＆ログ保存 |
| `SPEEDTEST_EVERY` | 10回 | 約10分ごとにSpeedtest実行 |
| `LOCATION_EVERY` | 30分 | 位置情報の定期再取得間隔 |
| `SLEEP_DETECT_SEC` | 30秒 | スリープ復帰判定の閾値 |
| `LOG_DIR` | `C:\WiFiMonitor\logs` | ログ保存先 |
| `NOISE_FLOOR_DBM` | -95 dBm | SNR推定用ノイズフロア |

---

## 実装済み機能

### 監視・記録
- 5秒ごとにWi-Fi接続状態を確認
- 60秒ごとに詳細ステータスを表示・ログ保存
- 切断検知時に自動診断（Ping・周辺APスキャン・WLANイベントログ）
- ローミング（BSSID変化）の記録

### 計測・取得データ
- SSID / BSSID / チャネル / 周波数帯 / 認証方式
- 電波強度（signal%）
- RSSI（signal% から逆算: `signal / 2 - 100`）
- SNR 推定値（`RSSI - NOISE_FLOOR_DBM`）
- リンク速度 Rx / Tx
- Speedtest（DL・UL・Ping・サーバー）バックグラウンドスレッド
- アクティブ機器数（`arp -a` の動的エントリ数・参考値）
- 現在地（WindowsのLocation API + OpenStreetMap逆ジオコーディング）

### 位置情報
- 起動時・30分ごと・スリープ復帰時に再取得
- PowerShell経由で `System.Device.Location` を呼び出し緯度経度を取得
- nominatim（OpenStreetMap）で日本語住所に変換
- GoogleマップURLをログに記録（ログビューアから別タブで開ける）

### スリープ復帰対応
- 前回ループからの経過時間が30秒超 → スリープ復帰と判定
- `WAKE` イベントをログに記録（スリープ時間も保存）
- 位置情報を即時再取得（事務所間移動に対応）

### UI・操作
- ×ボタン押下時に確認ポップアップ（WM_CLOSEサブクラス化方式）
  - 「はい」→ 終了 / 「いいえ」→ キャンセルして監視継続
- Ctrl+C で正常終了

### ログビューア（log_viewer.html）
- JSONファイルをドラッグ＆ドロップで読み込み
- Chart.js 時系列グラフ（7データセット・トグル切替）
  - 🟢 電波強度 / 🔵 リンク速度Rx / 🟠 DL実測 / 🟣 UL実測 / 🔴 Ping遅延 / 🟡 アクティブ機器数（デフォルトON）
  - ⚪ リンク速度Tx（デフォルトOFF）
- フィルターボタン（すべて / 切断 / 再接続 / ステータス / Speedtest）
- 統計カード（切断回数・平均電波強度・平均DL速度など）
- イベント種別：📊ステータス / ❌切断 / 🟢再接続 / 📶Speedtest / 💤スリープ復帰
- STATUSエントリ展開時に位置情報セクション＋「🗺️ 地図を開く」リンク

### アンインストール
- `アンインストール.bat` で以下を一括削除
  - `C:\WiFiMonitor\`（プログラム・ログ・ビューア）
  - タスクスケジューラ「WiFi監視」
  - speedtest-cli（pip uninstall）

---

## ログのJSONスキーマ（STATUSエントリ）

```json
{
  "timestamp": "2026-03-12T17:07:09.123456",
  "type": "STATUS",
  "wifi": {
    "state": "接続されました",
    "ssid": "BE001",
    "bssid": "9c:d5:7d:1f:12:0f",
    "signal": "96",
    "radio_type": "802.11ac",
    "channel": "100",
    "rx_rate": "200",
    "tx_rate": "200",
    "auth": "WPA2-パーソナル",
    "rssi_dbm": -52,
    "snr_db": 43
  },
  "speedtest": {
    "download_mbps": 31.28,
    "upload_mbps": 47.34,
    "ping_ms": 40.6,
    "server": "Tsukuba",
    "timestamp": "2026-03-12T17:01:18"
  },
  "active_devices": 47,
  "location": {
    "address": "Japan Fukushima Fukushima",
    "lat": 37.75,
    "lon": 140.46,
    "maps_url": "https://www.google.com/maps?q=37.75,140.46",
    "updated_at": "2026-03-12T17:07:05.123456"
  }
}
```

---

## 判明している既知事項・知見

- **ゲートウェイ**: `192.168.1.1`（以前は `192.168.21.254` と混在していたが現在は `192.168.1.1`）
- **ローミング**: 大川原内でBSSIDが `9c:d5:7d:1f:12:0f`（Ch52/100） ↔ `9c:d5:7d:23:da:af`（Ch132）間で切り替わることがある
- **Speedtest実測値**: 約28〜31Mbps（リンク速度120〜180Mbpsに対して低い → 会社回線の共有帯域制限の可能性）
- **Speedtest 403エラー**: 特定サーバーが会社ネットワークでブロックされる場合がある。次回計測で別サーバーが選ばれて自然解消することが多い
- **RSSI取得方式**: Windows WLAN APIのopcode 8はチャンネル番号を返すためRSSI取得に使えない。`signal% / 2 - 100` で逆算する方式を採用
- **アクティブ機器数の限界**: ARPテーブルはPCが通信した相手のみ記録。AP-Aに接続中はAP-B/C配下の機器は見えない。ネットワーク全体の正確な台数ではない
- **広野事務所との関係**: 同じSSID「BE001」だがBSSIDは異なる。別ネットワークセグメントの可能性あり
- **コマンドプロンプト選択モード**: クリックで出力一時停止 → 「簡易編集モード」をOFFに設定済み

---

## 保留・TODO

| 項目 | 状態 | メモ |
|------|------|------|
| **Speedtestの頻度変更** | 🔵 保留 | 現在は約10分ごと（`SPEEDTEST_EVERY = 10`）。会社ネットワークへの影響・情報システム部との調整を考慮して変更を検討。設定値1箇所の変更のみで対応可能 |

---

## GitHubリポジトリ

- **登録済みファイル**: `wifi_monitor.py` / `log_viewer.html` / `README.md`
- **除外ファイル**: `起動.bat`（個人環境依存のため）
- **Description**: "Windows用 Wi-Fi接続監視・切断原因診断ツール"
