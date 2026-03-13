"""
Wi-Fi 接続監視・原因特定システム (Windows版)
監視間隔ごとに接続状況・電波強度・Pingを記録し、
切断時に詳細な診断レポートを自動生成します。
定期的に実際の通信速度（Speedtest）も計測・記録します。
"""

import subprocess
import time
import datetime
import json
import re
import sys
import threading
import ctypes
import ctypes.wintypes as wt
import urllib.request
from pathlib import Path

# ===================== 設定 =====================
CHECK_INTERVAL    = 5     # 監視間隔（秒）
PING_COUNT        = 4     # Ping送信回数
LOG_DIR           = Path("C:/WiFiMonitor/logs")
STATUS_LOG_EVERY  = 12    # N回に1回詳細表示＆ログ保存（60秒ごと）
SPEEDTEST_EVERY   = 10    # N回の詳細表示ごとにSpeedtest実行（約10分ごと）
LOCATION_EVERY    = 30    # N分ごとに位置情報を再取得（スリープ復帰後にも即更新）
SLEEP_DETECT_SEC  = 30    # N秒以上ループが遅延したらスリープ復帰と判断
# ================================================

# Speedtest結果を保持するグローバル変数
_speedtest_result  = None
_speedtest_running = False
_speedtest_lock    = threading.Lock()

# 位置情報キャッシュ
_location_cache = None
_location_lock  = threading.Lock()


def _setup_close_handler():
    """コンソールウィンドウのWM_CLOSEをサブクラス化して×ボタンを制御する"""
    WM_CLOSE        = 0x0010
    GWL_WNDPROC     = -4
    MB_YESNO        = 0x04
    MB_ICONQUESTION = 0x20
    IDYES           = 6

    hwnd = ctypes.windll.kernel32.GetConsoleWindow()
    if not hwnd:
        return

    WNDPROCTYPE = ctypes.WINFUNCTYPE(
        ctypes.c_longlong,
        wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM
    )

    # オリジナルのウィンドウプロシージャを保存
    original = ctypes.windll.user32.GetWindowLongPtrW(hwnd, GWL_WNDPROC)

    def custom_wndproc(h, msg, wp, lp):
        if msg == WM_CLOSE:
            result = ctypes.windll.user32.MessageBoxW(
                h,
                "Wi-Fi監視システムが終了します。\n本当に閉じますか？",
                "Wi-Fi監視システム",
                MB_YESNO | MB_ICONQUESTION,
            )
            if result == IDYES:
                # 元のプロシージャを呼び出して正常終了
                ctypes.windll.user32.CallWindowProcW(original, h, msg, wp, lp)
                sys.exit(0)
            return 0  # 閉じるをキャンセル（0を返すことでWM_CLOSEを無視）
        return ctypes.windll.user32.CallWindowProcW(original, h, msg, wp, lp)

    # GC防止のためグローバルに保持
    _setup_close_handler._proc = WNDPROCTYPE(custom_wndproc)
    ctypes.windll.user32.SetWindowLongPtrW(
        hwnd, GWL_WNDPROC, _setup_close_handler._proc
    )


def run_cmd(cmd, timeout=15):
    try:
        r = subprocess.run(
            cmd, capture_output=True, text=True,
            encoding="utf-8", errors="ignore", timeout=timeout
        )
        return r.stdout
    except Exception:
        return ""


NOISE_FLOOR_DBM = -95  # 推定ノイズフロア（dBm）


def signal_to_rssi(signal_pct):
    """
    Windowsのsignal%(0-100)からRSSI(dBm)を逆算する。
    Windows内部式: signal% = 2 * (RSSI + 100)  [-100≦RSSI≦-50]
    逆算:          RSSI = signal% / 2 - 100
    """
    try:
        s = int(str(signal_pct).replace('%', ''))
        rssi = round(s / 2 - 100)
        return rssi
    except Exception:
        return None


def get_wifi_info():
    out = run_cmd(["netsh", "wlan", "show", "interfaces"])
    patterns = {
        "state":      r"(?:State|状態)\s*:\s(.+)",
        "ssid":       r"^\s+SSID\s*:\s(.+)",
        "bssid":      r"(?:BSSID|AP BSSID)\s*:\s(.+)",
        "signal":     r"(?:Signal|シグナル)\s*:\s(\d+)%",
        "radio_type": r"(?:Radio type|無線の種類)\s*:\s(.+)",
        "channel":    r"(?:Channel|チャネル)\s*:\s(\d+)",
        "rx_rate":    r"(?:Receive rate \(Mbps\)|受信速度 \(Mbps\))\s*:\s(.+)",
        "tx_rate":    r"(?:Transmit rate \(Mbps\)|送信速度 \(Mbps\))\s*:\s(.+)",
        "auth":       r"(?:Authentication|認証)\s*:\s(.+)",
    }
    info = {}
    for key, pat in patterns.items():
        m = re.search(pat, out, re.MULTILINE)
        if m:
            info[key] = m.group(1).strip()

    # RSSI / SNR を signal% から逆算
    if "signal" in info:
        rssi = signal_to_rssi(info["signal"])
        if rssi is not None:
            info["rssi_dbm"] = rssi
            info["snr_db"]   = rssi - NOISE_FLOOR_DBM  # 推定SNR

    return info


def get_location():
    """Windows Location APIをPowerShell経由で取得し、逆ジオコーディングで住所を得る"""
    try:
        # PowerShellでWindows Location APIを呼び出す
        ps = """
Add-Type -AssemblyName System.Device
$loc = New-Object System.Device.Location.GeoCoordinateWatcher
$loc.Start()
$timeout = 10
$elapsed = 0
while ($loc.Status -ne 'Ready' -and $elapsed -lt $timeout) {
    Start-Sleep -Milliseconds 500
    $elapsed += 0.5
}
$coord = $loc.Position.Location
if ($coord.IsUnknown) {
    Write-Output "UNKNOWN"
} else {
    Write-Output "$($coord.Latitude),$($coord.Longitude)"
}
$loc.Stop()
"""
        out = run_cmd(["powershell", "-Command", ps], timeout=20).strip()

        if not out or out == "UNKNOWN" or "," not in out:
            return None

        parts = out.split(",")
        lat = float(parts[0].strip())
        lon = float(parts[1].strip())

        # 逆ジオコーディング（nominatim）で住所を取得
        try:
            geo_url = f"https://nominatim.openstreetmap.org/reverse?lat={lat}&lon={lon}&format=json&accept-language=ja"
            req = urllib.request.Request(geo_url, headers={"User-Agent": "WiFiMonitor/1.0"})
            with urllib.request.urlopen(req, timeout=10) as res:
                geo = json.loads(res.read().decode())
            addr_parts = geo.get("address", {})
            addr = " ".join(filter(None, [
                addr_parts.get("country", ""),
                addr_parts.get("province") or addr_parts.get("state", ""),
                addr_parts.get("city") or addr_parts.get("town") or addr_parts.get("village", ""),
                addr_parts.get("suburb") or addr_parts.get("neighbourhood", ""),
            ]))
        except Exception:
            addr = f"{lat}, {lon}"

        maps_url = f"https://www.google.com/maps?q={lat},{lon}"
        return {
            "address":  addr,
            "lat":      round(lat, 6),
            "lon":      round(lon, 6),
            "maps_url": maps_url,
        }

    except Exception:
        return None


def update_location():
    """バックグラウンドで位置情報を更新する"""
    global _location_cache
    loc = get_location()
    with _location_lock:
        if loc:
            _location_cache = loc
            _location_cache["updated_at"] = datetime.datetime.now().isoformat()


def start_location_update():
    t = threading.Thread(target=update_location, daemon=True)
    t.start()


def get_active_devices():
    """ARPテーブルからアクティブな機器数を取得する"""
    try:
        out = run_cmd(["arp", "-a"])
        # 動的エントリ（dynamic）のみカウント（静的・不完全なものを除外）
        dynamic = [l for l in out.splitlines() if "dynamic" in l.lower() or "動的" in l.lower()]
        if dynamic:
            return len(dynamic)
        # dynamicの表記がない環境用フォールバック（IPアドレス行をカウント）
        ip_lines = [l for l in out.splitlines() if re.search(r'\d+\.\d+\.\d+\.\d+', l) and "インターフェイス" not in l and "Interface" not in l and "アドレス" not in l and "Address" not in l]
        return len(ip_lines)
    except Exception:
        return None


def get_gateway():
    out = run_cmd(["ipconfig"])
    m = re.search(r"Default Gateway[.\s]*:\s*([\d.]+)", out)
    return m.group(1) if m else "192.168.1.1"


def ping(host, count=4):
    out = run_cmd(["ping", "-n", str(count), host], timeout=20)
    loss_m = re.search(r"(\d+)% loss", out)
    avg_m  = re.search(r"Average = (\d+)ms", out)
    min_m  = re.search(r"Minimum = (\d+)ms", out)
    max_m  = re.search(r"Maximum = (\d+)ms", out)
    loss   = int(loss_m.group(1)) if loss_m else 100
    return {
        "host":      host,
        "loss_pct":  loss,
        "avg_ms":    int(avg_m.group(1)) if avg_m else None,
        "min_ms":    int(min_m.group(1)) if min_m else None,
        "max_ms":    int(max_m.group(1)) if max_m else None,
        "reachable": loss < 100,
    }


def get_nearby_aps():
    out = run_cmd(["netsh", "wlan", "show", "networks", "mode=bssid"])
    aps = []
    blocks = re.split(r"(?=SSID \d+)", out)
    for block in blocks:
        ssid_m   = re.search(r"SSID \d+\s*:\s(.+)", block)
        signal_m = re.search(r"Signal\s*:\s(\d+)%", block)
        chan_m   = re.search(r"Channel\s*:\s(\d+)", block)
        if ssid_m:
            aps.append({
                "ssid":    ssid_m.group(1).strip(),
                "signal":  int(signal_m.group(1)) if signal_m else None,
                "channel": int(chan_m.group(1)) if chan_m else None,
            })
    return aps


def get_wlan_events(minutes=5):
    ps = (
        f"Get-WinEvent -LogName 'Microsoft-Windows-WLAN-AutoConfig/Operational' "
        f"-MaxEvents 30 -ErrorAction SilentlyContinue | "
        f"Where-Object {{ $_.TimeCreated -gt (Get-Date).AddMinutes(-{minutes}) }} | "
        f"Select-Object @{{n='time';e={{$_.TimeCreated.ToString('HH:mm:ss')}}}}, "
        f"@{{n='id';e={{$_.Id}}}}, "
        f"@{{n='msg';e={{$_.Message -replace '`n',' '}}}} | "
        f"ConvertTo-Json -Depth 2"
    )
    out = run_cmd(["powershell", "-Command", ps], timeout=20)
    if not out.strip():
        return []
    try:
        data = json.loads(out)
        return data if isinstance(data, list) else [data]
    except Exception:
        return []


def run_speedtest_thread():
    """バックグラウンドでSpeedtestを実行する"""
    global _speedtest_result, _speedtest_running
    try:
        import speedtest as st
        print("\n  📶 Speedtest 計測中（バックグラウンド）...", flush=True)
        s = st.Speedtest()
        s.get_best_server()
        download = s.download() / 1_000_000   # Mbps
        upload   = s.upload()   / 1_000_000   # Mbps
        ping_ms  = s.results.ping

        result = {
            "timestamp":   datetime.datetime.now().isoformat(),
            "download_mbps": round(download, 2),
            "upload_mbps":   round(upload, 2),
            "ping_ms":       round(ping_ms, 1),
            "server":        s.results.server.get("name", "---"),
        }

        with _speedtest_lock:
            _speedtest_result = result

        print(f"\n  📶 Speedtest 完了 ▼{result['download_mbps']}Mbps ▲{result['upload_mbps']}Mbps Ping:{result['ping_ms']}ms", flush=True)

        save_log({
            "timestamp": result["timestamp"],
            "type":      "SPEEDTEST",
            "speedtest": result,
        })

    except ImportError:
        print("\n  ⚠ speedtest-cli が未インストール。'pip install speedtest-cli' を実行してください。", flush=True)
        with _speedtest_lock:
            _speedtest_result = {"error": "speedtest-cli not installed"}
    except Exception as e:
        print(f"\n  ⚠ Speedtest エラー: {e}", flush=True)
        with _speedtest_lock:
            _speedtest_result = {"error": str(e)}
    finally:
        _speedtest_running = False


def start_speedtest():
    """Speedtestをバックグラウンドスレッドで起動"""
    global _speedtest_running
    if _speedtest_running:
        return
    _speedtest_running = True
    t = threading.Thread(target=run_speedtest_thread, daemon=True)
    t.start()


def analyze_causes(wifi_before, ping_results, gateway):
    causes = []
    recommendations = []

    sig = wifi_before.get("signal", "")
    if sig:
        s = int(sig.replace("%", ""))
        if s < 20:
            causes.append(f"電波強度が極めて低い（{s}%）- 切断直前")
            recommendations.append("ルーターに近づくか、中継器の設置を検討してください")
        elif s < 40:
            causes.append(f"電波強度が低い（{s}%）- 切断直前")
            recommendations.append("PCの設置場所の変更や中継器の追加を検討してください")

    gw   = ping_results.get(gateway, {})
    ext1 = ping_results.get("8.8.8.8", {})

    if not gw.get("reachable") and not ext1.get("reachable"):
        causes.append("ゲートウェイ・外部サーバーともに到達不能")
        recommendations.append("ルーターの再起動を試みてください")
    elif gw.get("reachable") and not ext1.get("reachable"):
        causes.append("ゲートウェイには到達できるが、インターネットに出られない")
        recommendations.append("ISP側の障害、またはルーターのWAN側設定を確認してください")
    elif not gw.get("reachable") and ext1.get("reachable"):
        causes.append("ルーターがPingに一時的に無応答（通信自体は維持されていた可能性あり）")

    gw_avg = gw.get("avg_ms")
    if gw_avg and gw_avg > 50:
        causes.append(f"ゲートウェイへのPing遅延が大きい（{gw_avg}ms）")
        recommendations.append("ルーターの負荷が高い可能性があります")

    if not causes:
        causes.append("原因を自動特定できませんでした（一時的な障害の可能性）")

    return causes, recommendations


def save_log(entry):
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    date_str = datetime.datetime.now().strftime("%Y-%m-%d")
    log_file = LOG_DIR / f"wifi_log_{date_str}.json"
    logs = []
    if log_file.exists():
        try:
            with open(log_file, "r", encoding="utf-8") as f:
                logs = json.load(f)
        except Exception:
            logs = []
    logs.append(entry)
    with open(log_file, "w", encoding="utf-8") as f:
        json.dump(logs, f, ensure_ascii=False, indent=2)


def sep(char="=", width=62):
    print(char * width)


def signal_bar(val):
    try:
        n = int(str(val).replace("%", ""))
    except Exception:
        return str(val)
    filled = round(n / 10)
    bar = "█" * filled + "░" * (10 - filled)
    return f"{bar} {n}%"


def print_detail(ts, wifi, gateway, start_dt, detail_count, active_devices=None):
    """詳細ステータスを画面に表示"""
    elapsed = datetime.datetime.now() - start_dt
    h, rem  = divmod(int(elapsed.total_seconds()), 3600)
    m, s    = divmod(rem, 60)

    with _speedtest_lock:
        sp = _speedtest_result

    print()
    sep()
    print(f"  📡 Wi-Fi 詳細ステータス  [{ts}]")
    sep("-")
    print(f"  監視開始    : {start_dt.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  経過時間    : {h:02d}:{m:02d}:{s:02d}")
    print(f"  ゲートウェイ: {gateway}")
    dev_str = f"{active_devices} 台（ARPテーブル）" if active_devices is not None else "---"
    print(f"  アクティブ機器: {dev_str}")
    with _location_lock:
        loc = _location_cache
    if loc:
        print(f"  現在地      : {loc.get('address', '---')}")
        print(f"  緯度経度    : {loc.get('lat', '---')}, {loc.get('lon', '---')}")
        print(f"  Googleマップ: {loc.get('maps_url', '---')}")
    sep("-")
    print(f"  SSID        : {wifi.get('ssid', '---')}")
    print(f"  電波強度    : {signal_bar(wifi.get('signal', '---'))}")
    print(f"  チャネル    : {wifi.get('channel', '---')} ch")
    print(f"  周波数帯    : {wifi.get('radio_type', '---')}")
    print(f"  受信速度(L) : {wifi.get('rx_rate', '---')} Mbps  ← リンク速度")
    print(f"  送信速度(L) : {wifi.get('tx_rate', '---')} Mbps  ← リンク速度")
    print(f"  BSSID       : {wifi.get('bssid', '---')}")
    print(f"  認証方式    : {wifi.get('auth', '---')}")
    if wifi.get('rssi_dbm') is not None:
        print(f"  RSSI        : {wifi.get('rssi_dbm')} dBm")
        print(f"  SNR (推定)  : {wifi.get('snr_db')} dB")
    sep("-")

    # Speedtest結果表示
    if sp is None:
        print(f"  Speedtest   : 計測待ち（約10分ごとに自動計測）")
    elif "error" in sp:
        print(f"  Speedtest   : エラー - {sp['error']}")
    else:
        st_ts = sp.get("timestamp", "")[:19].replace("T", " ")
        print(f"  ▼ ダウンロード: {sp.get('download_mbps', '---')} Mbps  （実測値）")
        print(f"  ▲ アップロード: {sp.get('upload_mbps',   '---')} Mbps  （実測値）")
        print(f"  🏓 Ping       : {sp.get('ping_ms',        '---')} ms")
        print(f"  サーバー      : {sp.get('server',          '---')}")
        print(f"  計測時刻      : {st_ts}")

    sep("-")
    next_st_min = (SPEEDTEST_EVERY - (detail_count % SPEEDTEST_EVERY)) * CHECK_INTERVAL * STATUS_LOG_EVERY // 60
    print(f"  次の詳細更新まで {CHECK_INTERVAL * STATUS_LOG_EVERY} 秒  |  次のSpeedtestまで約{next_st_min}分  |  終了: Ctrl+C")
    sep()
    print()


def run_monitor():
    global _speedtest_result

    _setup_close_handler()  # ×ボタン確認ポップアップを有効化

    start_dt     = datetime.datetime.now()
    detail_count = 0   # 詳細表示の回数カウント

    sep()
    print("  Wi-Fi 接続監視システム  起動完了")
    print(f"  監視開始   : {start_dt.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  ログ保存先 : {LOG_DIR}")
    print(f"  監視間隔   : {CHECK_INTERVAL} 秒")
    print(f"  Speedtest  : 約{CHECK_INTERVAL * STATUS_LOG_EVERY * SPEEDTEST_EVERY // 60}分ごとに自動計測")
    print("  終了するには Ctrl + C を押してください")
    sep()

    # speedtest-cli のインストール確認
    try:
        import speedtest
        print("  [OK] speedtest-cli が使用可能です\n")
    except ImportError:
        print("  [!] speedtest-cli が未インストールです")
        print("  インストールするには次のコマンドを実行してください:")
        print(f"  %LocalAppData%\\Programs\\Python\\Python314\\python.exe -m pip install speedtest-cli\n")

    gateway   = get_gateway()
    print(f"  デフォルトゲートウェイ: {gateway}")

    # 起動時に位置情報を取得
    print("  📍 位置情報を取得中...")
    start_location_update()
    print()

    prev_wifi          = {}
    was_connected      = None
    loop_count         = 0
    last_loop_time     = datetime.datetime.now()
    last_location_time = datetime.datetime.now()

    # 起動時に最初のSpeedtestを実行
    start_speedtest()

    while True:
        try:
            now = datetime.datetime.now()

            # スリープ復帰検知
            elapsed_since_last = (now - last_loop_time).total_seconds()
            if elapsed_since_last > SLEEP_DETECT_SEC:
                ts_wake = now.strftime("%H:%M:%S")
                print(f"\n\n[{ts_wake}] 💤 スリープ復帰を検知しました。位置情報を更新中...")
                start_location_update()
                last_location_time = now
                save_log({
                    "timestamp": now.isoformat(),
                    "type":      "WAKE",
                    "sleep_duration_sec": round(elapsed_since_last),
                })
            last_loop_time = now

            # 定期的な位置情報更新
            if (now - last_location_time).total_seconds() >= LOCATION_EVERY * 60:
                start_location_update()
                last_location_time = now

            wifi = get_wifi_info()
            state_val    = wifi.get("state", "").strip()
            is_connected = state_val in ("connected", "接続済み", "接続", "接続されました")
            ts           = now.strftime("%H:%M:%S")

            # ---- 接続中 ----
            if is_connected:
                sig  = wifi.get("signal", "---")
                ssid = wifi.get("ssid", "---")
                ch   = wifi.get("channel", "---")
                rx   = wifi.get("rx_rate", "---")
                print(f"\r[{ts}] ✅ 接続中 | {ssid} | 電波:{sig}% | Ch:{ch} | Rx:{rx}Mbps   ", end="", flush=True)

                # 定期詳細表示＆ログ保存
                if loop_count % STATUS_LOG_EVERY == 0:
                    with _speedtest_lock:
                        sp = _speedtest_result
                    active_devices = get_active_devices()
                    print_detail(ts, wifi, gateway, start_dt, detail_count, active_devices)
                    with _location_lock:
                        loc = _location_cache
                    save_log({
                        "timestamp":      datetime.datetime.now().isoformat(),
                        "type":           "STATUS",
                        "wifi":           wifi,
                        "speedtest":      sp,
                        "active_devices": active_devices,
                        "location":       loc,
                    })

                    # Speedtest定期実行
                    if detail_count % SPEEDTEST_EVERY == 0 and detail_count > 0:
                        start_speedtest()

                    detail_count += 1

                # 再接続検出
                if was_connected is False:
                    print(f"\n\n[{ts}] 🟢 再接続を検出しました！")
                    save_log({
                        "timestamp": datetime.datetime.now().isoformat(),
                        "type":      "RECONNECTION",
                        "wifi":      wifi,
                    })
                    # 再接続後にもSpeedtestを実行
                    start_speedtest()

                was_connected = True

            # ---- 切断 ----
            else:
                if was_connected is True:
                    print(f"\n\n[{ts}] ❌ 切断を検出！ 診断を実行中...\n")

                    ping_results = {}
                    for host in [gateway, "8.8.8.8", "1.1.1.1"]:
                        print(f"  Ping → {host} ...", end="", flush=True)
                        r = ping(host, PING_COUNT)
                        ping_results[host] = r
                        status = "✅" if r["reachable"] else "❌"
                        avg    = f"{r['avg_ms']}ms" if r["avg_ms"] else "タイムアウト"
                        print(f" {status} {avg}  ロス:{r['loss_pct']}%")

                    print("\n  周辺アクセスポイントをスキャン中...")
                    nearby_aps = get_nearby_aps()

                    print("  Windowsイベントログを確認中...")
                    events = get_wlan_events(minutes=5)

                    causes, recommendations = analyze_causes(prev_wifi, ping_results, gateway)

                    print()
                    sep()
                    print(f"  切断診断レポート  [{ts}]")
                    sep()

                    print("\n【考えられる原因】")
                    for c in causes:
                        print(f"  ・{c}")

                    if recommendations:
                        print("\n【推奨対処】")
                        for r in recommendations:
                            print(f"  → {r}")

                    print("\n【切断直前のWi-Fi状態】")
                    for k, v in prev_wifi.items():
                        print(f"  {k:12s}: {v}")

                    if nearby_aps:
                        print(f"\n【周辺AP（{len(nearby_aps)}件）】")
                        same_ch = [ap for ap in nearby_aps if ap.get("channel") == int(prev_wifi.get("channel", 0) or 0)]
                        if same_ch:
                            print(f"  ⚠ 同チャンネルのAPが {len(same_ch)} 件あります（干渉の可能性）")
                        for ap in nearby_aps[:5]:
                            print(f"  SSID:{ap['ssid']}  Ch:{ap.get('channel','?')}  電波:{ap.get('signal','?')}%")

                    if events:
                        print(f"\n【直近のWLANイベント（{len(events)}件）】")
                        for e in events[:5]:
                            print(f"  [{e.get('time','')}] ID:{e.get('id','')} {str(e.get('msg',''))[:80]}")

                    sep()
                    print(f"  詳細ログ: {LOG_DIR}")
                    sep()
                    print()

                    save_log({
                        "timestamp":       datetime.datetime.now().isoformat(),
                        "type":            "DISCONNECTION",
                        "wifi_before":     prev_wifi,
                        "wifi_after":      wifi,
                        "ping_results":    ping_results,
                        "nearby_aps":      nearby_aps,
                        "wlan_events":     events,
                        "possible_causes": causes,
                        "recommendations": recommendations,
                    })

                else:
                    print(f"\r[{ts}] ❌ 未接続（接続待機中...）", end="", flush=True)

                was_connected = False

            prev_wifi = wifi
            loop_count += 1
            time.sleep(CHECK_INTERVAL)

        except KeyboardInterrupt:
            print("\n\n監視を終了しました。お疲れ様でした。\n")
            sys.exit(0)
        except Exception as e:
            print(f"\n[エラー] {e}")
            time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    run_monitor()
