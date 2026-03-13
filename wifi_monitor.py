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
import concurrent.futures
from pathlib import Path

# ===================== 設定 =====================
CHECK_INTERVAL    = 5     # 監視間隔（秒）
ARP_SCAN_EVERY    = 10    # N回の詳細表示ごとにARPスキャン実行（約10分ごと）
PING_COUNT        = 4     # Ping送信回数
LOG_DIR           = Path("C:/WiFiMonitor/logs")
STATUS_LOG_EVERY  = 12    # N回に1回詳細表示＆ログ保存（60秒ごと）
SPEEDTEST_EVERY   = 30    # N回の詳細表示ごとにSpeedtest実行（約30分ごと）
BSS_SCAN_EVERY    = 10    # N回の詳細表示ごとにBSSスキャン実行（約10分ごと）
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

# ARPスキャン結果キャッシュ
_arp_scan_result  = None
_arp_scan_running = False
_arp_scan_lock    = threading.Lock()


def _arp_scan_worker(gateway):
    global _arp_scan_result, _arp_scan_running
    try:
        result = arp_ping_sweep(gateway)
        result["timestamp"] = datetime.datetime.now().isoformat()
        # スキャン後にARPテーブルを再取得してカウント更新
        result["arp_count"] = get_active_devices()
        with _arp_scan_lock:
            _arp_scan_result = result
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        print(f"\n  🔍 ARPスキャン完了: {result['count']} 台が応答 (ARP台数: {result['arp_count']} 台)  [{ts}]", flush=True)
    except Exception as e:
        print(f"\n  ⚠ ARPスキャン エラー: {e}", flush=True)
    finally:
        _arp_scan_running = False


def start_arp_scan(gateway):
    global _arp_scan_running
    with _arp_scan_lock:
        if _arp_scan_running:
            return
        _arp_scan_running = True
    t = threading.Thread(target=_arp_scan_worker, args=(gateway,), daemon=True)
    t.start()


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
            cmd, capture_output=True, timeout=timeout
        )
        # 日本語WindowsはCP932（Shift-JIS）で出力するため、まずCP932でデコードを試みる
        for enc in ("cp932", "utf-8"):
            try:
                return r.stdout.decode(enc)
            except UnicodeDecodeError:
                continue
        return r.stdout.decode("utf-8", errors="ignore")
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


# ---- WlanApi 実RSSI取得 ----------------------------------------
import ctypes, ctypes.wintypes

def _get_real_rssi(target_bssid: str):
    """
    Wlanapi.dll の WlanGetNetworkBssList を呼び出し、
    接続中BSSIDの実RSSIを dBm で返す。
    取得失敗時は None（fallbackとして signal_to_rssi を使う）。
    """
    try:
        wlan = ctypes.windll.wlanapi

        ver    = ctypes.wintypes.DWORD()
        handle = ctypes.wintypes.HANDLE()
        if wlan.WlanOpenHandle(2, None, ctypes.byref(ver), ctypes.byref(handle)) != 0:
            return None

        class WLAN_INTERFACE_INFO(ctypes.Structure):
            _fields_ = [("InterfaceGuid",            ctypes.c_byte * 16),
                        ("strInterfaceDescription",   ctypes.c_wchar * 256),
                        ("isState",                   ctypes.c_uint)]

        class WLAN_INTERFACE_INFO_LIST(ctypes.Structure):
            _fields_ = [("dwNumberOfItems", ctypes.wintypes.DWORD),
                        ("dwIndex",         ctypes.wintypes.DWORD),
                        ("InterfaceInfo",   WLAN_INTERFACE_INFO * 64)]

        p_iface = ctypes.POINTER(WLAN_INTERFACE_INFO_LIST)()
        if wlan.WlanEnumInterfaces(handle, None, ctypes.byref(p_iface)) != 0:
            wlan.WlanCloseHandle(handle, None)
            return None

        iface_list = p_iface.contents
        if iface_list.dwNumberOfItems == 0:
            wlan.WlanFreeMemory(p_iface)
            wlan.WlanCloseHandle(handle, None)
            return None

        guid = (ctypes.c_byte * 16)(*iface_list.InterfaceInfo[0].InterfaceGuid)

        class WLAN_BSS_ENTRY(ctypes.Structure):
            _fields_ = [
                ("dot11Ssid",               ctypes.c_byte * 36),
                ("uPhyId",                  ctypes.wintypes.DWORD),
                ("dot11Bssid",              ctypes.c_byte * 6),
                ("dot11BssType",            ctypes.c_uint),
                ("dot11BssPhyType",         ctypes.c_uint),
                ("lRssi",                   ctypes.c_long),
                ("uLinkQuality",            ctypes.wintypes.DWORD),
                ("bInRegDomain",            ctypes.c_bool),
                ("usBeaconPeriod",          ctypes.c_ushort),
                ("ullTimestamp",            ctypes.c_ulonglong),
                ("ullHostTimestamp",        ctypes.c_ulonglong),
                ("usCapabilityInformation", ctypes.c_ushort),
                ("ulChCenterFrequency",     ctypes.wintypes.DWORD),
                ("wlanRateSet",             ctypes.c_byte * 130),
                ("ulIeOffset",              ctypes.wintypes.DWORD),
                ("ulIeSize",                ctypes.wintypes.DWORD),
            ]

        class WLAN_BSS_LIST(ctypes.Structure):
            _fields_ = [("dwTotalSize",     ctypes.wintypes.DWORD),
                        ("dwNumberOfItems", ctypes.wintypes.DWORD),
                        ("wlanBssEntries",  WLAN_BSS_ENTRY * 256)]

        p_bss = ctypes.POINTER(WLAN_BSS_LIST)()
        ret   = wlan.WlanGetNetworkBssList(
            handle, ctypes.byref(guid), None, 1, None, None, ctypes.byref(p_bss))

        rssi = None
        if ret == 0:
            bss_list = p_bss.contents
            target   = target_bssid.lower().replace(":", "")
            for i in range(bss_list.dwNumberOfItems):
                entry = bss_list.wlanBssEntries[i]
                mac   = "".join(f"{b:02x}" for b in entry.dot11Bssid)
                if mac == target:
                    rssi = int(entry.lRssi)
                    break
            wlan.WlanFreeMemory(p_bss)

        wlan.WlanFreeMemory(p_iface)
        wlan.WlanCloseHandle(handle, None)
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

    # RSSI / SNR: まず WlanApi で実値取得、失敗時は signal% から逆算
    if "signal" in info:
        bssid = info.get("bssid", "")
        real_rssi = _get_real_rssi(bssid) if bssid else None
        if real_rssi is not None:
            info["rssi_dbm"]      = real_rssi
            info["rssi_source"]   = "wlanapi"   # 実測値
        else:
            est = signal_to_rssi(info["signal"])
            if est is not None:
                info["rssi_dbm"]    = est
                info["rssi_source"] = "estimated"  # 逆算値
        if info.get("rssi_dbm") is not None:
            info["snr_db"] = info["rssi_dbm"] - NOISE_FLOOR_DBM

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


def arp_ping_sweep(gateway, timeout_ms=300):
    """
    サブネット全体にICMP pingを並列送信し、ARPテーブルを更新する。
    戻り値: { count: int, hosts: [str] }  ← 応答があったIPの一覧
    実行はバックグラウンドスレッドで行う想定。所要時間は約5〜15秒。
    """
    # ゲートウェイからサブネットプレフィックスを推定（例: 192.168.1.x）
    parts = gateway.split(".")
    if len(parts) != 4:
        return {"count": 0, "hosts": []}
    prefix = ".".join(parts[:3])

    def _ping_host(ip):
        try:
            r = subprocess.run(
                ["ping", "-n", "1", "-w", str(timeout_ms), ip],
                capture_output=True, timeout=2
            )
            out = r.stdout.decode("cp932", errors="ignore")
            return ip if ("TTL" in out or "ttl" in out) else None
        except Exception:
            return None

    with concurrent.futures.ThreadPoolExecutor(max_workers=64) as ex:
        targets = [f"{prefix}.{i}" for i in range(1, 255)]
        results = list(ex.map(_ping_host, targets))

    hosts = [h for h in results if h]
    return {"count": len(hosts), "hosts": sorted(hosts, key=lambda x: int(x.split(".")[-1]))}


def get_gateway():
    out = run_cmd(["ipconfig"])
    m = re.search(r"Default Gateway[.\s]*:\s*([\d.]+)", out)
    return m.group(1) if m else "192.168.1.1"


def ping(host, count=4):
    out = run_cmd(["ping", "-n", str(count), host], timeout=20)
    # 日本語: "0% の損失" / 英語: "0% loss"
    loss_m = re.search(r"(\d+)%\s*(?:の損失|loss)", out)
    avg_m  = re.search(r"(?:Average|平均)\s*=\s*(\d+)\s*ms", out, re.IGNORECASE)
    min_m  = re.search(r"(?:Minimum|最小)\s*=\s*(\d+)\s*ms", out, re.IGNORECASE)
    max_m  = re.search(r"(?:Maximum|最大)\s*=\s*(\d+)\s*ms", out, re.IGNORECASE)
    loss   = int(loss_m.group(1)) if loss_m else 100
    return {
        "host":      host,
        "loss_pct":  loss,
        "avg_ms":    int(avg_m.group(1)) if avg_m else None,
        "min_ms":    int(min_m.group(1)) if min_m else None,
        "max_ms":    int(max_m.group(1)) if max_m else None,
        "reachable": loss < 100,
    }


def quick_ping_loss(host, count=4):
    """
    定期パケットロス計測用の軽量版ping。
    戻り値: { loss_pct: int, avg_ms: int|None }
    失敗時: { loss_pct: 100, avg_ms: None }
    """
    try:
        out    = run_cmd(["ping", "-n", str(count), "-w", "1000", host], timeout=15)
        # 日本語: "0% の損失" / 英語: "0% loss"
        loss_m = re.search(r"(\d+)%\s*(?:の損失|loss)", out)
        avg_m  = re.search(r"(?:Average|平均)\s*=\s*(\d+)\s*ms", out, re.IGNORECASE)
        loss   = int(loss_m.group(1)) if loss_m else 100
        return {
            "loss_pct": loss,
            "avg_ms":   int(avg_m.group(1)) if avg_m else None,
        }
    except Exception:
        return {"loss_pct": 100, "avg_ms": None}


_bss_cache = None  # 直前の有効なBSSスキャン結果をキャッシュ


def get_bss_scan(current_bssid=""):
    """
    BSSスキャン: SSIDごと・BSSIDごとの接続台数・電波状況を取得する。
    current_bssid: 自分が現在接続しているBSSID（強調表示用）
    戻り値: { ssids: [...], total_stations: int, cached: bool }
    出力が空/不完全な場合は直前のキャッシュを返す（cached=True）。
    """
    global _bss_cache

    out = run_cmd(["netsh", "wlan", "show", "networks", "mode=bssid"])

    # SSIDブロックに分割
    # 注意: "BSSID \d+" にも "SSID \d+" が含まれるため負の後読みで除外
    ssid_blocks = re.split(r"(?<!B)(?=SSID \d+\s*:)", out)
    ssid_list = []

    for block in ssid_blocks:
        # 先頭行のSSID名を取得（"BSSID"行にはマッチしないよう行頭を想定）
        ssid_m = re.search(r"(?<!B)SSID \d+\s*:[ \t]*(.*)", block)
        if not ssid_m:
            continue
        ssid_name = ssid_m.group(1).strip()
        auth_m    = re.search(r"(?:Authentication|認証)\s*:\s*(.+)", block)
        auth      = auth_m.group(1).strip() if auth_m else ""

        # BSSIDブロックに分割
        bssid_blocks = re.split(r"(?=BSSID \d+\s*:)", block)
        bssid_list = []
        for bb in bssid_blocks:
            bssid_m   = re.search(r"BSSID \d+\s*:\s*([0-9a-f:]+)", bb, re.IGNORECASE)
            signal_m  = re.search(r"(?:Signal|シグナル)\s*:\s*(\d+)%", bb)
            band_m    = re.search(r"(?:Band|バンド)\s*:\s*(.+)", bb)
            chan_m    = re.search(r"(?:Channel|チャネル)\s*:\s*(\d+)", bb)
            # 「接続されているステーション」と「接続されていステーション」の両表記に対応
            station_m = re.search(r"(?:接続されてい[るい]?\s*ステーション|Connected\s+[Ss]tations?)\s*:\s*(\d+)", bb)
            radio_m   = re.search(r"(?:Radio type|無線タイプ)\s*:\s*(.+)", bb)
            # チャンネル使用率: "チャンネル使用率: 85 (33 %)" または "Channel Utilization: 85 (33 %)"
            util_m    = re.search(r"(?:チャンネル使用率|Channel [Uu]tilization)\s*:\s*(\d+)\s*\((\d+)\s*%\)", bb)
            if not bssid_m:
                continue
            bssid = bssid_m.group(1).strip().lower()
            bssid_list.append({
                "bssid":        bssid,
                "signal":       int(signal_m.group(1)) if signal_m else None,
                "band":         band_m.group(1).strip() if band_m else None,
                "channel":      int(chan_m.group(1)) if chan_m else None,
                "stations":     int(station_m.group(1)) if station_m else None,  # Noneは非対応AP
                "radio":        radio_m.group(1).strip() if radio_m else None,
                "is_mine":      bssid == current_bssid.lower(),
                "ch_util_pct":  int(util_m.group(2)) if util_m else None,
                "ch_util_raw":  int(util_m.group(1)) if util_m else None,
            })

        ssid_total = sum(b["stations"] for b in bssid_list if b["stations"] is not None)
        ssid_list.append({
            "ssid":     ssid_name,
            "auth":     auth,
            "bssids":   bssid_list,
            "total":    ssid_total,
        })

    # BSSIDエントリが1件も取れなかった場合はキャッシュにフォールバック
    total_bssid_entries = sum(len(s["bssids"]) for s in ssid_list)
    if total_bssid_entries == 0 and _bss_cache is not None:
        cached = dict(_bss_cache)
        cached["cached"] = True
        # is_mine フラグだけ現在のBSSIDで更新する
        if current_bssid:
            for s in cached["ssids"]:
                for b in s["bssids"]:
                    b["is_mine"] = (b["bssid"] == current_bssid.lower())
        return cached

    result = {"ssids": ssid_list, "total_stations": sum(s["total"] for s in ssid_list), "cached": False,
              "timestamp": datetime.datetime.now().isoformat()}
    # BSSIDエントリがあれば正式キャッシュ。SSIDのみ（BSSIDなし）でも初回キャッシュとして保存
    if total_bssid_entries > 0 or (ssid_list and _bss_cache is None):
        _bss_cache = result
    return result


def get_nearby_aps():
    """切断診断用（後方互換）: 簡易AP一覧を返す"""
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


def print_bss_scan(scan, label="定期"):
    """BSSスキャン結果をコンソールに表示する"""
    total      = scan.get("total_stations", 0)
    is_cached  = scan.get("cached", False)
    cache_note = "  ※ キャッシュ使用（スキャン不完全）" if is_cached else ""
    print(f"\n  ┌─ BSSスキャン結果（{label}）{cache_note}─────────────────────────")
    print(f"  │  エリア全体の接続台数: {total} 台")
    print(f"  │  ※ BSS台数=APが把握する実接続数  ARP台数=自PC経由の推定値")
    print(f"  │")
    for s in scan.get("ssids", []):
        ssid_label = s["ssid"] if s["ssid"] else f"(名称なし / {s['auth']})"
        all_unsupported = all(b["stations"] is None for b in s.get("bssids", [])) and s.get("bssids")
        total_str = "BSS Load非対応" if all_unsupported else f"合計 {s['total']} 台"
        print(f"  │  [{ssid_label}]  {total_str}")
        for b in s.get("bssids", []):
            mine  = "★ " if b["is_mine"] else "  "
            band  = b["band"] or "---"
            ch    = b["channel"] or "---"
            sig   = f"{b['signal']}%" if b["signal"] is not None else "---"
            st    = f"{b['stations']}台" if b["stations"] is not None else "非対応"
            util  = f"  Ch使用率:{b['ch_util_pct']}%" if b.get("ch_util_pct") is not None else ""
            mine_label = "  ← 自分が接続中" if b["is_mine"] else ""
            # 台数が非対応の場合はSSID合計行にも注記
            print(f"  │  {mine}{b['bssid']}  {band} Ch{ch}  {st}  シグナル:{sig}{util}{mine_label}")
    print(f"  └──────────────────────────────────────────────────")


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
    """バックグラウンドでSpeedtestを実行する（403時は別サーバーに自動切替）"""
    global _speedtest_result, _speedtest_running
    MAX_RETRY = 6  # 最大試行サーバー数
    try:
        import speedtest as st
        print("\n  📶 Speedtest 計測中（バックグラウンド）...", flush=True)
        s = st.Speedtest()

        # 全サーバーをレイテンシ順に取得
        s.get_servers()
        all_servers = sorted(
            [sv for slist in s.servers.values() for sv in slist],
            key=lambda x: x.get("latency", 9999)
        )

        tried_ids = set()
        last_error = None
        success = False

        for attempt, sv in enumerate(all_servers, start=1):
            if attempt > MAX_RETRY:
                break
            sv_id = sv.get("id")
            if sv_id in tried_ids:
                continue
            tried_ids.add(sv_id)

            try:
                s.get_best_server([sv])
                download = s.download() / 1_000_000
                upload   = s.upload()   / 1_000_000
                ping_ms  = s.results.ping
                last_error = None
                success = True
                break  # 成功
            except Exception as e:
                last_error = e
                err_str = str(e)
                if "403" in err_str or "Forbidden" in err_str:
                    print(f"\n  ⚠ Speedtest 403 [{sv.get('name','?')}] → 次のサーバーへ ({attempt}/{MAX_RETRY})", flush=True)
                    continue  # 次のサーバーを試す
                else:
                    # 403以外（ネットワーク断など）はリトライしない
                    break

        if not success:
            raise last_error or Exception("全サーバーで計測失敗")

        result = {
            "timestamp":     datetime.datetime.now().isoformat(),
            "download_mbps": round(download, 2),
            "upload_mbps":   round(upload, 2),
            "ping_ms":       round(ping_ms, 1),
            "server":        s.results.server.get("name", "---"),
        }

        with _speedtest_lock:
            _speedtest_result = result

        print(f"\n  📶 Speedtest 完了 ▼{result['download_mbps']}Mbps ▲{result['upload_mbps']}Mbps Ping:{result['ping_ms']}ms (サーバー: {result['server']})", flush=True)

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


def print_detail(ts, wifi, gateway, start_dt, detail_count, active_devices=None, packet_loss=None, arp_scan=None):
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
    # ARPスキャン済みならping応答数も併記
    if arp_scan:
        arp_ts   = arp_scan.get("timestamp", "")
        arp_time = arp_ts[11:16] if arp_ts else "---"
        dev_str  = f"{active_devices} 台（ARPテーブル）  ping応答: {arp_scan['count']} 台 / {arp_time} 取得"
    else:
        dev_str = f"{active_devices} 台（ARPテーブル）" if active_devices is not None else "---（ARPスキャン待ち）"
    print(f"  アクティブ機器: {dev_str}")
    # BSSキャッシュから接続台数を表示（自分が接続中のBSSIDの台数 + エリア合計）
    if _bss_cache:
        bss_total = _bss_cache.get("total_stations", 0)
        bss_ts    = _bss_cache.get("timestamp", "")
        bss_time  = bss_ts[11:16] if bss_ts else "---"
        mine_stations = None
        current_bssid = wifi.get("bssid", "").lower()
        for s in _bss_cache.get("ssids", []):
            for b in s.get("bssids", []):
                if b.get("bssid") == current_bssid:
                    mine_stations = b.get("stations")
                    break
        cached_note = "キャッシュ" if _bss_cache.get("cached") else f"{bss_time} 取得"
        mine_str = f"  自AP:{mine_stations}台" if mine_stations is not None else ""
        print(f"  BSS接続台数 : {bss_total} 台（エリア合計{mine_str} / {cached_note}）")
    else:
        print(f"  BSS接続台数 : --- （取得待ち・約10分ごとに自動更新）")
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
        src = "実測値" if wifi.get("rssi_source") == "wlanapi" else "推定値"
        snr_src = "実測RSSI基準" if wifi.get("rssi_source") == "wlanapi" else "推定"
        print(f"  RSSI        : {wifi.get('rssi_dbm')} dBm  ({src})")
        print(f"  SNR ({snr_src}): {wifi.get('snr_db')} dB")
    if packet_loss is not None:
        loss = packet_loss.get("loss_pct", 100)
        avg  = packet_loss.get("avg_ms")
        loss_icon = "✅" if loss == 0 else ("⚠" if loss < 50 else "❌")
        avg_str  = f"  Ping:{avg}ms" if avg is not None else ""
        print(f"  パケットロス: {loss_icon} {loss}%{avg_str}  (GW宛4回)")
    sep("-")

    # Speedtest結果表示
    if sp is None:
        print(f"  Speedtest   : 計測待ち（約{CHECK_INTERVAL * STATUS_LOG_EVERY * SPEEDTEST_EVERY // 60}分ごとに自動計測）")
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

    prev_wifi           = {}
    was_connected       = None
    bssid_at_disconnect = None   # 切断直前のBSSID（再接続時のローミング判定用）
    loop_count          = 0
    last_loop_time      = datetime.datetime.now()
    last_location_time  = datetime.datetime.now()
    last_bss_scan_count = 0      # BSSスキャン実行済みの detail_count

    # 起動時に最初のSpeedtestを実行
    start_speedtest()

    # 起動時にARPスキャンを実行（バックグラウンド）
    start_arp_scan(gateway)

    # 起動時にBSSスキャンを実行（接続済みの場合のみ）
    _startup_wifi = get_wifi_info()
    _startup_bssid = _startup_wifi.get("bssid", "")
    if _startup_bssid:
        bss = get_bss_scan(current_bssid=_startup_bssid)
        print_bss_scan(bss, label="起動時")
        save_log({
            "timestamp": datetime.datetime.now().isoformat(),
            "type":      "BSS_SCAN",
            "trigger":   "startup",
            "bss_scan":  bss,
        })

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
                    # ARPスキャン済みならその結果でactive_devicesを更新
                    with _arp_scan_lock:
                        arp_scan = _arp_scan_result
                    active_devices = arp_scan["arp_count"] if arp_scan else get_active_devices()
                    # パケットロス定期計測（ゲートウェイへ4回ping）
                    pkt = quick_ping_loss(gateway, count=4)
                    print_detail(ts, wifi, gateway, start_dt, detail_count,
                                 active_devices, packet_loss=pkt, arp_scan=arp_scan)
                    with _location_lock:
                        loc = _location_cache
                    save_log({
                        "timestamp":      datetime.datetime.now().isoformat(),
                        "type":           "STATUS",
                        "wifi":           wifi,
                        "speedtest":      sp,
                        "active_devices": active_devices,
                        "arp_scan":       arp_scan,
                        "packet_loss":    pkt,
                        "location":       loc,
                    })

                    # Speedtest定期実行
                    if detail_count % SPEEDTEST_EVERY == 0 and detail_count > 0:
                        start_speedtest()

                    # ARPスキャン定期実行（約10分ごと）
                    if detail_count % ARP_SCAN_EVERY == 0 and detail_count > 0:
                        start_arp_scan(gateway)

                    # BSSスキャン定期実行（約10分ごと）
                    if detail_count % BSS_SCAN_EVERY == 0 and detail_count > 0:
                        bss = get_bss_scan(current_bssid=wifi.get("bssid", ""))
                        print_bss_scan(bss, label="定期")
                        last_bss_scan_count = detail_count
                        save_log({
                            "timestamp":  datetime.datetime.now().isoformat(),
                            "type":       "BSS_SCAN",
                            "trigger":    "periodic",
                            "bss_scan":   bss,
                        })

                    detail_count += 1

                # 再接続検出
                if was_connected is False:
                    curr_bssid = wifi.get("bssid", "")
                    is_roaming_reconnect = (
                        bssid_at_disconnect and curr_bssid
                        and bssid_at_disconnect != curr_bssid.lower()
                    )
                    if is_roaming_reconnect:
                        print(f"\n\n[{ts}] 🔀 ローミング（切断経由）を検出！")
                        print(f"  BSSID: {bssid_at_disconnect} → {curr_bssid}")
                        # ローミング後にBSSスキャンを実行
                        bss = get_bss_scan(current_bssid=curr_bssid)
                        print_bss_scan(bss, label="ローミング後")
                        save_log({
                            "timestamp":    datetime.datetime.now().isoformat(),
                            "type":         "ROAMING",
                            "roaming_via":  "disconnect",
                            "bssid_from":   bssid_at_disconnect,
                            "bssid_to":     curr_bssid,
                            "channel_from": prev_wifi.get("channel", "---"),
                            "channel_to":   wifi.get("channel", "---"),
                            "wifi":         wifi,
                            "bss_scan":     bss,
                        })
                    else:
                        print(f"\n\n[{ts}] 🟢 再接続を検出しました！")
                        save_log({
                            "timestamp": datetime.datetime.now().isoformat(),
                            "type":      "RECONNECTION",
                            "wifi":      wifi,
                        })
                    bssid_at_disconnect = None
                    # 再接続後にもSpeedtestを実行
                    start_speedtest()

                # シームレスローミング検出（接続維持のままBSSIDが変化）
                elif was_connected is True:
                    prev_bssid = prev_wifi.get("bssid", "")
                    curr_bssid = wifi.get("bssid", "")
                    if prev_bssid and curr_bssid and prev_bssid != curr_bssid:
                        prev_ch = prev_wifi.get("channel", "---")
                        curr_ch = wifi.get("channel", "---")
                        print(f"\n\n[{ts}] 🔀 ローミングを検出！")
                        print(f"  BSSID: {prev_bssid} (Ch:{prev_ch}) → {curr_bssid} (Ch:{curr_ch})")
                        bss = get_bss_scan(current_bssid=curr_bssid)
                        print_bss_scan(bss, label="ローミング後")
                        save_log({
                            "timestamp":    datetime.datetime.now().isoformat(),
                            "type":         "ROAMING",
                            "bssid_from":   prev_bssid,
                            "bssid_to":     curr_bssid,
                            "channel_from": prev_ch,
                            "channel_to":   curr_ch,
                            "wifi":         wifi,
                            "bss_scan":     bss,
                        })

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

                bssid_at_disconnect = prev_wifi.get("bssid", "").lower()
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
