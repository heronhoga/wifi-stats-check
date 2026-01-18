import subprocess
import re
import statistics
import speedtest
from openpyxl import Workbook
from datetime import datetime

def run_ping(host="8.8.8.8", count=20):
    cmd = ["ping", host, "-n", str(count)]
    result = subprocess.run(cmd, capture_output=True, text=True)
    output = result.stdout

    times = [int(t) for t in re.findall(r'time[=<](\d+)ms', output)]
    packet_loss_match = re.search(r'(\d+)% loss', output)
    packet_loss = int(packet_loss_match.group(1)) if packet_loss_match else None

    avg = statistics.mean(times) if times else None
    jitter = statistics.stdev(times) if len(times) > 1 else 0

    return avg, jitter, packet_loss


def run_speedtest():
    st = speedtest.Speedtest()
    st.get_best_server()

    download = st.download() / 1_000_000
    upload = st.upload() / 1_000_000
    ping = st.results.ping

    return download, upload, ping


def get_wifi_info():
    cmd = ["netsh", "wlan", "show", "interfaces"]
    result = subprocess.run(cmd, capture_output=True, text=True)

    signal = None
    ssid = None

    for line in result.stdout.splitlines():
        if "SSID" in line and "BSSID" not in line:
            ssid = line.split(":")[1].strip()
        if "Signal" in line:
            signal = line.split(":")[1].strip()

    return ssid, signal


def save_to_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Network Test Results"

    headers = [
        "Run",
        "Avg Latency (ms)",
        "Jitter (ms)",
        "Packet Loss (%)",
        "Download Speed (Mbps)",
        "Upload Speed (Mbps)",
        "Speedtest Ping (ms)",
    ]

    ws.append(headers)

    for row in rows:
        ws.append(row)

    filename = f"network_test_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb.save(filename)

    print(f"Excel report saved as: {filename}")



if __name__ == "__main__":
    print("Running network diagnostics (5 runs)...")

    RUNS = 5
    excel_rows = []

    for i in range(RUNS):
        print(f"Run {i + 1}/{RUNS}")

        avg_latency, jitter, packet_loss = run_ping()
        download, upload, speedtest_ping = run_speedtest()

        excel_rows.append([
            i + 1,
            round(avg_latency, 2) if avg_latency else "N/A",
            round(jitter, 2),
            packet_loss,
            round(download, 2),
            round(upload, 2),
            round(speedtest_ping, 2),
        ])

    save_to_excel(excel_rows)


