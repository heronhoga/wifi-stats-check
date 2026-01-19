import subprocess
import re
import statistics
import speedtest
from openpyxl import Workbook
from datetime import datetime
import sys
import time


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def run_ping(host="8.8.8.8", count=20):
    log("PING: starting")
    cmd = ["ping", "-c", str(count), host]
    log(f"PING: running command -> {' '.join(cmd)}")

    start = time.time()
    result = subprocess.run(cmd, capture_output=True, text=True)
    log(f"PING: command finished in {round(time.time() - start, 2)}s")

    output = result.stdout

    times = [int(t) for t in re.findall(r'time=(\d+)', output)]
    packet_loss_match = re.search(r'(\d+)% packet loss', output)
    packet_loss = int(packet_loss_match.group(1)) if packet_loss_match else None

    avg = statistics.mean(times) if times else None
    jitter = statistics.stdev(times) if len(times) > 1 else 0

    log(f"PING: avg={avg}ms jitter={jitter}ms loss={packet_loss}%")

    return avg, jitter, packet_loss


def run_speedtest():
    log("SPEEDTEST: initializing Speedtest()")
    start_total = time.time()

    st = speedtest.Speedtest(timeout=10)

    log("SPEEDTEST: fetching server list")
    st.get_servers()

    log("SPEEDTEST: selecting best server")
    st.get_best_server()

    log("SPEEDTEST: starting download test")
    start = time.time()
    download = st.download() / 1_000_000
    log(f"SPEEDTEST: download finished in {round(time.time() - start, 2)}s")

    log("SPEEDTEST: starting upload test")
    start = time.time()
    upload = st.upload() / 1_000_000
    log(f"SPEEDTEST: upload finished in {round(time.time() - start, 2)}s")

    ping = st.results.ping

    log(f"SPEEDTEST: done in {round(time.time() - start_total, 2)}s")
    log(f"SPEEDTEST: download={round(download,2)}Mbps upload={round(upload,2)}Mbps ping={round(ping,2)}ms")

    return download, upload, ping


def save_to_excel(rows):
    log("EXCEL: creating workbook")
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
    log(f"EXCEL: saved file -> {filename}")


if __name__ == "__main__":
    log("PROGRAM: starting network diagnostics")

    RUNS = 5
    excel_rows = []

    for i in range(RUNS):
        log(f"PROGRAM: ===== RUN {i + 1}/{RUNS} =====")

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

        log(f"PROGRAM: run {i + 1} completed")

    save_to_excel(excel_rows)
    log("PROGRAM: all runs finished")
