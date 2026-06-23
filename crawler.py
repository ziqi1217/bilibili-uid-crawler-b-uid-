#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
B站UID极简爬虫 v6 - 多线程极速版
特性：多线程并发 / Ctrl+C保护 / 自动保存 / 断点续传 / 用时统计
"""

import re, time, random, csv, json, sys, signal, threading, queue as qmod
from pathlib import Path
import requests as req_lib

# ===== 配置区 =====
API_URL    = "https://api.bilibili.com/x/space/acc/info"
TIMEOUT    = 10
DELAY_MIN  = 0.08    # 每次请求最小延迟（秒）
DELAY_MAX  = 0.20    # 最大延迟
THREADS    = 8       # 并发线程数（想更快改这里）
AUTO_SAVE_EVERY = 15 # 找到N个自动保存一次

UA_POOL = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 Chrome/124.0.0.0 Safari/605.1",
]

# ===== 全局状态 =====
results     = []
found_lock  = threading.Lock()
scanned     = 0
scan_lock   = threading.Lock()
total_uids  = 0
running     = True
cookie_str  = ""
out_dir     = Path("bilibili_results")
progress_file = out_dir / "_progress.json"
start_time  = 0

# ===== 工具函数 =====
def has_chinese(text: str) -> bool:
    return bool(re.search(r'[\u4e00-\u9fff]', text))

def is_target(u: dict) -> bool:
    return (u.get('level', -1) == 0 and
            u.get('following', -1) == 0 and
            u.get('follower', -1) == 0 and
            not has_chinese(u.get('name', '')))

def fetch(uid: int):
    headers = {
        "User-Agent": random.choice(UA_POOL),
        "Referer": "https://www.bilibili.com/",
    }
    if cookie_str:
        headers["Cookie"] = cookie_str
    for attempt in range(3):
        try:
            r = req_lib.get(API_URL, params={"mid": uid}, headers=headers, timeout=TIMEOUT)
            d = r.json()
            if d.get("code") == 0:
                info = d["data"]
                return {
                    "uid": uid, "name": info.get("name", ""),
                    "level": info.get("level", 0),
                    "following": info.get("following", 0),
                    "follower": info.get("follower", 0),
                    "sign": info.get("sign", ""),
                    "jointime": info.get("jointime", 0),
                }
            elif d.get("code") in (-352, -799):
                time.sleep(random.uniform(2, 4))
                continue
            else:
                return None
        except Exception as e:
            if attempt < 2:
                time.sleep(random.uniform(1, 2))
                continue
            return None
    return None

def gen_uids(prefix: str, suffix: str, uid_len: int = 7):
    if not prefix and not suffix:
        raise ValueError("前缀和后缀至少填一个")
    if uid_len < 1 or uid_len > 10:
        raise ValueError("UID位数须在1~10之间")
    results = []
    if prefix and not suffix:
        pad = uid_len - len(prefix)
        if pad < 0:
            raise ValueError(f"前缀长度{len(prefix)}超过了UID位数{uid_len}")
        for i in range(10 ** pad):
            results.append(int(prefix + str(i).zfill(pad)))
    elif suffix and not prefix:
        pad = uid_len - len(suffix)
        if pad < 0:
            raise ValueError(f"后缀长度{len(suffix)}超过了UID位数{uid_len}")
        for i in range(10 ** pad):
            results.append(int(str(i).zfill(pad) + suffix))
    else:
        pre_pad = uid_len - len(prefix) - len(suffix)
        if pre_pad < 0:
            raise ValueError(f"前缀+后缀总长{len(prefix)+len(suffix)}超过了UID位数{uid_len}")
        for i in range(10 ** pre_pad):
            results.append(int(prefix + str(i).zfill(pre_pad) + suffix))
    return results

# ===== 保存 =====
def save_results(label=""):
    with found_lock:
        local = list(results)
    if not local:
        return
    ts = int(time.time())
    csv_path  = out_dir / f"result_{ts}{label}.csv"
    json_path = out_dir / f"result_{ts}{label}.json"

    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["uid","name","level","following","follower","sign","jointime"])
        for u in local:
            w.writerow([u["uid"],u["name"],u["level"],u["following"],u["follower"],u["sign"],u["jointime"]])

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(local, f, ensure_ascii=False, indent=2)

    print(f"\n  已保存 {len(local)} 条 → {csv_path.name}")

def save_progress(scanned_set):
    out_dir.mkdir(exist_ok=True)
    with open(progress_file, "w") as f:
        json.dump({"scanned": sorted(scanned_set), "found_count": len(results)}, f)

def load_progress():
    if progress_file.exists():
        with open(progress_file) as f:
            d = json.load(f)
        return set(d.get("scanned", [])), d.get("found_count", 0)
    return set(), 0

def format_time(sec):
    h, r = divmod(int(sec), 3600)
    m, s = divmod(r, 60)
    if h > 0:
        return f"{h}时{m}分{s}秒"
    elif m > 0:
        return f"{m}分{s}秒"
    else:
        return f"{s}秒"

# ===== 线程工作函数 =====
def worker(uid_queue: qmod.Queue, thread_id: int):
    global scanned
    while running:
        try:
            uid = uid_queue.get(timeout=1)
        except qmod.Empty:
            break

        user = fetch(uid)

        with scan_lock:
            scanned += 1

        if user and is_target(user):
            with found_lock:
                results.append(user)
                n = len(results)
            print(f"  [T{thread_id}] #{n}: UID={user['uid']} 昵称={user['name']}")

        uid_queue.task_done()

        # 随机小延迟防封
        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

# ===== Ctrl+C =====
def sigint_handler(signum, frame):
    global running
    running = False
    print("\n\n收到中断，正在保存结果...")
    save_results("_中断保存")
    print("保存完成。程序退出。")
    sys.exit(0)

signal.signal(signal.SIGINT, sigint_handler)

# ===== 主程序 =====
def main():
    global total_uids, running, cookie_str, start_time, results, scanned

    print("=" * 55)
    print("  B站UID爬虫 v7 - 多线程极速版")
    print(f"  默认 {THREADS} 线程并发 | Ctrl+C 安全退出")
    print("=" * 55)

    uid_len_input = input("\nUID位数（7/8/9/10，回车=7）: ").strip()
    uid_len = 7
    if uid_len_input.isdigit():
        uid_len = max(1, min(10, int(uid_len_input)))

    prefix = input("UID前缀（可不填）: ").strip()
    suffix = input("UID后缀（可不填）: ").strip()
    cookie = input("B站Cookie SESSDATA值（不填则慢）: ").strip()

    if cookie and "=" not in cookie:
        cookie = f"SESSDATA={cookie}"
    cookie_str = cookie

    # 可选自定义线程数
    t_input = input(f"线程数（回车=默认{THREADS}）: ").strip()
    n_threads = THREADS
    if t_input.isdigit():
        n_threads = max(1, min(30, int(t_input)))

    try:
        uid_list = gen_uids(prefix, suffix, uid_len)
    except ValueError as e:
        print(f"错误：{e}")
        input("按回车退出...")
        return

    if len(uid_list) > 500000:
        print(f"\n⚠ 警告：将扫描 {len(uid_list):,} 个UID，可能耗时较长")
        r = input("确认继续？(y/n): ").strip().lower()
        if r != "y":
            return

    # 断点续传
    scanned_set, prev_found = load_progress()
    if scanned_set:
        print(f"\n发现上次未完成任务（已扫描 {len(scanned_set)} 个，找到 {prev_found} 个）")
        r = input("跳过已扫描继续？(y/n): ").strip().lower()
        if r == "y":
            uid_list = [u for u in uid_list if u not in scanned_set]
            print(f"  剩余 {len(uid_list):,} 个UID")
        else:
            scanned_set.clear()
            prev_found = 0

    total_uids = len(uid_list)
    out_dir.mkdir(exist_ok=True)

    print(f"\n{'='*55}")
    print(f"开始扫描 {total_uids:,} 个UID | {uid_len}位数 | {n_threads}线程 | Cookie={'ON' if cookie_str else 'OFF'}")
    print(f"提示：按 Ctrl+C 随时中断并自动保存")
    print(f"{'='*55}")

    start_time = time.time()
    last_t = start_time
    last_save = 0
    found = 0

    # 建任务队列
    task_q = qmod.Queue()
    for u in uid_list:
        task_q.put(u)

    # 启动线程
    threads = []
    for i in range(n_threads):
        t = threading.Thread(target=worker, args=(task_q, i+1), daemon=True)
        t.start()
        threads.append(t)

    # 进度显示循环
    while True:
        # 检查是否完成
        with scan_lock:
            s = scanned

        if s >= total_uids or not any(t.is_alive() for t in threads):
            break

        now = time.time()
        if now - last_t >= 1.5:
            elapsed = now - start_time
            speed = s / elapsed if elapsed > 0 else 0
            pct = s / total_uids * 100
            remain = (total_uids - s) / speed if speed > 0 else 0

            bar_len = 40
            fill = int(pct / 100 * bar_len)
            bar = "█" * fill + "░" * (bar_len - fill)

            with found_lock:
                f = len(results)

            print(f"  [{bar}] {pct:.1f}% | {s:,}/{total_uids:,} | "
                  f"{speed:.1f}/s | 找到{f} | 剩余{format_time(remain)}")
            last_t = now

            # 定期保存
            if f >= AUTO_SAVE_EVERY and f - last_save >= AUTO_SAVE_EVERY:
                save_results()
                save_progress(set())  # 不做断点了，直接存结果
                last_save = f

        time.sleep(0.3)

    # 等所有线程结束
    for t in threads:
        t.join(timeout=5)

    # 最终保存
    save_results()
    save_progress(set())

    elapsed = time.time() - start_time
    avg_spd = scanned / elapsed if elapsed > 0 else 0

    print(f"\n{'='*55}")
    print(f"扫描完成！")
    print(f"  用时：{format_time(elapsed)}")
    print(f"  扫描：{scanned:,} 个UID")
    print(f"  找到：{len(results)} 个目标账号")
    print(f"  平均速度：{avg_spd:.1f} uid/s")
    print(f"  结果保存在：{out_dir}")
    print(f"{'='*55}")
    input("\n按回车退出...")

if __name__ == "__main__":
    main()
