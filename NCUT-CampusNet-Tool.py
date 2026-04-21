'''
===================== 简介 =====================
为 NCUT-AUTO 校园网环境开发，支持：
- 开机时自动抓取校园网流量与网络状态，并测速
- 记录使用数据、生成月度报告，进行流量异常检测

储存检测日志和月度报告的工作区：
默认为与用户文档文件夹中的NCUT_Campus_Network_Log文件夹

使用方式：
Win+R，输入shell:startup（自启动文件夹）
将本文件放入
开机时稍加等待

环境配置：
- 必须拥有Python环境
- 电脑里必须有Edge（应该都有吧）
- 连接了校园网
- 安装了这些库：
pip install selenium
pip install plyer
使用打包exe版本可以免于环境配置！
'''

import time
import os
import re
import subprocess
import urllib.request
import ssl
import socket
import datetime
import statistics
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from plyer import notification

# ===================== 配置项 =====================
MAX_RETRY = 5
RETRY_INTERVAL = 5
TARGET_SSID = "NCUT-AUTO"
CAMPUS_URL = "https://ip.ncut.edu.cn/srun_portal_success?ac_id=1&theme=pro"
CAMPUS_HOST = "ip.ncut.edu.cn"
EXTERNAL_TEST_HOST = "223.5.5.5" #阿里云公共DNS，用于测速
NOTICE_TIMEOUT = 15
LOW_FLOW_THRESHOLD_GB = 10.0
WORK_DIR_NAME = "NCUT_Campus_Network_Log"
PING_COUNT = 10

# 异常检测配置（程序不可能实时检测异常，也不可能处置异常，仅能在生成报告时从统计学角度尝试分析异常！）
ANOMALY_MAD_MULTIPLIER = 3.0
MIN_RECORDS_FOR_ANOMALY = 3
ABSOLUTE_DAILY_THRESHOLD_GB = 15.0
SAFE_DAILY_FLOOR_GB = 1.0
# ==================================================

ssl._create_default_https_context = ssl._create_unverified_context

# ===================== 路径与日志系统 =====================

def get_work_directory():
    doc_path = os.path.expanduser("~/Documents")
    work_path = os.path.join(doc_path, WORK_DIR_NAME)
    if not os.path.exists(work_path):
        os.makedirs(work_path)
    return work_path

def get_log_file_path(date=None):
    work_path = get_work_directory()
    if date is None:
        date = datetime.datetime.now()
    date_str = date.strftime("%Y-%m")
    return os.path.join(work_path, f"network_log_{date_str}.txt")

def append_log(content):
    log_path = get_log_file_path()
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {content}\n")
    except Exception as e:
        print(f"日志写入失败: {e}")

def check_and_generate_monthly_report():
    now = datetime.datetime.now()
    work_path = get_work_directory()
    
    last_month = now.replace(day=1) - datetime.timedelta(days=1)
    last_month_str = last_month.strftime("%Y-%m")
    report_filename = f"Report_{last_month_str}.txt"
    report_path = os.path.join(work_path, report_filename)
    
    if os.path.exists(report_path):
        return False, ""
    
    log_path = get_log_file_path(last_month)
    
    summary_content = f"=== 北方工业大学校园网月度报告 ({last_month_str}) ===\n"
    summary_content += f"生成时间: {now.strftime('%Y-%m-%d %H:%M:%S')}\n"
    summary_content += "----------------------------------------\n\n"
    
    report_notification_msg = ""
    has_anomaly = False
    anomalies = []
    
    if os.path.exists(log_path):
        summary_content += "成功...\n\n"
        
        records = []
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                for line in f:
                    log_pattern = re.compile(
                        r"\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\] "
                        r"用户:([^|]+) \| "
                        r".*已用流量:([\d.]+)\s*GB"
                    )
                    
                    match = log_pattern.search(line)
                    if match:
                        dt_str, username, flow_str = match.groups()
                        try:
                            dt = datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
                            flow = float(flow_str)
                            records.append({
                                "datetime": dt,
                                "date": dt.date(),
                                "username": username,
                                "flow": flow
                            })
                        except:
                            continue
        except Exception as e:
            summary_content += f"日志读取失败: {str(e)}\n"
            return True, "报告生成失败，无法读取日志"
        
        if not records:
            summary_content += "未找到有效的流量记录\n"
            return True, "报告已生成，但无有效记录"
        
        total_records = len(records)
        first_record = min(records, key=lambda x: x["datetime"])
        last_record = max(records, key=lambda x: x["datetime"])
        max_flow = max(r["flow"] for r in records)
        
        daily_records = {}
        for r in sorted(records, key=lambda x: x["datetime"]):
            daily_records[r["date"]] = r
        
        summary_content += "使用记录:\n"
        summary_content += f"{'日期':<12} | {'已用流量':>10}\n"
        summary_content += "-" * 28 + "\n"
        
        prev_date = None
        for date in sorted(daily_records.keys()):
            r = daily_records[date]
            if prev_date is None or (date - prev_date).days >= 1:
                summary_content += f"{date.strftime('%Y-%m-%d'):<12} | {r['flow']:>8.2f} GB\n"
                prev_date = date
        
        # ===================== 异常检测模块 =====================
        daily_dates = sorted(daily_records.keys())
        increments = []
        
        if len(daily_dates) >= MIN_RECORDS_FOR_ANOMALY:
            for i in range(1, len(daily_dates)):
                prev_date = daily_dates[i-1]
                curr_date = daily_dates[i]
                prev_flow = daily_records[prev_date]["flow"]
                curr_flow = daily_records[curr_date]["flow"]
                days_diff = (curr_date - prev_date).days
                
                if days_diff > 0 and curr_flow > prev_flow:
                    total_inc = curr_flow - prev_flow
                    daily_avg = total_inc / days_diff
                    increments.append({
                        "start": prev_date,
                        "end": curr_date,
                        "days": days_diff,
                        "total": total_inc,
                        "avg": daily_avg
                    })
            
            if increments:
                daily_avgs = [inc["avg"] for inc in increments]
                median_avg = statistics.median(daily_avgs)
                mad = statistics.median([abs(x - median_avg) for x in daily_avgs])
                threshold_avg = median_avg + ANOMALY_MAD_MULTIPLIER * mad
                
                anomalies = []
                for inc in increments:
                    if inc["avg"] < SAFE_DAILY_FLOOR_GB:
                        continue
                    # ================================
                    
                    is_anomaly = False
                    reason = []
                    
                    if inc["avg"] > threshold_avg:
                        is_anomaly = True
                        reason.append(f"日均({inc['avg']:.1f}GB)远超正常")
                    
                    if inc["avg"] > ABSOLUTE_DAILY_THRESHOLD_GB:
                        is_anomaly = True
                        reason.append(f"超单日阈值{ABSOLUTE_DAILY_THRESHOLD_GB}GB")
                    
                    if is_anomaly:
                        inc["reason"] = "；".join(reason)
                        anomalies.append(inc)
                
                summary_content += f"  日均中位数: {median_avg:.2f} GB/天\n"
                
                if not anomalies:
                    summary_content += "  未检测到异常流量消耗\n"
                else:
                    has_anomaly = True
                    summary_content += f"  ⚠️ 检测到 {len(anomalies)} 次异常流量消耗:\n"
                    summary_content += f"{'时间段':<22} | {'间隔':>6} | {'总消耗':>10} | {'日均消耗':>12} | 异常原因\n"
                    summary_content += "-" * 90 + "\n"
                    for anom in anomalies:
                        start_str = anom["start"].strftime("%Y-%m-%d")
                        end_str = anom["end"].strftime("%Y-%m-%d")
                        summary_content += (
                            f"{start_str} ~ {end_str:<10} | "
                            f"{anom['days']:>4}天 | "
                            f"{anom['total']:>10.2f} GB | "
                            f"{anom['avg']:>12.2f} GB/天 | "
                            f"{anom['reason']}\n"
                        )
            else:
                summary_content += "  未找到有效的流量增量数据\n"
        else:
            summary_content += f"  记录不足（需≥{MIN_RECORDS_FOR_ANOMALY}条），无法进行统计检测\n"
        
        summary_content += "\n\n月度综合统计:\n"
        summary_content += f"  • 总检测次数: {total_records} 次\n"
        summary_content += f"  • 首次记录: {first_record['datetime'].strftime('%Y-%m-%d %H:%M')}\n"
        summary_content += f"  • 末次记录: {last_record['datetime'].strftime('%Y-%m-%d %H:%M')}\n"
        summary_content += f"  • 本月累计使用: ~{max_flow:.2f} GB\n"
        
        report_notification_msg = f"生成了 {last_month_str} 月度报告\n"
        report_notification_msg += f"本月累计使用: ~{max_flow:.2f} GB\n"
        if has_anomaly:
            report_notification_msg += f"检测到 {len(anomalies)} 次流量异常！\n"
            report_notification_msg += f"最高日均消耗: {max(a['avg'] for a in anomalies):.1f} GB/天\n"
        report_notification_msg += f"报告路径: {report_path}"
        
    else:
        summary_content += f"未找到上月日志文件\n"
        summary_content += f"   预期路径: {log_path}\n"
        report_notification_msg = "尝试生成报告，但未找到上月日志文件"
    
    try:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(summary_content)
        append_log(f"系统: 已生成 {last_month_str} 月度报告，异常检测结果: {'发现异常' if has_anomaly else '正常'}")
        return True, report_notification_msg
    except Exception as e:
        print(f"报告生成失败: {e}")
        return False, ""

# ===================== 网络检测功能 =====================

def check_network_available() -> bool:
    targets = [
        (CAMPUS_HOST, 443, "socket"),
        (EXTERNAL_TEST_HOST, 53, "socket"),
        ("http://www.baidu.com", None, "http")
    ]
    for host, port, target_type in targets:
        try:
            if target_type == "socket":
                sock = socket.create_connection((host, port), timeout=5)
                sock.close()
                return True
            elif target_type == "http":
                urllib.request.urlopen(host, timeout=5)
                return True
        except:
            continue
    return False

def get_current_wifi_ssid() -> str:
    for attempt in range(2):
        try:
            startupinfo = None
            if os.name == 'nt':
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE

            result = subprocess.run(
                ["netsh", "wlan", "show", "interfaces"],
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore',
                startupinfo=startupinfo
            )
            ssid_match = re.search(r"^\s*SSID\s*[:：]\s*(.+)$", result.stdout, re.MULTILINE)
            if ssid_match:
                return ssid_match.group(1).strip()
        except:
            pass
        if attempt == 0: time.sleep(1)
    return ""

def check_campus_network_reachable() -> bool:
    for attempt in range(2):
        try:
            sock = socket.create_connection((CAMPUS_HOST, 443), timeout=5)
            sock.close()
            return True
        except:
            pass
        if attempt == 0: time.sleep(1)
    return False

def ping_host(host, count=PING_COUNT):
    try:
        startupinfo = None
        if os.name == 'nt':
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = subprocess.SW_HIDE
            cmd = ["ping", "-n", str(count), "-w", "1000", host]
        else:
            cmd = ["ping", "-c", str(count), "-W", "1", host]

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='gbk' if os.name == 'nt' else 'utf-8',
            errors='ignore',
            startupinfo=startupinfo
        )
        
        output = result.stdout
        if not output:
            return -1.0, 100.0

        loss_patterns = [r"(\d+)% 丢失", r"(\d+)% loss", r"(\d+)% packet loss"]
        loss = 100.0
        for pattern in loss_patterns:
            match = re.search(pattern, output, re.IGNORECASE)
            if match:
                loss = float(match.group(1))
                break

        time_matches = re.findall(r"(?:时间|time)[=<]\s*(\d+(?:\.\d+)?)ms", output, re.IGNORECASE)
        if time_matches:
            times = [float(t) for t in time_matches]
            avg_latency = sum(times) / len(times)
            return avg_latency, loss
        
        return -1.0, loss

    except Exception as e:
        return -1.0, 100.0

def get_network_quality():
    internal_latency, internal_loss = ping_host(CAMPUS_HOST, count=5)
    external_latency, external_loss = ping_host(EXTERNAL_TEST_HOST, count=PING_COUNT)
    return {
        "internal_latency": internal_latency,
        "internal_loss": internal_loss,
        "external_latency": external_latency,
        "external_loss": external_loss
    }

# ===================== 核心业务逻辑 =====================

def parse_flow_to_gb(flow_text):
    if not flow_text or flow_text == "N/A":
        return 0.0
    flow_text = flow_text.strip().upper()
    try:
        if "GB" in flow_text:
            return float(flow_text.replace("GB", "").strip())
        elif "MB" in flow_text:
            return float(flow_text.replace("MB", "").strip()) / 1024.0
        else:
            return float(flow_text) / 1024.0
    except:
        return 0.0

def get_campus_info():
    edge_options = webdriver.EdgeOptions()
    edge_options.add_argument("--headless=new")
    edge_options.add_argument("--disable-gpu")
    edge_options.add_argument("--window-size=1920,1080")
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument("--disable-extensions")
    edge_options.add_argument("--log-level=3")
    edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])

    from selenium.webdriver.edge.service import Service
    service = Service() 
    
    driver = None
    data = {
        "success": False,
        "username": "",
        "used_time": "",
        "used_flow": "",
        "used_flow_gb": 0.0,
        "remain_flow": "",
        "remain_flow_gb": 0.0,
        "total_flow_gb": 0.0
    }

    try:
        driver = webdriver.Edge(service=service, options=edge_options)
        driver = webdriver.Edge(options=edge_options)
        driver.get(CAMPUS_URL)
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.ID, "remain-bytes")))
        
        def safe_get(xpath):
            try:
                return driver.find_element(By.XPATH, xpath).text.strip()
            except:
                return "N/A"

        data["username"] = safe_get('//*[@id="username"]')
        data["used_time"] = safe_get('//*[@id="used-time"]')
        data["used_flow"] = safe_get('//*[@id="used-flow"]')
        data["remain_flow"] = safe_get('//*[@id="remain-bytes"]')
        
        data["used_flow_gb"] = parse_flow_to_gb(data["used_flow"])
        data["remain_flow_gb"] = parse_flow_to_gb(data["remain_flow"])
        
        if data["used_flow_gb"] > 0 and data["remain_flow_gb"] > 0:
            total = data["used_flow_gb"] + data["remain_flow_gb"]
            data["total_flow_gb"] = round(total / 10) * 10
        
        data["success"] = True

    except Exception as e:
        append_log(f"错误: Selenium 抓取失败 - {str(e)}")
    finally:
        if driver:
            driver.quit()
    
    return data

def main():
    report_generated, report_msg = check_and_generate_monthly_report()
    
    network_ok = False
    for _ in range(MAX_RETRY):
        if check_network_available():
            network_ok = True
            break
        time.sleep(RETRY_INTERVAL)
    
    if not network_ok:
        msg = '网络未连接，请检查网络后重试'
        notification.notify(title='网络异常', message=msg, app_name='校园网流量助手', timeout=NOTICE_TIMEOUT)
        append_log("状态: 网络连接失败")
        return

    current_ssid = get_current_wifi_ssid()
    is_campus_network = False

    if current_ssid == TARGET_SSID:
        is_campus_network = True
    else:
        if check_campus_network_reachable():
            is_campus_network = True
    
    if not is_campus_network:
        msg = f'当前非校园网状态 (SSID: {current_ssid})\n仅支持NCUT-AUTO查询'
        notification.notify(title='非校园网环境', message=msg, app_name='校园网流量助手', timeout=NOTICE_TIMEOUT)
        append_log(f"状态: 非校园网环境 (SSID: {current_ssid})")
        return

    quality = get_network_quality()
    info_data = get_campus_info()
    
    if info_data["success"]:
        if info_data["total_flow_gb"] > 0:
            flow_display = f"{info_data['remain_flow']} / {info_data['total_flow_gb']}GB"
        else:
            flow_display = info_data["remain_flow"]
        
        latency_display = ""
        if quality["external_latency"] > 0:
            latency_display = f"公网延迟: {quality['external_latency']:.1f}ms"
            if quality["external_loss"] > 0:
                latency_display += f" (丢包 {quality['external_loss']:.0f}%)"
        
        log_path = get_log_file_path()
        
        message = f"  流量: {flow_display}\n"
        if latency_display:
            message += f" {latency_display}\n"
        message += " 日志已生成"
        
        if report_generated and report_msg:
            message = report_msg + "\n\n" + message
        
        if 0 < info_data["remain_flow_gb"] < LOW_FLOW_THRESHOLD_GB:
            title = '⚠️ 校园网流量预警'
            warning_msg = f"剩余流量不足 {LOW_FLOW_THRESHOLD_GB}GB！\n\n"
            if info_data["total_flow_gb"] == 60:
                warning_msg += "可以在企业微信-服务大厅-上网流量申请中申请\n\n"
            message = warning_msg + message
        else:
            title = '北方工业大学校园网'

        notification.notify(
            title=title,
            message=message,
            app_name='校园网流量助手',
            timeout=NOTICE_TIMEOUT
        )
    else:
        notification.notify(
            title='流量查询失败',
            message='校园网页面加载失败，请检查登录状态',
            app_name='校园网流量助手',
            timeout=NOTICE_TIMEOUT
        )

    log_msg = (
        f"用户:{info_data['username']} | "
        f"已用时长:{info_data['used_time']} | "
        f"已用流量:{info_data['used_flow']} | "
        f"剩余流量:{info_data['remain_flow']} | "
        f"总流量:{info_data['total_flow_gb']}GB | "
        f"内网延迟:{quality['internal_latency']:.1f}ms | "
        f"内网丢包:{quality['internal_loss']:.0f}% | "
        f"公网延迟:{quality['external_latency']:.1f}ms | "
        f"公网丢包:{quality['external_loss']:.0f}%"
    )
    
    append_log(log_msg)

if __name__ == "__main__":
    main()
