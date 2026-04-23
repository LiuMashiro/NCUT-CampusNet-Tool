# 北方工业大学校园网流量助手 v1.5
# 项目地址: https://github.com/LiuMashiro/NCUT-CampusNet-Tool
# 适用于 NCUT-AUTO 校园网，支持流量查询、网络检测、低流量告警、月度报告生成

import time
import os
import re
import subprocess
import socket
import datetime
import statistics
import json
import sys
import traceback
from typing import Dict, Optional, Tuple, List
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from plyer import notification

# ===================== 默认配置（将被配置文件覆盖） =====================
DEFAULT_CONFIG = {
    "MAX_RETRY": 5,
    "RETRY_INTERVAL": 3,
    "TARGET_SSID": "NCUT-AUTO",
    "CAMPUS_URL": "https://ip.ncut.edu.cn/srun_portal_success?ac_id=1&theme=pro",
    "CAMPUS_HOST": "ip.ncut.edu.cn",
    "EXTERNAL_TEST_HOST": "223.5.5.5",
    "NOTICE_TIMEOUT": 15,
    "LOW_FLOW_THRESHOLD_GB": 10.0,
    "WORK_DIR_NAME": "NCUT_Campus_Network_Log",
    "PING_COUNT": 10,
    
    # 功能开关
    "LOG_ENABLED": True,
    "DEBUG_MODE": False,
    "SPEED_TEST_ENABLED": True,
    
    # 异常检测配置
    "ANOMALY_MAD_MULTIPLIER": 3.0,
    "MIN_RECORDS_FOR_ANOMALY": 3,
    "ABSOLUTE_DAILY_THRESHOLD_GB": 15.0,
    "SAFE_DAILY_FLOOR_GB": 1.5,
    
    # 报告配置
    "OPEN_REPORT_AFTER_GENERATE": True
}

# 配置文件说明
CONFIG_COMMENTS = """# 北方工业大学校园网流量助手 配置文件 v1.5
# 项目地址: https://github.com/LiuMashiro/NCUT-CampusNet-Tool
# 修改此文件后重启程序生效
# 如配置文件损坏，删除后重新运行程序将自动生成默认配置

# 基础配置
# MAX_RETRY: 网络连接失败重试次数
# RETRY_INTERVAL: 重试间隔(秒)
# TARGET_SSID: 校园网WiFi名称
# CAMPUS_URL: 校园网认证成功页面地址
# CAMPUS_HOST: 校园网服务器地址
# EXTERNAL_TEST_HOST: 公网连通性测试地址(默认阿里云DNS)
# NOTICE_TIMEOUT: 系统通知显示时长(秒)
# LOW_FLOW_THRESHOLD_GB: 低流量告警阈值(GB)
# WORK_DIR_NAME: 工作目录名称(位于"文档"下)
# PING_COUNT: 测速时发送的ping包数量

# 功能开关
# LOG_ENABLED: 是否启用日志记录(关闭后不生成日志和月度报告)
# DEBUG_MODE: 调试模式(开启后生成详细错误报告，用于排查问题)
# SPEED_TEST_ENABLED: 是否启用网络测速(关闭后不检测延迟和丢包)

# 异常检测配置
# ANOMALY_MAD_MULTIPLIER: 异常检测中位数绝对偏差倍数
# MIN_RECORDS_FOR_ANOMALY: 异常检测所需最少记录数
# ABSOLUTE_DAILY_THRESHOLD_GB: 单日流量绝对阈值(超过即判定为异常)
# SAFE_DAILY_FLOOR_GB: 安全流量下限(低于此值不判定为异常)

# 报告配置
# OPEN_REPORT_AFTER_GENERATE: 生成月度报告后是否自动打开
"""
# ======================================================================

# Windows专用子进程配置（全局复用，避免重复创建）
_STARTUPINFO = subprocess.STARTUPINFO()
_STARTUPINFO.dwFlags |= subprocess.STARTF_USESHOWWINDOW
_STARTUPINFO.wShowWindow = subprocess.SW_HIDE

# 全局配置变量
config = DEFAULT_CONFIG.copy()

# ===================== 配置文件管理 =====================

def load_config() -> None:
    """加载配置文件，不存在则创建默认配置"""
    global config
    work_path = get_work_directory()
    config_path = os.path.join(work_path, "config.ini")
    
    if not os.path.exists(config_path):
        # 生成带注释的配置文件
        with open(config_path, "w", encoding="utf-8") as f:
            f.write(CONFIG_COMMENTS)
            f.write("\n")
            json.dump(DEFAULT_CONFIG, f, indent=4, ensure_ascii=False)
        return
    
    try:
        with open(config_path, "r", encoding="utf-8") as f:
            # 跳过注释行
            lines = []
            for line in f:
                stripped = line.strip()
                if not stripped or stripped.startswith("#"):
                    continue
                lines.append(line)
            
            config_content = "".join(lines)
            user_config = json.loads(config_content)
            
            # 合并配置，保留默认值中用户未设置的项
            config.update(user_config)
            
    except Exception as e:
        config = DEFAULT_CONFIG.copy()

# ===================== 路径与日志系统 =====================

def get_work_directory() -> str:
    """获取工作目录，确保存在，首次创建返回True"""
    doc_path = os.path.expanduser("~/Documents")
    work_path = os.path.abspath(os.path.join(doc_path, config["WORK_DIR_NAME"]))
    
    if not os.path.exists(work_path):
        os.makedirs(work_path, exist_ok=True)
        # 强制自动打开工作目录,首次运行提示通知
        try:
            os.startfile(work_path)
            notification.notify(
                title='程序首次运行提示',
                message='程序正在后台静默运行，需等待延迟测试、开机配置等流程，不会立即显示通知，这是正常现象。\n若持续超过1分钟无响应，请开启Debug模式排查问题。',
                app_name='校园网流量助手',
                timeout=config["NOTICE_TIMEOUT"]
            )
        except Exception:
            pass
    return work_path

def get_log_file_path(date: Optional[datetime.datetime] = None) -> str:
    """获取指定日期的日志文件路径"""
    work_path = get_work_directory()
    if date is None:
        date = datetime.datetime.now()
    date_str = date.strftime("%Y-%m")
    return os.path.join(work_path, f"network_log_{date_str}.txt")

def append_log(content: str) -> None:
    """追加日志到文件（受LOG_ENABLED控制）"""
    if not config["LOG_ENABLED"]:
        return
    
    log_path = get_log_file_path()
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with open(log_path, "a", encoding="utf-8", newline="") as f:
            f.write(f"[{timestamp}] {content}\n")
    except Exception as e:
        pass

def generate_debug_report(exc: Exception) -> None:
    """生成调试错误报告（仅在DEBUG_MODE开启时）"""
    if not config["DEBUG_MODE"]:
        return
    
    work_path = get_work_directory()
    debug_dir = os.path.join(work_path, "debug")
    os.makedirs(debug_dir, exist_ok=True)
    
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    report_path = os.path.join(debug_dir, f"error_report_{timestamp}.txt")
    
    try:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(f"=== 错误报告 ===\n")
            f.write(f"生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"程序版本: v1.5\n")
            f.write(f"Python版本: {sys.version}\n")
            f.write(f"操作系统: {os.name}\n")
            f.write(f"工作目录: {get_work_directory()}\n")
            f.write("\n=== 异常信息 ===\n")
            f.write(f"异常类型: {type(exc).__name__}\n")
            f.write(f"异常信息: {str(exc)}\n")
            f.write("\n=== 堆栈跟踪 ===\n")
            f.write(traceback.format_exc())
            f.write("\n=== 当前配置 ===\n")
            json.dump(config, f, indent=4, ensure_ascii=False)
    except Exception:
        pass

def check_and_generate_monthly_report() -> Tuple[bool, str]:
    if not config["LOG_ENABLED"]:
        return False, ""
    
    now = datetime.datetime.now()
    work_path = get_work_directory()
    
    # 计算上月时间与文件路径
    last_month = now.replace(day=1) - datetime.timedelta(days=1)
    last_month_str = last_month.strftime("%Y-%m")
    report_filename = f"Report_{last_month_str}.txt"
    report_path = os.path.join(work_path, report_filename)
    log_path = get_log_file_path(last_month)

    # 条件1：报告已存在 → 直接跳过
    if os.path.exists(report_path):
        return False, ""
    
    # 条件2：无上月日志文件 → 跳过生成
    if not os.path.exists(log_path):
        append_log(f"系统: 无{last_month_str}月度日志，跳过报告生成")
        return False, ""

    # 读取日志，校验是否存在有效流量记录
    records: List[Dict] = []
    try:
        with open(log_path, "r", encoding="utf-8") as f:
            log_pattern = re.compile(
                r"\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\] "
                r"用户:([^|]+) \| "
                r".*已用流量:([\d.]+)\s*GB"
            )
            for line in f:
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
                    except (ValueError, TypeError):
                        continue
    except Exception as e:
        append_log(f"错误: 读取上月日志失败，跳过报告生成: {e}")
        return False, ""

    # 条件3：日志存在但无有效流量记录 → 跳过生成
    if not records:
        append_log(f"系统: {last_month_str}日志无有效流量数据，跳过报告生成")
        return False, ""

    # ========== 满足所有条件：开始生成月度报告 ==========
    summary_content = f"=== 北方工业大学校园网月度报告 ({last_month_str}) ===\n"
    summary_content += f"生成时间: {now.strftime('%Y-%m-%d %H:%M:%S')}\n"
    summary_content += f"程序版本: v1.5\n"
    summary_content += "----------------------------------------\n\n"
    
    report_notification_msg = ""
    has_anomaly = False
    anomalies: List[Dict] = []

    total_records = len(records)
    first_record = min(records, key=lambda x: x["datetime"])
    last_record = max(records, key=lambda x: x["datetime"])
    max_flow = max(r["flow"] for r in records)
    
    daily_records: Dict[datetime.date, Dict] = {}
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
    
    # 异常检测模块
    daily_dates = sorted(daily_records.keys())
    increments: List[Dict] = []
    
    if len(daily_dates) >= config["MIN_RECORDS_FOR_ANOMALY"]:
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
            threshold_avg = median_avg + config["ANOMALY_MAD_MULTIPLIER"] * mad
            
            anomalies = []
            for inc in increments:
                if inc["avg"] < config["SAFE_DAILY_FLOOR_GB"]:
                    continue
                is_anomaly = False
                reason = []
                if inc["avg"] > threshold_avg:
                    is_anomaly = True
                    reason.append(f"日均({inc['avg']:.1f}GB)远超正常")
                if inc["avg"] > config["ABSOLUTE_DAILY_THRESHOLD_GB"]:
                    is_anomaly = True
                    reason.append(f"超单日阈值{config['ABSOLUTE_DAILY_THRESHOLD_GB']}GB")
                if is_anomaly:
                    inc["reason"] = "；".join(reason)
                    anomalies.append(inc)
            
            summary_content += f"\n  日均中位数: {median_avg:.2f} GB/天\n"
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
    
    # 写入报告文件
    try:
        with open(report_path, "w", encoding="utf-8", newline="") as f:
            f.write(summary_content)
        append_log(f"系统: 已生成 {last_month_str} 月度报告，异常检测结果: {'发现异常' if has_anomaly else '正常'}")
        # 自动打开报告/工作目录
        if config["OPEN_REPORT_AFTER_GENERATE"]:
            try:
                os.startfile(report_path)
            except Exception:
                os.startfile(work_path)
        return True, report_notification_msg
    except Exception as e:
        append_log(f"错误: 月度报告写入失败 - {e}")
        return False, ""

# ===================== 网络检测功能 =====================

def check_network_available() -> bool:
    """检查网络是否可用"""
    for host, port in [(config["EXTERNAL_TEST_HOST"], 53), (config["CAMPUS_HOST"], 443)]:
        try:
            with socket.create_connection((host, port), timeout=3):
                pass
            return True
        except (socket.timeout, ConnectionRefusedError, OSError):
            continue
    try:
        with socket.create_connection(("www.baidu.com", 80), timeout=3):
            pass
        return True
    except (socket.timeout, ConnectionRefusedError, OSError):
        return False

def get_current_wifi_ssid() -> str:
    """获取当前连接的WiFi SSID"""
    for attempt in range(2):
        try:
            result = subprocess.run(
                ["netsh", "wlan", "show", "interfaces"],
                capture_output=True,
                text=True,
                encoding="gbk",
                errors="ignore",
                startupinfo=_STARTUPINFO,
                timeout=5
            )
            ssid_match = re.search(r"^\s*SSID\s*[:：]\s*(.+)$", result.stdout, re.MULTILINE)
            if ssid_match:
                return ssid_match.group(1).strip()
        except subprocess.TimeoutExpired:
            pass
        if attempt == 0:
            time.sleep(0.5)
    return ""

def check_campus_network_reachable() -> bool:
    """检查校园网主机是否可达"""
    try:
        with socket.create_connection((config["CAMPUS_HOST"], 443), timeout=3):
            pass
        return True
    except (socket.timeout, ConnectionRefusedError, OSError):
        return False

def ping_host(host: str, count: int = None) -> Tuple[float, float]:
    """Ping指定主机，返回平均延迟(ms)和丢包率(%)"""
    if count is None:
        count = config["PING_COUNT"]
    try:
        result = subprocess.run(
            ["ping", "-n", str(count), "-w", "1000", host],
            capture_output=True,
            text=True,
            encoding="gbk",
            errors="ignore",
            startupinfo=_STARTUPINFO,
            timeout=count + 2
        )
        output = result.stdout
        if not output:
            return -1.0, 100.0
        loss_match = re.search(r"(\d+)% 丢失", output, re.IGNORECASE)
        loss = float(loss_match.group(1)) if loss_match else 100.0
        time_matches = re.findall(r"时间[=<]\s*(\d+(?:\.\d+)?)ms", output, re.IGNORECASE)
        if time_matches:
            times = [float(t) for t in time_matches]
            avg_latency = sum(times) / len(times)
            return avg_latency, loss
        return -1.0, loss
    except Exception:
        return -1.0, 100.0

def get_network_quality() -> Dict[str, float]:
    """获取网络质量信息（受SPEED_TEST_ENABLED控制）"""
    if not config["SPEED_TEST_ENABLED"]:
        return {
            "internal_latency": -1.0,
            "internal_loss": -1.0,
            "external_latency": -1.0,
            "external_loss": -1.0
        }
    internal_latency, internal_loss = ping_host(config["CAMPUS_HOST"], count=5)
    external_latency, external_loss = ping_host(config["EXTERNAL_TEST_HOST"])
    return {
        "internal_latency": internal_latency,
        "internal_loss": internal_loss,
        "external_latency": external_latency,
        "external_loss": external_loss
    }

# ===================== 核心业务逻辑 =====================

def parse_flow_to_gb(flow_text: str) -> float:
    """将流量文本转换为GB单位"""
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
    except (ValueError, TypeError): 
        return 0.0

def get_campus_info() -> Dict:
    """获取校园网信息"""
    edge_options = webdriver.EdgeOptions()
    edge_options.add_argument("--headless=new")
    edge_options.add_argument("--disable-gpu")
    edge_options.add_argument("--window-size=1280,720")
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument("--disable-extensions")
    edge_options.add_argument("--disable-dev-shm-usage")
    edge_options.add_argument("--log-level=3")
    edge_options.add_argument("--silent")
    edge_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    edge_options.add_experimental_option("useAutomationExtension", False)

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
        driver = webdriver.Edge(options=edge_options)
        driver.set_page_load_timeout(15)
        driver.get(config["CAMPUS_URL"])
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.ID, "remain-bytes")))
        
        def safe_get(xpath: str) -> str:
            try:
                return driver.find_element(By.XPATH, xpath).text.strip()
            except Exception:
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
        append_log(f"错误: Selenium抓取失败 - {str(e)}")
        if config["DEBUG_MODE"]:
            generate_debug_report(e)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
    return data

def check_startup_location() -> None:
    """检查程序是否在启动文件夹运行，不在则提示"""
    try:
        if getattr(sys, 'frozen', False):
            current_exe = sys.executable
        else:
            current_exe = os.path.abspath(__file__)
        current_dir = os.path.dirname(current_exe)
        startup_dir = os.path.expanduser("~\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup")
        if not os.path.samefile(current_dir, startup_dir):
            notification.notify(
                title='⚠️ 提示',
                message='程序未设置为开机自启动\n建议放入启动文件夹实现开机检测',
                app_name='校园网流量助手',
                timeout=config["NOTICE_TIMEOUT"]
            )
    except Exception:
        pass

def main() -> None:
    """主函数"""
    load_config()
    check_startup_location()
    report_generated, report_msg = check_and_generate_monthly_report()

    network_ok = False
    for _ in range(config["MAX_RETRY"]):
        if check_network_available():
            network_ok = True
            break
        time.sleep(config["RETRY_INTERVAL"])
    
    if not network_ok:
        notification.notify(title='网络异常', message='网络未连接，请检查网络后重试', 
                            app_name='校园网流量助手', timeout=config["NOTICE_TIMEOUT"])
        append_log("状态: 网络连接失败")
        return

    current_ssid = get_current_wifi_ssid()
    is_campus_network = current_ssid == config["TARGET_SSID"] or check_campus_network_reachable()
    if not is_campus_network:
        notification.notify(title='非校园网环境', message=f'当前SSID: {current_ssid}，仅支持NCUT-AUTO查询', 
                            app_name='校园网流量助手', timeout=config["NOTICE_TIMEOUT"])
        append_log(f"状态: 非校园网环境 (SSID: {current_ssid})")
        return

    quality = get_network_quality()
    info_data = get_campus_info()
    if info_data["success"]:
        flow_display = f"{info_data['remain_flow']} / {info_data['total_flow_gb']}GB" if info_data["total_flow_gb"] > 0 else info_data["remain_flow"]
        latency_display = ""
        if config["SPEED_TEST_ENABLED"] and quality["external_latency"] > 0:
            latency_display = f"公网延迟: {quality['external_latency']:.1f}ms"
            if quality["external_loss"] > 0:
                latency_display += f" (丢包 {quality['external_loss']:.0f}%)"
        message = f"  流量: {flow_display}\n"
        if latency_display:
            message += f" {latency_display}\n"
        if report_generated and report_msg:
            message = report_msg + "\n\n" + message
        if 0 < info_data["remain_flow_gb"] < config["LOW_FLOW_THRESHOLD_GB"]:
            title = '⚠️ 校园网流量预警'
            warn = f"剩余流量不足 {config['LOW_FLOW_THRESHOLD_GB']}GB！\n\n"
            if info_data["total_flow_gb"] == 60:
                warn += "企业微信-服务大厅可申请流量\n\n"
            message = warn + message
        else:
            title = '北方工业大学校园网'
        notification.notify(title=title, message=message, app_name='校园网流量助手', timeout=config["NOTICE_TIMEOUT"])
    else:
        notification.notify(title='流量查询失败', message='校园网页面加载失败，请检查登录状态', 
                            app_name='校园网流量助手', timeout=config["NOTICE_TIMEOUT"])

    if config["LOG_ENABLED"]:
        log_msg = (
            f"用户:{info_data['username']} | 已用时长:{info_data['used_time']} | "
            f"已用流量:{info_data['used_flow']} | 剩余流量:{info_data['remain_flow']} | "
            f"总流量:{info_data['total_flow_gb']}GB | 内网延迟:{quality['internal_latency']:.1f}ms | "
            f"内网丢包:{quality['internal_loss']:.0f}% | 公网延迟:{quality['external_latency']:.1f}ms | "
            f"公网丢包:{quality['external_loss']:.0f}%"
        )
        append_log(log_msg)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        append_log(f"致命错误: 主程序异常 - {str(e)}")
        generate_debug_report(e)
        notification.notify(
            title='程序异常',
            message=f'校园网助手运行出错: {str(e)}',
            app_name='校园网流量助手',
            timeout=config["NOTICE_TIMEOUT"]
        )
