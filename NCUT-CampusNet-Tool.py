# 北方工业大学校园网流量助手 v1.7.1
# 项目地址: https://github.com/LiuMashiro/NCUT-CampusNet-Tool
# 适用于 NCUT-AUTO 校园网，支持流量查询、网络检测、低流量告警、月度报告生成

import time
import os
import re
import subprocess
import socket
import datetime
import statistics
import sys
import traceback
import threading
from typing import Dict, Optional, Tuple, List

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import yaml
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.font_manager as fm

from winotify import Notification, audio


def _print_error(prefix: str, e: Exception, with_traceback: bool = True) -> None:
    try:
        print(f"[错误] {prefix}: {type(e).__name__}: {e}", file=sys.stderr)
        if with_traceback:
            traceback.print_exc()
    except Exception:
        pass


# ===================== 配置管理类 =====================
class ConfigManager:
    DEFAULT_CONFIG = {
        "MAX_RETRY": 5,
        "RETRY_INTERVAL": 3,
        "TARGET_SSID": "NCUT-AUTO",
        "CAMPUS_URL": "https://ip.ncut.edu.cn/srun_portal_success?ac_id=1&theme=pro",
        "CAMPUS_HOST": "ip.ncut.edu.cn",
        "EXTERNAL_TEST_HOST": "223.5.5.5",
        "NOTICE_TIMEOUT": 0,
        "LOW_FLOW_THRESHOLD_GB": 10.0,
        "PING_COUNT": 10,
        "LOG_ENABLED": True,
        "DEBUG_MODE": False,
        "SPEED_TEST_ENABLED": True,
        "ANOMALY_MAD_MULTIPLIER": 3.0,
        "MIN_RECORDS_FOR_ANOMALY": 3,
        "ABSOLUTE_DAILY_THRESHOLD_GB": 15.0,
        "SAFE_DAILY_FLOOR_GB": 1.5,
        "OPEN_REPORT_AFTER_GENERATE": True,
        # 说明：阈值 <= 0 表示禁用该项判定
        "NETWORK_WARN_EXTERNAL_LATENCY_MS": 200.0,
        "NETWORK_WARN_EXTERNAL_LOSS_PERCENT": 10.0,
        "NETWORK_WARN_INTERNAL_LATENCY_MS": 200.0,
        "NETWORK_WARN_INTERNAL_LOSS_PERCENT": 10.0,
    }

    def __init__(self, work_dir: str):
        self.work_dir = work_dir
        self.config_path = os.path.join(work_dir, "config.yaml")
        self.config = self.DEFAULT_CONFIG.copy()

    def load(self) -> None:
        if not os.path.exists(self.config_path):
            self._create_default()
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                user_config = yaml.safe_load(f)
                if isinstance(user_config, dict) and user_config:
                    self.config.update(user_config)
        except Exception as e:
            _print_error("配置加载失败，将回退到默认配置", e)
            self.config = self.DEFAULT_CONFIG.copy()

    def _create_default(self) -> None:
        config_content = """# 北方工业大学校园网流量助手 配置文件 v1.7.1
# 项目地址: https://github.com/LiuMashiro/NCUT-CampusNet-Tool
# 修改此文件后重启程序生效
# 如配置文件损坏，删除后重新运行程序将自动生成默认配置

# ==================== 基础配置 ====================
MAX_RETRY: 5                    # 网络连接失败重试次数
RETRY_INTERVAL: 3               # 重试间隔(秒)
TARGET_SSID: "NCUT-AUTO"        # 校园网WiFi名称
CAMPUS_URL: "https://ip.ncut.edu.cn/srun_portal_success?ac_id=1&theme=pro"  # 校园网认证成功页面地址
CAMPUS_HOST: "ip.ncut.edu.cn"   # 校园网服务器地址
EXTERNAL_TEST_HOST: "223.5.5.5" # 公网连通性测试地址(默认阿里云DNS)
NOTICE_TIMEOUT: 0               # 0 => 普通通知 short；非0 => 普通通知 long
LOW_FLOW_THRESHOLD_GB: 10.0     # 低流量告警阈值(GB)
PING_COUNT: 10                  # 测速时发送的ping包数量

# ==================== 功能开关 ====================
LOG_ENABLED: true               # 是否启用日志记录(关闭后不生成日志和月度报告)
DEBUG_MODE: false               # 调试模式(开启后生成详细错误报告)
SPEED_TEST_ENABLED: true        # 是否启用网络测速(关闭后不检测延迟和丢包)

# ==================== 网络质量告警阈值 ====================
# 说明：阈值 <= 0 表示禁用该项判定
NETWORK_WARN_EXTERNAL_LATENCY_MS: 200.0      # 公网延迟告警阈值(ms)
NETWORK_WARN_EXTERNAL_LOSS_PERCENT: 10.0     # 公网丢包告警阈值(%)
NETWORK_WARN_INTERNAL_LATENCY_MS: 200.0      # 内网延迟告警阈值(ms)
NETWORK_WARN_INTERNAL_LOSS_PERCENT: 10.0     # 内网丢包告警阈值(%)

# ==================== 异常检测配置 ====================
ANOMALY_MAD_MULTIPLIER: 3.0     # 异常检测中位数绝对偏差倍数
MIN_RECORDS_FOR_ANOMALY: 3      # 异常检测所需最少记录数
ABSOLUTE_DAILY_THRESHOLD_GB: 15.0  # 单日流量绝对阈值(超过即判定为异常)
SAFE_DAILY_FLOOR_GB: 1.5        # 安全流量下限(低于此值不判定为异常)

# ==================== 报告配置 ====================
OPEN_REPORT_AFTER_GENERATE: true  # 生成月度报告后是否自动打开
"""
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                f.write(config_content)
        except Exception as e:
            _print_error("默认配置文件写入失败", e)


# ===================== 日志管理类 =====================
class Logger:
    def __init__(self, work_dir: str, config: Dict):
        self.work_dir = work_dir
        self.config = config

    def get_log_file_path(self, date: Optional[datetime.datetime] = None) -> str:
        if date is None:
            date = datetime.datetime.now()
        date_str = date.strftime("%Y-%m")
        return os.path.join(self.work_dir, f"network_log_{date_str}.txt")

    def append(self, content: str) -> None:
        if not self.config.get("LOG_ENABLED", True):
            return
        log_path = self.get_log_file_path()
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            with open(log_path, "a", encoding="utf-8", newline="") as f:
                f.write(f"[{timestamp}] {content}\n")
        except Exception as e:
            _print_error("日志写入失败", e)

    def generate_debug_report(self, exc: Exception) -> None:
        if not self.config.get("DEBUG_MODE", False):
            return
        debug_dir = os.path.join(self.work_dir, "debug")
        try:
            os.makedirs(debug_dir, exist_ok=True)
        except Exception as e:
            _print_error("调试目录创建失败", e)
            return

        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        report_path = os.path.join(debug_dir, f"error_report_{timestamp}.txt")
        try:
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(f"=== 错误报告 ===\n")
                f.write(f"生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"程序版本: v1.7.1\n")
                f.write(f"Python版本: {sys.version}\n")
                f.write(f"操作系统: Windows\n")
                f.write(f"工作目录: {self.work_dir}\n")
                f.write("\n=== 异常信息 ===\n")
                f.write(f"异常类型: {type(exc).__name__}\n")
                f.write(f"异常信息: {str(exc)}\n")
                f.write("\n=== 堆栈跟踪 ===\n")
                f.write(traceback.format_exc())
                f.write("\n=== 当前配置 ===\n")
                yaml.dump(self.config, f, default_flow_style=False, allow_unicode=True)
        except Exception as e:
            _print_error("调试报告生成失败", e)


# ===================== 网络检测类 =====================
class NetworkChecker:
    def __init__(self, config: Dict):
        self.config = config
        self._startupinfo = subprocess.STARTUPINFO()
        self._startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        self._startupinfo.wShowWindow = subprocess.SW_HIDE

    def is_available(self) -> bool:
        for host, port in [(self.config["EXTERNAL_TEST_HOST"], 53), (self.config["CAMPUS_HOST"], 443)]:
            try:
                with socket.create_connection((host, port), timeout=3):
                    return True
            except (socket.timeout, ConnectionRefusedError, OSError):
                continue
            except Exception as e:
                _print_error(f"网络可用性检测异常 ({host}:{port})", e, with_traceback=False)
                continue
        try:
            with socket.create_connection(("www.baidu.com", 80), timeout=3):
                return True
        except (socket.timeout, ConnectionRefusedError, OSError):
            return False
        except Exception as e:
            _print_error("网络可用性检测异常 (www.baidu.com:80)", e, with_traceback=False)
            return False

    def get_wifi_ssid(self) -> str:
        for attempt in range(2):
            try:
                result = subprocess.run(
                    ["netsh", "wlan", "show", "interfaces"],
                    capture_output=True,
                    text=True,
                    encoding="gbk",
                    errors="ignore",
                    startupinfo=self._startupinfo,
                    timeout=5
                )
                ssid_match = re.search(r"^\s*SSID\s*[:：]\s*(.+)$", result.stdout, re.MULTILINE)
                if ssid_match:
                    return ssid_match.group(1).strip()
            except subprocess.TimeoutExpired as e:
                _print_error("获取SSID超时", e, with_traceback=False)
            except Exception as e:
                _print_error("获取SSID失败", e, with_traceback=False)
            if attempt == 0:
                time.sleep(0.5)
        return ""

    def is_campus_reachable(self) -> bool:
        try:
            with socket.create_connection((self.config["CAMPUS_HOST"], 443), timeout=3):
                return True
        except (socket.timeout, ConnectionRefusedError, OSError):
            return False
        except Exception as e:
            _print_error("校园网服务器连通性检测异常", e, with_traceback=False)
            return False

    def ping(self, host: str, count: int = None) -> Tuple[float, float]:
        if count is None:
            count = self.config["PING_COUNT"]
        try:
            result = subprocess.run(
                ["ping", "-n", str(count), "-w", "1000", host],
                capture_output=True,
                text=True,
                encoding="gbk",
                errors="ignore",
                startupinfo=self._startupinfo,
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
        except Exception as e:
            _print_error(f"Ping失败 ({host})", e, with_traceback=False)
            return -1.0, 100.0

    def get_quality(self) -> Dict[str, float]:
        if not self.config["SPEED_TEST_ENABLED"]:
            return {
                "internal_latency": -1.0,
                "internal_loss": -1.0,
                "external_latency": -1.0,
                "external_loss": -1.0
            }
        internal_latency, internal_loss = self.ping(self.config["CAMPUS_HOST"], count=5)
        external_latency, external_loss = self.ping(self.config["EXTERNAL_TEST_HOST"])
        return {
            "internal_latency": internal_latency,
            "internal_loss": internal_loss,
            "external_latency": external_latency,
            "external_loss": external_loss
        }


# ===================== 校园网信息获取类 =====================
class CampusNetFetcher:
    def __init__(self, config: Dict, logger: Logger):
        self.config = config
        self.logger = logger

    @staticmethod
    def _parse_flow_to_gb(flow_text: str) -> float:
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

    def fetch(self) -> Dict:
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
            driver.get(self.config["CAMPUS_URL"])
            wait = WebDriverWait(driver, 10)
            wait.until(EC.presence_of_element_located((By.ID, "remain-bytes")))

            def safe_get(xpath: str) -> str:
                try:
                    return driver.find_element(By.XPATH, xpath).text.strip()
                except Exception as e:
                    _print_error(f"元素获取失败: {xpath}", e, with_traceback=False)
                    return "N/A"

            data["username"] = safe_get('//*[@id="username"]')
            data["used_time"] = safe_get('//*[@id="used-time"]')
            data["used_flow"] = safe_get('//*[@id="used-flow"]')
            data["remain_flow"] = safe_get('//*[@id="remain-bytes"]')

            data["used_flow_gb"] = self._parse_flow_to_gb(data["used_flow"])
            data["remain_flow_gb"] = self._parse_flow_to_gb(data["remain_flow"])

            if data["used_flow_gb"] >= 0 and data["remain_flow_gb"] >= 0:
                total = data["used_flow_gb"] + data["remain_flow_gb"]
                data["total_flow_gb"] = round(total)

            data["success"] = True
        except Exception as e:
            self.logger.append(f"错误: Selenium抓取失败 - {str(e)}")
            self.logger.generate_debug_report(e)
            _print_error("Selenium抓取失败", e)
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception as e:
                    _print_error("Driver关闭失败", e, with_traceback=False)
        return data


# ===================== 通知管理类 =====================
class Notifier:
    def __init__(self, config: Dict):
        self.config = config
        self._notification_count = 0

    def send(self, title: str, message: str, is_warning: bool = False) -> None:
        try:
            notice_timeout = int(self.config.get("NOTICE_TIMEOUT", 0))
            normal_duration = "short" if notice_timeout == 0 else "long"
            duration_mode = "long" if is_warning else normal_duration

            self._notification_count += 1
            toast = Notification(
                app_id="校园网流量助手",
                title=title,
                msg=message,
                duration=duration_mode
            )
            toast.set_audio(audio.Default, loop=False)
            if self._notification_count > 1:
                time.sleep(0.5)
            toast.show()
        except Exception as e:
            _print_error("通知发送失败", e)


# ===================== 中文字体工具函数 =====================
def _get_chinese_font() -> Optional[str]:
    candidates = [
        "msyh.ttc",
        "msyhbd.ttc",
        "simhei.ttf",
        "simsun.ttc",
        "simkai.ttf",
    ]
    font_dirs = [
        os.path.join(os.environ.get("SystemRoot", "C:\\Windows"), "Fonts"),
        os.path.expanduser("~/AppData/Local/Microsoft/Windows/Fonts"),
    ]
    for font_file in candidates:
        for font_dir in font_dirs:
            full_path = os.path.join(font_dir, font_file)
            if os.path.exists(full_path):
                return full_path
    return None


def _setup_matplotlib_chinese_font() -> None:
    font_path = _get_chinese_font()
    if font_path:
        try:
            fm.fontManager.addfont(font_path)
            prop = fm.FontProperties(fname=font_path)
            font_name = prop.get_name()
            plt.rcParams["font.family"] = "sans-serif"
            plt.rcParams["font.sans-serif"] = [font_name] + plt.rcParams["font.sans-serif"]
            plt.rcParams["axes.unicode_minus"] = False
        except Exception as e:
            _print_error("Matplotlib中文字体设置失败，回退默认字体", e, with_traceback=False)
            plt.rcParams["font.family"] = "DejaVu Sans"
            plt.rcParams["axes.unicode_minus"] = False
    else:
        plt.rcParams["font.family"] = "DejaVu Sans"
        plt.rcParams["axes.unicode_minus"] = False


# ===================== 报告生成类 =====================
class ReportGenerator:
    def __init__(self, work_dir: str, config: Dict, logger: Logger):
        self.work_dir = work_dir
        self.config = config
        self.logger = logger

    def _read_log_records(self, log_path: str) -> List[Dict]:
        records = []
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
                        except (ValueError, TypeError) as e:
                            _print_error(f"日志记录解析失败: {line.strip()}", e, with_traceback=False)
                            continue
        except Exception as e:
            self.logger.append(f"错误: 读取日志失败 - {e}")
            _print_error("日志读取失败", e)
        return records

    def _generate_line_chart(self, records: List[Dict], report_date_str: str) -> Optional[str]:
        if not records:
            return None
        try:
            _setup_matplotlib_chinese_font()

            daily_records = {}
            for r in sorted(records, key=lambda x: x["datetime"]):
                daily_records[r["date"]] = r["flow"]
            dates = list(daily_records.keys())
            flows = list(daily_records.values())

            fig, ax = plt.subplots(figsize=(12, 6))
            ax.plot(dates, flows, marker='o', linestyle='-', color='b')
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
            ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, len(dates) // 10)))
            fig.autofmt_xdate()
            ax.set_xlabel('日期')
            ax.set_ylabel('已用流量 (GB)')
            ax.set_title(f'北方工业大学校园网流量趋势 ({report_date_str})')
            ax.grid(True)

            chart_path = os.path.join(self.work_dir, f"Flow_Chart_{report_date_str}.png")
            fig.savefig(chart_path, dpi=300, bbox_inches='tight')
            plt.close(fig)
            return chart_path
        except Exception as e:
            self.logger.append(f"错误: 折线图生成失败 - {e}")
            _print_error("折线图生成失败", e)
            return None

    def check_and_generate(self) -> Tuple[bool, str, bool]:
        if not self.config["LOG_ENABLED"]:
            return False, "", False
        now = datetime.datetime.now()
        last_month = now.replace(day=1) - datetime.timedelta(days=1)
        last_month_str = last_month.strftime("%Y-%m")
        report_filename = f"Report_{last_month_str}.txt"
        report_path = os.path.join(self.work_dir, report_filename)
        log_path = self.logger.get_log_file_path(last_month)

        if os.path.exists(report_path):
            return False, "", False
        if not os.path.exists(log_path):
            self.logger.append(f"系统: 无{last_month_str}月度日志，跳过报告生成")
            return False, "", False

        records = self._read_log_records(log_path)
        if not records:
            self.logger.append(f"系统: {last_month_str}日志无有效流量数据，跳过报告生成")
            return False, "", False

        summary_content = f"=== 北方工业大学校园网月度报告 ({last_month_str}) ===\n"
        summary_content += f"生成时间: {now.strftime('%Y-%m-%d %H:%M:%S')}\n"
        summary_content += f"程序版本: v1.7.1\n"
        summary_content += "----------------------------------------\n\n"

        report_notification_msg = ""
        has_anomaly = False
        anomalies = []

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

        daily_dates = sorted(daily_records.keys())
        increments = []

        if len(daily_dates) >= self.config["MIN_RECORDS_FOR_ANOMALY"]:
            for i in range(1, len(daily_dates)):
                prev_date = daily_dates[i - 1]
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
                threshold_avg = median_avg + self.config["ANOMALY_MAD_MULTIPLIER"] * mad

                anomalies = []
                for inc in increments:
                    if inc["avg"] < self.config["SAFE_DAILY_FLOOR_GB"]:
                        continue
                    is_anomaly = False
                    reason = []
                    if inc["avg"] > threshold_avg:
                        is_anomaly = True
                        reason.append(f"日均({inc['avg']:.1f}GB)远超正常")
                    if inc["avg"] > self.config["ABSOLUTE_DAILY_THRESHOLD_GB"]:
                        is_anomaly = True
                        reason.append(f"超单日阈值{self.config['ABSOLUTE_DAILY_THRESHOLD_GB']}GB")
                    if is_anomaly:
                        inc["reason"] = "；".join(reason)
                        anomalies.append(inc)

                summary_content += f"\n  日均中位数: {median_avg:.2f} GB/天\n"

                if not anomalies:
                    summary_content += "  未检测到异常流量消耗\n"
                else:
                    has_anomaly = True
                    summary_content += f"  ⚠ 检测到 {len(anomalies)} 次异常流量消耗:\n"
                    summary_content += "\n"

                    for idx, anom in enumerate(anomalies, start=1):
                        start_str = anom["start"].strftime("%Y-%m-%d")
                        end_str = anom["end"].strftime("%Y-%m-%d")
                        summary_content += f"  【异常 {idx}】\n"
                        summary_content += f"    · 时间段：{start_str} ~ {end_str}\n"
                        summary_content += f"    · 间隔：  {anom['days']} 天\n"
                        summary_content += f"    · 总消耗：{anom['total']:.2f} GB\n"
                        summary_content += f"    · 日均消耗：{anom['avg']:.2f} GB/天\n"
                        summary_content += f"    · 异常原因：{anom['reason']}\n"
                        summary_content += "\n"

        summary_content += "\n月度综合统计:\n"
        summary_content += f"  • 总检测次数: {total_records} 次\n"
        summary_content += f"  • 首次记录: {first_record['datetime'].strftime('%Y-%m-%d %H:%M')}\n"
        summary_content += f"  • 末次记录: {last_record['datetime'].strftime('%Y-%m-%d %H:%M')}\n"
        summary_content += f"  • 本月累计使用: ~{max_flow:.2f} GB\n"

        chart_path = self._generate_line_chart(records, last_month_str)
        if chart_path:
            summary_content += f"  • 流量趋势图: {chart_path}\n"

        report_notification_msg = f"生成了 {last_month_str} 月度报告\n"
        report_notification_msg += f"本月累计使用: ~{max_flow:.2f} GB\n"
        if has_anomaly:
            report_notification_msg += f"检测到 {len(anomalies)} 次流量异常！\n"
            report_notification_msg += f"最高日均消耗: {max(a['avg'] for a in anomalies):.1f} GB/天\n"
        report_notification_msg += f"报告路径: {report_path}"

        try:
            with open(report_path, "w", encoding="utf-8", newline="") as f:
                f.write(summary_content)
            self.logger.append(
                f"系统: 已生成 {last_month_str} 月度报告，"
                f"异常检测结果: {'发现异常' if has_anomaly else '正常'}"
            )
            if self.config["OPEN_REPORT_AFTER_GENERATE"]:
                try:
                    os.startfile(report_path)
                except Exception:
                    os.startfile(self.work_dir)
            return True, report_notification_msg, has_anomaly
        except Exception as e:
            self.logger.append(f"错误: 月度报告写入失败 - {e}")
            _print_error("月度报告写入失败", e)
            return False, "", False


# ===================== 主程序类 =====================
class NCUTCampusNetTool:
    def __init__(self):
        self.work_dir = self._get_work_directory()
        self.config_manager = ConfigManager(self.work_dir)
        self.config_manager.load()
        self.config = self.config_manager.config
        self.logger = Logger(self.work_dir, self.config)
        self.network_checker = NetworkChecker(self.config)
        self.fetcher = CampusNetFetcher(self.config, self.logger)
        self.notifier = Notifier(self.config)
        self.report_generator = ReportGenerator(self.work_dir, self.config, self.logger)
        self._quality = None

    @staticmethod
    def _get_work_directory() -> str:
        doc_path = os.path.expanduser("~/Documents")
        work_path = os.path.abspath(os.path.join(doc_path, "NCUT_Campus_Network_Log"))
        if not os.path.exists(work_path):
            os.makedirs(work_path, exist_ok=True)
            try:
                os.startfile(work_path)
            except Exception as e:
                _print_error("打开工作目录失败", e, with_traceback=False)
        return work_path

    def _check_startup_location(self) -> str:
        try:
            if getattr(sys, 'frozen', False):
                current_exe = sys.executable
            else:
                current_exe = os.path.abspath(__file__)
            current_dir = os.path.dirname(current_exe)
            startup_dir = os.path.expanduser(
                "~\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup"
            )
            if not os.path.samefile(current_dir, startup_dir):
                return "\n\n💡 提示: 程序未设置为开机自启\n建议放入启动文件夹实现开机检测"
            return ""
        except Exception as e:
            _print_error("启动位置检查失败", e, with_traceback=False)
            return ""

    def _get_last_record(self) -> Optional[Dict]:
        log_path = self.logger.get_log_file_path()
        if not os.path.exists(log_path):
            return None
        try:
            with open(log_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
                log_pattern = re.compile(
                    r"\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\] "
                    r"用户:([^|]+) \| "
                    r".*已用流量:([\d.]+)\s*GB"
                )
                for line in reversed(lines):
                    match = log_pattern.search(line)
                    if match:
                        dt_str, username, flow_str = match.groups()
                        try:
                            dt = datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
                            flow = float(flow_str)
                            return {"datetime": dt, "flow": flow}
                        except (ValueError, TypeError):
                            continue
        except Exception as e:
            _print_error("获取上次记录失败", e)
        return None

    def _check_realtime_anomaly(
        self, current_flow: float, current_datetime: datetime.datetime
    ) -> Optional[str]:
        last_record = self._get_last_record()
        if not last_record:
            return None
        days_diff = (current_datetime - last_record["datetime"]).days
        if days_diff <= 0:
            return None
        flow_inc = current_flow - last_record["flow"]
        if flow_inc <= 0:
            return None
        daily_avg = flow_inc / days_diff

        thr = float(self.config.get("ABSOLUTE_DAILY_THRESHOLD_GB", 15.0))
        if thr > 0 and daily_avg > thr:
            return (
                f"检测到流量异常！\n"
                f"上次记录: {last_record['datetime'].strftime('%Y-%m-%d %H:%M')}, "
                f"已用: {last_record['flow']:.2f}GB\n"
                f"本次记录: {current_datetime.strftime('%Y-%m-%d %H:%M')}, "
                f"已用: {current_flow:.2f}GB\n"
                f"日均消耗: {daily_avg:.2f}GB/天，超过{thr:.0f}GB阈值！"
            )
        return None

    def _background_network_test(self) -> None:
        quality = self.network_checker.get_quality()

        if self.config["SPEED_TEST_ENABLED"]:
            # 从配置读取告警阈值（阈值 <= 0 视为禁用该项）
            ext_lat_thr = float(self.config.get("NETWORK_WARN_EXTERNAL_LATENCY_MS", 200.0))
            ext_loss_thr = float(self.config.get("NETWORK_WARN_EXTERNAL_LOSS_PERCENT", 10.0))
            int_lat_thr = float(self.config.get("NETWORK_WARN_INTERNAL_LATENCY_MS", 200.0))
            int_loss_thr = float(self.config.get("NETWORK_WARN_INTERNAL_LOSS_PERCENT", 10.0))

            reasons = []

            # external latency
            if ext_lat_thr > 0 and quality["external_latency"] >= 0 and quality["external_latency"] > ext_lat_thr:
                reasons.append(f"公网延迟 {quality['external_latency']:.1f}ms > {ext_lat_thr:.0f}ms")
            # external loss
            if ext_loss_thr > 0 and quality["external_loss"] >= 0 and quality["external_loss"] > ext_loss_thr:
                reasons.append(f"公网丢包 {quality['external_loss']:.0f}% > {ext_loss_thr:.0f}%")
            # internal latency
            if int_lat_thr > 0 and quality["internal_latency"] >= 0 and quality["internal_latency"] > int_lat_thr:
                reasons.append(f"内网延迟 {quality['internal_latency']:.1f}ms > {int_lat_thr:.0f}ms")
            # internal loss
            if int_loss_thr > 0 and quality["internal_loss"] >= 0 and quality["internal_loss"] > int_loss_thr:
                reasons.append(f"内网丢包 {quality['internal_loss']:.0f}% > {int_loss_thr:.0f}%")

            if reasons:
                msg = (
                    "网络质量异常！\n"
                    + "；".join(reasons)
                    + "\n\n"
                    f"公网延迟: {quality['external_latency']:.1f}ms (丢包 {quality['external_loss']:.0f}%)\n"
                    f"内网延迟: {quality['internal_latency']:.1f}ms (丢包 {quality['internal_loss']:.0f}%)"
                )
                self.notifier.send('⚠️ 网络质量警告', msg, is_warning=True)

        self._quality = quality

    def run(self) -> None:
        report_generated, report_msg, report_has_anomaly = self.report_generator.check_and_generate()

        network_ok = False
        for _ in range(self.config["MAX_RETRY"]):
            if self.network_checker.is_available():
                network_ok = True
                break
            time.sleep(self.config["RETRY_INTERVAL"])

        if not network_ok:
            self.notifier.send('网络异常', '网络未连接，请检查网络后重试', is_warning=True)
            self.logger.append("状态: 网络连接失败")
            return

        current_ssid = self.network_checker.get_wifi_ssid()
        is_campus_network = (
            current_ssid == self.config["TARGET_SSID"]
            or self.network_checker.is_campus_reachable()
        )
        if not is_campus_network:
            self.notifier.send(
                '非校园网环境',
                f'当前SSID: {current_ssid}，仅支持NCUT-AUTO查询',
                is_warning=False
            )
            self.logger.append(f"状态: 非校园网环境 (SSID: {current_ssid})")
            return

        info_data = self.fetcher.fetch()
        self._quality = None
        if info_data["success"]:
            flow_display = (
                f"{info_data['remain_flow']} / {info_data['total_flow_gb']}GB"
                if info_data["total_flow_gb"] > 0
                else info_data["remain_flow"]
            )
            message = f"流量: {flow_display}\n"

            title = '北方工业大学校园网'
            is_warning = False

            if 0 < info_data["remain_flow_gb"] < 1.0:
                title = '🚨 校园网流量紧急预警'
                is_warning = True
                warn = "剩余流量不足1GB！！！\n\n"
                if info_data["total_flow_gb"] == 60:
                    warn += "企业微信-服务大厅可申请流量\n\n"
                message = warn + message
            elif 0 < info_data["remain_flow_gb"] < self.config["LOW_FLOW_THRESHOLD_GB"]:
                title = '⚠️ 校园网流量预警'
                is_warning = True
                warn = f"剩余流量不足 {self.config['LOW_FLOW_THRESHOLD_GB']}GB！\n\n"
                if info_data["total_flow_gb"] == 60:
                    warn += "企业微信-服务大厅可申请流量\n\n"
                message = warn + message

            startup_hint = self._check_startup_location()
            if startup_hint:
                message += startup_hint

            anomaly_msg = self._check_realtime_anomaly(
                info_data["used_flow_gb"], datetime.datetime.now()
            )

            if report_generated and report_msg:
                self.notifier.send('📊 月度报告已生成', report_msg, is_warning=report_has_anomaly)

            if anomaly_msg:
                self.notifier.send('⚠️ 实时流量异常警告', anomaly_msg, is_warning=True)

            self.notifier.send(title, message, is_warning=is_warning)

            network_thread = threading.Thread(target=self._background_network_test, daemon=True)
            network_thread.start()
            network_thread.join(timeout=30)

            if self.config["LOG_ENABLED"]:
                quality = self._quality if self._quality else {
                    "internal_latency": -1.0,
                    "internal_loss": -1.0,
                    "external_latency": -1.0,
                    "external_loss": -1.0
                }
                log_msg = (
                    f"用户:{info_data['username']} | 已用时长:{info_data['used_time']} | "
                    f"已用流量:{info_data['used_flow_gb']:.2f} GB | "
                    f"剩余流量:{info_data['remain_flow']} | "
                    f"总流量:{info_data['total_flow_gb']}GB | "
                    f"内网延迟:{quality['internal_latency']:.1f}ms | "
                    f"内网丢包:{quality['internal_loss']:.0f}% | "
                    f"公网延迟:{quality['external_latency']:.1f}ms | "
                    f"公网丢包:{quality['external_loss']:.0f}%"
                )
                self.logger.append(log_msg)
        else:
            self.notifier.send('流量查询失败', '校园网页面加载失败，请检查登录状态', is_warning=True)
            self.logger.append("错误: 流量查询失败（fetch返回success=False）")


def main():
    try:
        tool = NCUTCampusNetTool()
        tool.run()
    except Exception as e:
        _print_error("致命错误: 主程序异常", e, with_traceback=True)

        try:
            work_dir = os.path.abspath(
                os.path.join(os.path.expanduser("~/Documents"), "NCUT_Campus_Network_Log")
            )
            os.makedirs(work_dir, exist_ok=True)
            config = ConfigManager.DEFAULT_CONFIG.copy()
            logger = Logger(work_dir, config)
            logger.append(f"致命错误: 主程序异常 - {str(e)}")
            logger.generate_debug_report(e)
        except Exception as log_e:
            _print_error("致命错误: 写入日志失败", log_e, with_traceback=False)

        # 打包运行一般无控制台，仍尝试弹Toast
        try:
            toast = Notification(
                app_id="校园网流量助手",
                title="程序异常",
                msg=f"校园网助手运行出错: {str(e)}",
                duration="long"
            )
            toast.show()
        except Exception as toast_e:
            _print_error("致命错误: 弹出通知失败", toast_e, with_traceback=False)


if __name__ == "__main__":
    main()
