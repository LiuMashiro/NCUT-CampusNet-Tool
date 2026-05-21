# 北方工业大学校园网流量助手 v1.7.6
# 项目地址: https://github.com/LiuMashiro/NCUT-CampusNet-Tool
# 适用于 NCUT-AUTO 校园网，支持流量查询、网络检测、低流量告警、月度报告生成及强制按需生成

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
    """统一错误输出工具函数，捕获写入 stderr 时的二次异常"""
    try:
        print(f"[错误] {prefix}: {type(e).__name__}: {e}", file=sys.stderr)
        if with_traceback:
            traceback.print_exc()
    except Exception:
        pass


# ===================== 配置管理类 =====================
class ConfigManager:
    """
    负责配置文件的读取、写入、版本兼容升级及强制标志位复位。
    """

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
        "ANOMALY_MAD_MULTIPLIER": 4.0,
        "MIN_RECORDS_FOR_ANOMALY": 3,
        "ABSOLUTE_DAILY_THRESHOLD_GB": 20.0,
        "SAFE_DAILY_FLOOR_GB": 3.0,
        "OPEN_REPORT_AFTER_GENERATE": True,
        "NETWORK_WARN_EXTERNAL_LATENCY_MS": 200.0,
        "NETWORK_WARN_EXTERNAL_LOSS_PERCENT": 10.0,
        "NETWORK_WARN_INTERNAL_LATENCY_MS": 200.0,
        "NETWORK_WARN_INTERNAL_LOSS_PERCENT": 10.0,
        "FORCE_GENERATE_LAST_MONTH_REPORT": False,
        "FORCE_GENERATE_THIS_MONTH_REPORT": False,
    }

    def __init__(self, work_dir: str):
        self.work_dir = work_dir
        self.config_path = os.path.join(work_dir, "config.yaml")
        self.config = self.DEFAULT_CONFIG.copy()

    def load(self) -> None:
        """加载配置文件，若不存在则创建默认配置"""
        if not os.path.exists(self.config_path):
            self._create_default()
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                user_config = yaml.safe_load(f)
                if isinstance(user_config, dict) and user_config:
                    self.config.update(user_config)
            # 兼容旧版本：向文件追加缺少的新配置项
            self._append_missing_configs_to_file()
        except Exception as e:
            _print_error("配置加载失败，将回退到默认配置", e)
            self.config = self.DEFAULT_CONFIG.copy()

    def _append_missing_configs_to_file(self) -> None:
        """检查并自动补全旧版缺失的配置，同时保留已有注释"""
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                content = f.read()

            appends = []
            if "FORCE_GENERATE_LAST_MONTH_REPORT:" not in content:
                appends.append(
                    "FORCE_GENERATE_LAST_MONTH_REPORT: false"
                    "  # 强制生成上月报告(覆盖原有)，运行一次后自动复位\n"
                )
            if "FORCE_GENERATE_THIS_MONTH_REPORT:" not in content:
                appends.append(
                    "FORCE_GENERATE_THIS_MONTH_REPORT: false"
                    "  # 强制生成本月报告(即使本月未结束)，运行一次后自动复位\n"
                )

            if appends:
                with open(self.config_path, "a", encoding="utf-8") as f:
                    f.write("\n# ==================== 报告强制生成配置 ====================\n")
                    for line in appends:
                        f.write(line)
        except Exception as e:
            _print_error("旧版本配置文件自动升级补全失败", e, with_traceback=False)

    def reset_force_flags(self) -> None:
        if not os.path.exists(self.config_path):
            return
        try:
            with open(self.config_path, "r", encoding="utf-8") as f:
                content = f.read()

            new_content = re.sub(
                r"(FORCE_GENERATE_LAST_MONTH_REPORT:\s*)true",
                r"\g<1>false",
                content,
                flags=re.IGNORECASE
            )
            new_content = re.sub(
                r"(FORCE_GENERATE_THIS_MONTH_REPORT:\s*)true",
                r"\g<1>false",
                new_content,
                flags=re.IGNORECASE
            )

            if content != new_content:
                with open(self.config_path, "w", encoding="utf-8") as f:
                    f.write(new_content)
                self.config["FORCE_GENERATE_LAST_MONTH_REPORT"] = False
                self.config["FORCE_GENERATE_THIS_MONTH_REPORT"] = False
        except Exception as e:
            _print_error("配置文件强制标志位复位失败", e, with_traceback=False)

    def _create_default(self) -> None:
        """写入带注释的默认配置文件"""
        config_content = """# 北方工业大学校园网流量助手 配置文件 v1.7.6
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
DEBUG_MODE: false               # 调试模式(开启后生成详细错误报告，并可能固定触发错误项)
SPEED_TEST_ENABLED: true        # 是否启用网络测速(关闭后不检测延迟和丢包)

# ==================== 网络质量告警阈值 ====================
# 说明：阈值 <= 0 表示禁用该项判定
NETWORK_WARN_EXTERNAL_LATENCY_MS: 200.0      # 公网延迟告警阈值(ms)
NETWORK_WARN_EXTERNAL_LOSS_PERCENT: 10.0     # 公网丢包告警阈值(%)
NETWORK_WARN_INTERNAL_LATENCY_MS: 200.0      # 内网延迟告警阈值(ms)
NETWORK_WARN_INTERNAL_LOSS_PERCENT: 10.0     # 内网丢包告警阈值(%)

# ==================== 异常检测配置 (v1.7.6 放宽) ====================
ANOMALY_MAD_MULTIPLIER: 4.0     # 异常检测中位数绝对偏差倍数
MIN_RECORDS_FOR_ANOMALY: 3      # 异常检测所需最少记录数
ABSOLUTE_DAILY_THRESHOLD_GB: 20.0  # 单日流量绝对阈值(超过即标记为疑似异常)
SAFE_DAILY_FLOOR_GB: 3.0        # 安全流量下限(低于此值不判定为疑似异常)

# ==================== 报告配置 ====================
OPEN_REPORT_AFTER_GENERATE: true  # 生成月度报告后是否自动打开

# ==================== 报告强制生成配置 ====================
# 提示：将其改为 true 并运行一次后，程序会自动生成对应报告并将其复位为 false
FORCE_GENERATE_LAST_MONTH_REPORT: false  # 强制生成上月报告
FORCE_GENERATE_THIS_MONTH_REPORT: false  # 强制生成本月报告
"""
        try:
            with open(self.config_path, "w", encoding="utf-8") as f:
                f.write(config_content)
        except Exception as e:
            _print_error("默认配置文件写入失败", e)


# ===================== 日志管理类 =====================
class Logger:
    """
    负责运行日志的追加写入与 Debug 错误报告的生成。
    """

    def __init__(self, work_dir: str, config: Dict):
        self.work_dir = work_dir
        self.config = config

    def get_log_file_path(self, date: Optional[datetime.datetime] = None) -> str:
        """返回指定月份的日志文件路径，默认为当前月"""
        if date is None:
            date = datetime.datetime.now()
        return os.path.join(self.work_dir, f"network_log_{date.strftime('%Y-%m')}.txt")

    def append(self, content: str) -> None:
        """追加一条日志记录，自动附加时间戳"""
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
        """在 debug 目录下生成详细错误报告文件，仅 DEBUG_MODE 开启时执行"""
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
                f.write("=== 错误报告 ===\n")
                f.write(f"生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"程序版本: v1.7.6\n")
                f.write(f"Python版本: {sys.version}\n")
                f.write(f"操作系统: {sys.platform}\n")
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
    """
    负责网络可用性探测、SSID 获取、Ping 测速及网络质量评估。
    """

    def __init__(self, config: Dict):
        self.config = config
        self._startupinfo = subprocess.STARTUPINFO()
        self._startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        self._startupinfo.wShowWindow = subprocess.SW_HIDE

    def is_available(self) -> bool:
        """多节点探测网络可用性，任意节点可达即返回 True"""
        for host, port in [
            (self.config["EXTERNAL_TEST_HOST"], 53),
            (self.config["CAMPUS_HOST"], 443)
        ]:
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
        """通过 netsh 命令获取当前 WiFi SSID，最多重试2次"""
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
        """TCP 探测校园网服务器是否可达"""
        try:
            with socket.create_connection((self.config["CAMPUS_HOST"], 443), timeout=3):
                return True
        except (socket.timeout, ConnectionRefusedError, OSError):
            return False
        except Exception as e:
            _print_error("校园网服务器连通性检测异常", e, with_traceback=False)
            return False

    def ping(self, host: str, count: int = None) -> Tuple[float, float]:
        """
        执行 Windows ping 命令并解析结果。
        返回 (平均延迟ms, 丢包率%)，失败时返回 (-1.0, 100.0)。
        """
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
            loss_match = re.search(r"(\d+)%\s*丢失", output, re.IGNORECASE)
            loss = float(loss_match.group(1)) if loss_match else 100.0
            time_matches = re.findall(r"时间[=<]\s*(\d+(?:\.\d+)?)ms", output, re.IGNORECASE)
            if time_matches:
                times = [float(t) for t in time_matches]
                return sum(times) / len(times), loss
            return -1.0, loss
        except Exception as e:
            _print_error(f"Ping 失败 ({host})", e, with_traceback=False)
            return -1.0, 100.0

    def get_quality(self) -> Dict[str, float]:
        """
        执行内外网 Ping 并返回质量指标字典。
        SPEED_TEST_ENABLED=false 时返回全 -1 占位值。
        """
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


# ===================== 通知管理类 =====================
class Notifier:
    """
    封装 winotify Windows 系统通知发送逻辑。
    维护通知计数以避免连续通知无间隔发送。
    """

    def __init__(self, config: Dict):
        self.config = config
        self._notification_count = 0

    def send(self, title: str, message: str, is_warning: bool = False) -> None:
        """发送 Windows 系统通知，警告类通知强制使用 long 持续时间"""
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


# ===================== 校园网信息获取类 =====================
class CampusNetFetcher:
    """
    使用 Selenium + Edge WebDriver 抓取校园网认证页面，解析流量信息。
    """

    def __init__(self, config: Dict, logger: "Logger", notifier: "Notifier"):
        self.config = config
        self.logger = logger
        self.notifier = notifier

    @staticmethod
    def _parse_flow_to_gb(flow_text: str) -> float:
        """将页面显示的流量字符串统一转换为 GB 浮点数"""
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
        """
        启动无头 Edge 浏览器访问校园网认证页面并抓取流量数据。
        返回包含 success 标志及各项流量字段的字典。
        """
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
            default_source_success = False

            def _attempt_default_source():
                nonlocal driver, default_source_success
                try:
                    if self.config.get("DEBUG_MODE", False):
                        # DEBUG: 强制睡眠超过 join timeout，触发镜像源切换逻辑
                        time.sleep(31)
                    driver = webdriver.Edge(options=edge_options)
                    default_source_success = True
                except Exception:
                    pass

            download_thread = threading.Thread(target=_attempt_default_source, daemon=True)
            start_time = time.time()
            download_thread.start()
            download_thread.join(timeout=30)

            if not default_source_success or driver is None:
                elapsed = time.time() - start_time
                self.logger.append(f"系统: 默认源下载超时({elapsed:.1f}s)，切换国内镜像")
                self.notifier.send(
                    "Edge 驱动下载提示",
                    "默认源下载超时，已切换国内镜像源，请稍候",
                    is_warning=False
                )
                os.environ["SE_CDN_URL"] = "https://registry.npmmirror.com/-/binary/edgedriver"
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

            data["username"]    = safe_get('//*[@id="username"]')
            data["used_time"]   = safe_get('//*[@id="used-time"]')
            data["used_flow"]   = safe_get('//*[@id="used-flow"]')
            data["remain_flow"] = safe_get('//*[@id="remain-bytes"]')

            data["used_flow_gb"]    = self._parse_flow_to_gb(data["used_flow"])
            data["remain_flow_gb"]  = self._parse_flow_to_gb(data["remain_flow"])

            if data["used_flow_gb"] >= 0 and data["remain_flow_gb"] >= 0:
                data["total_flow_gb"] = round(data["used_flow_gb"] + data["remain_flow_gb"])

            data["success"] = True

        except Exception as e:
            self.logger.append(f"错误: Selenium 抓取失败 - {str(e)}")
            self.logger.generate_debug_report(e)
            _print_error("Selenium 抓取失败", e)
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception as e:
                    _print_error("Driver 关闭失败", e, with_traceback=False)
        return data


# ===================== 中文字体工具函数 =====================
def _get_chinese_font() -> Optional[str]:
    """在常见 Windows 字体目录中查找可用的中文字体文件，返回第一个找到的路径"""
    candidates = ["msyh.ttc", "msyhbd.ttc", "simhei.ttf", "simsun.ttc", "simkai.ttf"]
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
    """
    配置 matplotlib 中文字体支持，优先使用系统字体文件注册，失败时安全降级。
    """
    font_path = _get_chinese_font()
    if font_path:
        try:
            fm.fontManager.addfont(font_path)
            prop = fm.FontProperties(fname=font_path)
            font_name = prop.get_name()
            # FIX-A: 原代码为 "font.font.family"（错误），修正为 "font.family"
            plt.rcParams["font.family"] = "sans-serif"
            plt.rcParams["font.sans-serif"] = (
                [font_name, "SimHei", "Microsoft YaHei"]
                + plt.rcParams["font.sans-serif"]
            )
            plt.rcParams["axes.unicode_minus"] = False
            return
        except Exception as e:
            _print_error("Matplotlib 中文字体注册失败，回退降级方案", e, with_traceback=False)

    # 降级：直接写入常用中文字体名称，依赖系统已安装
    plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei", "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False


# ===================== 报告生成类 =====================
class ReportGenerator:
    """
    负责从日志文件读取数据、执行异常检测、生成月度文本报告和折线图。
    """

    def __init__(self, work_dir: str, config: Dict, logger: "Logger"):
        self.work_dir = work_dir
        self.config = config
        self.logger = logger

    def _read_log_records(self, log_path: str) -> List[Dict]:
        """解析日志文件，提取所有有效的流量记录条目"""
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
                            records.append({
                                "datetime": dt,
                                "date": dt.date(),
                                "dt_hour": dt.replace(minute=0, second=0, microsecond=0),
                                "username": username.strip(),
                                "flow": float(flow_str)
                            })
                        except (ValueError, TypeError):
                            continue
        except Exception as e:
            self.logger.append(f"错误: 读取日志失败 - {e}")
            _print_error("日志读取失败", e)
        return records

    @staticmethod
    def _smooth_curve(dates: List[datetime.datetime], flows: List[float], num_points: int = 500, window_size: int = 15) -> Tuple[List[datetime.datetime], List[float]]:
        """严格单调平滑算法"""
        if len(dates) < 3:
            return dates, flows

        timestamps = [d.timestamp() for d in dates]
        start_t = timestamps[0]
        end_t = timestamps[-1]
        
        # 1. 建立高密度时间轴
        step = (end_t - start_t) / (num_points - 1)
        dense_t = [start_t + i * step for i in range(num_points)]
        dense_y = []

        # 2. 线性插值
        idx = 0
        for t in dense_t:
            while idx < len(timestamps) - 2 and timestamps[idx + 1] < t:
                idx += 1
            t0, t1 = timestamps[idx], timestamps[idx+1]
            y0, y1 = flows[idx], flows[idx+1]
            if t1 == t0:
                dense_y.append(y1)
            else:
                ratio = (t - t0) / (t1 - t0)
                dense_y.append(y0 + ratio * (y1 - y0))

        # 3. 滑动平均滤波
        smoothed_y = []
        half_window = window_size // 2
        for i in range(num_points):
            start_idx = max(0, i - half_window)
            end_idx = min(num_points, i + half_window + 1)
            window = dense_y[start_idx:end_idx]
            smoothed_y.append(sum(window) / len(window))

        # 强制首尾一致
        smoothed_y[0] = flows[0]
        smoothed_y[-1] = flows[-1]

        for i in range(1, len(smoothed_y)):
            if smoothed_y[i] < smoothed_y[i - 1]:
                smoothed_y[i] = smoothed_y[i - 1]  # 强制等于前值

        smoothed_dates = [datetime.datetime.fromtimestamp(t) for t in dense_t]
        return smoothed_dates, smoothed_y

    def _generate_line_chart(self, records: List[Dict], report_date_str: str) -> Optional[str]:
        """
        根据记录列表生成流量趋势折线图并保存为 PNG。
        """
        if not records:
            return None
        try:
            _setup_matplotlib_chinese_font()

            # 每小时取最后一条记录作为当日该小时的代表值
            hourly_records: Dict[datetime.datetime, float] = {}
            for r in sorted(records, key=lambda x: x["datetime"]):
                hourly_records[r["dt_hour"]] = r["flow"]

            raw_dates = sorted(hourly_records.keys())
            raw_flows = [hourly_records[d] for d in raw_dates]

            # 调用严格单调平滑算法
            dates, flows = self._smooth_curve(raw_dates, raw_flows, num_points=500, window_size=15)

            fig, ax = plt.subplots(figsize=(12, 6))
            
            # 绘制折线图：无 marker，加粗线条
            ax.plot(dates, flows, linestyle="-", color="#1a73e8", linewidth=2.5)
            
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))
            ax.xaxis.set_major_locator(mdates.AutoDateLocator())
            
            fig.autofmt_xdate()
            ax.set_xlabel("日期")
            ax.set_ylabel("已用流量 (GB)")
            ax.set_title(f"北方工业大学校园网流量趋势 ({report_date_str})")
            ax.grid(True, linestyle="--", alpha=0.6)

            chart_path = os.path.join(self.work_dir, f"Flow_Chart_{report_date_str}.png")
            fig.savefig(chart_path, dpi=300, bbox_inches="tight")
            plt.close(fig)
            return chart_path
        except Exception as e:
            self.logger.append(f"错误: 折线图生成失败 - {e}")
            _print_error("折线图生成失败", e)
            return None

    def _generate_for_period(
        self, target_date: datetime.datetime, force: bool = False
    ) -> Tuple[bool, str, bool]:
        """
        为指定月份生成月度报告。
        """
        target_month_str = target_date.strftime("%Y-%m")
        report_path = os.path.join(self.work_dir, f"Report_{target_month_str}.txt")
        log_path = self.logger.get_log_file_path(target_date)

        if not force and os.path.exists(report_path):
            return False, "", False
        if not os.path.exists(log_path):
            self.logger.append(f"系统: 无 {target_month_str} 月度日志，跳过报告生成")
            return False, "", False

        records = self._read_log_records(log_path)
        if not records:
            self.logger.append(f"系统: {target_month_str} 日志无有效流量数据，跳过报告生成")
            return False, "", False

        total_records = len(records)
        first_record = min(records, key=lambda x: x["datetime"])
        last_record  = max(records, key=lambda x: x["datetime"])
        max_flow     = max(r["flow"] for r in records)

        daily_records: Dict[datetime.date, Dict] = {}
        for r in sorted(records, key=lambda x: x["datetime"]):
            daily_records[r["date"]] = r

        lines = []
        lines.append(f"=== 北方工业大学校园网月度报告 ({target_month_str}) ===")
        lines.append(f"生成时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"程序版本: v1.7.6")
        lines.append("-" * 40)
        lines.append("")

        lines.append("使用记录:")
        lines.append(f"{'日期':<12} | {'已用流量':>10}")
        lines.append("-" * 28)

        iter_prev_date = None
        for date in sorted(daily_records.keys()):
            r = daily_records[date]
            if iter_prev_date is None or (date - iter_prev_date).days >= 1:
                lines.append(f"{date.strftime('%Y-%m-%d'):<12} | {r['flow']:>8.2f} GB")
                iter_prev_date = date

        has_anomaly = False
        anomalies: List[Dict] = []
        daily_dates = sorted(daily_records.keys())

        if len(daily_dates) >= self.config["MIN_RECORDS_FOR_ANOMALY"]:
            increments = []
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
                        "start":  prev_date,
                        "end":    curr_date,
                        "days":   days_diff,
                        "total":  total_inc,
                        "avg":    daily_avg
                    })

            if increments:
                daily_avgs  = [inc["avg"] for inc in increments]
                median_avg  = statistics.median(daily_avgs)
                mad         = statistics.median([abs(x - median_avg) for x in daily_avgs])
                threshold_avg = median_avg + self.config["ANOMALY_MAD_MULTIPLIER"] * mad

                for inc in increments:
                    if inc["avg"] < self.config["SAFE_DAILY_FLOOR_GB"]:
                        continue
                    reasons = []
                    if inc["avg"] > threshold_avg:
                        reasons.append(f"日均消耗 {inc['avg']:.1f} GB，可能不符合本月其他时段使用习惯")
                    if inc["avg"] > self.config["ABSOLUTE_DAILY_THRESHOLD_GB"]:
                        reasons.append(f"日均消耗 {inc['avg']:.1f} GB，超过设定参考阈值 {self.config['ABSOLUTE_DAILY_THRESHOLD_GB']:.0f} GB")
                    if reasons:
                        inc["reason"] = "；".join(reasons)
                        anomalies.append(inc)

                lines.append("")
                lines.append(f"  日均用量中位数参考: {median_avg:.2f} GB/天")

                if not anomalies:
                    lines.append("  未发现明显异常流量消耗。")
                else:
                    has_anomaly = True
                    lines.append(f"  注意: 以下 {len(anomalies)} 个时间段的流量消耗可能存在异常，请结合实际使用情况判断：\n")
                    for idx, anom in enumerate(anomalies, start=1):
                        lines.append(f"  【疑似异常 {idx}】")
                        lines.append(f"    · 时间段：{anom['start'].strftime('%Y-%m-%d')} ~ {anom['end'].strftime('%Y-%m-%d')}")
                        lines.append(f"    · 间隔：  {anom['days']} 天")
                        lines.append(f"    · 总消耗：{anom['total']:.2f} GB")
                        lines.append(f"    · 日均消耗：{anom['avg']:.2f} GB/天")
                        lines.append(f"    · 参考说明：{anom['reason']}\n")

        lines.append("")
        lines.append("月度综合统计:")
        lines.append(f"  • 总检测次数: {total_records} 次")
        lines.append(f"  • 首次记录: {first_record['datetime'].strftime('%Y-%m-%d %H:%M')}")
        lines.append(f"  • 末次记录: {last_record['datetime'].strftime('%Y-%m-%d %H:%M')}")
        lines.append(f"  • 本月累计使用: 约 {max_flow:.2f} GB")

        chart_path = self._generate_line_chart(records, target_month_str)
        if chart_path:
            lines.append(f"  • 流量趋势图: {chart_path}")

        summary_content = "\n".join(lines) + "\n"

        try:
            with open(report_path, "w", encoding="utf-8", newline="") as f:
                f.write(summary_content)
            self.logger.append(f"系统: 已生成 {target_month_str} 月度报告（覆盖模式={force}）")
            if self.config.get("OPEN_REPORT_AFTER_GENERATE", True):
                try:
                    os.startfile(report_path)
                except Exception:
                    pass
        except Exception as e:
            self.logger.append(f"错误: 月度报告写入失败 - {e}")
            _print_error("月度报告写入失败", e)
            return False, "", False

        notice_lines = [f"已生成 {target_month_str} 月度报告。", f"本期累计使用: 约 {max_flow:.2f} GB"]
        if has_anomaly:
            notice_lines.append(f"检测到 {len(anomalies)} 个时间段的流量消耗可能存在异常，请查阅报告。")
        notice_lines.append(f"报告路径: {report_path}")

        return True, "\n".join(notice_lines), has_anomaly

    def check_and_generate(self) -> List[Tuple[str, bool]]:
        """报告生成入口"""
        if not self.config.get("LOG_ENABLED", True):
            return []

        notifications: List[Tuple[str, bool]] = []
        now = datetime.datetime.now()
        last_month = now.replace(day=1) - datetime.timedelta(days=1)

        force_last = self.config.get("FORCE_GENERATE_LAST_MONTH_REPORT", False)
        force_this = self.config.get("FORCE_GENERATE_THIS_MONTH_REPORT", False)
        executed_force = False

        if force_last:
            success, msg, is_warn = self._generate_for_period(last_month, force=True)
            if success: notifications.append((msg, is_warn))
            executed_force = True

        if force_this:
            success, msg, is_warn = self._generate_for_period(now, force=True)
            if success: notifications.append((msg, is_warn))
            executed_force = True

        if not executed_force:
            success, msg, is_warn = self._generate_for_period(last_month, force=False)
            if success: notifications.append((msg, is_warn))

        return notifications

# ===================== 主程序类 =====================
class NCUTCampusNetTool:
    """
    主控类，协调所有模块的执行顺序，实现完整的一次检测流程。
    """

    def __init__(self):
        self.work_dir        = self._get_work_directory()
        self.config_manager  = ConfigManager(self.work_dir)
        self.config_manager.load()
        self.config          = self.config_manager.config
        self.logger          = Logger(self.work_dir, self.config)
        self.network_checker = NetworkChecker(self.config)
        self.notifier        = Notifier(self.config)
        self.fetcher         = CampusNetFetcher(self.config, self.logger, self.notifier)
        self.report_generator = ReportGenerator(self.work_dir, self.config, self.logger)
        self._quality: Optional[Dict[str, float]] = None

    @staticmethod
    def _get_work_directory() -> str:
        """确保工作目录存在，首次创建时自动打开"""
        work_path = os.path.abspath(
            os.path.join(os.path.expanduser("~/Documents"), "NCUT_Campus_Network_Log")
        )
        if not os.path.exists(work_path):
            os.makedirs(work_path, exist_ok=True)
            try:
                os.startfile(work_path)
            except Exception as e:
                _print_error("打开工作目录失败", e, with_traceback=False)
        return work_path

    def _check_startup_location(self) -> str:
        """检测程序是否在开机启动目录中运行，否则提示用户配置自启"""
        try:
            current_exe = sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
            current_dir = os.path.dirname(current_exe)
            startup_dir = os.path.expanduser(
                "~\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup"
            )
            if not os.path.samefile(current_dir, startup_dir):
                return "\n\n提示: 程序未设置为开机自启，建议放入启动文件夹。"
            return ""
        except Exception:
            return ""

    def _get_last_record(self) -> Optional[Dict]:
        """从当月日志文件中读取最后一条有效流量记录"""
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
                    dt_str, _, flow_str = match.groups()
                    try:
                        return {
                            "datetime": datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S"),
                            "flow": float(flow_str)
                        }
                    except (ValueError, TypeError):
                        continue
        except Exception as e:
            _print_error("获取上次记录失败", e)
        return None

    def _check_realtime_anomaly(
        self, current_flow: float, current_datetime: datetime.datetime
    ) -> Optional[str]:
        """
        与上次日志记录对比，检测当前是否存在疑似异常的流量增长。
        """
        last_record = self._get_last_record()
        if not last_record:
            return None

        flow_inc = current_flow - last_record["flow"]
        if flow_inc <= 0:
            return None

        thr = float(self.config.get("ABSOLUTE_DAILY_THRESHOLD_GB", 20.0))
        if thr <= 0:
            return None

        days_diff = (current_datetime.date() - last_record["datetime"].date()).days

        # 场景1：跨天，检测日均消耗是否超标
        if days_diff > 0:
            daily_avg = flow_inc / days_diff
            if daily_avg > thr:
                # FIX-W: 措辞克制化
                return (
                    f"流量变动提示\n"
                    f"上次记录: {last_record['datetime'].strftime('%m-%d %H:%M')}，"
                    f"{last_record['flow']:.2f} GB\n"
                    f"本次记录: {current_datetime.strftime('%m-%d %H:%M')}，"
                    f"{current_flow:.2f} GB\n"
                    f"区间日均消耗约 {daily_avg:.2f} GB/天，"
                    f"超过参考阈值 {thr:.0f} GB，请留意是否存在异常。"
                )

        # 场景2：同日内，检测单日增量是否超标
        elif days_diff == 0 and flow_inc > thr:
            # FIX-W: 措辞克制化，去除"狂飙""击穿"等情绪化词语
            return (
                f"流量变动提示\n"
                f"自今日 {last_record['datetime'].strftime('%H:%M')} 起，"
                f"当日流量增量约 {flow_inc:.2f} GB，"
                f"已超过单日参考阈值 {thr:.0f} GB，请留意是否存在异常。"
            )

        return None

    def _background_network_test(self) -> None:
        """
        在独立线程中执行网络质量测试，结果写入 self._quality。
        DEBUG_MODE 下注入固定的高延迟高丢包数据以验证告警逻辑。
        """
        if self.config.get("DEBUG_MODE", False):
            quality = {
                "internal_latency": 500.0,
                "internal_loss": 50.0,
                "external_latency": 800.0,
                "external_loss": 60.0
            }
        else:
            quality = self.network_checker.get_quality()

        if self.config.get("SPEED_TEST_ENABLED", True):
            ext_lat_thr  = float(self.config.get("NETWORK_WARN_EXTERNAL_LATENCY_MS", 200.0))
            ext_loss_thr = float(self.config.get("NETWORK_WARN_EXTERNAL_LOSS_PERCENT", 10.0))
            int_lat_thr  = float(self.config.get("NETWORK_WARN_INTERNAL_LATENCY_MS", 200.0))
            int_loss_thr = float(self.config.get("NETWORK_WARN_INTERNAL_LOSS_PERCENT", 10.0))

            reasons = []
            if ext_lat_thr  > 0 and quality["external_latency"]  >= 0 and quality["external_latency"]  > ext_lat_thr:
                reasons.append(f"公网延迟 {quality['external_latency']:.1f} ms（参考阈值 {ext_lat_thr:.0f} ms）")
            if ext_loss_thr > 0 and quality["external_loss"]      >= 0 and quality["external_loss"]      > ext_loss_thr:
                reasons.append(f"公网丢包 {quality['external_loss']:.0f}%（参考阈值 {ext_loss_thr:.0f}%）")
            if int_lat_thr  > 0 and quality["internal_latency"]   >= 0 and quality["internal_latency"]   > int_lat_thr:
                reasons.append(f"内网延迟 {quality['internal_latency']:.1f} ms（参考阈值 {int_lat_thr:.0f} ms）")
            if int_loss_thr > 0 and quality["internal_loss"]      >= 0 and quality["internal_loss"]      > int_loss_thr:
                reasons.append(f"内网丢包 {quality['internal_loss']:.0f}%（参考阈值 {int_loss_thr:.0f}%）")

            if reasons:
                msg = (
                    "当前网络质量可能存在异常，请留意：\n"
                    + "；".join(reasons)
                    + "\n\n"
                    f"公网: {quality['external_latency']:.1f} ms"
                    f"（丢包 {quality['external_loss']:.0f}%）\n"
                    f"内网: {quality['internal_latency']:.1f} ms"
                    f"（丢包 {quality['internal_loss']:.0f}%）"
                )
                self.notifier.send("网络质量提示", msg, is_warning=True)

        self._quality = quality

    def run(self) -> None:
        """主执行流程入口，按固定顺序协调所有模块"""

        # Step 1: 报告生成（含强制生成逻辑）
        report_notices = self.report_generator.check_and_generate()

        # Step 2: 强制标志复位（有 force 标志被触发时才写文件）
        if (self.config.get("FORCE_GENERATE_LAST_MONTH_REPORT")
                or self.config.get("FORCE_GENERATE_THIS_MONTH_REPORT")):
            self.config_manager.reset_force_flags()

        # Step 3: 网络可用性检测
        network_ok = False
        for _ in range(self.config["MAX_RETRY"]):
            if self.network_checker.is_available():
                network_ok = True
                break
            time.sleep(self.config["RETRY_INTERVAL"])

        if not network_ok:
            self.notifier.send("网络连接失败", "当前网络不可用，请检查网络连接后重试。", is_warning=True)
            self.logger.append("状态: 网络连接检测失败")
            return

        # Step 4: 校园网环境检测
        current_ssid = self.network_checker.get_wifi_ssid()
        is_campus_network = (
            current_ssid == self.config["TARGET_SSID"]
            or self.network_checker.is_campus_reachable()
        )
        if not is_campus_network:
            self.notifier.send(
                "非校园网环境",
                f"当前 SSID: {current_ssid}，本工具仅支持 NCUT-AUTO 网络环境。",
                is_warning=False
            )
            self.logger.append(f"状态: 非校园网环境（SSID: {current_ssid}）")
            return

        # Step 5: 流量数据抓取
        info_data = self.fetcher.fetch()
        self._quality = None

        if not info_data["success"]:
            self.notifier.send(
                "流量查询失败",
                "校园网页面加载失败，请确认已登录校园网后重试。",
                is_warning=True
            )
            self.logger.append("错误: 流量查询失败（fetch 返回 success=False）")
            return

        # Step 6: 构建流量通知
        flow_display = (
            f"{info_data['remain_flow']} / {info_data['total_flow_gb']} GB"
            if info_data["total_flow_gb"] > 0
            else info_data["remain_flow"]
        )
        message = f"流量: {flow_display}\n"
        title = "北方工业大学校园网"
        is_warning = False

        remain_gb = info_data["remain_flow_gb"]
        if 0 < remain_gb < 1.0:
            title = "校园网流量余量不足"
            is_warning = True
            warn = "剩余流量不足 1 GB，请及时处理。\n\n"
            if info_data["total_flow_gb"] == 60:
                warn += "可通过企业微信-服务大厅申请流量。\n\n"
            message = warn + message
        elif 0 < remain_gb < self.config["LOW_FLOW_THRESHOLD_GB"]:
            title = "校园网流量余量提示"
            is_warning = True
            warn = f"剩余流量低于 {self.config['LOW_FLOW_THRESHOLD_GB']} GB，请注意使用。\n\n"
            if info_data["total_flow_gb"] == 60:
                warn += "可通过企业微信-服务大厅申请流量。\n\n"
            message = warn + message

        startup_hint = self._check_startup_location()
        if startup_hint:
            message += startup_hint

        # Step 7: 实时异常检测
        anomaly_msg = self._check_realtime_anomaly(
            info_data["used_flow_gb"], datetime.datetime.now()
        )

        # Step 8: 统一发送所有通知
        for msg, is_warn in report_notices:
            self.notifier.send("月度报告已生成", msg, is_warning=is_warn)

        if anomaly_msg:
            self.notifier.send("流量变动提示", anomaly_msg, is_warning=True)

        self.notifier.send(title, message, is_warning=is_warning)

        # Step 9: 后台网络质量测试（等待最多30秒）
        network_thread = threading.Thread(target=self._background_network_test, daemon=True)
        network_thread.start()
        network_thread.join(timeout=30)

        # Step 10: 写入日志
        if self.config["LOG_ENABLED"]:
            quality = self._quality or {
                "internal_latency": -1.0,
                "internal_loss": -1.0,
                "external_latency": -1.0,
                "external_loss": -1.0
            }
            log_msg = (
                f"用户:{info_data['username']} | "
                f"已用时长:{info_data['used_time']} | "
                f"已用流量:{info_data['used_flow_gb']:.2f} GB | "
                f"剩余流量:{info_data['remain_flow']} | "
                f"总流量:{info_data['total_flow_gb']} GB | "
                f"内网延迟:{quality['internal_latency']:.1f}ms | "
                f"内网丢包:{quality['internal_loss']:.0f}% | "
                f"公网延迟:{quality['external_latency']:.1f}ms | "
                f"公网丢包:{quality['external_loss']:.0f}%"
            )
            self.logger.append(log_msg)


# ===================== 程序入口 =====================
def main():
    """
    程序入口，捕获顶层异常并尽力发出错误通知和日志。
    确保任何崩溃都有迹可查。
    """
    try:
        tool = NCUTCampusNetTool()
        tool.run()
    except Exception as e:
        _print_error("致命错误: 主程序异常", e, with_traceback=True)

        # 尝试写入错误日志
        try:
            work_dir = os.path.abspath(
                os.path.join(os.path.expanduser("~/Documents"), "NCUT_Campus_Network_Log")
            )
            os.makedirs(work_dir, exist_ok=True)
            fallback_logger = Logger(work_dir, ConfigManager.DEFAULT_CONFIG.copy())
            fallback_logger.append(f"致命错误: 主程序异常 - {str(e)}")
            fallback_logger.generate_debug_report(e)
        except Exception as log_e:
            _print_error("致命错误: 写入日志失败", log_e, with_traceback=False)

        # 尝试弹出错误通知
        try:
            toast = Notification(
                app_id="校园网流量助手",
                title="程序运行异常",
                msg=f"校园网助手运行时遇到错误: {type(e).__name__}，请查阅日志。",
                duration="long"
            )
            toast.show()
        except Exception as toast_e:
            _print_error("致命错误: 弹出通知失败", toast_e, with_traceback=False)


if __name__ == "__main__":
    main()
