# NCUT-CampusNet-Tool
适用于北方工业大学NCUT的校园网自动化流量查询工具。

学生社区中常见学生对校园网流量消耗存在疑惑，且流量查询较为不便。本程序有助于便捷、清晰掌握并规划校园网剩余流量。

开机时自动：
- 网络检测、剩余流量查询、网速延迟和丢包检测
- 低流量阈值告警
- 日志记录、月度报告生成、流量异常分析
- 全程静默无需手动操作。

通中位数绝对偏差值算法配合绝对阈值实现流量异常分析，或可解决对流量消耗速度的困惑？

v1.5已更新。


## 三步快速使用
- 在Releases中下载 .exe 文件
-  Win + R 输入 shell:startup 打开系统开机自启动文件夹
- 将 .exe 文件粘贴到该文件夹
- 完成！

也可以手动运行程序。

程序运行后，结果将通过Windows系统通知推送。日志、报告等需在工作区中查看。


## 环境要求
仅支持Windows 10或更新系统。

使用.exe版本(推荐）：
- Microsoft Edge（应当自带）

使用.py版本:
- Python 3
- pip install selenium plyer
- Microsoft Edge（应当自带）


## 提示
- 仅支持 NCUT-AUTO
- 本工具完全开源，不会收集、上传、存储任何用户校园网账号密码，所有数据仅在本地计算机生成与保存，不上传至任何第三方服务器。
- 本工具通过正常网页访问方式读取校园网已公开的流量信息，不破解、不篡改、不绕过校园网认证机制，不影响校园网正常运行与安全策略。
- 本工具仅用于北方工业大学 NCUT-AUTO 校园网个人流量查询、网络状态检测与本地日志记录，仅限本人合法校园网账号使用，严禁用于他人账号信息获取、网络攻击、流量盗用等违规违法活动。使用者因违规使用、不当操作、网络环境异常等导致的一切后果，由使用者自行承担，项目作者不承担任何法律及连带责任。
- 本工具为学生自制开源学习项目，非官方软件，与北方工业大学无关联。
- 水平有限，欢迎指正。


## 工作区：
我的电脑 → 文档 → NCUT_Campus_Network_Log
```
📁 我的文档 /
└─ 📁 NCUT_Campus_Network_Log # 【工作区】日志与报告存储目录
    ├─ 📄 network_log_YYYY-MM.txt # 当月网络使用详细日志
    └─ 📄 Report_YYYY-MM.txt # 月度流量统计与异常检测报告
```
## 自定义：
通过修改代码或配置文件实现自定义：
```

MAX_RETRY: 网络连接失败重试次数

RETRY_INTERVAL: 重试间隔(秒)

TARGET_SSID: 校园网WiFi名称

CAMPUS_URL: 校园网认证成功页面地址

CAMPUS_HOST: 校园网服务器地址

EXTERNAL_TEST_HOST: 公网连通性测试地址

NOTICE_TIMEOUT: 系统通知显示时长(秒)

LOW_FLOW_THRESHOLD_GB: 低流量告警阈值(GB)

WORK_DIR_NAME: 工作目录名称

PING_COUNT: 测速时发送的ping包数量

LOG_ENABLED: 是否启用日志记录(关闭后不生成日志和月度报告)

DEBUG_MODE: 调试模式(开启后生成详细错误报告，用于排查问题)

SPEED_TEST_ENABLED: 是否启用网络测速(关闭后不检测延迟和丢包)

ANOMALY_MAD_MULTIPLIER: 异常检测中位数绝对偏差倍数

MIN_RECORDS_FOR_ANOMALY: 异常检测所需最少记录数

ABSOLUTE_DAILY_THRESHOLD_GB: 单日流量绝对阈值(超过即判定为异常)

SAFE_DAILY_FLOOR_GB: 安全流量下限(低于此值不判定为异常)

OPEN_REPORT_AFTER_GENERATE: 生成月度报告后是否自动打开
```

## 问题：
程序在后台静默运行，且需要等待延迟测试、开机配置等，运行时不会立即显示内容，这是正常的。如果持续超过1分钟仍未响应，请打开Debug模式排查问题。
