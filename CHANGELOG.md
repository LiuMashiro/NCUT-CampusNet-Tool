# Changelog
All notable changes to this project will be documented in this file.

## [Unreleased] 待发布

### Added
- 新增独立配置文件系统，存放于程序工作目录，支持外部直接修改参数。
- 新增全局配置项：支持手动开关日志记录、网络测速（延迟丢包检测）功能。
- 增加程序启动路径检测，未放置在启动目录时自动弹出告警提示。
- 新增 DEBUG 错误调试模式，开启后自动在工作区生成错误报告。
- 优化首次运行与报告生成体验：首次启动自动打开工作区，月度报告生成后自动打开所在目录。

### Improved
- 底层逻辑重构优化，提升运行效率。
- 增强程序稳定性、兼容性。
