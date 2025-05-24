# LLMcyberworker

[![Python 3.8+](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

一个支持多平台大模型API调用的自动化文本处理和分析工具，提供可视化界面和灵活配置。

An automated text processing tool with GUI support for multiple LLM APIs, featuring flexible configuration and batch processing.


## 📌 主要功能

- ✅ 支持主流大模型API（兼容OpenAI格式API）
- 📊 支持Excel/CSV文件批量处理
- ⚙️ 可视化配置管理
- 🔄 断点续传功能
- 🚦 智能速率控制
- 📝 自定义系统/用户指令模板
- 📈 实时进度监控与统计

- ✅ Supports mainstream LLM APIs (OpenAI API compatible)
- 📊 Excel/CSV batch processing
- ⚙️ Visual configuration management
- 🔄 Resume from interruption
- 🚦 Intelligent rate control
- 📝 Customizable system/user prompts
- 📈 Real-time progress monitoring


## 📦 安装依赖

```
pandas>=2.0.0
requests>=2.28.0
openpyxl>=3.0.10
tkinter>=0.1.0
```

## 🚀 快速开始

1. **配置API信息**Configure API Settings
    - 点击菜单栏 `配置 -> API配置`Navigate to Config -> API Settings
    - 填写您的API地址和密钥  Fill in your API information:
    - 示例配置：
        API地址: https://api.your-llm-service.com/v1
        API密钥: sk-xxxxxxxxxxxxxxxxxxxxxxxx
        模型名称: your-model-name
2. **设置处理参数**
    - 通过菜单栏配置：
        - `列名配置`：指定输入文件的ID列和文本列
        - `处理配置`：设置并发数、重试策略等
3. **运行处理任务**
    1. 选择输入文件（支持.xlsx/.csv）
    2. 指定输出文件路径（.csv格式）
    3. 点击"开始处理"按钮

## ⚙️ 高级配置

配置文件 `config.json` 说明：
```
{
    "api_url": "API端点地址",
    "api_key": "API认证密钥",
    "model_name": "模型名称",
    "system_prompt": "系统级指令模板",
    "user_prompt": "用户指令模板（使用{text}占位符）",
    "id_column": "输入文件ID列名",
    "text_column": "输入文件文本列名",
    "max_retries": 3,
    "max_text_length": 4000,
    "max_workers": 10,
    "requests_per_second": 10,
    "rate_bucket": 15
}
```

> 注意：配置文件自动生成，请勿手动修改`config_hash`字段
> Note: Config file is auto-generated. Do NOT modify config_hash manually.

## 🛠️ 常见问题

### Q: 如何处理大文件？
- 程序支持断点续传，意外中断后可重新打开继续处理
- 检查点文件 `.progress_checkpoint` 会自动记录进度
### Q: API请求速度太慢？
- 您可以调整 `处理配置 -> 请求速率` 和 `并发线程数`
- API请求的上限往往受到相关大模型服务商限制，请在使用前核对是否存在此类限制
- 请尽可能选择较好网络环境，不要使用代理或其他网络工具
### Q: 如何验证配置有效性？
- 配置文件使用SHA256哈希校验
- 非法修改配置会自动恢复默认设置
### Q: 429 Too Many Requests错误？
- 这一报错往往源于触发了大模型服务平台的限速策略或风控
- 程序会自动重试并调整请求节奏
- 建议修改请求频率，适当降低请求参数，避免过多报错。

## 📄 许可证

本项目采用 MIT License

## 🤝 参与贡献

欢迎通过以下方式参与项目：
1. 提交Issue报告问题
2. Fork仓库后提交Pull Request
3. 完善文档或翻译
4. 分享使用案例

We welcome contributions through:
1. Submitting issues
2. Creating pull requests
3. Improving documentation
4. Sharing use cases
