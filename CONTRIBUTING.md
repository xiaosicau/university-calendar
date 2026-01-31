# 贡献指南

感谢您有兴趣为校历助手项目做出贡献！本文档提供了有关如何参与该项目的指南。

## 项目概述

校历助手是一款基于PyQt5开发的校历桌面应用程序，支持自定义导入校历和课表，提供系统托盘、开机自启动、课程表、上课提醒等功能。适用于各类高校教师和学生使用。

**开发者**：肖诗顺 - 四川农业大学经济学院金融系  
**许可证**：本程序为开源程序，供教师和学生免费使用

> 注意：当学校公布新的校历后，本程序会发布新的学年版本

## 贡献方式

### 报告问题 (Bug Reports)

当您发现bug时，请按以下步骤操作：

1. 搜索现有的 Issues，确保该问题尚未被报告
2. 创建新 Issue，提供以下信息：
   - 问题的详细描述
   - 复现步骤
   - 预期行为与实际行为
   - 您使用的系统和版本信息

### 功能请求 (Feature Requests)

我们欢迎功能请求！请在 Issue 中详细描述：
- 您想要的功能
- 为什么需要这个功能
- 使用场景

### 代码贡献

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some amazing feature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 创建 Pull Request

## 开发环境设置

1. 克隆仓库：
   ```bash
   git clone https://github.com/your-username/sicau-calendar.git
   ```

2. 安装依赖：
   ```bash
   pip install -r requirements.txt
   ```

3. 运行项目：
   ```bash
   python sicau_calendar.py
   ```

## 代码规范

- 遵循 PEP 8 Python 代码风格
- 使用有意义的变量和函数名
- 添加适当的注释和文档字符串
- 确保代码在不同平台上兼容

## 项目结构

```
sicau-calendar/
├── sicau_calendar.py        # 主程序文件
├── requirements.txt         # 项目依赖
├── README.md               # 项目说明
├── USAGE.md                # 使用说明
├── CONTRIBUTING.md         # 贡献指南
├── RELEASE_GUIDE.md        # 发布指南
├── VERSION_INFO.md         # 版本信息
├── LICENSE                 # 许可证
├── .gitignore              # Git忽略文件配置
├── .gitattributes          # Git属性配置
└── data/                   # 数据目录（自动创建）
    └── calendar_data.json  # 用户数据文件
```

## 发布版本

本项目提供两种使用方式：

1. **预编译版本**：从 [Releases](https://github.com/your-username/your-repo/releases) 页面下载预编译的可执行文件
2. **源代码版本**：下载源代码并自行编译运行

## 联系方式

如有问题请联系开发者。