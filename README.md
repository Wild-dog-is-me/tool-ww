# 幼儿饮食与运动记录表自动填写工具

这是一个基于 Flask 开发的本地网页工具，用于辅助填写幼儿园体弱儿个案管理相关 Excel 表格。

工具支持上传原始 Excel 模板、录入周次与日期、上传菜谱图片，并通过 OCR 识别当天食谱，自动填写“早点”“午餐”“午点”“进食情况”等内容，最后导出保持原文件名的 Excel 文件。

## 功能特点

- 支持通过网页界面上传 Excel 模板
- 支持按周次、月份、日期批量添加记录
- 支持上传菜谱图片并进行 OCR 识别
- 根据填写日期自动计算周几，并提取当天菜谱
- 自动填写表格中的早点、午餐、午点内容
- 自动填写“进食情况”字段
- 自动填写部分运动护理相关字段
- 导出文件时保留原始上传文件名
- 支持打包为 EXE，在其他老师电脑上直接运行
- 支持通过 `config.json` 配置工作目录与启动端口

## 技术栈

- Python 3
- Flask
- 百度 OCR（`baidu-aip`）
- xlrd
- xlwt
- xlutils

## 项目结构

```text
tool/
├─ app.py                  # 主程序入口
├─ templates/
│  └─ index.html           # 前端页面
├─ config.json             # 运行配置文件
├─ build.spec              # PyInstaller 打包配置
├─ requirements.txt        # Python 依赖
├─ data/                   # 示例数据目录
└─ .gitignore
```

## 运行方式

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 启动项目

```bash
python app.py
```

启动后会自动打开浏览器，默认访问：

```text
http://127.0.0.1:5000
```

## 使用流程

1. 打开网页工具
2. 上传 Excel 模板文件
3. 填写周次、月份、日期、记录人
4. 上传当天菜谱图片
5. 添加一条或多条记录
6. 点击填写并导出 Excel
7. 下载自动填充完成的文件

## 配置说明

项目根目录下的 `config.json` 用于控制程序运行参数。

示例：

```json
{
    "work_dir": "data",
    "port": 5000,
    "host": "127.0.0.1"
}
```

字段说明：

- `work_dir`：程序运行时的数据目录，支持相对路径或绝对路径
- `port`：Web 服务端口
- `host`：Web 服务监听地址

程序会在 `work_dir` 下自动创建以下目录：

- `uploads`：上传的 Excel 文件
- `images`：上传的菜谱图片
- `log`：运行日志

## 打包说明

项目已提供 `build.spec`，可直接使用 PyInstaller 打包：

```bash
python -m PyInstaller build.spec --clean --noconfirm
```

打包完成后，可执行文件位于：

```text
dist/diet_record/
```

建议将以下文件一起分发给其他使用者：

- `diet_record.exe`
- `config.json`
- `_internal` 目录

## OCR 说明

当前项目使用百度 OCR 识别菜谱图片，并基于日期对应的周几提取当天菜单内容。

识别后的数据会进一步整理为：

- 早点
- 午餐
- 午点
- 蔬菜

其中“午餐”会组合主食、荤菜、素菜和汤等内容。

## 日志说明

程序运行日志默认保存在：

```text
work_dir/log/app.log
```

如果遇到 OCR 识别异常、导出失败、路径问题等，可优先查看该日志定位问题。

## 注意事项

- 本项目主要面向本地 Windows 环境使用
- Excel 模板格式需要与当前填写逻辑保持一致
- 菜谱图片建议清晰、完整，避免裁剪和模糊
- `config.json` 建议与 EXE 放在同一目录下
- 如果部署到其他电脑，建议将工作目录配置到非系统敏感目录

## 后续可扩展方向

- 支持更多 Excel 模板格式
- 支持 OCR 结果人工校正
- 支持历史记录查看与复用
- 支持更多导出命名规则
- 支持图形化配置页面
