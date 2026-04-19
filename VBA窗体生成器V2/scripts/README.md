# 脚本说明 — VBA窗体生成器 V2

## create_vba_app.py

**用途：** 根据JSON配置生成包含UserForm + 功能模块的 .xlsm 文件

**依赖环境：**
```
Python >= 3.6
pywin32     → pip install pywin32
Excel 2010+  （必须安装）
Windows
```

### 使用方法

```bash
# 基本用法
python create_vba_app.py --config config.json

# 覆盖输出文件名
python create_vba_app.py --config config.json --output "我的工具.xlsm"
```

### 参数说明

| 参数 | 必需 | 说明 |
|------|------|------|
| `--config` | ✓ | 配置文件路径 |
| `--output` |  | 覆盖输出文件名 |

### 工作流程

```
解析配置 → 启动Excel(后台) → 导入模块(.bas) → 创建UserForm(设计时) 
→ 添加按钮控件 → 注入事件代码 → 添加启动宏 → 保存.xlsm
```

### 核心函数

| 函数 | 职责 |
|------|------|
| `load_config()` | 加载并校验 JSON 配置（含必填项检查）|
| `read_module_file()` | 读取 .bas 并清理 VB_Name 行 |
| `create_userform()` | 创建 UserForm + 设计时添加控件 + 注入代码 |
| `_add_button()` | 单个按钮控件的创建与属性设置 |
| `build_event_code()` | 构建 VBA 事件处理代码字符串 |
| `add_launcher_macro()` | 添加 Workbook_Open 和手动启动宏 |

### 错误排查

| 错误现象 | 解决方案 |
|----------|---------|
| `pywintypes.com_error` 无效类字符串 | Excel未安装或COM注册异常 → 以管理员运行 / 重装pywin32 |
| `PermissionError` 权限不足 | 关闭已打开的 Excel 文件后再执行 |
| `FileNotFoundError` 找不到模块 | 检查 modules 路径（支持绝对路径和相对路径）|

### 扩展开发

在 `create_userform()` 中添加新控件：

```python
# 添加标签
label = designer.Controls.Add("Forms.Label.1", "lblTitle")
label.Caption = "功能菜单"
label.Font.Size = 14

# 添加文本框
txt = designer.Controls.Add("Forms.TextBox.1", "txtInput")
txt.Width = 180
txt.Height = 24

# 添加图片控件（用于现代化背景）
img = designer.Controls.Add("Forms.Image.1", "imgBg")
img.Picture = LoadPicture("C:\\path\\to\\bg.png")
```
