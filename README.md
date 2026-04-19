# VBA窗体生成器 V2 SKILL

> 🎨 现代化 VBA 窗体应用生成SKILL - 支持图片背景、样式按钮、专业布局，一键输出 `.xlsm`

## 📋 项目简介

**VBA窗体生成器 V2** SKILL 是一个用于快速生成现代化 Excel VBA 应用程序的工具。通过简单的 JSON 配置文件，即可生成包含精美 UserForm 界面和功能模块的 `.xlsm` 文件，彻底告别老旧的 Win32 土气界面。

### ✨ 核心特性

- 🎯 **配置驱动** - 通过 JSON 配置即可定制，零代码修改
- 🎨 **现代化 UI** - 支持图片背景、自定义样式、无边框窗口
- 🔧 **设计时控件** - 控件永久保存在文件中，设计器画布可见
- 🚀 **一键生成** - 自动创建 UserForm、导入模块、添加启动宏
- 📦 **开箱即用** - 生成文件可直接分发，支持自动启动

---

## 🚀 快速开始

### 环境要求

- Python >= 3.6
- pywin32 (`pip install pywin32`)
- Microsoft Excel 2010+
- Windows 操作系统

### 安装依赖

```bash
pip install pywin32
```

### 基本用法

**Step 1: 创建配置文件 `config.json`**

```json
{
  "output_file": "我的工具.xlsm",
  "form": {
    "title": "数据分析工具箱",
    "width": 280,
    "height": 420,
    "font_name": "微软雅黑",
    "font_size": 11
  },
  "modules": ["module1.bas", "module2.bas"],
  "buttons": [
    {"name": "btn1", "caption": "数据清洗", "module": "数据清洗", "top": 50},
    {"name": "btn2", "caption": "报表生成", "module": "生成报表", "top": 95},
    {"name": "btnExit", "caption": "退出", "module": "Unload Me", "top": 340}
  ],
  "auto_start": true
}
```

**Step 2: 运行生成脚本**

```bash
python templates/create_vba_app.py --config config.json
```

**Step 3: 打开生成的文件验证**

打开生成的 `.xlsm` 文件 → 窗体自动弹出 → 测试每个按钮功能

---

## 📝 配置文件详解

### 基础配置

| 配置项 | 类型 | 必填 | 说明 |
|--------|------|------|------|
| `output_file` | string | ✓ | 输出文件名 (`.xlsm`) |
| `form` | object | ✓ | 窗体配置 |
| `form.title` | string | ✓ | 窗体标题 |
| `form.width` | number | | 窗体宽度 (默认 240) |
| `form.height` | number | | 窗体高度 (默认 300) |
| `form.font_name` | string | | 字体名称 (默认 "微软雅黑") |
| `form.font_size` | number | | 字体大小 (默认 11) |
| `modules` | array | ✓ | 模块文件路径数组 |
| `buttons` | array | ✓ | 按钮配置数组 |
| `auto_start` | boolean | | 打开时自动显示窗体 (默认 true) |

### 按钮配置

| 属性 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `name` | string | ✓ | 控件名称 (如 `btn1`) |
| `caption` | string | ✓ | 显示文本 |
| `module` | string | ✓ | 调用的函数名，或 `"Unload Me"` 表示退出 |
| `top` | number | ✓ | 距顶部位置 (像素) |
| `left` | number | | 距左侧位置 (默认 30) |
| `width` | number | | 按钮宽度 (默认 200) |
| `height` | number | | 按钮高度 (默认 32) |

### 现代化配置选项

```json
{
  "form": {
    "title": "现代化工具",
    "width": 520,
    "height": 480,
    "background_color": "2D2D2D",
    "background_image": "assets/bg.png",
    "_modern": {
      "border_style": "none"
    }
  }
}
```

| 选项 | 说明 |
|------|------|
| `background_color` | 背景色 (十六进制，如 `2D2D2D`) |
| `background_image` | 背景图片路径 (相对路径) |
| `_modern.border_style` | 边框样式 (`"standard"` 或 `"none"`) |

---

## 🎨 现代化 UI 特性

### 1. 图片背景支持

在窗体底层放置 Image 控件作为背景，瞬间提升高级感：

```json
{
  "form": {
    "background_image": "assets/bg_gradient.png"
  }
}
```

**推荐背景风格：**
- 磨砂渐变
- 浅色科技底纹
- 半透明 PNG

### 2. 无边框悬浮模式

隐藏原生标题栏和边框，完全自定义界面：

```json
{
  "form": {
    "_modern": {
      "border_style": "none"
    }
  }
}
```

**效果：** 完全不像 Excel 窗体，像独立桌面软件。

### 3. 专业布局规范

| 规则 | 说明 |
|------|------|
| **对齐统一** | 标签左对齐，输入框右对齐 |
| **尺寸统一** | 同类控件高度/宽度完全一致 |
| **间距统一** | 控件间留相等空隙 |
| **分隔清晰** | 用细线/浅色块区分功能区 |

---

## 📁 项目结构

```
VBA窗体生成器V2/
├── README.md                 # 项目说明文档
├── SKILL.md                  # Skill 详细文档
├── templates/
│   ├── create_vba_app.py     # 主生成脚本
│   └── config.json           # 配置示例
├── scripts/
│   └── README.md             # 脚本使用说明
└── references/
    ├── userform_patterns.md  # 布局模式参考
    └── common_errors.md      # 常见错误排查
```

---

## 💡 触发 Skill 的提示词

在支持 Skill 的 AI 助手环境中，使用以下提示词可以触发 **VBA窗体生成器 V2** Skill：

### 直接触发词

```
/VBA窗体生成器V2
VBA窗体生成器V2
vba-form-generator-v2
```

### 功能描述触发

```
生成 VBA 窗体
创建 Excel 窗体应用
制作 VBA 用户界面
生成 .xlsm 文件
创建现代化 VBA 界面
```

### 自然语言触发示例

```
帮我生成一个带按钮的 Excel 窗体
创建一个数据分析工具的 VBA 界面
制作一个现代化的 Excel 工具箱
生成包含 UserForm 的 .xlsm 文件
帮我做一个有图片背景的 VBA 窗体
```

### 场景化触发词

| 场景 | 触发提示词 |
|------|-----------|
| 基础窗体 | "生成 VBA 窗体" / "创建 Excel 工具界面" |
| 现代化界面 | "现代化 VBA 界面" / "无边框窗体" / "图片背景" |
| 功能集成 | "集成模块到窗体" / "按钮调用宏" |
| 自动启动 | "打开 Excel 自动显示窗体" |

---

## 🔧 高级用法

### 命令行参数

```bash
python create_vba_app.py --config config.json --output "自定义名称.xlsm"
```

| 参数 | 说明 |
|------|------|
| `--config` | 配置文件路径 (必填) |
| `--output` | 覆盖输出文件名 (可选) |

### 自定义按钮样式

通过控件名称前缀自动应用样式：

| 前缀 | 样式 |
|------|------|
| `btnExit` | 退出按钮样式 (灰色加粗) |
| `btnClose` | 关闭按钮样式 |
| `lbl` 开头 | 标签样式 (蓝色背景白字) |
| `lblTitleBar` | 标题栏样式 (深灰背景) |

### 分隔线

创建空标题且高度 <= 5 的按钮会自动变为分隔线：

```json
{"name": "sep1", "caption": "", "top": 150, "height": 2}
```

---

## ❗ 常见问题

| 问题 | 解决方案 |
|------|---------|
| 窗体不显示 | 检查宏是否被禁用：信任中心 → 启用宏 |
| 按钮点击无反应 | 检查 `module` 名称与 `.bas` 文件中函数名是否一致 |
| 找不到模块 | 使用绝对路径或相对于 `config.json` 的相对路径 |
| 文件无法保存 | 先关闭所有 Excel 进程再运行脚本 |
| `pywintypes.com_error` | Excel 未安装或 COM 注册异常，尝试以管理员运行或重装 pywin32 |

---

## 📄 许可证

MIT License

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

---

> 💡 **提示：** 详细的使用文档请参考 [SKILL.md](./SKILL.md)
