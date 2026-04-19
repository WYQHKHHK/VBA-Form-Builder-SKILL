# UserForm 设计模式与规范

## 一、布局模式

### 1. 单列按钮布局（功能 < 10个）

```
┌──────────────────────┐
│   ▎ 标题区域         │
├──────────────────────┤
│                      │
│  [ 按钮1           ] │
│  [ 按钮2           ] │
│  [ 按钮3           ] │
│                      │
│  [ 退出            ] │
└──────────────────────┘
```

**参数：** 宽240 / 按钮宽180 高28 / 垂直间距40pt

### 2. 双列按钮布局（10-20个功能）

```
┌─────────────────────────────┐
│       ▎ 标题区域             │
├─────────────────────────────┤
│ [按钮1]        [按钮2]      │
│ [按钮3]        [按钮4]      │
│ [按钮5]        [按钮6]      │
│                             │
│        [ 退出 ]             │
└─────────────────────────────┘
```

**参数：** 宽360 / 按钮宽140 高28 / 列间距20pt / 行间距40pt

### 3. 四区专业布局（推荐用于复杂应用）

```
┌──────────────────────────────┐
│ ███▎ 顶部标题栏 (深色背景)    │  ← Logo + 应用名称 + 可选关闭按钮
├────────┬─────────────────────┤
│        │                     │
│ 左侧   │   中间操作区          │  ← TextBox/ComboBox 集中摆放
│ 导航   │   输入框/下拉框       │
│ 图标   │   表格/列表          │
│ 菜单   │                     │
│        ├─────────────────────┤
│        │  底部按钮栏          │  ← 保存 / 取消 / 确定 固定底部
└────────┴─────────────────────┘
```

### 4. 分组布局（功能分类明确）

```
┌──────────────────────┐
│  ▎ 数据分析工具箱     │
├──────────────────────┤
│ ━━ 数据处理 ━━━      │
│ [ 清洗数据 ]         │
│ [ 格式转换 ]         │
│                      │
│ ━━ 报表生成 ━━━      │
│ [ 日报生成 ]         │
│ [ 月报汇总 ]         │
│                      │
│ ━━━━━━━━━━━━━       │
│ [      退出        ] │
└──────────────────────┘
```

---

## 二、现代化样式规范

### 字体系统

| 元素 | 字体 | 字号 | 字重 |
|------|------|------|------|
| 窗体标题 | 微软雅黑 | 12pt | Bold |
| 区域标题 | 微软雅黑 | 10pt | Bold |
| 按钮文本 | 微软雅黑 | 11pt | Regular |
| 标签文本 | 微软雅黑 | 10pt | Regular |
| 输入框文字 | 微软雅黑 | 10pt | Regular |

> **硬性规定：全程使用微软雅黑，禁止宋体/Tahoma。**

### 配色系统

#### 极简配色原则：2主色 + 黑白灰

```
主色系:    #4472C4 (专业蓝) 或 #5B9BD5 (浅蓝) 或 #70AD47 (自然绿)
辅助色:    #BDD7EE (浅蓝灰)
文字主色:  #333333 (深灰)
文字辅色:  #999999 (中灰) — 提示文字用
背景色:    #F5F5F5 (极浅灰) 或白色
边框色:    #E0E0E0 (浅灰)
危险/退出: #C00000 (暗红) 或 #E74C3C (亮红)
```

#### 按钮色彩规范

| 类型 | 背景色 | 文字色 | 边框 |
|------|--------|--------|------|
| **主操作** | `#4472C4` | 白色 | 无 |
| **次操作** | `#F0F0F0` | `#333` | `#D0D0D0` |
| **成功** | `#70AD47` | 白色 | 无 |
| **危险** | `#E74C3C` | 白色 | 无 |
| **禁用** | `#F5F5F5` | `#CCC` | `#EEE` |

### 尺寸规范

| 元素 | 宽度 | 高度 | 圆角(视觉) |
|------|------|------|-----------|
| 主按钮 | 180-200pt | 32pt | 大圆角(PNG实现) |
| 次按钮 | 140-160pt | 28pt | 小圆角 |
| 图标按钮 | 80-100pt | 28pt | 小圆角 |
| 输入框 | 180-200pt | 26pt | 2px圆角边框 |
| 下拉框 | 180-200pt | 26pt | 同输入框 |
| 组标题 | 自动宽度 | 18pt | 无 |

### 间距规范（留白是现代化的关键）

```
窗体内边距:     20pt (四边统一)
按钮垂直间距:   40pt (紧凑模式 32pt)
按钮水平间距:   20pt (双列时)
组间间距:       50-60pt
组内标题-控件:  12pt
标签-输入框:    8pt
底部栏-底边:    16pt
```

---

## 三、事件处理模板

### 标准事件（最常用）

```vba
Private Sub btn_Click()
    On Error Resume Next
    Call 功能函数名
    On Error GoTo 0
End Sub
```

### 带确认的事件

```vba
Private Sub btnDelete_Click()
    If MsgBox("确认执行此操作？", vbYesNo + vbQuestion, "确认") = vbYes Then
        On Error Resume Next
        Call 删除函数
        On Error GoTo 0
    End If
End Sub
```

### 带进度提示的事件（耗时操作）

```vba
Private Sub btnProcess_Click()
    On Error GoTo ErrorHandler
    
    ' 显示进度状态
    Me.Caption = "正在处理..."
    DoEvents
    
    ' 执行操作
    Application.ScreenUpdating = False
    Call 长时间处理函数
    Application.ScreenUpdating = True
    
    ' 恢复并提示
    Me.Caption = "数据分析工具箱"
    MsgBox "处理完成！", vbInformation
    Exit Sub
    
ErrorHandler:
    Me.Caption = "数据分析工具箱"
    Application.ScreenUpdating = True
    MsgBox "错误：" & Err.Description, vbCritical
End Sub
```

---

## 四、性能优化

### 窗体加载优化

```vba
Private Sub UserForm_Initialize()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 初始化控件...
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
```

### 延迟加载（大量控件时）

```vba
Private m_blnInitialized As Boolean

Private Sub UserForm_Activate()
    If Not m_blnInitialized Then
        Application.ScreenUpdating = False
        InitializeControls
        Application.ScreenUpdating = True
        m_blnInitialized = True
    End If
End Sub
```

### 按钮状态管理

```vba
' 禁用全部功能按钮（防止重复点击）
Private Sub DisableAllButtons()
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is MSForms.CommandButton Then
            If ctrl.Name <> "btnExit" Then ctrl.Enabled = False
        End If
    Next ctrl
End Sub
```

---

## 五、交互模式

| 模式 | 代码 | 说明 |
|------|------|------|
| **模态**（默认） | `frmMain.Show` | 阻塞Excel，必须关闭窗体才能操作 |
| **非模态** | `frmMain.Show vbModeless` | 不阻塞，可同时操作Excel和窗体 |
| **隐藏** | `Me.Hide` | 隐藏但不卸载（保留状态）|
| **卸载** | `Unload Me` | 完全释放资源 |

### 数据回传方式

```vba
' 方式1：公共变量（简单场景）
Public g_strResult As String

' 方式2：Property（推荐）
Private m_strData As String
Public Property Let Data(ByVal v As String): m_strData = v: End Property
Public Property Get Data() As String: Data = m_strData: End Property
```

---

## 六、DO & DON'T

### ✅ DO
- 使用标准布局模式（单列/双列/分组/四区）
- 统一字体（微软雅黑）、尺寸、间距
- 每个事件加 `On Error Resume Next` 保护
- 有意义的控件命名（`btnSaveData` > `CommandButton1`）
- 大量控件时启用延迟加载

### ❌ DON'T
- 混用多种字体或字号
- 不一致的按钮尺寸/间距
- 缺少错误处理
- 长时间操作不显示进度（用户以为卡死）
- 使用无意义名称（Text1, Label1, CommandButton1）
