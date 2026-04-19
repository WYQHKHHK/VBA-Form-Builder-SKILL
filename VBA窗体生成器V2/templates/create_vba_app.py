# -*- coding: utf-8 -*-
"""
VBA Form Application Generator V2 - Modern Edition
====================================================
Features: Create Excel VBA application with UserForm and modules from JSON config
Core capabilities:
  - Add controls at design-time (persist in file)
  - Modern UI support (theme/style/background)
  - Correct MSForms property access
Dependencies: Python 3.6+, pywin32, Microsoft Excel
Usage: python create_vba_app.py --config config.json [--output filename.xlsm]
"""

import win32com.client as win32
import os
import json
import argparse
import time


def load_config(config_file):
    """Load and validate JSON config file"""
    with open(config_file, 'r', encoding='utf-8') as f:
        config = json.load(f)

    required = ['output_file', 'form', 'modules', 'buttons']
    for key in required:
        if key not in config:
            raise ValueError(f"Missing required config: {key}")
    if 'title' not in config.get('form', {}):
        raise ValueError("Missing required config: form.title")

    return config


def read_module_file(filepath):
    """Read .bas module file, remove VB_Name attribute lines"""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.split('\n')
    cleaned = [
        line for line in lines
        if not line.strip().startswith('Attribute VB_Name')
    ]
    return '\n'.join(cleaned)


def create_excel_app():
    """Create background Excel application instance"""
    print("[1/6] Starting Excel ...")
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel


def create_workbook(excel):
    """Create new workbook"""
    print("[2/6] Creating new workbook ...")
    return excel.Workbooks.Add()


def add_vba_modules(wb, modules, base_dir):
    """Import .bas module files to VBA project"""
    print(f"[3/6] Importing {len(modules)} modules...")
    project = wb.VBProject

    for mod_file in modules:
        path = mod_file if os.path.isabs(mod_file) else os.path.join(base_dir, mod_file)
        code = read_module_file(path)

        comp = project.VBComponents.Add(1)  # vbext_ct_StdModule
        comp.CodeModule.AddFromString(code)
        print(f"       OK {mod_file}")


def create_userform(wb, config, base_dir='.'):
    """Create UserForm + add controls at design-time + inject event code"""
    print("[4/6] Creating UserForm ...")
    project = wb.VBProject

    # Create UserForm
    uf = project.VBComponents.Add(3)  # vbext_ct_MSForm
    uf.Name = "frmMain"
    designer = uf.Designer

    # Set form properties
    frm = config['form']
    modern = frm.get('_modern', {})
    
    try:
        designer.Properties("Caption").Value = frm['title']
    except Exception:
        pass
    try:
        designer.Properties("Width").Value = frm.get('width', 240)
    except Exception:
        pass
    try:
        designer.Properties("Height").Value = frm.get('height', 300)
    except Exception:
        pass
    
    # Set border style (0 = none for borderless window)
    border_style = modern.get('border_style', 'standard')
    if border_style == 'none':
        try:
            designer.Properties("BorderStyle").Value = 0  # fmBorderStyleNone
        except Exception:
            pass
    
    # Set background color using the form object
    try:
        bg_color = frm.get('background_color', '')
        if bg_color:
            # Convert hex string to integer
            if isinstance(bg_color, str):
                bg_color = int(bg_color, 16)
            # Try to access the form object directly
            form_obj = designer.vbControls.Parent
            form_obj.BackColor = bg_color
    except Exception:
        # Alternative: set via Properties if available
        try:
            if bg_color:
                designer.BackColor = bg_color
        except Exception:
            pass

    # Add background image control if configured
    bg_image = frm.get('background_image', '')
    if bg_image:
        _add_background_image_control(designer, frm.get('width', 520), frm.get('height', 480))

    # Add button controls (design-time creation)
    font_name = frm.get('font_name', '微软雅黑')
    font_size = frm.get('font_size', 11)

    print("       Adding controls:")
    for btn in config['buttons']:
        _add_button(designer, btn, font_name, font_size)

    # Inject event code
    code = build_event_code(config)
    uf.CodeModule.AddFromString(code)


def _add_background_image_control(designer, width, height):
    """Add background image control to designer canvas (picture loaded via VBA code)"""
    try:
        # Add Image control for background
        img_ctrl = designer.Controls.Add("Forms.Image.1", "bgImage")
        
        # Set position and size (cover entire form including borders/title bar)
        img_ctrl.Left = 0
        img_ctrl.Top = 0
        # Make it significantly larger to ensure full coverage
        # VBA UserForm has internal margins, so we need extra space
        img_ctrl.Width = width + 100
        img_ctrl.Height = height + 100
        
        # Set picture mode to stretch
        img_ctrl.PictureSizeMode = 3  # fmPictureSizeModeStretch
        
        # Send to back so other controls are on top
        try:
            img_ctrl.ZOrder(1)  # fmSendToBack
        except Exception:
            pass
        
        print("       Background image control added")
        
    except Exception as e:
        print(f"       Warning: Could not add background image control: {e}")


def _add_button(designer, btn_cfg, font_name, font_size):
    """Add single button control to designer canvas"""
    name = btn_cfg['name']
    caption = btn_cfg.get('caption', '')
    
    # Check if separator (empty caption and small height)
    is_separator = caption == "" and btn_cfg.get('height', 32) <= 5
    
    try:
        if is_separator:
            # Use Label as separator line
            ctrl = designer.Controls.Add("Forms.Label.1", name)
        else:
            ctrl = designer.Controls.Add("Forms.CommandButton.1", name)
    except Exception as e:
        print(f"       FAIL {name}: {e}")
        return

    # Direct property assignment
    try:
        if not is_separator:
            ctrl.Caption = caption
        ctrl.Left = btn_cfg.get('left', 30)
        ctrl.Top = btn_cfg['top']
        ctrl.Width = btn_cfg.get('width', 200)
        ctrl.Height = btn_cfg.get('height', 32)

        # Font settings (with fallback)
        try:
            ctrl.Font.Name = font_name
            ctrl.Font.Size = font_size
        except Exception:
            pass
            
        # Separator line styling
        if is_separator:
            try:
                ctrl.BackColor = 0x808080
                ctrl.BorderStyle = 1
                ctrl.SpecialEffect = 0
            except Exception:
                pass
        # Custom title bar styling
        elif name == "lblTitleBar":
            try:
                ctrl.BackColor = 0x2D2D2D  # Dark gray for title bar
                ctrl.ForeColor = 0xFFFFFF  # White text
                ctrl.Font.Bold = True
                ctrl.Font.Size = font_size + 2
                ctrl.TextAlign = 2  # Center align
            except Exception:
                pass
        # Custom close button styling
        elif name == "btnClose":
            try:
                ctrl.BackColor = 0xC0C0C0
                ctrl.ForeColor = 0x000000
                ctrl.Font.Bold = True
                ctrl.Font.Size = font_size + 4
            except Exception:
                pass
        # Exit button styling
        elif name == "btnExit":
            try:
                ctrl.BackColor = 0xE0E0E0
                ctrl.Font.Bold = True
            except Exception:
                pass
        # Other label styling
        elif name.startswith("lbl"):
            try:
                ctrl.BackColor = 0x4472C4
                ctrl.ForeColor = 0xFFFFFF
                ctrl.Font.Bold = True
                ctrl.Font.Size = font_size + 1
            except Exception:
                pass
        else:
            # Normal button styling
            try:
                ctrl.BackColor = 0xF0F0F0
            except Exception:
                pass

        label = caption if not is_separator else '(separator)'
        print(f"       OK [{name}] {label}")

    except Exception as e:
        print(f"       FAIL {name}: {e}")


def build_event_code(config):
    """Build VBA event code for the form"""
    frm = config['form']
    title = frm['title']
    width = frm.get('width', 520)
    height = frm.get('height', 480)
    bg_color = frm.get('background_color', '')
    bg_image = frm.get('background_image', '')
    modern = frm.get('_modern', {})
    border_style = modern.get('border_style', 'standard')
    is_borderless = border_style == 'none'
    
    # Convert hex to VBA color value
    bg_color_vba = ''
    if bg_color:
        if isinstance(bg_color, str):
            bg_color = int(bg_color, 16)
        bg_color_vba = f"Me.BackColor = &H{bg_color:06X}"

    lines = [
        "Option Explicit",
        "",
    ]
    
    # Add Windows API declarations for borderless window
    if is_borderless:
        lines.extend([
            "' ============================================",
            "'  Windows API Declarations",
            "' ============================================",
            "#If VBA7 Then",
            "    Private Declare PtrSafe Function FindWindow Lib \"user32\" Alias \"FindWindowA\" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr",
            "    Private Declare PtrSafe Function SetWindowLong Lib \"user32\" Alias \"SetWindowLongA\" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long",
            "    Private Declare PtrSafe Function GetWindowLong Lib \"user32\" Alias \"GetWindowLongA\" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long",
            "#Else",
            "    Private Declare Function FindWindow Lib \"user32\" Alias \"FindWindowA\" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long",
            "    Private Declare Function SetWindowLong Lib \"user32\" Alias \"SetWindowLongA\" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long",
            "    Private Declare Function GetWindowLong Lib \"user32\" Alias \"GetWindowLongA\" (ByVal hWnd As Long, ByVal nIndex As Long) As Long",
            "#End If",
            "",
            "Private Const GWL_STYLE As Long = -16",
            "Private Const WS_CAPTION As Long = &HC00000",
            "Private Const WS_THICKFRAME As Long = &H40000",
            "",
        ])
    
    lines.extend([
        "' ============================================",
        "'  Form Initialization",
        "' ============================================",
        "Private Sub UserForm_Initialize()",
        f"    Me.Caption = \"{title}\"",
        f"    Me.Width = {width}",
        f"    Me.Height = {height}",
    ])
    
    if bg_color_vba:
        lines.append(f"    {bg_color_vba}")
    
    # Add borderless window code
    if is_borderless:
        lines.extend([
            "    ",
            "    ' Remove title bar and border using Windows API",
            "    #If VBA7 Then",
            "        Dim hWnd As LongPtr",
            "    #Else",
            "        Dim hWnd As Long",
            "    #End If",
            "    Dim lStyle As Long",
            "    hWnd = FindWindow(\"ThunderDFrame\", Me.Caption)",
            "    If hWnd <> 0 Then",
            "        lStyle = GetWindowLong(hWnd, GWL_STYLE)",
            "        lStyle = lStyle And Not WS_CAPTION ' Remove title bar",
            "        lStyle = lStyle And Not WS_THICKFRAME ' Remove border",
            "        SetWindowLong hWnd, GWL_STYLE, lStyle",
            "    End If",
        ])
    
    # Add background image loading code
    if bg_image:
        lines.append(f"    ")
        lines.append(f"    On Error Resume Next")
        lines.append(f"    Me.bgImage.Picture = LoadPicture(ThisWorkbook.Path & \"\\{bg_image}\")")
        lines.append(f"    Me.bgImage.PictureSizeMode = 3 ' Stretch")
        lines.append(f"    ' Set background to cover entire window")
        lines.append(f"    Me.bgImage.Left = -10")
        lines.append(f"    Me.bgImage.Top = -10")
        lines.append(f"    Me.bgImage.Width = Me.InsideWidth + 50")
        lines.append(f"    Me.bgImage.Height = Me.InsideHeight + 50")
        lines.append(f"    Me.bgImage.ZOrder 1 ' Send to back")
        lines.append(f"    On Error GoTo 0")
    
    lines.extend([
        "End Sub",
        "",
    ])

    for btn in config['buttons']:
        name = btn['name']
        module = btn.get('module', '')
        
        # Skip decorative controls
        if name.startswith("lbl") or name.startswith("btnSeparator"):
            continue

        lines.append(f"Private Sub {name}_Click()")

        if module == "Unload Me" or name == "btnExit":
            lines.append("    Unload Me")
        elif module:
            lines.append("    On Error Resume Next")
            lines.append(f"    Call {module}")
            lines.append("    On Error GoTo 0")
        else:
            lines.append("    ' TODO: Add functionality")

        lines.append("End Sub")
        lines.append("")

    return '\n'.join(lines)


def add_launcher_macro(wb, auto_start=True):
    """Add Workbook_Open auto-start + manual launcher macro"""
    print("[5/6] Adding launcher macros...")
    project = wb.VBProject

    if auto_start:
        twb = project.VBComponents("ThisWorkbook")
        code = (
            "Private Sub Workbook_Open()\n"
            "    On Error Resume Next\n"
            "    frmMain.Show\n"
            "    On Error GoTo 0\n"
            "End Sub\n"
        )
        twb.CodeModule.AddFromString(code)

    launcher = project.VBComponents.Add(1)  # StdModule
    launcher.Name = "Launcher"
    launcher.CodeModule.AddFromString(
        "Option Explicit\n\n"
        "Sub ShowForm()\n"
        "    frmMain.Show\n"
        "End Sub\n"
    )


def save_workbook(wb, filepath, base_dir='.'):
    """Save as .xlsm format (xlOpenXMLWorkbookMacroEnabled = 52)"""
    if not os.path.isabs(filepath):
        filepath = os.path.abspath(os.path.join(base_dir, filepath))

    if os.path.exists(filepath):
        try:
            os.remove(filepath)
        except PermissionError:
            base, ext = os.path.splitext(filepath)
            filepath = f"{base}_{int(time.time())}{ext}"

    print(f"[6/6] Saving to {filepath}")
    wb.SaveAs(filepath, FileFormat=52)
    return filepath


def main():
    parser = argparse.ArgumentParser(
        description='VBA Form Application Generator V2 - Modern Edition',
        epilog='Example: python create_vba_app.py --config config.json'
    )
    parser.add_argument('--config', required=True, help='Config file path (.json)')
    parser.add_argument('--output', help='Output filename (overrides config)')
    args = parser.parse_args()

    print("=" * 56)
    print("  VBA Form Generator V2  -  Modern UI Edition")
    print("=" * 56)

    config = load_config(args.config)
    base_dir = os.path.dirname(os.path.abspath(args.config))

    if args.output:
        config['output_file'] = args.output

    excel = None
    wb = None

    try:
        excel = create_excel_app()
        wb = create_workbook(excel)
        add_vba_modules(wb, config['modules'], base_dir)
        create_userform(wb, config, base_dir)
        add_launcher_macro(wb, config.get('auto_start', True))
        output = save_workbook(wb, config['output_file'], base_dir)

        print("\n" + "=" * 56)
        print("  SUCCESS!")
        print(f"  File: {output}")
        print("=" * 56)

    except Exception as e:
        print(f"\n  ERROR: {e}")
        import traceback
        traceback.print_exc()

    finally:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()
        import pythoncom
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
