# AuditFlow

Windows 本地香港审计资料整理与数据填列工作台。客户资料留在本机，软件目前属于辅助工具初版。

## 功能

- 客户资料库：整包资料放入 `客户资料` 文件夹，按名称、表名与有限文字摘要建议分类；支持人工改类、处理状态和文件更新提示。
- 客户期末余额导入、账户映射、管理账 TB 草稿与独立明细核对。
- Excel 保存值与部分公式检查、来源追踪、选定加总及舍入尾差检查。
- DOCX 明确金额位置填列、比较年度保护、Todd 修订及来源记录导出。

## 安装与启动

Windows 安装包见仓库 Releases，内含 Python；需要系统安装 Microsoft Edge WebView2 Runtime。安装后生成桌面入口。安装版的数据保存在 `%LOCALAPPDATA%\AuditFlow\data`，不会自动迁移源码目录下已有的资料和数据库。升级、卸载不会主动删除此数据目录。首次发行尚未进行代码签名。

构建安装包：安装本文件所列 Python 依赖、PyInstaller 和 Inno Setup 6，运行 `./packaging/build.ps1`。输出为 `dist/AuditFlow-Setup-0.1.0.exe` 与校验文件。

需要 Windows、Python 3.11+ 和 Microsoft Edge WebView2 Runtime；独立窗口使用 pywebview，窗口按钮与导航栏合并为一行。左侧红、黄、绿按钮分别关闭、最小化、最大化／还原；拖动导航栏空白处移动窗口，双击空白处最大化／还原。右下角可拖动调整窗口大小。更新后关闭旧窗口，再从桌面图标打开。

```powershell
python -m pip install -r requirements.txt
python start_workbench.py
```

仅启动本机服务：`python server.py --port 8776`，地址为 `http://127.0.0.1:8776`。
演示样例启动：`python docs/run_demo.py`，地址为 `http://127.0.0.1:8777`。

## 使用

1. 将客户资料包复制到项目目录下的 `客户资料`，每个客户／年度可使用一个子文件夹。压缩包先解压。
2. 在客户资料库检查分类，按需要修改类别和处理状态。界面显示期间每 10 秒检查更新，返回界面时重新检查。
3. Excel 点击“用于 TB／明细”，确认实体、期间、币种、金额列及取数范围，然后映射账户。
4. FS 填列页选择 working 和 DOCX 模板，逐项绑定本期金额，核对后导出预览副本。

状态保存在本机 `workspace.sqlite3`，输出位于 `outputs`。内容更新会使旧确认失效并重置处理状态；来源原件不会被改写。

## 当前范围

自动分类是建议；扫描件、图片及复杂资料需要人工查看。总账流水推导期末余额、审计调整、审定 TB、完整专项底稿、原 Excel 写回和整份 FS 勾稽尚未自动实现。
公式重算仅支持部分运算；不支持的公式须人工核实。旧 DOC 不能直接精确填列，需转为受支持的 DOCX。导出为待复核预览，最终排版及审计判断需要人工检查。

## 测试

```powershell
python -m unittest discover -s tests -v
```

仓库仅包含程序、测试与合成演示样例。客户资料、历史提取内容、原始文件目录清单、工作数据库、缓存和输出均不纳入版本控制。
