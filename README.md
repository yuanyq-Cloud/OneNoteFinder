# OneFinder — OneNote 全文搜索工具

一个轻量 Windows 桌面工具，通过 **OneNote COM API** 直接遍历所有笔记本页面并进行全文搜索，无需 Windows Search 索引。

## 功能

- 通过 OneNote COM API 获取所有笔记本 / 节 / 页面层次结构
- 对每页的完整 XML 内容做大小写不敏感全文匹配
- 返回"笔记本 › 节 › 页面"三级路径列表
- 双击或 Enter 键直接在 OneNote 中跳转到对应页面
- 搜索在后台线程运行，不阻塞 UI

## 前提条件

- Windows 10/11 x64
- **已安装 Microsoft OneNote 2016 或 Microsoft 365 OneNote 桌面版**（OneNote COM 服务器必须存在）
- .NET 8 Runtime（Windows）
- Visual Studio 2022 或 MSBuild 17+ 用于编译

## 构建

```bash
# 在仓库根目录
dotnet restore .\OneFinder\OneFinder.csproj
dotnet build .\OneFinder\OneFinder.csproj -c Release
```

或直接用 Visual Studio 打开 `OneFinder\OneFinder.csproj` 后按 F5/Ctrl+F5。

## 使用

1. 启动 `OneFinder.exe`（OneNote 桌面版须已安装）
2. 在搜索框输入关键词，按 Enter 或点击"搜索"
3. 等待扫描完成（底部状态栏显示当前扫描进度）
4. 双击结果列表中的条目，OneNote 会自动跳转到对应页面

## 注意事项

- 受密码保护的节会被自动跳过
- 回收站中的页面同样被跳过
- 笔记本越多、页面越多，首次扫描越慢；建议关键词尽量精确

## 项目结构

```
<repo-root>/
├── README.md
├── installer/
│   └── Package.wxs          # MSI 安装包定义
└── OneFinder/
    ├── OneFinder.csproj     # 项目文件（net8.0-windows, x64）
    ├── Program.cs           # 入口点
    ├── MainForm.cs          # WinForms 主窗口 + 搜索 UI
    ├── MainForm.Designer.cs # WinForms 设计器脚手架
    ├── OneNoteService.cs    # OneNote COM API 封装 + 搜索逻辑
    ├── USER_GUIDE.md        # 用户指南
    └── CHANGELOG.md         # 更新日志
```
