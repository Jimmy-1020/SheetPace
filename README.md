# SheetPace

Windows Excel 插件：筛选值计数、鼠标行列光影，以及可调整的颜色与透明度。

## 安装

下载 [SheetPace-Setup-1.0.3.exe](dist/SheetPace-Setup-1.0.3.exe)，保存工作簿并关闭所有 Excel 进程后双击安装。当前用户安装，无需管理员权限。重新打开 Excel，在功能区找到 **SheetPace**。已安装旧版时，直接运行 1.0.3 覆盖更新；请使用此 Setup 文件安装。

- Windows 10 / 11，Microsoft Excel 桌面版，.NET Framework 4.8。
- 使用 AnyCPU COM 类库，在 32 / 64 位注册视图分别注册。目标 Office 2016 及以上；目前实际测试为 64 位 Microsoft 365，其他版本和 32 位 Office 需要回归验证。
- 无 VSTO、Office PIA 或 NuGet 依赖。Mac、网页版 Excel、WPS 不适用。
- 安装位置：`%LOCALAPPDATA%/Programs/SheetPace`。在 Windows“已安装的应用”中卸载，或运行安装目录下的 `SheetPace.Uninstall.exe /uninstall`。

## 使用

**筛选计数**：打开列标题的筛选按钮，受支持的原生下拉列表会在各项旁显示次数。也可选中目标列中的一个单元格，再点击 **SheetPace → 计数筛选**，搜索、勾选并应用筛选。应用只改变当前列的条件，同一区域其他列的筛选保留。

次数统计当前自动筛选区域 / Excel 表格的整列数据，排除标题，包含筛选隐藏的行；空白单独计数。没有筛选区域时使用当前连续数据区域，并在点击“应用”后创建自动筛选。一个工作表只能有一个普通范围的自动筛选；已有筛选在其他区域时会提示，避免改错数据。

![鼠标光影图标](docs/images/hover-icon.png) **鼠标光影**：点击带有网格与指针图标的“鼠标光影”按钮，可快速开关十字行列光影。默认使用 **点击选中** 模式，以当前选中的单元格为中心，移动鼠标不改变光影位置；点击其他单元格或用键盘改变活动单元格时，光影随选中格更新。也可使用 **鼠标跟随** 模式，鼠标移到哪个格子，光影就跟随到该格，无需点击。覆盖层不接收鼠标操作，也不修改内容、选区或单元格格式。离开 Excel、进入对话框或拖动时隐藏；选中格不在可视区域时不显示。

![光影设置图标](docs/images/settings-icon.png) **光影设置**：点击带有网格与齿轮图标的 **光影设置**，切换点击选中 / 鼠标跟随模式，选择颜色、调整透明度并预览。100% 透明时光影不可见。保存后模式、颜色和透明度会在下次启动时保留；取消不会应用修改。旧配置首次升级默认采用点击选中模式，保留已有颜色、透明度和启用开关。

![原生筛选数量标签](docs/images/native-filter.png)

![半透明行列光影](docs/images/hover.png)

## 原生菜单的兼容性边界

Excel 没有公开接口让插件修改原生筛选列表的渲染。本插件采用 UI Automation 读取可见项目，再用鼠标穿透的计数标签增强显示。它依赖 Office 可访问性树、语言、主题与显示缩放，不能保证所有版本的原生菜单均可识别。识别失败时不会改写菜单，请使用功能区的计数筛选窗口。

数字、日期采用 Excel 的统一列显示格式统计和筛选。混合单元格格式按原始值显示；复杂格式、日期层级、本地化、多显示器 DPI、拆分/冻结与超大数据区域需要按具体环境回归。Excel 值筛选最多支持 10000 个筛选值。现有颜色、比较表达式或日期分组条件无法准确转成多选框时，窗口会提示应用将替换本列条件。

## 源码构建

```powershell
python scripts/build.py
./dist/SheetPace.Tests.exe tests/output/core
./dist/SheetPace.ComAbiTests.exe
./dist/SheetPace.ComAbiTests.x86.exe
./dist/SheetPace-Setup-1.0.3.exe /selftest "$PWD/tests/output/installer"
./dist/SheetPace.ExcelTests.exe tests/output/excel
./dist/SheetPace.HoverTests.exe tests/output/selection --selection
./dist/SheetPace.HoverTests.exe tests/output/hover --packet
```

构建使用 Windows 自带 .NET Framework 编译器，Python 3.8 及以上，不联网、不恢复 NuGet。产物位于 `dist/`，安装器内嵌插件 DLL、独立界面进程和原生菜单识别进程，自动安装全部文件。`SHA256SUMS.txt` 可用于校验文件完整性。

`/selftest` 验证三个内嵌文件、COM 类和注册/卸载计划，**不写入注册表**。Excel 集成测试需要安装 Excel，仅读写自己创建的测试工作簿。`SheetPace.InstalledTests.exe` 验证安装后自动加载、真实文件启动、界面回调和正常退出；它仅打开指定工作簿的只读副本。`SheetPace.UiProbe.exe` 是交互显示测试，会暂时打开自己的 Excel 窗口并移动鼠标，结束后关闭该实例；不要在正在操作 Excel 时运行。

## 排障

若功能区未出现，在 Excel **文件 → 选项 → 加载项 → 管理 COM 加载项 → 转到** 中启用 SheetPace。1.0.1 修复了原版的 COM 加载接口和 Excel 退出异常；1.0.2 将鼠标坐标查询改为 Excel 内部采样并缓存，只传输坐标，光影改为窄条窗口，减少移动时的延迟。1.0.3 增加默认点击选中与鼠标跟随两种模式、专用图标及光影设置入口。界面及菜单识别在独立进程中运行。更新安装会备份并恢复明确匹配 SheetPace 的禁用记录；若仍被禁用，检查“禁用项目”并重新启用。组织安全策略可能限制 COM 加载项；本版安装程序未签名。

设置及诊断日志：`%LOCALAPPDATA%/SheetPace/`。卸载保留设置和日志。安装程序不会自动关闭 Excel。

更多说明：[需求与设计](docs/需求与设计.md)、[测试记录](docs/测试记录.md)。
