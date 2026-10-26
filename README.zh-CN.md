# Office Reader

[English README](README.md) | [迭代路线](docs/ROADMAP.md)

> **上游项目：** 本项目维护于
> [`yulimfish/obsidian-word-reader`](https://github.com/yulimfish/obsidian-word-reader)，
> 基于上游项目
> [`qianwei4712/obsidian-word-reader`](https://github.com/qianwei4712/obsidian-word-reader)
> 开发，原作者保留原始工作的署名权。

Office Reader 是一个 Obsidian 桌面端插件，用于在 Obsidian 内直接打开
`.docx`、`.pptx` 和 `.xlsx` 文件，并以只读方式阅读 Office 文档。

插件不是 Office 编辑器，而是提供完全本地的阅读流程，同时保持原始文件不变。

需要 Obsidian Desktop 1.12.7 或更新版本。

## 功能

- 在 Obsidian 标签页中打开 `.docx`，支持标题、段落、列表、表格、图片、
  分页、目录大纲、搜索、缩放和纯文本/Markdown 复制。
- 在 Obsidian 标签页中本地打开 `.pptx`，支持文本、图片、常见形状、表格、
  主题、版式和母版，并提供缩略图、幻灯片搜索、演讲者备注、全屏阅读和
  摘要笔记。
- 在有界虚拟网格中打开 `.xlsx`，支持工作表标签、冻结窗格、合并单元格、
  基础样式、公式缓存结果、工作簿级搜索、隐藏表提示、TSV/Markdown 导出和
  摘要笔记。
- 为最近使用的 100 个 Office 文档恢复缩放、适配模式、导航状态和阅读位置。
- 支持插件界面在中文和英文之间切换。
- 支持创建与原 Office 文件链接的同名 Markdown 摘要笔记。
- 支持复制隐私安全诊断，用于兼容性反馈，不包含文档正文、内部 XML 或 vault
  绝对路径。

## 支持的文件

| 扩展名 | 状态 | 说明 |
| --- | --- | --- |
| `.docx` | 支持 | 通过 `docx-preview` 在 Obsidian 内渲染。 |
| `.pptx` | 支持 | 在 Obsidian 内本地渲染为只读演示文稿。 |
| `.xlsx` | 支持 | 在 Obsidian 内本地渲染为只读虚拟电子表格。 |
| `.doc` | 说明页面 | 在 Obsidian 中显示外部打开和 `.docx` 转换说明。 |
| `.xls` | 说明页面 | 在 Obsidian 中显示外部打开和 `.xlsx` 转换说明。 |
| `.xlsm` | 不注册 | 含宏工作簿不在支持范围内。 |

## 使用注意事项

- 将 `.docx`、`.pptx` 或 `.xlsx` 文件放入 vault 后，在文件列表中点击即可打开
  只读标签页。
- 插件是只读的，不会把修改保存回原始 Office 文件。
- `.xlsm` 不注册；`.doc` 和 `.xls` 会显示说明页面，而不是直接渲染。
- 加密或受密码保护的文档会显示专门的错误页面。
- 手动安装或升级时，四个发布文件（`main.js`、`docx-engine.cjs`、
  `manifest.json`、`styles.css`）必须成套使用。
- 工作簿超链接仅在用户明确点击并确认后打开。

详细使用、兼容性、安全和已知限制说明见
[docs/usage-notes.zh-CN.md](docs/usage-notes.zh-CN.md)。

## 开发注意事项

- 使用 `npm install` 安装依赖。
- 使用 `npm run check` 执行完整本地质量门禁。
- 使用 `npm run release:check` 校验发布元数据。
- 生产构建强制执行 `dist/main.js` 不超过 500 KiB 的体积预算。
- CI 和发布工作流使用 Node.js 20.19.0。
- 发布包只包含 `main.js`、`docx-engine.cjs`、`manifest.json` 和 `styles.css`。

开发、安装和发布细节见
[docs/development-release.zh-CN.md](docs/development-release.zh-CN.md)。

## 上游与协议说明

- 上游署名见上文；本仓库基于 MIT License 继续维护该项目。
- 运行时依赖的第三方协议清单见
  [docs/THIRD_PARTY_LICENSES.md](docs/THIRD_PARTY_LICENSES.md)。
- 稳定版范围、支持边界和手动测试清单维护在
  [docs/STABILITY.md](docs/STABILITY.md)。
