# 开发与发布注意事项

## 从源码构建

1. 安装依赖：

   ```bash
   npm install
   ```

2. 构建插件：

   ```bash
   npm run build
   ```

3. 将 `dist` 中生成的文件复制到 Obsidian vault 的插件目录：

   ```text
   dist/main.js
   dist/docx-engine.cjs
   dist/manifest.json
   dist/styles.css
   ```

## 安装与升级

### 通过发布 zip 安装

1. 从 GitHub Release 页面下载 `obsidian-word-reader-X.Y.Z.zip`。
2. 在 vault 中创建插件目录：

   ```text
   YourVault/.obsidian/plugins/word-reader/
   ```

3. 将 zip 解压到该目录。四个文件都必须直接位于插件目录根部：

   ```text
   YourVault/.obsidian/plugins/word-reader/main.js
   YourVault/.obsidian/plugins/word-reader/docx-engine.cjs
   YourVault/.obsidian/plugins/word-reader/manifest.json
   YourVault/.obsidian/plugins/word-reader/styles.css
   ```

   `docx-engine.cjs` 是经过校验的 DOCX 渲染引擎，必须与 `main.js` 放在同一
   目录，不要重命名、移动或删除。

4. 重启 Obsidian，或重新加载第三方插件。
5. 在 Obsidian 设置中启用第三方插件，并启用 Office Reader。
6. 打开一个 `.docx` 文件验证引擎加载。若 `docx-engine.cjs` 缺失或损坏，
   插件会显示加载错误，而不是继续渲染。

### 安装单个发布文件

从同一个 GitHub Release 下载全部四个文件（`main.js`、`docx-engine.cjs`、
`manifest.json`、`styles.css`），直接放入上面的插件目录。不要混用不同版本的
文件。

> 如果通过 Obsidian 社区插件浏览器安装或更新，它可能只下载 `main.js`、
> `manifest.json` 和 `styles.css`。请从同一个 Release 下载
> `docx-engine.cjs` 并放到 `main.js` 旁边，或改用发布 zip。

### 从旧版本升级

1. 在 Obsidian 中停用 Office Reader。
2. 用新版本的文件整体替换插件目录中的四个文件。
3. 如果存在 `data.json`，请保留。
4. 重新启用插件。

不要只替换 `main.js`；`main.js` 和 `docx-engine.cjs` 必须来自同一个版本。

## 本地发布

创建并校验本地发布包：

```bash
npm run release
```

该命令会执行 TypeScript 检查、构建插件、校验版本一致性、生成可安装 zip，
并提取当前版本的 changelog 作为 release notes。

预期输出：

```text
release/obsidian-word-reader-X.Y.Z.zip
release/CHANGELOG-X.Y.Z.md
```

zip 根目录只包含 Obsidian 需要的文件：

```text
main.js
docx-engine.cjs
manifest.json
styles.css
```

发布产物会被 Git 忽略，不需要提交。

## GitHub 自动发布

推送不带 `v` 前缀的版本 tag 后，GitHub Actions 会自动创建 release：

```bash
git tag 3.6.0
git push origin 2.5.0
```

workflow 会校验 tag 是否与 `package.json`、`manifest.json` 和
`package-lock.json` 一致，检查 500 KiB bundle 预算，然后构建插件、生成
`release/obsidian-word-reader-3.6.0.zip`、提取对应版本的 `CHANGELOG.md`
内容，并把 `main.js`、`docx-engine.cjs`、`manifest.json`、`styles.css`
和 zip 一起上传到 GitHub Release。

## 开发命令

安装依赖：

```bash
npm install
```

执行生产构建：

```bash
npm run build
```

执行 TypeScript 检查：

```bash
npm run typecheck
```

执行单元测试和集成测试：

```bash
npm test
```

构建后执行生产安全扫描：

```bash
npm run build
npm run security:scan
```

执行全部自动化质量门禁：

```bash
npm run check
```

打包后校验发布元数据和 zip 内容：

```bash
npm run release:zip
npm run release:check
```

首次提交社区插件目录和后续发布审查步骤见
[COMMUNITY_PLUGIN_CHECKLIST.md](COMMUNITY_PLUGIN_CHECKLIST.md)。

CI 和发布工作流使用 Node.js 20.19.0。
