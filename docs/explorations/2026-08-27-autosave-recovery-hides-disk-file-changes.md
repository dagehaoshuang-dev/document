# autosave 会话恢复会吞掉"改磁盘文件"的修改（排查经验）

日期：2026-08-27
相关：`lib/history/session.ts`、`lib/history/autosave.ts`、`lib/history/db.ts`、`index.ts`
触发场景：改 `public/*.pptx`（如改主题字体）后刷新编辑器，页面仍是旧版。

## 症状

1. 磁盘源文件确实改了：本地 WPS / Office 打开正常（字体、内容都对）。
2. 编辑器里刷新 N 次都还是旧版：字体一直是旧 Calibri、乱码依旧。
3. `getDefaultFontFamily()` 返回新值（如 Arial）——但它只是"新建默认字体"配置，**不代表文档实际用的字体**，容易误导排查。

## 根因

编辑器打开文档后会把 `?saved=<id>` 写进地址栏（`stampDocumentIdInUrl`），并持续把文档快照写进 IndexedDB（autosave）。之后**每次刷新，`index.ts` 都优先从 IndexedDB 恢复旧快照**（注释原话："A stored snapshot wins over every other way of opening"），根本不重新 fetch 磁盘文件。所以：

- 磁盘文件改得再对也没用——编辑器读的是 IndexedDB 里的旧快照。
- 判断命中：**看编辑器地址栏 URL 是否带 `?saved=`**。

## 修法（禁用会话恢复）

改 `lib/history/session.ts` 的 `stampDocumentIdInUrl`，不写 `?saved=`（函数体置空即可）。改完 vite 热重载 + iframe 重载，刷新改走 `?src=` 重新 fetch 磁盘文件，立即生效。

```ts
export function stampDocumentIdInUrl(docId: string): void {
  // Autosave session recovery is disabled: keep the URL untouched so a reload
  // re-fetches the source file (`?src=`) instead of restoring an IndexedDB
  // snapshot that may hold an older copy of the document.
  void docId;
  if (typeof window === 'undefined') return;
}
```

## 排查顺序建议

**"改文件不生效" → 先怀疑 autosave 会话恢复，不要先怀疑文件或字体注册表。**
在 OnlyOffice + 自建 autosave 的项目里，这条优先于所有"文件解析/字体映射"类猜测。

## 代价

禁用后，刷新不再恢复到上次编辑状态（丢失"关网页重开接着改"的便利）。要恢复便利，需在保存时主动落盘快照并重新接入 `?saved=`。
