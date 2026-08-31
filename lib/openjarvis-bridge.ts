/**
 * OpenJarvis 画布桥接：让 document 编辑页在 OpenJarvis 画布 iframe 里被直接加载，
 * 而不是再套一个 office-work viewer 壳。
 *
 * 目标（spec 2026-08-31-direct-editor-bridge-design）：
 *   画布 Iframe → document 编辑页（本工程）→ frameEditor（OnlyOffice 引擎，同源直调）
 *   AI 的 call_iframe(region, "office.<method>", [...]) 直接命中本页 window.office，
 *   不再经过 office-work 那套 viewer 壳 + postMessage 跨窗。
 *
 * 这个模块只在「画布嵌入模式」下由 index.ts 调用（有 window.JarvisSDK，或 query 带 canvas 标记）。
 * 它做：
 *   1) 画布桥握手（hello → ack → skill-data-token），拿 JarvisSDK；
 *   2) 从 skill-data 读 current.<ext>（revision 锁），经 openLocalFile 打开；
 *   3) 挂 window.office.* 命令面（基础版：status / exec / getState，POC 先通基础闭环）；
 *   4) 编辑器 aiEvent → dataModelUpdate(/interaction) 双工。
 *
 * 依赖的 document 原语：
 *   - lib/document.ts openLocalFile(file)
 *   - lib/converter.ts loadEditorApi()
 *   - lib/agent-plugin/editor-bridge.ts getEditorApi() / requireEditorApi()
 *   - lib/onlyoffice-editor.ts requestSaveDocument()
 */

import { openLocalFile } from './document';
import { loadEditorApi } from './converter';
import { getEditorApi } from './agent-plugin/editor-bridge';
import { requestSaveDocument } from './onlyoffice-editor';

/** skill-data 通道上限（与 office-work project-file-store 一致）。 */
const SKILL_DATA_LIMIT = 32 * 1024 * 1024;

/** JarvisSDK 全局（平台经 errorShim=1 注入，token 由父页握手下发）。 */
interface JarvisSdkLike {
  fetchData: (path: string) => Promise<ArrayBuffer>;
  saveData: (path: string, bytes: ArrayBuffer) => Promise<void>;
  fetchDataVersioned: (path: string) => Promise<{ bytes: ArrayBuffer; etag: string }>;
  saveDataVersioned: (
    path: string,
    bytes: ArrayBuffer,
    opts: { ifMatch?: string },
  ) => Promise<{ etag: string }>;
}

declare global {
  interface Window {
    JarvisSDK?: JarvisSdkLike;
    office?: unknown;
  }
}

/** 从 URL query 读参数（项目名 / 画布标记）。 */
function queryParam(name: string): string | null {
  return new URLSearchParams(window.location.search).get(name);
}

/** 是否处于画布嵌入模式：有 JarvisSDK，或 query 带 canvas=1 / office=1 显式标记。 */
export function isCanvasEmbedded(): boolean {
  const flagged = queryParam('canvas') === '1' || queryParam('office') === '1';
  return Boolean(window.JarvisSDK) || flagged;
}

/** 画布桥：hello → ack → skill-data-token 握手，拿到 JarvisSDK 后 resolve。 */
function waitForCanvasBridge(timeoutMs = 6000): Promise<void> {
  return new Promise((resolve) => {
    let done = false;
    const finish = (): void => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      window.removeEventListener('message', onMessage);
      resolve();
    };
    const timer = setTimeout(finish, timeoutMs);
    const onMessage = (e: MessageEvent): void => {
      const m = e.data;
      if (m && m.jarvis === 1 && m.type === 'skill-data-token' && m.value?.token) finish();
    };
    window.addEventListener('message', onMessage);
    // 握手探测：发 hello，父页支持则回 ack + 补发 skill-data-token 帧
    try {
      window.parent.postMessage({ jarvis: 1, type: 'hello', spotlight: true }, '*');
    } catch {
      finish();
    }
  });
}

/** 从 query 读项目名（画布 Iframe src 里带 &project=xxx）。 */
function projectName(): string | null {
  return queryParam('project');
}

/** 读 skill-data 的 current.<ext> + revision.json → File 对象。 */
async function loadProjectFile(sdk: JarvisSdkLike, project: string, kind: 'pptx' | 'docx'): Promise<File | null> {
  const currentPath = `/${project}/current.${kind}`;
  const metaPath = `/${project}/revision.json`;
  let bytes: ArrayBuffer | null = null;
  let title = `${project}.${kind}`;
  try {
    const current = await sdk.fetchData(currentPath);
    bytes = current;
  } catch {
    bytes = null;
  }
  if (!bytes) return null;
  try {
    const meta = await sdk.fetchData(metaPath);
    const parsed = JSON.parse(new TextDecoder().decode(meta)) as { title?: string };
    if (typeof parsed.title === 'string' && parsed.title.trim()) title = parsed.title;
  } catch {
    // revision.json 缺失/损坏 → 用默认名
  }
  const mime =
    kind === 'docx'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      : 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  return new File([bytes], title.endsWith(`.${kind}`) ? title : `${title}.${kind}`, { type: mime });
}

/** office.* 命令面（POC：status / getState / exec / save，同源直调 frameEditor + 保存回 skill-data）。 */
/* eslint-disable @typescript-eslint/no-explicit-any */
function makeOfficeApi(ctx: {
  getRevision: () => number;
  getEtag: () => string | null;
  save: (targetExt?: string) => Promise<number>;
}): Record<string, unknown> {
  const result = (op: string, error?: unknown): Record<string, unknown> => ({
    ok: !error,
    op,
    error: error ? sanitizeError(op, error) : undefined,
  });

  return {
    async status(): Promise<Record<string, unknown>> {
      const api = getEditorApi();
      return {
        ok: true,
        op: 'status',
        data: {
          revision: ctx.getRevision(),
          etag: ctx.getEtag(),
          editorReady: Boolean(api),
          embedded: isCanvasEmbedded(),
        },
      };
    },
    async getState(): Promise<Record<string, unknown>> {
      const api = getEditorApi();
      if (!api) return result('getState', new Error('OnlyOffice editor not ready'));
      try {
        const selType =
          typeof (api as any).pluginMethod_GetSelectionType === 'function'
            ? String((api as any).pluginMethod_GetSelectionType())
            : 'n/a';
        const selText =
          typeof (api as any).pluginMethod_GetSelectedText === 'function'
            ? String((api as any).pluginMethod_GetSelectedText())
            : '';
        return { ok: true, op: 'getState', data: { selectionType: selType, selectedText: selText } };
      } catch (error) {
        return result('getState', error);
      }
    },
    async exec(method: string, args: unknown[] = []): Promise<Record<string, unknown>> {
      if (typeof method !== 'string' || !method) return result('exec', new Error('需要 method'));
      if (!Array.isArray(args)) return result('exec', new Error('args 需为数组'));
      const api = getEditorApi();
      if (!api) return result('exec', new Error('OnlyOffice editor not ready'));
      const fn = (api as any)[method];
      if (typeof fn !== 'function') return result('exec', new Error(`no such method on editor: ${method}`));
      try {
        const ret = fn.apply(api, args);
        let serialized: unknown = null;
        if (ret == null) serialized = null;
        else if (typeof ret === 'object') {
          try {
            serialized = JSON.stringify(ret);
          } catch {
            serialized = '[unserializable]';
          }
        } else serialized = String(ret);
        return { ok: true, op: 'exec', data: { method, ret: serialized } };
      } catch (error) {
        return result('exec', error);
      }
    },
    /** 显式保存：把编辑器当前 bytes 条件写回 skill-data current.<ext>，推进 revision（ETag 条件写）。 */
    async save(targetExt?: string): Promise<Record<string, unknown>> {
      try {
        const revision = await ctx.save(targetExt);
        return { ok: true, op: 'save', data: { revision } };
      } catch (error) {
        return result('save', error);
      }
    },
  } as Record<string, unknown>;
}

function sanitizeError(op: string, cause: unknown): string {
  const message = cause instanceof Error ? cause.message : String(cause ?? '');
  if (message && message.length > 0 && /[\u4e00-\u9fff]/.test(message)) return message;
  return message ? `${op} 失败：${message}` : `${op} 失败`;
}

/** 保存 current.<ext> 到 skill-data（revision 锁：`saveDataVersioned(ifMatch)` 条件写）。
 *  返回 {revision, etag}（新 ETag 供下次保存作为 If-Match）。revision 真值来自 revision.json，不是 etag。 */
async function saveToSkillData(
  sdk: JarvisSdkLike,
  project: string,
  kind: 'pptx' | 'docx',
  ifMatch: string,
  currentRevision: number,
  targetExt: string,
): Promise<{ revision: number; etag: string }> {
  const file = await requestSaveDocument(targetExt, {});
  const bytes = await file.arrayBuffer();
  if (bytes.byteLength > SKILL_DATA_LIMIT) throw new Error('文档超过 skill-data 通道上限');
  const currentPath = `/${project}/current.${kind}`;
  // 条件写：ifMatch 不匹配（他人已改）→ SDK 抛 revision_conflict，禁止覆盖。
  const saved = await sdk.saveDataVersioned(currentPath, bytes, { ifMatch });
  const newEtag = saved.etag;
  // 推进 revision.json（保留 kind/createdAt/title/sourceName）。
  const metaPath = `/${project}/revision.json`;
  let previousMeta: Record<string, unknown> = {};
  try {
    const raw = await sdk.fetchData(metaPath);
    previousMeta = JSON.parse(new TextDecoder().decode(raw));
  } catch {
    // revision.json 缺/坏 → 用默认
  }
  const meta: Record<string, unknown> = {
    schemaVersion: 1,
    revision: currentRevision + 1,
    title: previousMeta.title ?? `${project}.${kind}`,
    sourceName: previousMeta.sourceName,
    kind: previousMeta.kind ?? kind,
    createdAt: previousMeta.createdAt,
    updatedAt: new Date().toISOString(),
  };
  await sdk.saveData(
    metaPath,
    new TextEncoder().encode(JSON.stringify(meta)).buffer as ArrayBuffer,
  );
  return { revision: currentRevision + 1, etag: newEtag };
}

/**
 * 初始化 OpenJarvis 画布桥。在 index.ts 检测到 isCanvasEmbedded() 时调用。
 * 幂等：只在首次调用时运行。
 */
export async function initOpenJarvisBridge(): Promise<void> {
  const project = projectName();
  if (!project) return;

  await waitForCanvasBridge();
  const sdk = window.JarvisSDK;
  if (!sdk) return;

  // 读 skill-data 并打开文档
  const kindRaw = queryParam('kind');
  const kind: 'pptx' | 'docx' = kindRaw === 'docx' ? 'docx' : 'pptx';
  const file = await loadProjectFile(sdk, project, kind);
  if (!file) {
    console.warn('[openjarvis-bridge] 未找到 skill-data 项目文件', project, kind);
    return;
  }
  await loadEditorApi();
  await openLocalFile(file);

  // 从 skill-data 读 revision + etag（真值来源：revision.json + current.<ext> 的 ETag）。
  let revision = 0;
  let etag: string | null = null;
  try {
    const metaRaw = await sdk.fetchData(`/${project}/revision.json`);
    const meta = JSON.parse(new TextDecoder().decode(metaRaw)) as { revision?: number };
    if (Number.isSafeInteger(meta.revision) && (meta.revision as number) >= 0) revision = meta.revision as number;
  } catch {
    // revision.json 缺/坏 → revision=0，首次保存即可推进
  }
  try {
    const ver = await sdk.fetchDataVersioned(`/${project}/current.${kind}`);
    etag = ver.etag;
  } catch {
    // current 缺/无 ETag → 无 ifMatch，保存走无锁分支（首次创建）
  }
  let currentRevision = revision;
  let currentEtag = etag;

  // 挂 window.office.*（同源直调 frameEditor + 保存回 skill-data）。
  window.office = makeOfficeApi({
    getRevision: () => currentRevision,
    getEtag: () => currentEtag,
    save: async (targetExt?: string): Promise<number> => {
      const ext = targetExt === 'docx' || targetExt === 'pptx' ? targetExt : kind;
      const saved = await saveToSkillData(sdk, project, kind, currentEtag ?? '', currentRevision, ext);
      currentRevision = saved.revision;
      currentEtag = saved.etag;   // 保存后更新 ETag，下次保存用作 If-Match（防 stale ifMatch 误判 conflict）
      return saved.revision;
    },
  });
}
