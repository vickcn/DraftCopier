"use client";

import Script from "next/script";
import { useEffect, useMemo, useRef, useState } from "react";
import HowItWorks from "./components/HowItWorks";

declare global {
  interface Window {
    gapi?: {
      load: (library: string, callback: () => void) => void;
    };
    google?: any;
  }
}

type PreviewPayload = {
  total_records: number;
  headers: string[];
  preview_first_row: string;
  template_html?: string;
  first_row?: Record<string, string>;
  preview_rows?: Array<Record<string, string>>;
  preview_pagination?: {
    page: number;
    page_size: number;
    total_pages: number;
  };
  sheet_names?: string[];
  selected_sheet?: string;
  cache_id?: string;
  cached_files?: {
    docx_name: string;
    xlsx_name: string;
  };
  detected_fields?: {
    email: string | null;
    cc?: string | null;
    bcc?: string | null;
    subject: string | null;
    attachments?: string[];
  };
};

type UploadState = "idle" | "uploading" | "done" | "error";

type DroppedFiles = {
  docx?: File;
  xlsx?: File;
};

type CachedUpload = {
  cacheId: string;
  docxName: string;
  xlsxName: string;
  font?: string;
  selectedSheet?: string;
};

type BatchFailedItem = {
  row?: number;
  field?: string;
  name?: string;
  message?: string;
};

type BatchSaveResponse = {
  status?: "ok" | "partial" | "failed";
  draft_count?: number;
  draft_ids?: string[];
  failed_items?: BatchFailedItem[];
};

type DriveKind = "docx" | "xlsx" | "attachment" | "folder";

type DriveFile = {
  id: string;
  name: string;
  mime_type: string;
  modified_time?: string;
  kind: DriveKind;
  is_google_workspace: boolean;
};

type PickerConfig = {
  enabled: boolean;
  client_id?: string | null;
  api_key?: string | null;
  app_id?: string | null;
  scope?: string | null;
  login_hint?: string | null;
};

type FolderBrowserPurpose = "lookup" | "create-parent";

const apiBase = process.env.NEXT_PUBLIC_API_BASE || "";
const uploadCacheKey = "draftcopier_upload_cache_v1";
const userKeyCacheKey = "draftcopier_user_key_v1";
const userEmailCacheKey = "draftcopier_user_email_v1";
const userNameCacheKey = "draftcopier_user_name_v1";
const previewPageSizeOptions = [20, 50, 100];
const pickerMimeTypes: Record<DriveKind, string[]> = {
  docx: [
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/vnd.google-apps.document",
  ],
  xlsx: [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    "application/vnd.google-apps.spreadsheet",
  ],
  attachment: [],
  folder: ["application/vnd.google-apps.folder"],
};
const fontOptions = [
  { label: "Sans Serif", value: "Sans Serif" },
  { label: "Serif", value: "Serif" },
  { label: "等寬字型", value: "等寬字型" },
  { label: "微軟正黑體", value: "微軟正黑體" },
  { label: "新細明體", value: "新細明體" },
  { label: "細明體", value: "細明體" },
  { label: "寬", value: "寬" },
  { label: "窄", value: "窄" },
  { label: "Comic Sans MS", value: "Comic Sans MS" },
  { label: "Garamond", value: "Garamond" },
  { label: "Georgia", value: "Georgia" },
  { label: "Tahoma", value: "Tahoma" },
  { label: "Trebuchet MS", value: "Trebuchet MS" },
  { label: "Verdana", value: "Verdana" },
];

const emailFieldCandidates = new Set([
  "email",
  "e-mail",
  "mail",
  "email address",
  "e-mail address",
  "電子郵件",
  "信箱",
]);

const ccFieldCandidates = new Set([
  "cc",
  "copy",
  "carbon copy",
  "抄送",
  "副本",
  "副本收件人",
]);

const bccFieldCandidates = new Set([
  "bcc",
  "blind carbon copy",
  "密件副本",
  "秘密副本",
]);

const subjectFieldCandidates = new Set([
  "subject",
  "email subject",
  "mail subject",
  "title",
  "subject line",
  "主旨",
  "標題",
  "信件主旨",
]);

const defaultTemplateHtml =
  '<div style="font-family: Arial, Helvetica, sans-serif; line-height: 1.6; color: #333;"><p>請在這裡編輯郵件內容</p></div>';
const placeholderPattern = /\{\{\s*([^{}]+?)\s*\}\}/g;

function stripTemplatePlaceholderMarkup(html: string): string {
  if (typeof document === "undefined") {
    return html.replace(/<span class="template-placeholder"[^>]*>([\s\S]*?)<\/span>/g, "$1");
  }
  const container = document.createElement("div");
  container.innerHTML = html;
  container.querySelectorAll("span.template-placeholder").forEach((node) => {
    const parent = node.parentNode;
    if (!parent) return;
    while (node.firstChild) {
      parent.insertBefore(node.firstChild, node);
    }
    parent.removeChild(node);
  });
  return container.innerHTML;
}

function decorateTemplatePlaceholders(html: string): string {
  if (typeof document === "undefined") return html;
  const container = document.createElement("div");
  container.innerHTML = stripTemplatePlaceholderMarkup(html);
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
  const textNodes: Text[] = [];
  let current = walker.nextNode();
  while (current) {
    if (current.nodeType === Node.TEXT_NODE) {
      textNodes.push(current as Text);
    }
    current = walker.nextNode();
  }

  for (const textNode of textNodes) {
    const text = textNode.textContent ?? "";
    placeholderPattern.lastIndex = 0;
    if (!placeholderPattern.test(text)) continue;
    placeholderPattern.lastIndex = 0;
    const fragment = document.createDocumentFragment();
    let lastIndex = 0;
    let match: RegExpExecArray | null;
    while ((match = placeholderPattern.exec(text)) !== null) {
      const [fullMatch] = match;
      const start = match.index;
      if (start > lastIndex) {
        fragment.appendChild(document.createTextNode(text.slice(lastIndex, start)));
      }
      const span = document.createElement("span");
      span.className = "template-placeholder";
      span.textContent = fullMatch;
      span.setAttribute("data-placeholder", fullMatch);
      fragment.appendChild(span);
      lastIndex = start + fullMatch.length;
    }
    if (lastIndex < text.length) {
      fragment.appendChild(document.createTextNode(text.slice(lastIndex)));
    }
    textNode.parentNode?.replaceChild(fragment, textNode);
  }

  return container.innerHTML;
}

function findHeader(headers: Array<string | number>, candidates: Set<string>): string | null {
  const normalized = headers.map((h) => String(h).trim());
  const lowerMap = new Map(normalized.map((h) => [h.toLowerCase(), h]));
  for (const candidate of candidates) {
    const match = lowerMap.get(candidate);
    if (match) return match;
  }
  for (const header of normalized) {
    const lower = header.toLowerCase();
    for (const candidate of candidates) {
      if (lower.includes(candidate)) return header;
    }
  }
  return null;
}


function classifyFiles(files: FileList | File[]): DroppedFiles {
  const result: DroppedFiles = {};
  Array.from(files).forEach((file) => {
    const lower = file.name.toLowerCase();
    if (lower.endsWith(".docx")) {
      result.docx = file;
    } else if (lower.endsWith(".xlsx")) {
      result.xlsx = file;
    }
  });
  return result;
}

function renderTemplateWithRow(templateHtml: string, rowData?: Record<string, string>): string {
  const cleanTemplateHtml = stripTemplatePlaceholderMarkup(templateHtml);
  if (!rowData) return cleanTemplateHtml;
  return cleanTemplateHtml.replace(/\{\{\s*(.*?)\s*\}\}/g, (_, rawKey: string) => {
    const key = String(rawKey || "").replace(/<[^>]+>/g, "").trim();
    return rowData[key] ?? `{{${key}}}`;
  });
}

function collectTemplateFieldKeys(templateHtml: string): Set<string> {
  const keys = new Set<string>();
  const cleanTemplateHtml = stripTemplatePlaceholderMarkup(templateHtml);
  cleanTemplateHtml.replace(/\{\{\s*(.*?)\s*\}\}/g, (_, rawKey: string) => {
    const key = String(rawKey || "").replace(/<[^>]+>/g, "").trim();
    if (key) {
      keys.add(key);
    }
    return "";
  });
  return keys;
}

export default function Home() {
  const [docxFile, setDocxFile] = useState<File | null>(null);
  const [xlsxFile, setXlsxFile] = useState<File | null>(null);
  const [driveDocx, setDriveDocx] = useState<DriveFile | null>(null);
  const [driveXlsx, setDriveXlsx] = useState<DriveFile | null>(null);
  const [selectedDriveAttachments, setSelectedDriveAttachments] = useState<DriveFile[]>([]);
  const [selectedDriveFolder, setSelectedDriveFolder] = useState<DriveFile | null>(null);
  const [cloudUploadFolder, setCloudUploadFolder] = useState<DriveFile | null>(null);
  const [uploadedCloudFolderFiles, setUploadedCloudFolderFiles] = useState<DriveFile[]>([]);
  const [selectedLocalAttachmentFiles, setSelectedLocalAttachmentFiles] = useState<File[]>([]);
  const [cachedUpload, setCachedUpload] = useState<CachedUpload | null>(null);
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState<UploadState>("idle");
  const [error, setError] = useState<string | null>(null);
  const [preview, setPreview] = useState<PreviewPayload | null>(null);
  const [isDragActive, setIsDragActive] = useState(false);
  const [selectedFont, setSelectedFont] = useState(fontOptions[0].value);
  const [draftStatus, setDraftStatus] = useState<"idle" | "saving" | "done" | "error">("idle");
  const [draftMessage, setDraftMessage] = useState<string | null>(null);
  const [draftFailedItems, setDraftFailedItems] = useState<BatchFailedItem[]>([]);
  const [attachmentsDir, setAttachmentsDir] = useState("");
  const [templateHtml, setTemplateHtml] = useState(defaultTemplateHtml);
  const [templateDirty, setTemplateDirty] = useState(false);
  const [exportingTemplate, setExportingTemplate] = useState(false);
  const [reloadingTemplate, setReloadingTemplate] = useState(false);
  const [textColor, setTextColor] = useState("#1c2230");
  const [gmailEmail, setGmailEmail] = useState<string | null>(null);
  const [gmailName, setGmailName] = useState<string | null>(null);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [driveBrowserKind, setDriveBrowserKind] = useState<DriveKind | null>(null);
  const [folderBrowserPurpose, setFolderBrowserPurpose] = useState<FolderBrowserPurpose>("lookup");
  const [driveFiles, setDriveFiles] = useState<DriveFile[]>([]);
  const [driveQuery, setDriveQuery] = useState("");
  const [driveLoading, setDriveLoading] = useState(false);
  const [driveError, setDriveError] = useState<string | null>(null);
  const [uploadingCloudFiles, setUploadingCloudFiles] = useState(false);
  const [creatingCloudFolder, setCreatingCloudFolder] = useState(false);
  const [isHostedFrontend, setIsHostedFrontend] = useState<boolean | null>(null);
  const [pickerConfig, setPickerConfig] = useState<PickerConfig | null>(null);
  const [pickerApiReady, setPickerApiReady] = useState(false);
  const [pickerIdentityReady, setPickerIdentityReady] = useState(false);
  const [step, setStep] = useState<1 | 2 | 3>(1);
  const [pendingAdvance, setPendingAdvance] = useState(false);
  const [openMenu, setOpenMenu] = useState<DriveKind | null>(null);
  const [showGuide, setShowGuide] = useState(false);
  const [previewPage, setPreviewPage] = useState(1);
  const [previewPageSize, setPreviewPageSize] = useState(previewPageSizeOptions[0]);

  const statusLabel: Record<UploadState, string> = {
    idle: "待命",
    uploading: "上傳中",
    done: "完成",
    error: "失敗",
  };

  const docxInputRef = useRef<HTMLInputElement>(null);
  const xlsxInputRef = useRef<HTMLInputElement>(null);
  const localAttachmentFolderInputRef = useRef<HTMLInputElement>(null);
  const cloudAttachmentInputRef = useRef<HTMLInputElement>(null);
  const editorRef = useRef<HTMLDivElement>(null);
  const restoreAttemptedRef = useRef<string | null>(null);
  const pickerTokenRef = useRef<string | null>(null);

  const persistCachedUpload = (nextCached: CachedUpload | null) => {
    setCachedUpload(nextCached);
    if (nextCached) {
      localStorage.setItem(uploadCacheKey, JSON.stringify(nextCached));
    } else {
      localStorage.removeItem(uploadCacheKey);
    }
  };

  const resetPreparedState = () => {
    setStatus("idle");
    setProgress(0);
    setError(null);
    setPreview(null);
    setDraftStatus("idle");
    setDraftMessage(null);
    setDraftFailedItems([]);
    setSelectedSheet("");
    setTemplateDirty(false);
    setPreviewPage(1);
    setPreviewPageSize(previewPageSizeOptions[0]);
    persistCachedUpload(null);
    restoreAttemptedRef.current = null;
  };

  const applyPreviewPayload = (data: PreviewPayload) => {
    if (data.template_html) {
      setTemplateHtml(data.template_html);
      setTemplateDirty(false);
    }
    setPreview(data);
    if (data.preview_pagination) {
      setPreviewPage(data.preview_pagination.page);
      setPreviewPageSize(data.preview_pagination.page_size);
    }
    setSelectedSheet(data.selected_sheet ?? "");
    if (data.cache_id && data.cached_files) {
      const nextCached: CachedUpload = {
        cacheId: data.cache_id,
        docxName: data.cached_files.docx_name,
        xlsxName: data.cached_files.xlsx_name,
        font: selectedFont,
        selectedSheet: data.selected_sheet,
      };
      persistCachedUpload(nextCached);
      restoreAttemptedRef.current = data.cache_id;
    }
    setStatus("done");
    setProgress(100);
  };

  // 每次載入時，嘗試從 localStorage 還原 session（跨 session 保持登入狀態）
  useEffect(() => {
    const restoreSession = async () => {
      const cachedEmail = localStorage.getItem(userEmailCacheKey);
      if (cachedEmail) setGmailEmail(cachedEmail);
      const cachedName = localStorage.getItem(userNameCacheKey);
      if (cachedName) setGmailName(cachedName);

      const savedKey = localStorage.getItem(userKeyCacheKey);
      if (!savedKey) return;
      try {
        const res = await fetch(`${apiBase}/api/session/restore`, {
          method: "POST",
          credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ user_key: savedKey }),
        });
        const data = (await res.json()) as { ok: boolean; restored: boolean; email?: string | null; name?: string | null };
        if (!data.restored) {
          localStorage.removeItem(userKeyCacheKey);
          localStorage.removeItem(userEmailCacheKey);
          localStorage.removeItem(userNameCacheKey);
          setGmailEmail(null);
          setGmailName(null);
        } else {
          if (data.email) {
            setGmailEmail(data.email);
            localStorage.setItem(userEmailCacheKey, data.email);
          }
          if (data.name) {
            setGmailName(data.name);
            localStorage.setItem(userNameCacheKey, data.name);
          }
        }
      } catch {
        // 網路錯誤，不影響後續操作
      }
    };
    void restoreSession();
  }, []);

  // OAuth 授權成功後，從 session 取得 user_key 並存入 localStorage
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    if (params.get("auth") !== "success") return;
    const saveUserKey = async () => {
      try {
        const res = await fetch(`${apiBase}/api/session/info`, {
          credentials: "include",
        });
        const data = (await res.json()) as { user_key?: string; email?: string; name?: string };
        if (data.user_key) {
          localStorage.setItem(userKeyCacheKey, data.user_key);
        }
        if (data.email) {
          setGmailEmail(data.email);
          localStorage.setItem(userEmailCacheKey, data.email);
        }
        if (data.name) {
          setGmailName(data.name);
          localStorage.setItem(userNameCacheKey, data.name);
        }
      } catch {
        // ignore
      }
      window.history.replaceState({}, "", window.location.pathname);
    };
    void saveUserKey();
  }, []);

  useEffect(() => {
    const raw = localStorage.getItem(uploadCacheKey);
    if (!raw) return;
    try {
      const parsed = JSON.parse(raw) as CachedUpload;
      if (parsed?.cacheId && parsed?.docxName && parsed?.xlsxName) {
        setCachedUpload(parsed);
        if (parsed.font) {
          setSelectedFont(parsed.font);
        }
        if (parsed.selectedSheet) {
          setSelectedSheet(parsed.selectedSheet);
        }
      }
    } catch {
      localStorage.removeItem(uploadCacheKey);
    }
  }, []);

  useEffect(() => {
    const loadPickerConfig = async () => {
      try {
        const response = await fetch(`${apiBase}/api/google/picker/config`, {
          credentials: "include",
        });
        if (!response.ok) return;
        const data = (await response.json()) as PickerConfig;
        setPickerConfig(data);
      } catch {
        // ignore config load failures; picker stays disabled
      }
    };
    void loadPickerConfig();
  }, [gmailEmail]);

  useEffect(() => {
    const hostname = window.location.hostname;
    setIsHostedFrontend(!["localhost", "127.0.0.1"].includes(hostname));
  }, []);

  useEffect(() => {
    const editor = editorRef.current;
    if (!editor) return;
    const decoratedHtml = decorateTemplatePlaceholders(templateHtml);
    if (editor.innerHTML !== decoratedHtml) {
      editor.innerHTML = decoratedHtml;
    }
  }, [templateHtml]);

  useEffect(() => {
    const restorePreview = async () => {
      if (!cachedUpload) return;
      if (preview) return;
      if (restoreAttemptedRef.current === cachedUpload.cacheId) return;
      restoreAttemptedRef.current = cachedUpload.cacheId;
      try {
        setStatus("uploading");
        setProgress(15);
        const url = new URL(`${apiBase}/api/process/cache`, window.location.origin);
        url.searchParams.set("cache_id", cachedUpload.cacheId);
        url.searchParams.set("font", cachedUpload.font || selectedFont);
        if (cachedUpload.selectedSheet) {
          url.searchParams.set("sheet", cachedUpload.selectedSheet);
        }
        url.searchParams.set("preview_page", String(previewPage));
        url.searchParams.set("preview_page_size", String(previewPageSize));
        const response = await fetch(url.toString(), {
          method: "GET",
          credentials: "include",
        });
        if (!response.ok) {
          if (response.status === 400) {
            setCachedUpload(null);
            localStorage.removeItem(uploadCacheKey);
            setDraftMessage("快取已失效，請重新上傳檔案。");
          }
          setStatus("idle");
          setProgress(0);
          return;
        }
        const data = (await response.json()) as PreviewPayload;
        applyPreviewPayload(data);
      } catch {
        setStatus("idle");
        setProgress(0);
      }
    };
    void restorePreview();
  }, [cachedUpload, preview, selectedFont, previewPage, previewPageSize]);

  // 引導動畫：首次造訪自動開啟一次
  useEffect(() => {
    const seenKey = "draftcopier_guide_seen_v1";
    if (!localStorage.getItem(seenKey)) {
      setShowGuide(true);
      localStorage.setItem(seenKey, "1");
    }
  }, []);

  // 引導 Modal：Esc 關閉
  useEffect(() => {
    if (!showGuide) return;
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") setShowGuide(false);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [showGuide]);

  // 精靈：處理完成後自動前進到「產生草稿」步驟
  useEffect(() => {
    if (!pendingAdvance) return;
    if (status === "done" && preview) {
      setPendingAdvance(false);
      setStep(3);
    } else if (status === "error") {
      setPendingAdvance(false);
    }
  }, [pendingAdvance, status, preview]);

  const emailHeader = useMemo(() => {
    if (!preview?.headers) return null;
    return findHeader(preview.headers, emailFieldCandidates);
  }, [preview]);

  const ccHeader = useMemo(() => {
    if (!preview?.headers) return null;
    return findHeader(preview.headers, ccFieldCandidates);
  }, [preview]);

  const bccHeader = useMemo(() => {
    if (!preview?.headers) return null;
    return findHeader(preview.headers, bccFieldCandidates);
  }, [preview]);

  const subjectHeader = useMemo(() => {
    if (!preview?.headers) return null;
    return findHeader(preview.headers, subjectFieldCandidates);
  }, [preview]);

  const missingHeaders = useMemo(() => {
    const missing: string[] = [];
    if (!emailHeader) missing.push("email");
    if (!subjectHeader) missing.push("subject");
    return missing;
  }, [emailHeader, subjectHeader]);

  const hasDocxSource = !!docxFile || !!driveDocx;
  const hasXlsxSource = !!xlsxFile || !!driveXlsx;

  const previewHtml = useMemo(
    () => renderTemplateWithRow(templateHtml, preview?.first_row),
    [templateHtml, preview?.first_row]
  );

  const templateFieldKeys = useMemo(() => collectTemplateFieldKeys(templateHtml), [templateHtml]);

  const previewSummaryItems = useMemo(() => {
    if (!preview?.first_row) return [];
    const firstRow = preview.first_row;
    const items: Array<{ label: string; header: string | null; value: string; tone: "to" | "subject" | "copy" | "attachment" }> = [];
    const appendItem = (label: string, header: string | null) => {
      if (!header) return;
      items.push({
        label,
        header,
        value: firstRow[header] ?? "",
        tone: label === "主旨" ? "subject" : label === "收件人" ? "to" : "copy",
      });
    };

    appendItem("收件人", emailHeader);
    appendItem("副本", ccHeader);
    appendItem("密件副本", bccHeader);
    appendItem("主旨", subjectHeader);

    for (const header of preview.detected_fields?.attachments ?? []) {
      items.push({
        label: header.startsWith("附件") ? header : "附件",
        header,
        value: firstRow[header] ?? "",
        tone: "attachment",
      });
    }

    return items;
  }, [preview, emailHeader, ccHeader, bccHeader, subjectHeader]);

  const previewRowItems = useMemo(() => {
    if (!preview?.headers || !preview.first_row) return [];
    const firstRow = preview.first_row;
    const attachmentHeaders = new Set(preview.detected_fields?.attachments ?? []);
    return preview.headers.map((header) => {
      const value = firstRow[header] ?? "";
      const tags: string[] = [];
      if (templateFieldKeys.has(header)) {
        tags.push("模板");
      }
      if ([emailHeader, ccHeader, bccHeader, subjectHeader].includes(header)) {
        tags.push("寄件");
      }
      if (attachmentHeaders.has(header)) {
        tags.push("附件");
      }
      return {
        header,
        value,
        tags,
      };
    });
  }, [preview, templateFieldKeys, emailHeader, ccHeader, bccHeader, subjectHeader]);

  const previewTableRows = preview?.preview_rows ?? [];
  const previewPagination = preview?.preview_pagination;

  const loadDriveFiles = async (kind: DriveKind, query = driveQuery) => {
    setDriveLoading(true);
    setDriveError(null);
    try {
      const url = new URL(`${apiBase}/api/drive/files`, window.location.origin);
      url.searchParams.set("kind", kind);
      if (query.trim()) {
        url.searchParams.set("q", query.trim());
      }
      const response = await fetch(url.toString(), {
        credentials: "include",
      });
      if (response.status === 401) {
        setDriveFiles([]);
        setDriveError("請先連結 Gmail，才能讀取雲端硬碟。");
        return;
      }
      if (!response.ok) {
        setDriveFiles([]);
        setDriveError("讀取雲端硬碟檔案失敗。");
        return;
      }
      const data = (await response.json()) as { files?: DriveFile[] };
      setDriveFiles(Array.isArray(data.files) ? data.files : []);
    } catch {
      setDriveFiles([]);
      setDriveError("讀取雲端硬碟檔案時發生錯誤。");
    } finally {
      setDriveLoading(false);
    }
  };

  const openDriveBrowser = (kind: DriveKind) => {
    setDriveBrowserKind(kind);
    setDriveQuery("");
    void loadDriveFiles(kind, "");
  };

  const openFolderBrowser = (purpose: FolderBrowserPurpose) => {
    setFolderBrowserPurpose(purpose);
    setDriveBrowserKind("folder");
    setDriveQuery("");
    void loadDriveFiles("folder", "");
  };

  const openGooglePicker = (kind: DriveKind) => {
    if (!gmailEmail) {
      setError("請先連結 Gmail，再從雲端硬碟挑選檔案。");
      return;
    }
    if (!pickerConfig?.enabled || !pickerConfig.client_id || !pickerConfig.api_key || !pickerConfig.app_id) {
      setError("Google Picker 尚未設定完成，請補上 API key / client id / app id。");
      return;
    }
    if (!pickerApiReady || !pickerIdentityReady || !window.google?.picker || !window.google?.accounts?.oauth2) {
      setError("Google Picker 載入中，請稍後再試。");
      return;
    }

    const mimeTypes = pickerMimeTypes[kind];
    const showPicker = (accessToken: string) => {
      const view = new window.google.picker.DocsView(window.google.picker.ViewId.DOCS);
      view.setMode(window.google.picker.DocsViewMode.LIST);
      if (mimeTypes.length > 0) {
        view.setMimeTypes(mimeTypes.join(","));
      }

      let picker = new window.google.picker.PickerBuilder()
        .addView(view)
        .setOAuthToken(accessToken)
        .setDeveloperKey(pickerConfig.api_key)
        .setAppId(pickerConfig.app_id)
        .setOrigin(window.location.origin)
        .setTitle(
          kind === "docx"
            ? "選擇模板"
            : kind === "xlsx"
            ? "選擇收件人清單"
            : kind === "folder"
            ? "選擇資料夾"
            : "選擇附件"
        )
        .setCallback((data: any) => {
          if (data[window.google.picker.Response.ACTION] !== window.google.picker.Action.PICKED) {
            return;
          }
          const docs = data[window.google.picker.Response.DOCUMENTS] ?? [];
          if (!Array.isArray(docs) || docs.length === 0) return;
          const pickedFiles = docs.map((doc: any) => {
            const mimeType = String(doc[window.google.picker.Document.MIME_TYPE] ?? "");
            return {
              id: String(doc[window.google.picker.Document.ID] ?? ""),
              name: String(doc[window.google.picker.Document.NAME] ?? "未命名檔案"),
              mime_type: mimeType,
              kind,
              is_google_workspace: mimeType.startsWith("application/vnd.google-apps."),
            } satisfies DriveFile;
          });
          selectDriveFiles(kind, pickedFiles);
        })
        ;
      if (mimeTypes.length > 0) {
        picker = picker.setSelectableMimeTypes(mimeTypes.join(","));
      }
      if (kind === "attachment") {
        picker = picker.enableFeature(window.google.picker.Feature.MULTISELECT_ENABLED);
      }
      picker = picker.build();

      picker.setVisible(true);
    };

    const tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: pickerConfig.client_id,
      scope: pickerConfig.scope || "https://www.googleapis.com/auth/drive.file",
      login_hint: pickerConfig.login_hint || gmailEmail || undefined,
      callback: (response: any) => {
        if (response?.error || !response?.access_token) {
          setError("無法取得 Google Picker 存取權杖。");
          return;
        }
        pickerTokenRef.current = response.access_token;
        showPicker(response.access_token);
      },
      error_callback: () => {
        setError("Google Picker 視窗未完成授權。");
      },
    });

    tokenClient.requestAccessToken({
      prompt: pickerTokenRef.current ? "" : "consent",
      login_hint: pickerConfig.login_hint || gmailEmail || undefined,
    });
  };

  const selectDriveFiles = (kind: DriveKind, files: DriveFile[]) => {
    if (kind === "attachment") {
      setSelectedDriveAttachments((current) => {
        const next = [...current];
        for (const file of files) {
          if (!next.some((item) => item.id === file.id)) {
            next.push(file);
          }
        }
        return next;
      });
      setDriveBrowserKind(null);
      return;
    }
    if (kind === "folder") {
      const file = files[0];
      if (!file) return;
      if (folderBrowserPurpose === "create-parent") {
        void createCloudUploadFolder(file);
      } else {
        setSelectedDriveFolder(file);
        setCloudUploadFolder(null);
        setDriveBrowserKind(null);
      }
      return;
    }
    const file = files[0];
    if (!file) return;
    resetPreparedState();
    if (kind === "docx") {
      setDriveDocx(file);
      setDocxFile(null);
      if (docxInputRef.current) docxInputRef.current.value = "";
    } else {
      setDriveXlsx(file);
      setXlsxFile(null);
      if (xlsxInputRef.current) xlsxInputRef.current.value = "";
    }
    setDriveBrowserKind(null);
  };

  const removeSelectedDriveAttachment = (fileId: string) => {
    setSelectedDriveAttachments((current) => current.filter((file) => file.id !== fileId));
  };

  const clearSelectedDriveFolder = () => {
    setSelectedDriveFolder(null);
    setCloudUploadFolder(null);
    setUploadedCloudFolderFiles([]);
  };

  const clearSelectedLocalAttachmentFiles = () => {
    setSelectedLocalAttachmentFiles([]);
    if (localAttachmentFolderInputRef.current) localAttachmentFolderInputRef.current.value = "";
  };

  const chooseLocalAttachmentFolder = (files: FileList | File[]) => {
    setSelectedLocalAttachmentFiles(Array.from(files));
  };

  const createCloudUploadFolder = async (parentFolder: DriveFile | null) => {
    setCreatingCloudFolder(true);
    setDriveError(null);
    try {
      const response = await fetch(`${apiBase}/api/drive/folders/create`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          parent_folder_id: parentFolder?.id ?? null,
        }),
      });
      if (!response.ok) {
        setDriveError("建立雲端資料夾失敗，請重新連結 Gmail 後再試。");
        return;
      }
      const data = (await response.json()) as { folder?: DriveFile };
      if (!data.folder) {
        setDriveError("建立雲端資料夾失敗。");
        return;
      }
      setCloudUploadFolder(data.folder);
      setSelectedDriveFolder(data.folder);
      setUploadedCloudFolderFiles([]);
      setDriveBrowserKind(null);
    } catch {
      setDriveError("建立雲端資料夾時發生錯誤。");
    } finally {
      setCreatingCloudFolder(false);
    }
  };

  const uploadFilesToCloudFolder = async (files: FileList | File[]) => {
    if (!cloudUploadFolder) return;
    const pendingFiles = Array.from(files);
    if (pendingFiles.length === 0) return;
    setUploadingCloudFiles(true);
    setError(null);
    try {
      const formData = new FormData();
      formData.append("folder_id", cloudUploadFolder.id);
      pendingFiles.forEach((file) => formData.append("files", file));
      const response = await fetch(`${apiBase}/api/drive/folders/upload`, {
        method: "POST",
        credentials: "include",
        body: formData,
      });
      if (!response.ok) {
        setError("上傳檔案到雲端資料夾失敗。");
        return;
      }
      const data = (await response.json()) as { files?: DriveFile[] };
      const uploadedFiles = Array.isArray(data.files) ? data.files : [];
      setUploadedCloudFolderFiles((current) => [...current, ...uploadedFiles]);
    } catch {
      setError("上傳檔案到雲端資料夾時發生錯誤。");
    } finally {
      setUploadingCloudFiles(false);
      if (cloudAttachmentInputRef.current) cloudAttachmentInputRef.current.value = "";
    }
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragActive(false);
    const files = classifyFiles(event.dataTransfer.files);
    const hasUnsupportedXls = Array.from(event.dataTransfer.files).some((file) =>
      file.name.toLowerCase().endsWith(".xls")
    );
    if (!files.docx && !files.xlsx) {
      if (hasUnsupportedXls) {
        setError("只支援上傳 .docx 與 .xlsx 檔案，.xls 不支援。");
      }
      return;
    }
    resetPreparedState();
    if (files.docx) {
      setDocxFile(files.docx);
      setDriveDocx(null);
    }
    if (files.xlsx) {
      setXlsxFile(files.xlsx);
      setDriveXlsx(null);
    }
  };

  const handleUpload = () => {
    if (!hasXlsxSource) {
      setError("請至少選擇收件人清單。");
      return;
    }

    setError(null);
    setStatus("uploading");
    setProgress(0);
    setPreview(null);
    setDraftStatus("idle");
    setDraftMessage(null);
    setDraftFailedItems([]);
    setPreviewPage(1);

    const formData = new FormData();
    if (docxFile) {
      formData.append("docx_file", docxFile);
    } else if (driveDocx) {
      formData.append("docx_drive_file_id", driveDocx.id);
    }
    if (xlsxFile) {
      formData.append("xlsx_file", xlsxFile);
    } else if (driveXlsx) {
      formData.append("xlsx_drive_file_id", driveXlsx.id);
    }
    if (templateDirty || !hasDocxSource) {
      formData.append("template_html", templateHtml);
    }

    const url = new URL(`${apiBase}/api/process`, window.location.origin);
    url.searchParams.set("font", selectedFont);
    url.searchParams.set("preview_page", "1");
    url.searchParams.set("preview_page_size", String(previewPageSize));
    if (selectedSheet) {
      url.searchParams.set("sheet", selectedSheet);
    }
    const xhr = new XMLHttpRequest();
    xhr.open("POST", url.toString());
    xhr.withCredentials = true;

    xhr.upload.addEventListener("progress", (event) => {
      if (event.lengthComputable) {
        const pct = Math.round((event.loaded / event.total) * 100);
        setProgress(pct);
      }
    });

    xhr.addEventListener("load", () => {
      if (xhr.status >= 200 && xhr.status < 300) {
        try {
          const data = JSON.parse(xhr.responseText) as PreviewPayload;
          applyPreviewPayload(data);
        } catch {
          setStatus("error");
          setError("伺服器回應解析失敗。");
        }
      } else {
        setStatus("error");
        setError(`上傳失敗：${xhr.status} ${xhr.statusText}`);
      }
    });

    xhr.addEventListener("error", () => {
      setStatus("error");
      setError("上傳時發生網路錯誤。");
    });

    xhr.send(formData);
  };

  const resetFiles = () => {
    setDocxFile(null);
    setXlsxFile(null);
    setDriveDocx(null);
    setDriveXlsx(null);
    setProgress(0);
    setStatus("idle");
    setError(null);
    setPreview(null);
    setDraftStatus("idle");
    setDraftMessage(null);
    setAttachmentsDir("");
    setTemplateHtml(defaultTemplateHtml);
    setTemplateDirty(false);
    setSelectedDriveAttachments([]);
    setSelectedDriveFolder(null);
    setCloudUploadFolder(null);
    setUploadedCloudFolderFiles([]);
    setSelectedLocalAttachmentFiles([]);
    setSelectedSheet("");
    persistCachedUpload(null);
    restoreAttemptedRef.current = null;
    setDriveBrowserKind(null);
    setDriveFiles([]);
    setDriveQuery("");
    setDriveError(null);
    if (localAttachmentFolderInputRef.current) localAttachmentFolderInputRef.current.value = "";
    if (cloudAttachmentInputRef.current) cloudAttachmentInputRef.current.value = "";
    if (docxInputRef.current) docxInputRef.current.value = "";
    if (xlsxInputRef.current) xlsxInputRef.current.value = "";
  };

  const handleSheetChange = async (nextSheet: string) => {
    setSelectedSheet(nextSheet);
    setPreviewPage(1);
    setStatus("uploading");
    setProgress(25);
    setError(null);
    try {
      let response: globalThis.Response;
      const canRepostCurrentSources = (!!xlsxFile || !!driveXlsx) && (templateDirty || !cachedUpload?.cacheId);
      if (cachedUpload?.cacheId && !canRepostCurrentSources) {
        const url = new URL(`${apiBase}/api/process/cache`, window.location.origin);
        url.searchParams.set("cache_id", cachedUpload.cacheId);
        url.searchParams.set("font", selectedFont);
        if (nextSheet) {
          url.searchParams.set("sheet", nextSheet);
        }
        url.searchParams.set("preview_page", "1");
        url.searchParams.set("preview_page_size", String(previewPageSize));
        response = await fetch(url.toString(), {
          method: "GET",
          credentials: "include",
        });
      } else {
        const formData = new FormData();
        if (xlsxFile) {
          formData.append("xlsx_file", xlsxFile);
        } else if (driveXlsx) {
          formData.append("xlsx_drive_file_id", driveXlsx.id);
        } else {
          setStatus("idle");
          setProgress(0);
          return;
        }
        if (docxFile) {
          formData.append("docx_file", docxFile);
        } else if (driveDocx) {
          formData.append("docx_drive_file_id", driveDocx.id);
        }
        formData.append("template_html", templateHtml);
        const url = new URL(`${apiBase}/api/process`, window.location.origin);
        url.searchParams.set("font", selectedFont);
        if (nextSheet) {
          url.searchParams.set("sheet", nextSheet);
        }
        url.searchParams.set("preview_page", "1");
        url.searchParams.set("preview_page_size", String(previewPageSize));
        response = await fetch(url.toString(), {
          method: "POST",
          credentials: "include",
          body: formData,
        });
      }
      if (!response.ok) {
        setStatus("error");
        setError("切換工作表失敗。");
        return;
      }
      const data = (await response.json()) as PreviewPayload;
      applyPreviewPayload(data);
    } catch {
      setStatus("error");
      setError("切換工作表時發生錯誤。");
    }
  };

  const refreshPreviewPage = async (nextPage: number, nextPageSize = previewPageSize) => {
    setStatus("uploading");
    setProgress(25);
    setError(null);
    try {
      let response: globalThis.Response;
      if (cachedUpload?.cacheId) {
        const url = new URL(`${apiBase}/api/process/cache`, window.location.origin);
        url.searchParams.set("cache_id", cachedUpload.cacheId);
        url.searchParams.set("font", selectedFont);
        url.searchParams.set("preview_page", String(nextPage));
        url.searchParams.set("preview_page_size", String(nextPageSize));
        if (selectedSheet) {
          url.searchParams.set("sheet", selectedSheet);
        }
        response = await fetch(url.toString(), {
          method: "GET",
          credentials: "include",
        });
      } else {
        const formData = new FormData();
        if (xlsxFile) {
          formData.append("xlsx_file", xlsxFile);
        } else if (driveXlsx) {
          formData.append("xlsx_drive_file_id", driveXlsx.id);
        } else {
          setStatus("idle");
          setProgress(0);
          return;
        }
        if (docxFile) {
          formData.append("docx_file", docxFile);
        } else if (driveDocx) {
          formData.append("docx_drive_file_id", driveDocx.id);
        }
        formData.append("template_html", templateHtml);
        const url = new URL(`${apiBase}/api/process`, window.location.origin);
        url.searchParams.set("font", selectedFont);
        url.searchParams.set("preview_page", String(nextPage));
        url.searchParams.set("preview_page_size", String(nextPageSize));
        if (selectedSheet) {
          url.searchParams.set("sheet", selectedSheet);
        }
        response = await fetch(url.toString(), {
          method: "POST",
          credentials: "include",
          body: formData,
        });
      }
      if (!response.ok) {
        setStatus("error");
        setError("切換 XLSX 分頁失敗。");
        return;
      }
      const data = (await response.json()) as PreviewPayload;
      applyPreviewPayload(data);
    } catch {
      setStatus("error");
      setError("切換 XLSX 分頁時發生錯誤。");
    }
  };

  const connectGmail = async () => {
    setDraftMessage(null);
    try {
      const response = await fetch(`${apiBase}/api/auth/google`, {
        method: "GET",
        credentials: "include",
      });
      if (!response.ok) {
        setDraftStatus("error");
        setDraftMessage("取得授權連結失敗。");
        setDraftFailedItems([]);
        return;
      }
      const data = (await response.json()) as { auth_url?: string };
      if (!data.auth_url) {
        setDraftStatus("error");
        setDraftMessage("授權連結格式錯誤。");
        setDraftFailedItems([]);
        return;
      }
      window.location.href = data.auth_url;
    } catch (err) {
      setDraftStatus("error");
      setDraftMessage("啟動授權流程失敗，請稍後再試。");
      setDraftFailedItems([]);
    }
  };

  const syncEditorHtml = () => {
    const plainHtml = stripTemplatePlaceholderMarkup(editorRef.current?.innerHTML || defaultTemplateHtml);
    const decoratedHtml = decorateTemplatePlaceholders(plainHtml);
    if (editorRef.current && editorRef.current.innerHTML !== decoratedHtml) {
      editorRef.current.innerHTML = decoratedHtml;
    }
    setTemplateDirty(true);
    setTemplateHtml(plainHtml || defaultTemplateHtml);
  };

  const runEditorCommand = (
    command:
      | "bold"
      | "italic"
      | "underline"
      | "insertUnorderedList"
      | "insertOrderedList"
      | "indent"
      | "outdent"
      | "unlink"
      | "removeFormat"
  ) => {
    editorRef.current?.focus();
    document.execCommand(command);
    syncEditorHtml();
  };

  const runEditorValueCommand = (command: "foreColor" | "fontName", value: string) => {
    editorRef.current?.focus();
    document.execCommand(command, false, value);
    syncEditorHtml();
  };

  const applyEditorFont = (font: string) => {
    setSelectedFont(font);
    runEditorValueCommand("fontName", font);
  };

  const createEditorLink = () => {
    const currentSelection = window.getSelection()?.toString().trim() || "";
    const rawUrl = window.prompt("輸入連結網址", currentSelection.startsWith("http") ? currentSelection : "https://");
    if (!rawUrl) return;
    const url = rawUrl.trim();
    if (!url) return;
    editorRef.current?.focus();
    document.execCommand("createLink", false, url);
    syncEditorHtml();
  };

  const reloadTemplateFromDocx = async () => {
    if (!hasDocxSource) return;
    setReloadingTemplate(true);
    setError(null);
    try {
      const formData = new FormData();
      if (docxFile) {
        formData.append("docx_file", docxFile);
      } else if (driveDocx) {
        formData.append("docx_drive_file_id", driveDocx.id);
      } else {
        setError("找不到目前的 DOCX 來源。");
        return;
      }

      const url = new URL(`${apiBase}/api/template/load`, window.location.origin);
      url.searchParams.set("font", selectedFont);
      const response = await fetch(url.toString(), {
        method: "POST",
        credentials: "include",
        body: formData,
      });
      if (!response.ok) {
        setError("重新載入模板失敗。");
        return;
      }
      const data = (await response.json()) as { template_html?: string };
      if (!data.template_html) {
        setError("DOCX 模板內容為空。");
        return;
      }
      setTemplateHtml(data.template_html);
      setTemplateDirty(false);
    } catch {
      setError("重新載入模板時發生錯誤。");
    } finally {
      setReloadingTemplate(false);
    }
  };

  const exportTemplate = async () => {
    setExportingTemplate(true);
    setError(null);
    try {
      const filenameBase = hasDocxSource
        ? (docxFile?.name || driveDocx?.name || cachedUpload?.docxName || "template").replace(/\.docx$/i, "")
        : "template";
      const response = await fetch(`${apiBase}/api/template/export`, {
        method: "POST",
        credentials: "include",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          template_html: templateHtml,
          filename: filenameBase,
        }),
      });
      if (!response.ok) {
        setError("匯出 DOCX 失敗。");
        return;
      }
      const blob = await response.blob();
      const downloadUrl = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = `${filenameBase || "template"}.docx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(downloadUrl);
    } catch {
      setError("匯出 DOCX 時發生錯誤。");
    } finally {
      setExportingTemplate(false);
    }
  };

  const saveDrafts = async () => {
    if (!preview && !cachedUpload) {
      setDraftStatus("error");
      setDraftMessage("尚未產生預覽，無法儲存草稿。");
      setDraftFailedItems([]);
      return;
    }
    if (preview && missingHeaders.length > 0) {
      setDraftStatus("error");
      setDraftMessage("找不到必要欄位（email / subject），請檢查 Excel 標題列。");
      setDraftFailedItems([]);
      return;
    }
    setDraftStatus("saving");
    setDraftMessage(null);
    setDraftFailedItems([]);
    try {
      const formData = new FormData();
      if (docxFile && xlsxFile) {
        formData.append("docx_file", docxFile);
        formData.append("xlsx_file", xlsxFile);
      } else if (xlsxFile) {
        formData.append("xlsx_file", xlsxFile);
      } else if (cachedUpload?.cacheId) {
        formData.append("cache_id", cachedUpload.cacheId);
      } else {
        if (driveXlsx) {
          formData.append("xlsx_drive_file_id", driveXlsx.id);
          if (driveDocx) {
            formData.append("docx_drive_file_id", driveDocx.id);
          }
        } else {
          setDraftStatus("error");
          setDraftMessage("請先選擇收件清單。");
          return;
        }
      }
      formData.append("template_html", templateHtml);
      if (selectedDriveAttachments.length > 0) {
        formData.append(
          "attachment_drive_file_ids_json",
          JSON.stringify(selectedDriveAttachments.map((file) => file.id))
        );
      }
      if (selectedDriveFolder) {
        formData.append("attachment_drive_folder_id", selectedDriveFolder.id);
      }
      if (selectedLocalAttachmentFiles.length > 0) {
        selectedLocalAttachmentFiles.forEach((file) => formData.append("attachment_local_files", file));
      }

      const url = new URL(`${apiBase}/api/drafts/batch`, window.location.origin);
      url.searchParams.set("font", selectedFont);
      if (selectedSheet) {
        url.searchParams.set("sheet", selectedSheet);
      }
      if (attachmentsDir.trim()) {
        url.searchParams.set("attachments_dir", attachmentsDir.trim());
      }
      const response = await fetch(url.toString(), {
        method: "POST",
        credentials: "include",
        body: formData,
      });

      if (response.status === 401) {
        setDraftStatus("error");
        setDraftMessage("尚未連結 Gmail，請先完成授權。");
        setDraftFailedItems([]);
        return;
      }
      if (!response.ok) {
        const text = await response.text();
        setDraftStatus("error");
        setDraftMessage(`儲存草稿失敗：${text}`);
        setDraftFailedItems([]);
        return;
      }

      const data = (await response.json()) as BatchSaveResponse;
      const failedItems = Array.isArray(data.failed_items) ? data.failed_items : [];
      setDraftFailedItems(failedItems);

      const draftCount = data.draft_count ?? 0;
      if ((data.status === "failed" || draftCount === 0) && failedItems.length > 0) {
        setDraftStatus("error");
        setDraftMessage(`未建立草稿。失敗項目：${failedItems.length}。`);
        return;
      }

      setDraftStatus("done");
      if (failedItems.length > 0) {
        setDraftMessage(`已建立 Gmail 草稿：${draftCount} 封；失敗項目：${failedItems.length}。`);
      } else {
        setDraftMessage(`已建立 Gmail 草稿：${draftCount} 封。`);
      }
    } catch (err) {
      setDraftStatus("error");
      setDraftMessage("儲存草稿時發生錯誤。");
      setDraftFailedItems([]);
    }
  };

  const docxSourceName = docxFile
    ? docxFile.name
    : driveDocx?.name ?? cachedUpload?.docxName ?? "尚未選擇檔案";
  const xlsxSourceName = xlsxFile
    ? xlsxFile.name
    : driveXlsx?.name ?? cachedUpload?.xlsxName ?? "尚未選擇檔案";
  const docxOrigin = docxFile
    ? "本機檔案"
    : driveDocx
    ? "雲端硬碟"
    : cachedUpload?.docxName
    ? "已快取"
    : "—";
  const xlsxOrigin = xlsxFile
    ? "本機檔案"
    : driveXlsx
    ? "雲端硬碟"
    : cachedUpload?.xlsxName
    ? "已快取"
    : "—";

  const steps: Array<{ id: 1 | 2 | 3; label: string }> = [
    { id: 1, label: "選擇檔案" },
    { id: 2, label: "合併設定" },
    { id: 3, label: "產生草稿" },
  ];

  const canLeaveStep1 = true;

  const goToStep = (target: 1 | 2 | 3) => {
    if (target < step) setStep(target);
  };

  const goBack = () => {
    if (step > 1) setStep((prev) => (prev - 1) as 1 | 2 | 3);
  };

  const goNext = () => {
    if (step === 1) {
      if (canLeaveStep1) setStep(2);
      return;
    }
    if (step === 2) {
      if (preview && status === "done") {
        setStep(3);
      } else if (status !== "uploading") {
        setPendingAdvance(true);
        handleUpload();
      }
    }
  };

  const handleReset = () => {
    resetFiles();
    setPendingAdvance(false);
    setOpenMenu(null);
    setStep(1);
  };

  const chooseSource = (kind: DriveKind, method: "local" | "picker" | "recent") => {
    setOpenMenu(null);
    if (method === "local") {
      (kind === "docx" ? docxInputRef : xlsxInputRef).current?.click();
    } else if (method === "picker") {
      openGooglePicker(kind);
    } else {
      openDriveBrowser(kind);
    }
  };

  return (
    <div className="page">
      <div className="bg-fx" aria-hidden="true">
        <div className="bg-grid" />
        <svg className="bg-lines" viewBox="0 0 1440 900" preserveAspectRatio="xMidYMid slice">
          <path d="M-60 120 C 320 40, 700 200, 1040 110 S 1520 -20, 1620 150" />
          <path d="M-60 300 C 300 220, 720 420, 1080 300 S 1540 200, 1620 360" />
          <path d="M-60 500 C 340 420, 700 620, 1060 500 S 1520 400, 1620 560" />
          <path d="M-60 700 C 320 620, 740 820, 1100 700 S 1540 620, 1620 780" />
        </svg>
      </div>

      <input
        ref={docxInputRef}
        type="file"
        accept=".docx"
        className="visually-hidden"
        onChange={(event) => {
          resetPreparedState();
          setDocxFile(event.target.files?.[0] ?? null);
          setDriveDocx(null);
        }}
      />
      <input
        ref={xlsxInputRef}
        type="file"
        accept=".xlsx"
        className="visually-hidden"
        onChange={(event) => {
          resetPreparedState();
          const nextFile = event.target.files?.[0] ?? null;
          if (nextFile && !nextFile.name.toLowerCase().endsWith(".xlsx")) {
            setXlsxFile(null);
            setError("只支援上傳 .docx 與 .xlsx 檔案，.xls 不支援。");
            if (xlsxInputRef.current) xlsxInputRef.current.value = "";
            return;
          }
          setXlsxFile(nextFile);
          setDriveXlsx(null);
        }}
      />
      <input
        ref={cloudAttachmentInputRef}
        type="file"
        multiple
        className="visually-hidden"
        onChange={(event) => {
          if (event.target.files) {
            void uploadFilesToCloudFolder(event.target.files);
          }
        }}
      />
      <input
        ref={localAttachmentFolderInputRef}
        type="file"
        multiple
        className="visually-hidden"
        {...({ webkitdirectory: "", directory: "" } as Record<string, string>)}
        onChange={(event) => {
          if (event.target.files) {
            chooseLocalAttachmentFolder(event.target.files);
          }
        }}
      />

      <Script
        src="https://apis.google.com/js/api.js"
        strategy="afterInteractive"
        onLoad={() => {
          window.gapi?.load("picker", () => {
            setPickerApiReady(true);
          });
        }}
      />
      <Script
        src="https://accounts.google.com/gsi/client"
        strategy="afterInteractive"
        onLoad={() => setPickerIdentityReady(true)}
      />

      <header className="topbar">
        <div className="brand">
          <div className="logo">DC</div>
          <div className="brand-text">
            <p className="title">郵件批量寄送系統</p>
            <p className="subtitle">DOCX × Excel → Gmail Drafts</p>
          </div>
        </div>
        <div className="topbar-right">
          <button
            className="btn btn-ghost btn-sm guide-btn"
            onClick={() => setShowGuide(true)}
            title="播放運作原理引導動畫"
          >
            <span className="guide-play" aria-hidden="true">▶</span>
            運作原理
          </button>
          <div className="account">
            {gmailEmail ? (
              <button className="account-chip" onClick={connectGmail} title="點擊可重新連結 Gmail">
                <span className="account-dot" aria-hidden="true" />
                <span className="account-meta">
                  <span className="account-label">Google 帳號</span>
                  {gmailName && <span className="account-name">{gmailName}</span>}
                  <span className="account-email">{gmailEmail}</span>
                </span>
              </button>
            ) : (
              <button className="btn btn-outline btn-sm" onClick={connectGmail}>
                連結 Gmail
              </button>
            )}
          </div>
        </div>
      </header>

      <nav className="stepper" aria-label="流程步驟">
        {steps.flatMap((s, i) => {
          const state = step === s.id ? "active" : step > s.id ? "done" : "idle";
          const item = (
            <button
              key={`step-${s.id}`}
              type="button"
              className={`step-item ${state}`}
              onClick={() => goToStep(s.id)}
              disabled={s.id >= step}
            >
              <span className="step-dot">{step > s.id ? "✓" : s.id}</span>
              <span className="step-name">{s.label}</span>
            </button>
          );
          if (i === 0) return [item];
          const line = (
            <span
              key={`line-${s.id}`}
              className={`step-line ${step > steps[i - 1].id ? "done" : ""}`}
              aria-hidden="true"
            />
          );
          return [line, item];
        })}
      </nav>

      <main className="wizard">
        <div className="wizard-card">
          {step === 1 && (
            <div className="step-body">
              <div className="step-head">
                <div>
                  <h2 className="step-title">選擇檔案</h2>
                  <p className="step-desc">
                    Word 模板決定內文與樣式，Excel 清單提供收件人與代換欄位。
                  </p>
                </div>
                <div className="step-head-actions">
                  {gmailEmail ? (
                    <div className="step-account-card">
                      <div className="step-account-meta">
                        <span className="step-account-label">目前連結帳號</span>
                        {gmailName && <span className="step-account-name">{gmailName}</span>}
                        <span className="step-account-email">{gmailEmail}</span>
                      </div>
                      <button className="btn btn-outline btn-sm" onClick={connectGmail}>
                        重新連結
                      </button>
                    </div>
                  ) : (
                    <button className="btn btn-primary btn-sm" onClick={connectGmail}>
                      先連結 Gmail
                    </button>
                  )}
                  <button
                    className="link-btn step-guide-link"
                    onClick={() => setShowGuide(true)}
                  >
                    <span aria-hidden="true">▶</span> 看運作原理
                  </button>
                  {(hasDocxSource || hasXlsxSource) && (
                    <button className="link-btn" onClick={handleReset}>
                      清除重來
                    </button>
                  )}
                </div>
              </div>

              <div
                className={`dropzone ${isDragActive ? "active" : ""}`}
                onDragOver={(event) => {
                  event.preventDefault();
                  setIsDragActive(true);
                }}
                onDragLeave={() => setIsDragActive(false)}
                onDrop={handleDrop}
              >
                <span className="dropzone-icon" aria-hidden="true">
                  ↓
                </span>
                <p className="dropzone-title">將檔案拖放到此處</p>
                <p className="dropzone-sub">支援 .docx 模板與 .xlsx 清單</p>
              </div>

              <div className="source-list">
                <div className={`src-row ${hasDocxSource ? "is-set" : ""}`}>
                  <div className="src-info">
                    <span className="src-kind">模板 · DOCX</span>
                    <span className="src-file" title={docxSourceName}>
                      {docxSourceName}
                    </span>
                    {hasDocxSource && <span className="src-origin">{docxOrigin}</span>}
                  </div>
                  <div className="src-actions">
                    <span className={`src-flag ${hasDocxSource ? "ok" : ""}`}>
                      {hasDocxSource ? "已選擇" : "未選擇"}
                    </span>
                    <div className="menu-wrap">
                      <button
                        className="btn btn-outline btn-select"
                        onClick={() => setOpenMenu(openMenu === "docx" ? null : "docx")}
                        aria-haspopup="menu"
                        aria-expanded={openMenu === "docx"}
                      >
                        選擇<span className="caret" aria-hidden="true">▾</span>
                      </button>
                      {openMenu === "docx" && (
                        <>
                          <div className="menu-backdrop" onClick={() => setOpenMenu(null)} />
                          <div className="menu" role="menu">
                            <button className="menu-item" onClick={() => chooseSource("docx", "local")}>
                              本機檔案
                            </button>
                            <button
                              className="menu-item"
                              onClick={() => chooseSource("docx", "picker")}
                              disabled={!gmailEmail}
                            >
                              Google 雲端挑選
                            </button>
                            <button
                              className="menu-item"
                              onClick={() => chooseSource("docx", "recent")}
                              disabled={!gmailEmail}
                            >
                              最近使用
                            </button>
                          </div>
                        </>
                      )}
                    </div>
                  </div>
                </div>

                <div className={`src-row ${hasXlsxSource ? "is-set" : ""}`}>
                  <div className="src-info">
                    <span className="src-kind">收件清單 · XLSX</span>
                    <span className="src-file" title={xlsxSourceName}>
                      {xlsxSourceName}
                    </span>
                    {hasXlsxSource && <span className="src-origin">{xlsxOrigin}</span>}
                  </div>
                  <div className="src-actions">
                    <span className={`src-flag ${hasXlsxSource ? "ok" : ""}`}>
                      {hasXlsxSource ? "已選擇" : "未選擇"}
                    </span>
                    <div className="menu-wrap">
                      <button
                        className="btn btn-outline btn-select"
                        onClick={() => setOpenMenu(openMenu === "xlsx" ? null : "xlsx")}
                        aria-haspopup="menu"
                        aria-expanded={openMenu === "xlsx"}
                      >
                        選擇<span className="caret" aria-hidden="true">▾</span>
                      </button>
                      {openMenu === "xlsx" && (
                        <>
                          <div className="menu-backdrop" onClick={() => setOpenMenu(null)} />
                          <div className="menu up" role="menu">
                            <button className="menu-item" onClick={() => chooseSource("xlsx", "local")}>
                              本機檔案
                            </button>
                            <button
                              className="menu-item"
                              onClick={() => chooseSource("xlsx", "picker")}
                              disabled={!gmailEmail}
                            >
                              Google 雲端挑選
                            </button>
                            <button
                              className="menu-item"
                              onClick={() => chooseSource("xlsx", "recent")}
                              disabled={!gmailEmail}
                            >
                              最近使用
                            </button>
                          </div>
                        </>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              {!gmailEmail && (
                <p className="inline-note">連結 Gmail 後即可從雲端硬碟挑選檔案。</p>
              )}

              <div className="editor-panel">
                <div className="editor-head">
                  <div>
                    <p className="editor-title">模板編輯器</p>
                    <p className="field-hint">
                      不上傳模板也能直接編輯；後續預覽、Gmail 草稿與 DOCX 匯出都會使用這份內容。
                    </p>
                    <p className="field-hint">
                      在 editor 內，<span className="editor-placeholder-inline">{"{{欄位}}"}</span> 會自動標示成 XLSX 對照欄位。
                    </p>
                    {hasDocxSource && (
                      <p className="field-hint">
                        已上傳 DOCX 後，editor 不會自動再次覆蓋；要以目前 DOCX 內容重載，請按「重新載入 DOCX」。
                      </p>
                    )}
                  </div>
                  <div className="editor-actions">
                    {hasDocxSource && (
                      <button
                        className="btn btn-outline btn-sm"
                        type="button"
                        onMouseDown={(event) => event.preventDefault()}
                        onClick={reloadTemplateFromDocx}
                        disabled={reloadingTemplate}
                      >
                        {reloadingTemplate ? "重載中…" : "重新載入 DOCX"}
                      </button>
                    )}
                    <label className="editor-select-control">
                      <span className="editor-select-label">字型</span>
                      <select
                        className="editor-font-select"
                        value={selectedFont}
                        onChange={(event) => applyEditorFont(event.target.value)}
                      >
                        {fontOptions.map((option) => (
                          <option key={option.value} value={option.value}>
                            {option.label}
                          </option>
                        ))}
                      </select>
                    </label>
                    <button
                      className="btn btn-outline btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => runEditorCommand("bold")}
                    >
                      粗體
                    </button>
                    <button
                      className="btn btn-outline btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => runEditorCommand("italic")}
                    >
                      斜體
                    </button>
                    <button
                      className="btn btn-outline btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => runEditorCommand("underline")}
                    >
                      底線
                    </button>
                    <label className="editor-color-control" onMouseDown={(event) => event.preventDefault()}>
                      <span className="editor-color-label">字色</span>
                      <input
                        className="editor-color-input"
                        type="color"
                        value={textColor}
                        onChange={(event) => {
                          const nextColor = event.target.value;
                          setTextColor(nextColor);
                          runEditorValueCommand("foreColor", nextColor);
                        }}
                      />
                    </label>
                    <button
                      className="btn btn-outline btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={createEditorLink}
                    >
                      連結
                    </button>
                    <button
                      className="btn btn-outline btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => runEditorCommand("unlink")}
                    >
                      取消連結
                    </button>
                    <button
                      className="btn btn-outline btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => runEditorCommand("insertUnorderedList")}
                    >
                      項目符號
                    </button>
                    <button
                      className="btn btn-outline btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => runEditorCommand("insertOrderedList")}
                    >
                      編號
                    </button>
                    <button
                      className="btn btn-outline btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => runEditorCommand("outdent")}
                    >
                      減少縮排
                    </button>
                    <button
                      className="btn btn-outline btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => runEditorCommand("indent")}
                    >
                      增加縮排
                    </button>
                    <button
                      className="btn btn-outline btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={() => runEditorCommand("removeFormat")}
                    >
                      清除格式
                    </button>
                    <button
                      className="btn btn-primary btn-sm"
                      type="button"
                      onMouseDown={(event) => event.preventDefault()}
                      onClick={exportTemplate}
                      disabled={exportingTemplate}
                    >
                      {exportingTemplate ? "匯出中…" : "匯出 DOCX"}
                    </button>
                  </div>
                </div>
                <div
                  ref={editorRef}
                  className="template-editor"
                  contentEditable
                  suppressContentEditableWarning
                  onInput={(event) => {
                    syncEditorHtml();
                  }}
                />
              </div>

              {(driveBrowserKind === "docx" || driveBrowserKind === "xlsx") && (
                <div className="drive-browser">
                  <div className="drive-browser-head">
                    <div>
                      <p className="drive-browser-title">
                        選擇{driveBrowserKind === "docx" ? "模板" : "收件清單"}
                      </p>
                      <p className="field-hint">顯示最近 20 筆支援檔案，可搜尋檔名。</p>
                    </div>
                    <button className="link-btn" onClick={() => setDriveBrowserKind(null)}>
                      關閉
                    </button>
                  </div>
                  <div className="drive-browser-search">
                    <input
                      className="text-input"
                      type="text"
                      placeholder="輸入檔名關鍵字"
                      value={driveQuery}
                      onChange={(event) => setDriveQuery(event.target.value)}
                      onKeyDown={(event) => {
                        if (event.key === "Enter") {
                          event.preventDefault();
                          void loadDriveFiles(driveBrowserKind, driveQuery);
                        }
                      }}
                    />
                    <button
                      className="btn btn-outline btn-sm"
                      onClick={() => void loadDriveFiles(driveBrowserKind, driveQuery)}
                    >
                      搜尋
                    </button>
                  </div>
                  {driveError && <p className="error">{driveError}</p>}
                  <div className="drive-file-list">
                    {driveLoading ? (
                      <p className="field-hint">讀取中…</p>
                    ) : driveFiles.length > 0 ? (
                      driveFiles.map((file) => (
                        <button
                          key={file.id}
                          className="drive-file-item"
                          onClick={() => selectDriveFiles(driveBrowserKind, [file])}
                        >
                          <span className="drive-file-name">{file.name}</span>
                          <span className="field-hint">
                            {file.is_google_workspace ? "Google 工作區" : "Office 檔案"}
                          </span>
                        </button>
                      ))
                    ) : (
                      <p className="field-hint">找不到可用檔案。</p>
                    )}
                  </div>
                </div>
              )}
            </div>
          )}

          {step === 2 && (
            <div className="step-body">
              <div className="step-head">
                <div>
                  <h2 className="step-title">合併設定</h2>
                  <p className="step-desc">選擇字型與附件來源，套用到每一封草稿。</p>
                </div>
              </div>

              <div className="settings-grid">
                <div className="field">
                  <label className="field-label" htmlFor="font-select">
                    字型（Gmail 支援）
                  </label>
                  <select
                    id="font-select"
                    className="select"
                    value={selectedFont}
                    onChange={(event) => setSelectedFont(event.target.value)}
                  >
                    {fontOptions.map((option) => (
                      <option key={option.value} value={option.value}>
                        {option.label}
                      </option>
                    ))}
                  </select>
                </div>

                {preview?.sheet_names && preview.sheet_names.length > 0 && (
                  <div className="field">
                    <label className="field-label" htmlFor="sheet-select">
                      收件人工作表
                    </label>
                    <select
                      id="sheet-select"
                      className="select"
                      value={selectedSheet}
                      onChange={(event) => void handleSheetChange(event.target.value)}
                      disabled={!cachedUpload?.cacheId || status === "uploading"}
                    >
                      {preview.sheet_names.map((sheetName) => (
                        <option key={sheetName} value={sheetName}>
                          {sheetName}
                        </option>
                      ))}
                    </select>
                  </div>
                )}

                <div className="field field-wide">
                  <label className="field-label" htmlFor="attachments-dir">
                    附件資料夾路徑（本機）
                  </label>
                  <div className="drive-browser-search">
                    <input
                      id="attachments-dir"
                      className="text-input"
                      type="text"
                      placeholder="例如 /Users/you/Desktop/attachments"
                      value={attachmentsDir}
                      onChange={(event) => setAttachmentsDir(event.target.value)}
                      disabled={isHostedFrontend === true}
                    />
                    {isHostedFrontend === false && (
                      <button
                        className="btn btn-outline btn-sm"
                        onClick={() => localAttachmentFolderInputRef.current?.click()}
                      >
                        瀏覽
                      </button>
                    )}
                  </div>
                  <span className="field-hint">
                    Excel 的「附件1／附件2」填檔名；系統會優先比對你手動挑選的 Drive
                    附件、這次瀏覽的本機資料夾，再找本機資料夾路徑、指定的 Drive 資料夾，最後才做 Drive fallback。
                  </span>
                  {isHostedFrontend === true && (
                    <span className="field-hint">Vercel 部署環境不提供本機資料夾瀏覽。</span>
                  )}
                  {selectedLocalAttachmentFiles.length > 0 && (
                    <div className="attachment-tray" aria-label="已選附件">
                      <div className="attachment-pill">
                        <span className="attachment-icon" aria-hidden="true">附</span>
                        <span className="attachment-name">本機檔案 {selectedLocalAttachmentFiles.length} 份</span>
                        <span className="attachment-meta">已載入</span>
                        <button className="attachment-remove" onClick={clearSelectedLocalAttachmentFiles}>
                          清除
                        </button>
                      </div>
                    </div>
                  )}
                  <div className="segmented" role="group" aria-label="附件來源">
                    <button
                      className="seg"
                      onClick={() => openGooglePicker("attachment")}
                      disabled={!gmailEmail}
                    >
                      雲端挑選附件
                    </button>
                    <button
                      className="seg"
                      onClick={() => openFolderBrowser("lookup")}
                      disabled={!gmailEmail}
                    >
                      雲端資料夾
                    </button>
                    <button
                      className="seg"
                      onClick={() => openFolderBrowser("create-parent")}
                      disabled={!gmailEmail || creatingCloudFolder}
                    >
                      建立雲端上傳夾
                    </button>
                    <button
                      className="seg"
                      onClick={() => openDriveBrowser("attachment")}
                      disabled={!gmailEmail}
                    >
                      最近附件
                    </button>
                  </div>
                  {cloudUploadFolder && (
                    <div className="attachment-cloud-box">
                      <div className="attachment-cloud-head">
                        <div>
                          <p className="attachment-cloud-title">{cloudUploadFolder.name}</p>
                          <p className="field-hint">此資料夾會作為 Excel 附件檔名的限定範圍。</p>
                        </div>
                        <button
                          className="btn btn-outline btn-sm"
                          type="button"
                          onClick={() => cloudAttachmentInputRef.current?.click()}
                          disabled={uploadingCloudFiles}
                        >
                          {uploadingCloudFiles ? "上傳中…" : "上傳檔案"}
                        </button>
                      </div>
                      <div
                        className={`attachment-dropline ${uploadingCloudFiles ? "active" : ""}`}
                        onDragOver={(event) => {
                          event.preventDefault();
                        }}
                        onDrop={(event) => {
                          event.preventDefault();
                          void uploadFilesToCloudFolder(event.dataTransfer.files);
                        }}
                      >
                        拖曳檔案到這一列即可上傳
                      </div>
                      {uploadedCloudFolderFiles.length > 0 && (
                        <div className="attachment-tray" aria-label="已上傳附件">
                          {uploadedCloudFolderFiles.map((file) => (
                            <div key={file.id} className="attachment-pill">
                              <span className="attachment-icon" aria-hidden="true">附</span>
                              <span className="attachment-name">{file.name}</span>
                              <span className="attachment-meta">已上傳</span>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                  {selectedDriveFolder && (!cloudUploadFolder || selectedDriveFolder.id !== cloudUploadFolder.id) && (
                    <div className="attachment-tray" aria-label="限定資料夾">
                      <div className="attachment-pill attachment-pill-folder">
                        <span className="attachment-icon" aria-hidden="true">夾</span>
                        <span className="attachment-name">{selectedDriveFolder.name}</span>
                        <span className="attachment-meta">資料夾</span>
                        <button className="attachment-remove" onClick={clearSelectedDriveFolder}>
                          移除資料夾
                        </button>
                      </div>
                    </div>
                  )}
                  {selectedDriveAttachments.length > 0 && (
                    <div className="attachment-tray" aria-label="已選 Drive 附件">
                      {selectedDriveAttachments.map((file) => (
                        <div key={file.id} className="attachment-pill">
                          <span className="attachment-icon" aria-hidden="true">附</span>
                          <span className="attachment-name">{file.name}</span>
                          <span className="attachment-meta">Drive</span>
                          <button
                            className="attachment-remove"
                            onClick={() => removeSelectedDriveAttachment(file.id)}
                          >
                            移除
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              {driveBrowserKind === "attachment" && (
                <div className="drive-browser">
                  <div className="drive-browser-head">
                    <div>
                      <p className="drive-browser-title">選擇附件</p>
                      <p className="field-hint">顯示最近 20 筆支援檔案，可搜尋檔名。</p>
                    </div>
                    <button className="link-btn" onClick={() => setDriveBrowserKind(null)}>
                      關閉
                    </button>
                  </div>
                  <div className="drive-browser-search">
                    <input
                      className="text-input"
                      type="text"
                      placeholder="輸入檔名關鍵字"
                      value={driveQuery}
                      onChange={(event) => setDriveQuery(event.target.value)}
                      onKeyDown={(event) => {
                        if (event.key === "Enter") {
                          event.preventDefault();
                          void loadDriveFiles("attachment", driveQuery);
                        }
                      }}
                    />
                    <button
                      className="btn btn-outline btn-sm"
                      onClick={() => void loadDriveFiles("attachment", driveQuery)}
                    >
                      搜尋
                    </button>
                  </div>
                  {driveError && <p className="error">{driveError}</p>}
                  <div className="drive-file-list">
                    {driveLoading ? (
                      <p className="field-hint">讀取中…</p>
                    ) : driveFiles.length > 0 ? (
                      driveFiles.map((file) => (
                        <button
                          key={file.id}
                          className="drive-file-item"
                          onClick={() => selectDriveFiles("attachment", [file])}
                        >
                          <span className="drive-file-name">{file.name}</span>
                          <span className="field-hint">
                            {file.is_google_workspace ? "Google 工作區" : "Office 檔案"}
                          </span>
                        </button>
                      ))
                    ) : (
                      <p className="field-hint">找不到可用檔案。</p>
                    )}
                  </div>
                </div>
              )}

              {driveBrowserKind === "folder" && (
                <div className="drive-browser">
                  <div className="drive-browser-head">
                    <div>
                      <p className="drive-browser-title">
                        {folderBrowserPurpose === "create-parent" ? "選擇建立位置" : "選擇雲端資料夾"}
                      </p>
                      <p className="field-hint">
                        {folderBrowserPurpose === "create-parent"
                          ? "先決定新資料夾要建立在雲端哪裡，建立後就能拖曳上傳多份本機檔案。"
                          : "Excel 附件檔名會先限定在這個資料夾內查找。"}
                      </p>
                    </div>
                    <button className="link-btn" onClick={() => setDriveBrowserKind(null)}>
                      關閉
                    </button>
                  </div>
                  {folderBrowserPurpose === "create-parent" && (
                    <button
                      className="drive-file-item"
                      onClick={() => void createCloudUploadFolder(null)}
                      disabled={creatingCloudFolder}
                    >
                      <span className="drive-file-name">我的雲端硬碟根目錄</span>
                      <span className="field-hint">{creatingCloudFolder ? "建立中…" : "建立在根目錄"}</span>
                    </button>
                  )}
                  <div className="drive-browser-search">
                    <input
                      className="text-input"
                      type="text"
                      placeholder={folderBrowserPurpose === "create-parent" ? "輸入建立位置資料夾名稱" : "輸入資料夾名稱"}
                      value={driveQuery}
                      onChange={(event) => setDriveQuery(event.target.value)}
                      onKeyDown={(event) => {
                        if (event.key === "Enter") {
                          event.preventDefault();
                          void loadDriveFiles("folder", driveQuery);
                        }
                      }}
                    />
                    <button
                      className="btn btn-outline btn-sm"
                      onClick={() => void loadDriveFiles("folder", driveQuery)}
                    >
                      搜尋
                    </button>
                  </div>
                  {driveError && <p className="error">{driveError}</p>}
                  <div className="drive-file-list">
                    {driveLoading ? (
                      <p className="field-hint">讀取中…</p>
                    ) : driveFiles.length > 0 ? (
                      driveFiles.map((file) => (
                        <button
                          key={file.id}
                          className="drive-file-item"
                          onClick={() => selectDriveFiles("folder", [file])}
                          disabled={creatingCloudFolder}
                        >
                          <span className="drive-file-name">{file.name}</span>
                          <span className="field-hint">
                            {folderBrowserPurpose === "create-parent" ? "建立在此資料夾下" : "Google 雲端資料夾"}
                          </span>
                        </button>
                      ))
                    ) : (
                      <p className="field-hint">找不到可用資料夾。</p>
                    )}
                  </div>
                </div>
              )}

              {status === "uploading" && (
                <div className="process-strip">
                  <div className="progress">
                    <div className="progress-bar" style={{ width: `${progress}%` }} />
                  </div>
                  <div className="progress-meta">
                    <span>狀態：{statusLabel[status]}</span>
                    <span>{progress}%</span>
                  </div>
                </div>
              )}
              {error && <p className="error">{error}</p>}
            </div>
          )}

          {step === 3 && (
            <div className="step-body">
              <div className="step-head">
                <div>
                  <h2 className="step-title">產生草稿</h2>
                  <p className="step-desc">確認第一筆預覽無誤後，批次寫入 Gmail 草稿匣。</p>
                </div>
                {preview && (
                  <div className="preview-meta">
                    <span className="meta-item">
                      <span className="meta-key">工作表</span>
                      {preview.selected_sheet ?? "未指定"}
                    </span>
                    <span className="meta-item">
                      <span className="meta-key">總筆數</span>
                      {preview.total_records}
                    </span>
                  </div>
                )}
              </div>

              {status === "uploading" && (
                <div className="process-strip">
                  <div className="progress">
                    <div className="progress-bar" style={{ width: `${progress}%` }} />
                  </div>
                  <div className="progress-meta">
                    <span>狀態：{statusLabel[status]}</span>
                    <span>{progress}%</span>
                  </div>
                </div>
              )}
              {error && <p className="error">{error}</p>}

              {preview?.sheet_names && preview.sheet_names.length > 0 && (
                <div className="field">
                  <label className="field-label" htmlFor="sheet-select-3">
                    收件人工作表
                  </label>
                  <select
                    id="sheet-select-3"
                    className="select"
                    value={selectedSheet}
                    onChange={(event) => void handleSheetChange(event.target.value)}
                    disabled={!cachedUpload?.cacheId || status === "uploading"}
                  >
                    {preview.sheet_names.map((sheetName) => (
                      <option key={sheetName} value={sheetName}>
                        {sheetName}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {preview && (
                <div className="detect">
                  <span className={`chip ${emailHeader ? "" : "warn"}`}>
                    收件人：{emailHeader ?? "未找到"}
                  </span>
                  <span className="chip subtle">副本：{ccHeader ?? "—"}</span>
                  <span className="chip subtle">密件副本：{bccHeader ?? "—"}</span>
                  <span className={`chip ${subjectHeader ? "" : "warn"}`}>
                    主旨：{subjectHeader ?? "未找到"}
                  </span>
                  {preview?.detected_fields?.attachments &&
                    preview.detected_fields.attachments.length > 0 && (
                      <span className="chip subtle">
                        附件：{preview.detected_fields.attachments.join("、")}
                      </span>
                    )}
                </div>
              )}
              {preview && missingHeaders.length > 0 && (
                <p className="error">缺少必要欄位：{missingHeaders.join("、")}</p>
              )}

              {preview && (
                <div className="xlsx-preview-panel">
                  <div className="xlsx-preview-head">
                    <div>
                      <h3 className="xlsx-preview-title">XLSX 首筆資料預覽</h3>
                      <p className="field-hint">這是目前信件全文預覽實際套用的第一列資料。</p>
                    </div>
                    <span className="xlsx-preview-count">第 1 / {preview.total_records} 筆</span>
                  </div>

                  {previewRowItems.length > 0 ? (
                    <>
                      {previewSummaryItems.length > 0 && (
                        <div className="xlsx-preview-summary">
                          {previewSummaryItems.map((item) => (
                            <div
                              key={`${item.label}-${item.header ?? "na"}`}
                              className={`xlsx-preview-card tone-${item.tone}`}
                            >
                              <span className="xlsx-preview-card-title">
                                <span>{item.label}</span>
                                <span>{item.header}</span>
                              </span>
                              <span
                                className={`xlsx-preview-card-value${
                                  item.value ? "" : " is-empty"
                                }`}
                              >
                                {item.value || "空白"}
                              </span>
                            </div>
                          ))}
                        </div>
                      )}

                      <div className="xlsx-field-list">
                        {previewRowItems.map((item) => (
                          <div key={item.header} className="xlsx-preview-item">
                            <div className="xlsx-preview-item-head">
                              <span className="xlsx-preview-item-label">{item.header}</span>
                              {item.tags.length > 0 && (
                                <span className="xlsx-preview-badges">
                                  {item.tags.map((tag) => (
                                    <span key={`${item.header}-${tag}`} className="xlsx-preview-badge">
                                      {tag}
                                    </span>
                                  ))}
                                </span>
                              )}
                            </div>
                            <div
                              className={`xlsx-preview-item-value${
                                item.value ? "" : " is-empty"
                              }`}
                            >
                              {item.value || "空白"}
                            </div>
                          </div>
                        ))}
                      </div>
                    </>
                  ) : (
                    <p className="field-hint">目前工作表沒有可預覽的資料列。</p>
                  )}
                </div>
              )}

              <div className="preview-card">
                {preview ? (
                  <div
                    className="preview-html"
                    dangerouslySetInnerHTML={{ __html: previewHtml }}
                  />
                ) : (
                  <div className="preview-empty">
                    <span className="preview-empty-mark" aria-hidden="true">
                      ✉
                    </span>
                    <p>尚未產生預覽</p>
                    <span>返回上一步後點「下一步」即可產生第一筆合併結果。</span>
                  </div>
                )}
              </div>

              {preview && (
                <div className="xlsx-table-panel">
                  <div className="xlsx-table-head">
                    <div>
                      <h3 className="xlsx-preview-title">XLSX 全列資料表</h3>
                      <p className="field-hint">顯示將用來產生草稿的每一列資料。</p>
                    </div>
                    <span className="xlsx-preview-count">
                      {previewTableRows.length} / {preview.total_records} 筆
                    </span>
                  </div>

                  <div className="xlsx-table-toolbar">
                    <label className="xlsx-page-size">
                      <span>每頁</span>
                      <select
                        className="select xlsx-page-size-select"
                        value={previewPagination?.page_size ?? previewPageSize}
                        onChange={(event) => {
                          const nextPageSize = Number(event.target.value);
                          void refreshPreviewPage(1, nextPageSize);
                        }}
                        disabled={status === "uploading"}
                      >
                        {previewPageSizeOptions.map((size) => (
                          <option key={size} value={size}>
                            {size}
                          </option>
                        ))}
                      </select>
                    </label>

                    {previewPagination && previewPagination.total_pages > 1 && (
                      <div className="xlsx-pagination">
                        <button
                          className="btn btn-outline btn-sm"
                          onClick={() => void refreshPreviewPage(previewPagination.page - 1)}
                          disabled={status === "uploading" || previewPagination.page <= 1}
                        >
                          上一頁
                        </button>
                        <span className="xlsx-pagination-status">
                          第 {previewPagination.page} / {previewPagination.total_pages} 頁
                        </span>
                        <button
                          className="btn btn-outline btn-sm"
                          onClick={() => void refreshPreviewPage(previewPagination.page + 1)}
                          disabled={
                            status === "uploading" ||
                            previewPagination.page >= previewPagination.total_pages
                          }
                        >
                          下一頁
                        </button>
                      </div>
                    )}
                  </div>

                  {previewTableRows.length > 0 && preview.headers?.length > 0 ? (
                    <div className="xlsx-table-scroll">
                      <table className="xlsx-table">
                        <thead>
                          <tr>
                            <th className="xlsx-row-number">#</th>
                            {preview.headers.map((header) => (
                              <th key={header}>{header}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {previewTableRows.map((row, rowIndex) => (
                            <tr key={`xlsx-row-${rowIndex}`}>
                              <td className="xlsx-row-number">
                                {((previewPagination?.page ?? 1) - 1) *
                                  (previewPagination?.page_size ?? previewPageSize) +
                                  rowIndex +
                                  1}
                              </td>
                              {preview.headers.map((header) => {
                                const value = row[header] ?? "";
                                return (
                                  <td key={`${rowIndex}-${header}`} className={value ? "" : "is-empty"}>
                                    {value || "空白"}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    <p className="field-hint">目前沒有可顯示的 XLSX 資料。</p>
                  )}
                </div>
              )}

              {preview && preview.headers?.length > 0 && (
                <div className="headers-strip">
                  <span className="headers-label">欄位</span>
                  {preview.headers.map((header) => (
                    <span key={header} className="header-tag">
                      {header}
                    </span>
                  ))}
                </div>
              )}

              {draftMessage && (
                <p className={draftStatus === "error" ? "error" : "hint"}>{draftMessage}</p>
              )}
              {draftFailedItems.length > 0 && (
                <div className="failed-items">
                  <p className="failed-title">失敗項目</p>
                  <ul className="failed-list">
                    {draftFailedItems.map((item, index) => (
                      <li key={`${item.row ?? "na"}-${item.field ?? "na"}-${item.name ?? index}-${index}`}>
                        第 {item.row ?? "?"} 列 / {item.field ?? "unknown"}
                        {item.name ? ` / ${item.name}` : ""}
                        {item.message ? `：${item.message}` : ""}
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </div>
          )}
        </div>

        <div className="wizard-nav">
          {step > 1 && (
            <button className="btn btn-ghost btn-back" onClick={goBack}>
              上一步
            </button>
          )}
          {step < 3 ? (
            <button
              className="btn btn-primary btn-next"
              onClick={goNext}
              disabled={(step === 1 && !canLeaveStep1) || (step === 2 && status === "uploading")}
            >
              {step === 2 && status === "uploading" ? "處理中…" : "下一步"}
            </button>
          ) : (
            <button
              className="btn btn-primary btn-next"
              onClick={saveDrafts}
              disabled={draftStatus === "saving" || missingHeaders.length > 0 || !preview}
            >
              {draftStatus === "saving" ? "批次建立中…" : "批次建立 Gmail 草稿"}
            </button>
          )}
        </div>
      </main>

      {showGuide && (
        <div
          className="guide-overlay"
          onClick={() => setShowGuide(false)}
          role="dialog"
          aria-modal="true"
          aria-label="運作原理引導動畫"
        >
          <div className="guide-modal" onClick={(event) => event.stopPropagation()}>
            <div className="guide-modal-head">
              <div>
                <p className="guide-modal-title">運作原理</p>
                <p className="guide-modal-sub">
                  Excel 的姓名 / 職稱 / 公司等欄位會代入 Word 模板的 <code>{"{{代換欄位}}"}</code>；收件人與主旨直接來自
                  Excel 欄位，一列資料生成一封草稿；附件依 <code>附件1/附件2…</code> 欄逐列指派。
                </p>
              </div>
              <button className="guide-close" onClick={() => setShowGuide(false)} aria-label="關閉">
                ✕
              </button>
            </div>
            <HowItWorks />
          </div>
        </div>
      )}
    </div>
  );
}
