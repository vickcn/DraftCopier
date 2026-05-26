"use client";

import { useEffect, useMemo, useRef, useState } from "react";

function parseDriveFileId(input: string): string | null {
  const s = input.trim();
  if (/^[a-zA-Z0-9_-]{25,}$/.test(s)) return s;
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/document\/d\/([a-zA-Z0-9_-]+)/,
    /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/,
  ];
  for (const p of patterns) {
    const m = s.match(p);
    if (m) return m[1];
  }
  return null;
}

type PreviewPayload = {
  total_records: number;
  headers: string[];
  preview_first_row: string;
  first_row?: Record<string, string>;
  cache_id?: string;
  cached_files?: {
    docx_name: string;
    xlsx_name: string;
  };
  detected_fields?: {
    email: string | null;
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

const apiBase = process.env.NEXT_PUBLIC_API_BASE || "";
const googlePickerApiKey = process.env.NEXT_PUBLIC_GOOGLE_PICKER_API_KEY || "";
const uploadCacheKey = "draftcopier_upload_cache_v1";
const userKeyCacheKey = "draftcopier_user_key_v1";
const userEmailCacheKey = "draftcopier_user_email_v1";
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
    } else if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
      result.xlsx = file;
    }
  });
  return result;
}

export default function Home() {
  const [docxFile, setDocxFile] = useState<File | null>(null);
  const [xlsxFile, setXlsxFile] = useState<File | null>(null);
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
  const [gmailEmail, setGmailEmail] = useState<string | null>(null);
  const [docxDriveId, setDocxDriveId] = useState<string | null>(null);
  const [xlsxDriveId, setXlsxDriveId] = useState<string | null>(null);
  const [docxDriveName, setDocxDriveName] = useState<string | null>(null);
  const [xlsxDriveName, setXlsxDriveName] = useState<string | null>(null);
  const [docxLinkInput, setDocxLinkInput] = useState("");
  const [xlsxLinkInput, setXlsxLinkInput] = useState("");

  const statusLabel: Record<UploadState, string> = {
    idle: "待命",
    uploading: "上傳中",
    done: "完成",
    error: "失敗",
  };

  const docxInputRef = useRef<HTMLInputElement>(null);
  const xlsxInputRef = useRef<HTMLInputElement>(null);
  const restoreAttemptedRef = useRef<string | null>(null);

  // 每次載入時，嘗試從 localStorage 還原 session（跨 session 保持登入狀態）
  useEffect(() => {
    const restoreSession = async () => {
      const cachedEmail = localStorage.getItem(userEmailCacheKey);
      if (cachedEmail) setGmailEmail(cachedEmail);

      const savedKey = localStorage.getItem(userKeyCacheKey);
      if (!savedKey) return;
      try {
        const res = await fetch(`${apiBase}/api/session/restore`, {
          method: "POST",
          credentials: "include",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ user_key: savedKey }),
        });
        const data = (await res.json()) as { ok: boolean; restored: boolean; email?: string | null };
        if (!data.restored) {
          localStorage.removeItem(userKeyCacheKey);
          localStorage.removeItem(userEmailCacheKey);
          setGmailEmail(null);
        } else if (data.email) {
          setGmailEmail(data.email);
          localStorage.setItem(userEmailCacheKey, data.email);
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
        const data = (await res.json()) as { user_key?: string; email?: string };
        if (data.user_key) {
          localStorage.setItem(userKeyCacheKey, data.user_key);
        }
        if (data.email) {
          setGmailEmail(data.email);
          localStorage.setItem(userEmailCacheKey, data.email);
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
      }
    } catch {
      localStorage.removeItem(uploadCacheKey);
    }
  }, []);

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
        setPreview(data);
        setStatus("done");
        setProgress(100);
      } catch {
        setStatus("idle");
        setProgress(0);
      }
    };
    void restorePreview();
  }, [cachedUpload, preview, selectedFont]);

  const emailHeader = useMemo(() => {
    if (!preview?.headers) return null;
    return findHeader(preview.headers, emailFieldCandidates);
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

  const canSubmit = useMemo(
    () =>
      (!!docxFile || !!docxDriveId) &&
      (!!xlsxFile || !!xlsxDriveId) &&
      status !== "uploading",
    [docxFile, xlsxFile, docxDriveId, xlsxDriveId, status]
  );

  const loadPickerApi = (): Promise<void> =>
    new Promise((resolve, reject) => {
      const g = (window as Record<string, unknown>).google as { picker?: unknown } | undefined;
      if (g?.picker) { resolve(); return; }
      if (document.getElementById("gapi-script")) {
        const wait = setInterval(() => {
          const gw = (window as Record<string, unknown>).google as { picker?: unknown } | undefined;
          if (gw?.picker) { clearInterval(wait); resolve(); }
        }, 100);
        return;
      }
      const script = document.createElement("script");
      script.id = "gapi-script";
      script.src = "https://apis.google.com/js/api.js";
      script.onload = () => {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (window as any).gapi.load("picker", resolve);
      };
      script.onerror = () => reject(new Error("Failed to load Google API script"));
      document.head.appendChild(script);
    });

  const openDrivePicker = async (type: "docx" | "xlsx") => {
    if (!gmailEmail) {
      setError("請先連結 Gmail 才能使用 Drive 選取器。");
      return;
    }
    try {
      const res = await fetch(`${apiBase}/api/auth/access-token`, { credentials: "include" });
      if (!res.ok) {
        setError("無法取得 Drive 授權，請重新連結 Gmail。");
        return;
      }
      const { access_token } = (await res.json()) as { access_token: string };
      await loadPickerApi();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const google = (window as any).google;
      const view = new google.picker.DocsView();
      if (type === "docx") {
        view.setMimeTypes(
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.google-apps.document"
        );
      } else {
        view.setMimeTypes(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,application/vnd.google-apps.spreadsheet"
        );
      }
      const builder = new google.picker.PickerBuilder()
        .addView(view)
        .setOAuthToken(access_token)
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .setCallback((data: any) => {
          if (data.action === google.picker.Action.PICKED) {
            const file = data.docs[0];
            if (type === "docx") {
              setDocxDriveId(file.id);
              setDocxDriveName(file.name);
              setDocxFile(null);
            } else {
              setXlsxDriveId(file.id);
              setXlsxDriveName(file.name);
              setXlsxFile(null);
            }
            setStatus("idle");
            setProgress(0);
            setPreview(null);
            setDraftStatus("idle");
            setDraftMessage(null);
            setDraftFailedItems([]);
          }
        });
      if (googlePickerApiKey) builder.setDeveloperKey(googlePickerApiKey);
      builder.build().setVisible(true);
    } catch {
      setError("無法開啟 Google Drive 選取器，請改用貼上連結方式。");
    }
  };

  const confirmDriveLink = (type: "docx" | "xlsx") => {
    const input = type === "docx" ? docxLinkInput : xlsxLinkInput;
    const fileId = parseDriveFileId(input);
    if (!fileId) {
      setError("無法解析 Drive 連結，請確認格式正確（例如 drive.google.com/file/d/...）。");
      return;
    }
    const shortName = `Drive 檔案（${fileId.slice(0, 8)}…）`;
    if (type === "docx") {
      setDocxDriveId(fileId);
      setDocxDriveName(shortName);
      setDocxFile(null);
      setDocxLinkInput("");
    } else {
      setXlsxDriveId(fileId);
      setXlsxDriveName(shortName);
      setXlsxFile(null);
      setXlsxLinkInput("");
    }
    setStatus("idle");
    setProgress(0);
    setPreview(null);
    setDraftStatus("idle");
    setDraftMessage(null);
    setDraftFailedItems([]);
  };

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragActive(false);
    const files = classifyFiles(event.dataTransfer.files);
    if (files.docx) {
      setDocxFile(files.docx);
      setDocxDriveId(null);
      setDocxDriveName(null);
    }
    if (files.xlsx) {
      setXlsxFile(files.xlsx);
      setXlsxDriveId(null);
      setXlsxDriveName(null);
    }
    setStatus("idle");
    setProgress(0);
    setPreview(null);
    setDraftStatus("idle");
    setDraftMessage(null);
    setDraftFailedItems([]);
  };

  const handleUpload = () => {
    if ((!docxFile && !docxDriveId) || (!xlsxFile && !xlsxDriveId)) {
      setError("請同時提供 .docx 模板與 .xlsx 清單（本機或 Drive）。");
      return;
    }

    setError(null);
    setStatus("uploading");
    setProgress(0);
    setPreview(null);
    setDraftStatus("idle");
    setDraftMessage(null);
    setDraftFailedItems([]);

    const formData = new FormData();
    if (docxFile) formData.append("docx_file", docxFile);
    if (xlsxFile) formData.append("xlsx_file", xlsxFile);
    if (docxDriveId) formData.append("docx_drive_id", docxDriveId);
    if (xlsxDriveId) formData.append("xlsx_drive_id", xlsxDriveId);

    const url = new URL(`${apiBase}/api/process`, window.location.origin);
    url.searchParams.set("font", selectedFont);
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
          setPreview(data);
          if (data.cache_id && data.cached_files) {
            const nextCached: CachedUpload = {
              cacheId: data.cache_id,
              docxName: data.cached_files.docx_name,
              xlsxName: data.cached_files.xlsx_name,
              font: selectedFont,
            };
            setCachedUpload(nextCached);
            localStorage.setItem(uploadCacheKey, JSON.stringify(nextCached));
            restoreAttemptedRef.current = data.cache_id;
          }
          setStatus("done");
          setProgress(100);
        } catch (err) {
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
    setDocxDriveId(null);
    setXlsxDriveId(null);
    setDocxDriveName(null);
    setXlsxDriveName(null);
    setDocxLinkInput("");
    setXlsxLinkInput("");
    setProgress(0);
    setStatus("idle");
    setError(null);
    setPreview(null);
    setDraftStatus("idle");
    setDraftMessage(null);
    setAttachmentsDir("");
    setCachedUpload(null);
    localStorage.removeItem(uploadCacheKey);
    restoreAttemptedRef.current = null;
    if (docxInputRef.current) docxInputRef.current.value = "";
    if (xlsxInputRef.current) xlsxInputRef.current.value = "";
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
      if (cachedUpload?.cacheId) {
        formData.append("cache_id", cachedUpload.cacheId);
      } else if (docxFile && xlsxFile) {
        formData.append("docx_file", docxFile);
        formData.append("xlsx_file", xlsxFile);
      } else if (docxDriveId && xlsxDriveId) {
        formData.append("docx_drive_id", docxDriveId);
        formData.append("xlsx_drive_id", xlsxDriveId);
      } else {
        setDraftStatus("error");
        setDraftMessage("請先上傳 DOCX 與 XLSX。");
        return;
      }

      const url = new URL(`${apiBase}/api/drafts/batch`, window.location.origin);
      url.searchParams.set("font", selectedFont);
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

  return (
    <div className="page">
      <div className="ambient" aria-hidden="true" />
      <header className="topbar">
        <div className="brand">
          <div className="logo">DC</div>
          <div>
            <p className="title">DraftCopier</p>
            <p className="subtitle">DOCX + Excel 轉 Gmail 草稿流程</p>
          </div>
        </div>
        <div className="pill">測試中</div>
      </header>

      <main className="main">
        <section className="hero">
          <div className="hero-text">
            <p className="eyebrow">批次寄信，從模板開始</p>
            <h1>一次完成草稿</h1>
            <p className="lead">
              上傳 Word 模板與 Excel 清單，系統會轉換樣式、合併欄位，
              並立即顯示第一筆預覽。
            </p>
            <div className="checks">
              <span>保留粗體、底線與文字顏色</span>
              <span>自動辨識欄位標題做合併</span>
              <span>已準備 Gmail 草稿整合</span>
            </div>
          </div>

          <div className="card">
            <div className="actions">
              <button className="ghost" onClick={connectGmail}>
                {gmailEmail ? "重新連結 Gmail" : "先連結 Gmail"}
              </button>
              {gmailEmail && (
                <span className="field-label" style={{ alignSelf: "center" }}>
                  已連結：{gmailEmail}
                </span>
              )}
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
              <div>
                <p className="drop-title">拖放檔案到此</p>
                <p className="drop-subtitle">.docx 模板 + .xlsx 收件人清單</p>
              </div>
              <div className="drop-actions">
                <label className="file-btn">
                  選擇 DOCX
                  <input
                    ref={docxInputRef}
                    type="file"
                    accept=".docx"
                    onChange={(event) => {
                      setDocxFile(event.target.files?.[0] ?? null);
                      setDocxDriveId(null);
                      setDocxDriveName(null);
                      setStatus("idle");
                      setProgress(0);
                      setPreview(null);
                      setDraftStatus("idle");
                      setDraftMessage(null);
                    }}
                  />
                </label>
                <label className="file-btn">
                  選擇 XLSX
                  <input
                    ref={xlsxInputRef}
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={(event) => {
                      setXlsxFile(event.target.files?.[0] ?? null);
                      setXlsxDriveId(null);
                      setXlsxDriveName(null);
                      setStatus("idle");
                      setProgress(0);
                      setPreview(null);
                      setDraftStatus("idle");
                      setDraftMessage(null);
                    }}
                  />
                </label>
              </div>

              <div className="drive-section">
                <p className="drop-subtitle" style={{ marginBottom: "0.5rem" }}>或從 Google Drive 匯入</p>

                <div className="drive-row">
                  <span className="drive-label">DOCX</span>
                  <input
                    className="text-input drive-link-input"
                    type="text"
                    placeholder="貼上 Drive 連結或 File ID"
                    value={docxLinkInput}
                    onChange={(e) => setDocxLinkInput(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter") confirmDriveLink("docx"); }}
                  />
                  <button className="ghost drive-confirm-btn" onClick={() => confirmDriveLink("docx")}>確認</button>
                  <button className="ghost drive-confirm-btn" onClick={() => void openDrivePicker("docx")}>選取</button>
                </div>

                <div className="drive-row">
                  <span className="drive-label">XLSX</span>
                  <input
                    className="text-input drive-link-input"
                    type="text"
                    placeholder="貼上 Drive 連結或 File ID"
                    value={xlsxLinkInput}
                    onChange={(e) => setXlsxLinkInput(e.target.value)}
                    onKeyDown={(e) => { if (e.key === "Enter") confirmDriveLink("xlsx"); }}
                  />
                  <button className="ghost drive-confirm-btn" onClick={() => confirmDriveLink("xlsx")}>確認</button>
                  <button className="ghost drive-confirm-btn" onClick={() => void openDrivePicker("xlsx")}>選取</button>
                </div>
              </div>
            </div>

            <div className="file-grid">
              <div className="file-tile">
                <p className="file-label">模板</p>
                <p className="file-name">
                  {docxFile
                    ? docxFile.name
                    : docxDriveName
                    ? docxDriveName
                    : cachedUpload?.docxName ?? "尚未選擇 DOCX"}
                </p>
              </div>
              <div className="file-tile">
                <p className="file-label">收件人清單</p>
                <p className="file-name">
                  {xlsxFile
                    ? xlsxFile.name
                    : xlsxDriveName
                    ? xlsxDriveName
                    : cachedUpload?.xlsxName ?? "尚未選擇 XLSX"}
                </p>
              </div>
            </div>

            <div className="field">
              <label className="field-label" htmlFor="font-select">
                字型（Gmail 支援）
              </label>
              <select
                id="font-select"
                className="font-select"
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

            <div className="field">
              <label className="field-label" htmlFor="attachments-dir">
                附件資料夾路徑（本機）
              </label>
              <input
                id="attachments-dir"
                className="text-input"
                type="text"
                placeholder="例如 /Users/you/Desktop/attachments"
                value={attachmentsDir}
                onChange={(event) => setAttachmentsDir(event.target.value)}
              />
              <span className="field-label">
                Excel 的「附件1/附件2」填檔名；若填的是絕對路徑，可留空。
              </span>
            </div>

            <div className="actions">
              <button className="primary" onClick={handleUpload} disabled={!canSubmit}>
                {status === "uploading" ? "上傳中..." : "上傳並預覽"}
              </button>
              <button className="ghost" onClick={resetFiles}>
                清除
              </button>
            </div>

            <div className="progress">
              <div className="progress-bar" style={{ width: `${progress}%` }} />
            </div>
            <div className="progress-meta">
              <span>狀態：{statusLabel[status]}</span>
              <span>{progress}%</span>
            </div>

            {error && <p className="error">{error}</p>}

            <div className="divider" />

            <div className="actions">
              <button className="primary" onClick={saveDrafts} disabled={draftStatus === "saving"}>
                {draftStatus === "saving" ? "批次建立中..." : "批次建立 Gmail 草稿"}
              </button>
            </div>

            <div className="field">
              <span className="field-label">
                收件人欄位：{emailHeader ?? "未找到"}
              </span>
              <span className="field-label">
                主旨欄位：{subjectHeader ?? "未找到"}
              </span>
              {preview?.detected_fields?.attachments &&
                preview.detected_fields.attachments.length > 0 && (
                  <span className="field-label">
                    附件欄位：{preview.detected_fields.attachments.join("、")}
                  </span>
                )}
              {preview && missingHeaders.length > 0 && (
                <span className="error">缺少必要欄位：{missingHeaders.join("、")}</span>
              )}
            </div>

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
        </section>

        <section className="preview">
          <div className="preview-head">
            <div>
              <h2>預覽</h2>
              <p>顯示第一筆合併結果（保留格式）。</p>
            </div>
            {preview && (
              <div className="preview-meta">
                <span>總筆數：{preview.total_records}</span>
                <span>欄位：{preview.headers.join("、")}</span>
              </div>
            )}
          </div>

          <div className="preview-card">
            {preview ? (
              <div
                className="preview-html"
                dangerouslySetInnerHTML={{ __html: preview.preview_first_row }}
              />
            ) : (
              <div className="preview-empty">
                <p>尚未產生預覽，請上傳檔案。</p>
              </div>
            )}
          </div>
        </section>
      </main>
    </div>
  );
}
