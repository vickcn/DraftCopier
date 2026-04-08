"use client";

import { useMemo, useRef, useState } from "react";

type PreviewPayload = {
  total_records: number;
  headers: string[];
  preview_first_row: string;
  first_row?: Record<string, string>;
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

const apiBase = process.env.NEXT_PUBLIC_API_BASE || "http://localhost:6311";
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
  const [progress, setProgress] = useState(0);
  const [status, setStatus] = useState<UploadState>("idle");
  const [error, setError] = useState<string | null>(null);
  const [preview, setPreview] = useState<PreviewPayload | null>(null);
  const [isDragActive, setIsDragActive] = useState(false);
  const [selectedFont, setSelectedFont] = useState(fontOptions[0].value);
  const [draftStatus, setDraftStatus] = useState<"idle" | "saving" | "done" | "error">("idle");
  const [draftMessage, setDraftMessage] = useState<string | null>(null);
  const [attachmentsDir, setAttachmentsDir] = useState("");

  const statusLabel: Record<UploadState, string> = {
    idle: "待命",
    uploading: "上傳中",
    done: "完成",
    error: "失敗",
  };

  const docxInputRef = useRef<HTMLInputElement>(null);
  const xlsxInputRef = useRef<HTMLInputElement>(null);

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
    () => !!docxFile && !!xlsxFile && status !== "uploading",
    [docxFile, xlsxFile, status]
  );

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setIsDragActive(false);
    const files = classifyFiles(event.dataTransfer.files);
    if (files.docx) setDocxFile(files.docx);
    if (files.xlsx) setXlsxFile(files.xlsx);
    setStatus("idle");
    setProgress(0);
    setPreview(null);
    setDraftStatus("idle");
    setDraftMessage(null);
  };

  const handleUpload = () => {
    if (!docxFile || !xlsxFile) {
      setError("請同時選擇 .docx 與 .xlsx 檔案。");
      return;
    }

    setError(null);
    setStatus("uploading");
    setProgress(0);
    setPreview(null);
    setDraftStatus("idle");
    setDraftMessage(null);

    const formData = new FormData();
    formData.append("docx_file", docxFile);
    formData.append("xlsx_file", xlsxFile);

    const url = new URL(`${apiBase}/api/process`);
    url.searchParams.set("font", selectedFont);
    const xhr = new XMLHttpRequest();
    xhr.open("POST", url.toString());

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
    setProgress(0);
    setStatus("idle");
    setError(null);
    setPreview(null);
    setDraftStatus("idle");
    setDraftMessage(null);
    setAttachmentsDir("");
    if (docxInputRef.current) docxInputRef.current.value = "";
    if (xlsxInputRef.current) xlsxInputRef.current.value = "";
  };

  const connectGmail = async () => {
    setDraftMessage(null);
    try {
      await fetch(`${apiBase}/api/dev/login`, {
        method: "POST",
        credentials: "include",
      });
      const response = await fetch(`${apiBase}/api/auth/google`, {
        method: "GET",
        credentials: "include",
      });
      if (!response.ok) {
        setDraftStatus("error");
        setDraftMessage("取得授權連結失敗。");
        return;
      }
      const data = (await response.json()) as { auth_url?: string };
      if (!data.auth_url) {
        setDraftStatus("error");
        setDraftMessage("授權連結格式錯誤。");
        return;
      }
      window.location.href = data.auth_url;
    } catch (err) {
      setDraftStatus("error");
      setDraftMessage("啟動授權流程失敗，請稍後再試。");
    }
  };

  const saveDrafts = async () => {
    if (!preview) {
      setDraftStatus("error");
      setDraftMessage("尚未產生預覽，無法儲存草稿。");
      return;
    }
    if (missingHeaders.length > 0) {
      setDraftStatus("error");
      setDraftMessage("找不到必要欄位（email / subject），請檢查 Excel 標題列。");
      return;
    }
    setDraftStatus("saving");
    setDraftMessage(null);
    try {
      const formData = new FormData();
      if (!docxFile || !xlsxFile) {
        setDraftStatus("error");
        setDraftMessage("請先上傳 DOCX 與 XLSX。");
        return;
      }
      formData.append("docx_file", docxFile);
      formData.append("xlsx_file", xlsxFile);

      const url = new URL(`${apiBase}/api/drafts/batch`);
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
        return;
      }
      if (!response.ok) {
        const text = await response.text();
        setDraftStatus("error");
        setDraftMessage(`儲存草稿失敗：${text}`);
        return;
      }

      const data = (await response.json()) as { draft_count?: number };
      setDraftStatus("done");
      setDraftMessage(`已建立 Gmail 草稿：${data.draft_count ?? 0} 封。`);
    } catch (err) {
      setDraftStatus("error");
      setDraftMessage("儲存草稿時發生錯誤。");
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
                      setStatus("idle");
                      setProgress(0);
                      setPreview(null);
                      setDraftStatus("idle");
                      setDraftMessage(null);
                    }}
                  />
                </label>
              </div>
            </div>

            <div className="file-grid">
              <div className="file-tile">
                <p className="file-label">模板</p>
                <p className="file-name">
                  {docxFile ? docxFile.name : "尚未選擇 DOCX"}
                </p>
              </div>
              <div className="file-tile">
                <p className="file-label">收件人清單</p>
                <p className="file-name">
                  {xlsxFile ? xlsxFile.name : "尚未選擇 XLSX"}
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
              <button className="ghost" onClick={connectGmail}>
                連結 Gmail
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
              {missingHeaders.length > 0 && (
                <span className="error">缺少必要欄位：{missingHeaders.join("、")}</span>
              )}
            </div>

            {draftMessage && (
              <p className={draftStatus === "error" ? "error" : "hint"}>{draftMessage}</p>
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
