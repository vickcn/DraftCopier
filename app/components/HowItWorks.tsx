"use client";

import { CSSProperties, useEffect, useRef, useState } from "react";

type Props = {
  /** 動畫速度倍率（時長 = 14 / speed 秒） */
  speed?: number;
  primary?: string;
  accent?: string;
};

/** 需要連線對應的欄位（收件人/主旨直接作為草稿欄位，不屬於 Word 模板代換） */
const KEYS = ["姓名", "職稱", "公司"] as const;
const MONO = "'JetBrains Mono', ui-monospace, monospace";

/** Excel 表頭：姓名/職稱/公司會映射到 Word 模板；收件人/主旨直接作為草稿欄位 */
const HEADERS: { label: string; map?: boolean; att?: boolean }[] = [
  { label: "收件人" },
  { label: "主旨" },
  { label: "姓名", map: true },
  { label: "職稱", map: true },
  { label: "公司", map: true },
  { label: "附件1", att: true },
  { label: "附件2", att: true },
];

const GRID_COLS = "1.5fr 1.15fr .8fr .95fr .75fr .95fr .95fr";

/**
 * 聯動原理引導動畫：Excel 欄位資料 → Word 模板代換欄位 + 草稿欄位 → 逐列指派附件 → 批次 Gmail 草稿。
 * 純 CSS keyframes 驅動 14s 無縫循環；JS 僅負責計數器與 SVG 連接線量測。
 * 尊重 prefers-reduced-motion：顯示已填好的最終狀態，不跑循環。
 */
export default function HowItWorks({ speed = 1, primary, accent }: Props) {
  const scalerRef = useRef<HTMLDivElement>(null);
  const stageRef = useRef<HTMLDivElement>(null);
  const bodyRef = useRef<HTMLDivElement>(null);
  const svgRef = useRef<SVGSVGElement>(null);
  const scaleRef = useRef(1);

  const [count, setCount] = useState(1);
  const [reduced, setReduced] = useState(false);

  const durSec = 14 / (speed || 1);

  useEffect(() => {
    const mq = window.matchMedia("(prefers-reduced-motion: reduce)");
    const apply = () => setReduced(mq.matches);
    apply();
    mq.addEventListener?.("change", apply);
    return () => mq.removeEventListener?.("change", apply);
  }, []);

  useEffect(() => {
    if (reduced) {
      setCount(24);
      return;
    }
    const period = durSec * 1000;
    const t0 = performance.now();
    let raf = 0;
    const tick = (now: number) => {
      const p = ((((now - t0) / period) % 1) + 1) % 1;
      let c: number;
      if (p < 0.66) c = 1;
      else if (p > 0.9) c = 24;
      else c = Math.round(1 + ((p - 0.66) / 0.24) * 23);
      setCount((prev) => (prev === c ? prev : c));
      raf = requestAnimationFrame(tick);
    };
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [reduced, durSec]);

  useEffect(() => {
    const scaler = scalerRef.current;
    const stage = stageRef.current;
    if (!scaler || !stage) return;

    const buildWires = () => {
      const svg = svgRef.current;
      const host = bodyRef.current;
      if (!svg || !host) return;
      const scale = scaleRef.current || 1;
      const hr = svg.getBoundingClientRect();
      if (hr.width < 10) return;
      while (svg.firstChild) svg.removeChild(svg.firstChild);
      const NS = "http://www.w3.org/2000/svg";
      const dur = `${durSec}s`;
      KEYS.forEach((k, i) => {
        const a = host.querySelector(`[data-map="${k}"]`);
        const b = host.querySelector(`[data-token="${k}"]`);
        if (!a || !b) return;
        const ra = a.getBoundingClientRect();
        const rb = b.getBoundingClientRect();
        const sx = (ra.right - hr.left) / scale;
        const sy = (ra.top + ra.height / 2 - hr.top) / scale;
        const ex = (rb.left - hr.left) / scale;
        const ey = (rb.top + rb.height / 2 - hr.top) / scale;
        const mx = (sx + ex) / 2;
        const d = `M ${sx} ${sy} C ${mx} ${sy}, ${mx} ${ey}, ${ex} ${ey}`;
        const stroke = i % 2 ? "var(--hiw-accent)" : "var(--hiw-primary)";

        const path = document.createElementNS(NS, "path");
        path.setAttribute("d", d);
        path.setAttribute("fill", "none");
        path.setAttribute("stroke", stroke);
        path.setAttribute("stroke-width", "1.5");
        path.setAttribute("pathLength", "1");
        path.setAttribute("stroke-linecap", "round");
        path.style.strokeDasharray = "1";
        path.style.strokeDashoffset = "1";
        path.style.opacity = "0";
        path.style.filter = "drop-shadow(0 0 3px rgba(122,162,255,.4))";
        if (!reduced) {
          path.style.animation = `hiwConnDraw ${dur} cubic-bezier(.7,0,.2,1) infinite`;
          path.style.animationDelay = `${i * 0.09}s`;
        }
        svg.appendChild(path);

        const dot = document.createElementNS(NS, "circle");
        dot.setAttribute("cx", String(ex));
        dot.setAttribute("cy", String(ey));
        dot.setAttribute("r", "2.6");
        dot.setAttribute("fill", stroke);
        dot.style.opacity = "0";
        dot.style.filter = "drop-shadow(0 0 4px rgba(95,224,207,.5))";
        if (!reduced) {
          dot.style.animation = `hiwConnDraw ${dur} cubic-bezier(.7,0,.2,1) infinite`;
          dot.style.animationDelay = `${i * 0.09 + 0.14}s`;
        }
        svg.appendChild(dot);
      });
    };

    const fit = () => {
      const w = scaler.clientWidth;
      const viewportW = window.innerWidth;
      const viewportH = window.innerHeight;
      const widthScale = w / 1080;
      const heightReserve = viewportW <= 480 ? 240 : viewportW <= 720 ? 220 : 170;
      const heightScale = Math.max(0.45, (viewportH - heightReserve) / 660);
      const s = Math.min(1, widthScale, heightScale);
      scaleRef.current = s;
      stage.style.transform = `scale(${s})`;
      scaler.style.height = `${660 * s}px`;
      buildWires();
    };

    fit();
    const ro = new ResizeObserver(fit);
    ro.observe(scaler);

    let cancelled = false;
    const fonts = (document as Document & { fonts?: FontFaceSet }).fonts;
    fonts?.ready?.then(() => {
      if (!cancelled) buildWires();
    });
    const late = window.setTimeout(buildWires, 400);

    return () => {
      cancelled = true;
      ro.disconnect();
      window.clearTimeout(late);
    };
  }, [reduced, durSec, primary, accent]);

  const stageStyle: CSSProperties = {
    ...(primary ? ({ "--hiw-primary": primary } as CSSProperties) : {}),
    ...(accent ? ({ "--hiw-accent": accent } as CSSProperties) : {}),
    ...({ "--dur": `${durSec}s` } as CSSProperties),
  };

  /** 有動態時回傳 animation 樣式，reduced 時回傳空物件 */
  const A = (value: string, delay?: string): CSSProperties =>
    reduced ? {} : { animation: value, animationDelay: delay };

  type TokenProps = {
    token: (typeof KEYS)[number];
    ph: string;
    value: string;
    sizer: string;
    sizerMono?: boolean;
    valueMono?: boolean;
    valueWeight?: number;
    small?: boolean;
    delay: string;
  };

  const Token = ({
    token,
    ph,
    value,
    sizer,
    sizerMono,
    valueMono,
    valueWeight,
    small,
    delay,
  }: TokenProps) => {
    const px = small ? 7 : 8;
    const py = small ? 1 : 2;
    return (
      <span
        data-token={token}
        style={{
          position: "relative",
          display: "inline-block",
          padding: small ? "1px 7px" : "2px 8px",
          borderRadius: 6,
          border: "1px solid rgba(255,255,255,.12)",
          background: "rgba(122,162,255,.04)",
          ...A(`hiwTokPulse var(--dur) linear infinite`, delay),
        }}
      >
        <span
          style={{ opacity: 0, fontFamily: sizerMono ? MONO : undefined, fontSize: sizerMono ? 11 : undefined }}
        >
          {sizer}
        </span>
        <span
          style={{
            position: "absolute",
            left: px,
            top: py,
            opacity: reduced ? 0 : 1,
            color: "#9aa4bd",
            fontFamily: MONO,
            fontSize: 11,
            whiteSpace: "nowrap",
            ...A(`hiwTokPh var(--dur) linear infinite`, delay),
          }}
        >
          {ph}
        </span>
        <span
          style={{
            position: "absolute",
            left: px,
            top: py,
            opacity: reduced ? 1 : 0,
            color: "#eaf0ff",
            fontWeight: valueWeight,
            fontFamily: valueMono ? MONO : undefined,
            fontSize: valueMono ? 11 : undefined,
            whiteSpace: "nowrap",
            ...A(`hiwTokVal var(--dur) linear infinite`, delay),
          }}
        >
          {value}
        </span>
      </span>
    );
  };

  const cellStyle = (delay: string): CSSProperties => ({
    padding: 8,
    fontFamily: "'Noto Sans TC', sans-serif",
    color: "#c8cfdd",
    ...A(`hiwCellFill var(--dur) linear infinite`, delay),
  });

  const hdrBase: CSSProperties = {
    padding: "7px 6px",
    fontWeight: 600,
    fontFamily: "'Noto Sans TC', sans-serif",
    background: "rgba(255,255,255,.03)",
    color: "#8a93a8",
    whiteSpace: "nowrap",
  };

  /** Excel 一般儲存格（第 2/3 列，靜態） */
  const plainCell: CSSProperties = { padding: 8, fontFamily: "'Noto Sans TC', sans-serif" };
  const monoCell: CSSProperties = {
    padding: 8,
    fontFamily: MONO,
    fontSize: 10.5,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  };
  const blankCell: CSSProperties = { padding: 8, fontFamily: MONO, fontSize: 10.5, color: "#3a4254" };

  const stackCard = (
    tx: string,
    ty: string,
    rot: string,
    bg: string,
    border: string,
    z: number,
    delay: string
  ): CSSProperties => ({
    ["--tx" as string]: tx,
    ["--ty" as string]: ty,
    ["--rot" as string]: rot,
    position: "absolute",
    inset: 0,
    borderRadius: 14,
    background: bg,
    border: `1px solid ${border}`,
    boxShadow: "0 14px 34px rgba(0,0,0,.4)",
    zIndex: z,
    ...A(`hiwCardStack var(--dur) cubic-bezier(.7,0,.2,1) infinite`, delay),
  });

  /** 卡片右上角的附件數徽章 */
  const CountBadge = ({ label, tone, animate }: { label: string; tone: "has" | "none"; animate?: boolean }) => (
    <span
      style={{
        position: "absolute",
        top: 12,
        right: 12,
        zIndex: 5,
        display: "inline-flex",
        alignItems: "center",
        gap: 5,
        padding: "3px 9px",
        borderRadius: 999,
        fontSize: 11,
        fontWeight: 600,
        whiteSpace: "nowrap",
        color: tone === "none" ? "#8a93a8" : "#d8fff6",
        background: tone === "none" ? "rgba(255,255,255,.05)" : "rgba(95,224,207,.16)",
        border: `1px solid ${tone === "none" ? "rgba(255,255,255,.12)" : "rgba(95,224,207,.45)"}`,
        ...(animate ? A(`hiwAttachIn var(--dur) linear infinite`) : {}),
      }}
    >
      {label}
    </span>
  );

  const Chip = ({ name }: { name: string }) => (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 5,
        padding: "3px 9px",
        borderRadius: 7,
        fontSize: 11.5,
        color: "#dbe6ff",
        background: "rgba(122,162,255,.1)",
        border: "1px solid rgba(122,162,255,.32)",
      }}
    >
      <span aria-hidden="true">📎</span>
      <span style={{ fontFamily: MONO, fontSize: 11 }}>{name}</span>
    </span>
  );

  const pill = (label: string, mark: string, markColor: string, glow: string): CSSProperties => ({
    display: "flex",
    alignItems: "center",
    gap: 6,
    padding: "5px 11px",
    borderRadius: 20,
    background: markColor.includes("accent") ? "rgba(95,224,207,.1)" : "rgba(122,162,255,.1)",
    border: `1px solid ${markColor.includes("accent") ? "rgba(95,224,207,.28)" : "rgba(122,162,255,.28)"}`,
    opacity: reduced ? 1 : undefined,
    ...A(`${glow} var(--dur) linear infinite`),
  });

  return (
    <div className="hiw-scaler" ref={scalerRef}>
      <div className="hiw-stage" ref={stageRef} style={stageStyle}>
        <div
          style={{
            position: "absolute",
            inset: 0,
            pointerEvents: "none",
            background:
              "radial-gradient(520px 300px at 12% 108%, rgba(95,224,207,.06), transparent 70%),radial-gradient(600px 320px at 92% 8%, rgba(122,162,255,.08), transparent 70%)",
          }}
        />
        <svg
          width="1080"
          height="660"
          viewBox="0 0 1080 660"
          preserveAspectRatio="none"
          style={{ position: "absolute", inset: 0, opacity: 0.5, pointerEvents: "none", ...A(`hiwFloatLine calc(var(--dur) * 2.4) ease-in-out infinite`) }}
        >
          <path d="M-40 210 C 240 150, 520 300, 1120 180" fill="none" stroke="rgba(122,162,255,.14)" strokeWidth="1" />
          <path d="M-40 470 C 300 560, 640 380, 1120 520" fill="none" stroke="rgba(95,224,207,.1)" strokeWidth="1" />
        </svg>

        <div style={{ position: "absolute", inset: 0, ...A(`hiwSceneLoop var(--dur) linear infinite`) }}>
          {/* 頂部列 */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "26px 40px 0" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
              <div
                style={{
                  width: 26,
                  height: 26,
                  borderRadius: 7,
                  background: "linear-gradient(135deg,var(--hiw-primary),var(--hiw-accent))",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontWeight: 700,
                  fontSize: 13,
                  color: "#06080e",
                }}
              >
                D
              </div>
              <div style={{ fontFamily: "'Fraunces',serif", fontSize: 17, fontWeight: 600, letterSpacing: 0.2 }}>
                DraftCopier
              </div>
              <div
                style={{ fontSize: 12, color: "#6a7488", letterSpacing: 0.5, paddingLeft: 9, borderLeft: "1px solid rgba(255,255,255,.12)" }}
              >
                運作原理
              </div>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12.5, fontWeight: 500 }}>
              <div style={pill("對應欄位", "primary", "primary", "hiwStep1Glow")}>
                <span style={{ color: "var(--hiw-primary)" }}>①</span>對應欄位
              </div>
              <span style={{ color: "#4a5266" }}>→</span>
              <div style={pill("填入資料", "accent", "accent", "hiwStep2Glow")}>
                <span style={{ color: "var(--hiw-accent)" }}>②</span>填入資料
              </div>
              <span style={{ color: "#4a5266" }}>→</span>
              <div style={pill("指派附件", "accent", "accent", "hiwStep3Glow")}>
                <span style={{ color: "var(--hiw-accent)" }}>③</span>指派附件
              </div>
              <span style={{ color: "#4a5266" }}>→</span>
              <div style={pill("批次生成", "primary", "primary", "hiwStep4Glow")}>
                <span style={{ color: "var(--hiw-primary)" }}>④</span>批次生成
              </div>
            </div>
          </div>

          {/* 主體 */}
          <div
            ref={bodyRef}
            style={{ position: "relative", display: "flex", alignItems: "center", justifyContent: "space-between", padding: "22px 44px 0", height: 520 }}
          >
            {/* Excel 卡片 */}
            <div
              style={{
                width: 500,
                borderRadius: 14,
                background: "linear-gradient(180deg,#131a27,#0f1420)",
                border: "1px solid rgba(255,255,255,.08)",
                boxShadow: "0 18px 44px rgba(0,0,0,.4)",
                overflow: "hidden",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 9,
                  padding: "11px 14px",
                  borderBottom: "1px solid rgba(255,255,255,.07)",
                  background: "rgba(255,255,255,.02)",
                }}
              >
                <div
                  style={{
                    width: 22,
                    height: 22,
                    borderRadius: 6,
                    background: "#1f7a4d",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 11,
                    fontWeight: 700,
                    color: "#eafff4",
                  }}
                >
                  XLS
                </div>
                <span style={{ fontSize: 13, fontWeight: 600 }}>recipients.xlsx</span>
                <span style={{ marginLeft: "auto", fontSize: 11, color: "#6a7488", fontFamily: MONO }}>工作表1 · 24 列</span>
              </div>
              <div style={{ padding: "12px 12px 14px" }}>
                {/* 表頭 */}
                <div style={{ display: "grid", gridTemplateColumns: GRID_COLS, fontSize: 11 }}>
                  {HEADERS.map((h, i) => (
                    <div
                      key={h.label}
                      data-map={h.map ? h.label : undefined}
                      style={{
                        ...hdrBase,
                        borderRadius: i === 0 ? "6px 0 0 6px" : i === HEADERS.length - 1 ? "0 6px 6px 0" : undefined,
                        ...(h.map
                          ? A(`hiwHdrGlow var(--dur) linear infinite`, `${i * 0.09}s`)
                          : A(`hiwAttColGlow var(--dur) linear infinite`)),
                      }}
                    >
                      {h.label}
                    </div>
                  ))}
                </div>

                {/* 第 1 列（來源列，2 個附件） */}
                <div style={{ position: "relative", borderRadius: 7, marginTop: 5, ...A(`hiwRowGlow var(--dur) linear infinite`) }}>
                  <div style={{ display: "grid", gridTemplateColumns: GRID_COLS, fontSize: 11, borderBottom: "1px solid rgba(255,255,255,.05)" }}>
                    <div style={{ ...cellStyle("0s"), fontFamily: MONO, fontSize: 10.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                      alex@acme.co
                    </div>
                    <div style={cellStyle(".09s")}>產品發表邀請</div>
                    <div style={cellStyle(".18s")}>陳世恩</div>
                    <div style={cellStyle(".27s")}>產品總監</div>
                    <div style={cellStyle(".36s")}>Acme</div>
                    <div style={{ ...cellStyle(".45s"), fontFamily: MONO, fontSize: 10.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                      簡章.pdf
                    </div>
                    <div style={{ ...cellStyle(".54s"), fontFamily: MONO, fontSize: 10.5, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                      報名表.pdf
                    </div>
                  </div>
                </div>

                {/* 第 2 列（1 個附件） */}
                <div style={{ borderRadius: 7, ...A(`hiwRowTick var(--dur) linear infinite`) }}>
                  <div style={{ display: "grid", gridTemplateColumns: GRID_COLS, fontSize: 11, borderBottom: "1px solid rgba(255,255,255,.05)", color: "#9aa4bd" }}>
                    <div style={monoCell}>mei@nova.io</div>
                    <div style={plainCell}>產品發表邀請</div>
                    <div style={plainCell}>林美惠</div>
                    <div style={plainCell}>行銷經理</div>
                    <div style={plainCell}>Nova</div>
                    <div style={{ ...monoCell, color: "#9aa4bd" }}>簡章.pdf</div>
                    <div style={blankCell}>—</div>
                  </div>
                </div>

                {/* 第 3 列（無附件） */}
                <div style={{ borderRadius: 7, ...A(`hiwRowTick var(--dur) linear infinite`, ".3s") }}>
                  <div style={{ display: "grid", gridTemplateColumns: GRID_COLS, fontSize: 11, color: "#9aa4bd" }}>
                    <div style={monoCell}>john@peak.com</div>
                    <div style={plainCell}>產品發表邀請</div>
                    <div style={plainCell}>王建國</div>
                    <div style={plainCell}>執行長</div>
                    <div style={plainCell}>Peak</div>
                    <div style={blankCell}>—</div>
                    <div style={blankCell}>—</div>
                  </div>
                </div>

                <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "9px 6px 2px", fontSize: 10.5, color: "#5a6376" }}>
                  <span style={{ width: 5, height: 5, borderRadius: "50%", background: "#3a4254" }} />… 其餘 21 列，附件欄各列獨立
                </div>
              </div>
            </div>

            {/* 草稿卡區 */}
            <div style={{ position: "relative", width: 420, height: 448 }}>
              {!reduced && (
                <>
                  <div style={stackCard("26px", "20px", "4deg", "#0f1420", "rgba(255,255,255,.06)", 1, ".24s")} />
                  <div style={stackCard("16px", "12px", "2.4deg", "#111725", "rgba(255,255,255,.06)", 2, ".16s")}>
                    <CountBadge label="無附件" tone="none" />
                  </div>
                  <div style={stackCard("8px", "5px", "1.1deg", "#131a29", "rgba(255,255,255,.07)", 3, ".08s")}>
                    <CountBadge label="📎 1 附件" tone="has" />
                  </div>
                </>
              )}

              {/* 主卡 */}
              <div
                style={{
                  position: "absolute",
                  inset: 0,
                  zIndex: 4,
                  borderRadius: 14,
                  background: "linear-gradient(180deg,#151b2a,#10151f)",
                  border: "1px solid rgba(255,255,255,.1)",
                  boxShadow: "0 20px 50px rgba(0,0,0,.5)",
                  overflow: "hidden",
                  ...A(`hiwMainAssemble var(--dur) cubic-bezier(.7,0,.2,1) infinite`),
                }}
              >
                <CountBadge label="📎 2 附件" tone="has" animate />
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 9,
                    padding: "11px 14px",
                    borderBottom: "1px solid rgba(255,255,255,.07)",
                    background: "rgba(255,255,255,.02)",
                  }}
                >
                  <div
                    style={{
                      width: 22,
                      height: 22,
                      borderRadius: 6,
                      background: "#2b57c4",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 10,
                      fontWeight: 700,
                      color: "#eaf0ff",
                    }}
                  >
                    DOC
                  </div>
                  <span style={{ fontSize: 13, fontWeight: 600 }}>invitation.docx</span>
                  <span style={{ marginLeft: "auto", fontSize: 11, color: "#6a7488", fontFamily: MONO }}>模板</span>
                </div>
                <div style={{ padding: "16px 20px", fontFamily: "'Noto Sans TC',sans-serif" }}>
                  <div style={{ display: "flex", alignItems: "baseline", gap: 8, fontSize: 12.5, marginBottom: 9 }}>
                    <span style={{ color: "#6a7488", width: 34, flex: "none" }}>收件</span>
                    <span style={{ color: "#d7deeb", fontFamily: MONO, fontSize: 11 }}>alex@acme.co</span>
                  </div>
                  <div
                    style={{
                      display: "flex",
                      alignItems: "baseline",
                      gap: 8,
                      fontSize: 12.5,
                      paddingBottom: 12,
                      borderBottom: "1px solid rgba(255,255,255,.07)",
                    }}
                  >
                    <span style={{ color: "#6a7488", width: 34, flex: "none" }}>主旨</span>
                    <span style={{ color: "#d7deeb", fontWeight: 500 }}>產品發表邀請</span>
                  </div>
                  <div style={{ fontSize: 13.5, lineHeight: 1.9, color: "#dfe6f2", marginTop: 13 }}>
                    親愛的{" "}
                    <Token token="姓名" ph="{{姓名}}" value="陳世恩" sizer="陳世恩" small delay="0s" />{" "}
                    <Token token="職稱" ph="{{職稱}}" value="產品總監" sizer="產品總監" small delay=".09s" />
                    ，您好：
                    <div style={{ color: "#9aa4bd", fontSize: 12.5, lineHeight: 1.9, marginTop: 6 }}>
                      誠摯邀請{" "}
                      <Token token="公司" ph="{{公司}}" value="Acme" sizer="Acme" small delay=".18s" />{" "}
                      蒞臨本公司年度產品發表會。
                    </div>
                  </div>

                  {/* 附件列：由 附件1/附件2 欄逐列指派 */}
                  <div
                    data-attach
                    style={{
                      marginTop: 13,
                      paddingTop: 12,
                      borderTop: "1px solid rgba(255,255,255,.07)",
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                      flexWrap: "wrap",
                      opacity: reduced ? 1 : undefined,
                      ...A(`hiwAttachIn var(--dur) linear infinite`),
                    }}
                  >
                    <span style={{ color: "#6a7488", fontSize: 12, flex: "none" }}>附件</span>
                    <Chip name="簡章.pdf" />
                    <Chip name="報名表.pdf" />
                  </div>

                  <div style={{ color: "#6a7488", fontSize: 12, marginTop: 12 }}>DraftCopier 團隊　敬上</div>
                </div>
              </div>
            </div>

            {/* 連接線 overlay */}
            <svg ref={svgRef} width="1080" height="660" style={{ position: "absolute", inset: 0, pointerEvents: "none", overflow: "visible", zIndex: 6 }} />
          </div>

          {/* 附件規則字幕 */}
          <div
            style={{
              position: "absolute",
              left: 0,
              right: 0,
              bottom: 66,
              display: "flex",
              justifyContent: "center",
              pointerEvents: "none",
              opacity: reduced ? 1 : undefined,
              ...A(`hiwCaptionIn var(--dur) linear infinite`),
            }}
          >
            <span
              style={{
                fontSize: 12,
                color: "#c8cfdd",
                background: "rgba(9,13,22,.72)",
                border: "1px solid rgba(95,224,207,.28)",
                padding: "6px 15px",
                borderRadius: 999,
              }}
            >
              📎 附件依「附件1・附件2…」欄<b style={{ color: "var(--hiw-accent)", fontWeight: 600 }}>逐列指派</b>，留空即略過 —— 每封信附件數量可不同
            </span>
          </div>

          {/* 底部狀態列 */}
          <div style={{ position: "absolute", left: 0, right: 0, bottom: 22, display: "flex", alignItems: "center", justifyContent: "center", gap: 16 }}>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 11,
                padding: "9px 16px",
                borderRadius: 12,
                background: "rgba(122,162,255,.1)",
                border: "1px solid rgba(122,162,255,.3)",
                opacity: reduced ? 1 : undefined,
                ...A(`hiwBadgeIn var(--dur) cubic-bezier(.7,0,.2,1) infinite`),
              }}
            >
              <div
                style={{
                  width: 22,
                  height: 22,
                  borderRadius: 6,
                  background: "linear-gradient(135deg,var(--hiw-primary),var(--hiw-accent))",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: 12,
                  color: "#06080e",
                  fontWeight: 700,
                }}
              >
                ✉
              </div>
              <span style={{ fontSize: 13, fontWeight: 500, color: "#c8cfdd" }}>Gmail 草稿匣</span>
              <span style={{ fontFamily: MONO, fontSize: 16, fontWeight: 600, color: "#eaf0ff", minWidth: 52, textAlign: "right" }}>
                {count}
                <span style={{ color: "#6a7488", fontSize: 12 }}> / 24</span>
              </span>
            </div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                fontSize: 13,
                fontWeight: 500,
                color: "var(--hiw-accent)",
                opacity: reduced ? 1 : undefined,
                ...A(`hiwSuccessIn var(--dur) cubic-bezier(.7,0,.2,1) infinite`),
              }}
            >
              <span
                style={{
                  width: 19,
                  height: 19,
                  borderRadius: "50%",
                  background: "var(--hiw-accent)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: 12,
                  color: "#06080e",
                  fontWeight: 700,
                }}
              >
                ✓
              </span>
              24 封草稿已建立
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
