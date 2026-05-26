import { useCallback, useEffect, useRef, useState } from "react";
import Reveal from "@/components/Reveal";

/**
 * BeforeAfter — drag-to-compare block.
 * LEFT (Before):  chaotic spreadsheet — auto-scrolls rows like a video.
 * RIGHT (After):  kozai-built dispatch console — a fake cursor clicks
 *                 through rows + actions on a loop, so it feels alive.
 *
 * Mobile (<768px): tab toggle. Both sides re-flow into a taller, single
 * column composition so nothing is squished.
 *
 * Both animations pause when the section is off-screen via IO, and
 * respect prefers-reduced-motion (static fallback).
 */

// 20 rows of "data" — duplicated in the track to make scrolling seamless.
const SHEET_ROWS = [
  ["SKU-1042", "Aurora 12pk",        "Northgate Distrib",  "ON",  "47",   "$1,284.50"],
  ["SKU-1043", "Aurora 24pk",        "Northgate Distrib",  "ON",  "12",   "$642.20"],
  ["SKU-2091", "Halcyon Tonic",      "Merrick Medical",    "QC",  "—",    "#REF!"],
  ["SKU-2092", "Halcyon Tonic L",    "Merrick Medical",    "QC",  "8",    "$214.00"],
  ["SKU-3310", "Ironwood Bracket",   "Ironwood Mfg",       "ON",  "184",  "$3,941.60"],
  ["SKU-3311", "Ironwood Plate",     "Ironwood Mfg",       "ON",  "62",   "$1,302.40"],
  ["SKU-4001", "Atrium Filter",      "Atrium Partners",    "BC",  "—",    "#ERROR!"],
  ["SKU-4002", "Atrium Filter HD",   "Atrium Partners",    "BC",  "23",   "$1,012.50"],
  ["SKU-5512", "Harlan Pack",        "Harlan Foods",       "AB",  "401",  "$8,422.00"],
  ["SKU-5513", "Harlan Pack XL",     "Harlan Foods",       "AB",  "98",   "$3,140.80"],
  ["SKU-6720", "Fairlane Strap",     "Fairlane Freight",   "ON",  "??",   "TBD"],
  ["SKU-6721", "Fairlane Strap XL",  "fairlane freight",   "on",  "44",   "#N/A"],
  ["SKU-7102", "Mariner Hook",       "Mariner Supply",     "NS",  "120",  "$2,402.00"],
  ["SKU-7103", "Mariner Hook XL",    "Mariner Supply",     "NS",  "14",   "$298.00"],
  ["SKU-8210", "Cedar Cone",         "Cedar Outdoors",     "BC",  "—",    "#REF!"],
  ["SKU-8211", "Cedar Cone (sm)",    "Cedar Outdoors",     "BC",  "82",   "$1,114.10"],
  ["SKU-9001", "Quintile Pen",       "Quintile Office",    "ON",  "1,204","$2,408.00"],
  ["SKU-9002", "Quintile Pen XL",    "quintile office",    "on",  "24",   "$120.00"],
  ["SKU-9501", "Pinecrest Tag",      "Pinecrest Realty",   "ON",  "55",   "$1,375.00"],
  ["SKU-9502", "Pinecrest Tag (alt)","Pinecrest Realty",   "ON",  "—",    "#N/A"],
] as const;

// Column grid template — tuned so values aren't clipped on the right.
// Item and Client both take fractional space; Client is *smaller* than Item.
const COL_GRID = "60px minmax(0,1.3fr) minmax(0,1fr) 32px 48px 88px";

const Spreadsheet = ({ playing }: { playing: boolean }) => {
  // Render the rows twice so a -50% translateY loop is seamless.
  const rows = [...SHEET_ROWS, ...SHEET_ROWS];

  return (
    <div className="relative flex h-full w-full flex-col overflow-hidden bg-[#FBFAF6] font-mono text-[10px] text-ink/85 md:text-[11px]">
      {/* Faux Excel title bar */}
      <div className="flex shrink-0 items-center gap-3 border-b border-ink/15 bg-[#E8E4D8] px-2.5 py-1 text-[9px] uppercase tracking-[0.16em] text-ink/55 md:px-3 md:py-1.5 md:text-[10px]">
        <span className="truncate">inventory_2026_05_v17_FINAL_v2 (copy)(jen-edits).xlsx</span>
        <span className="ml-auto shrink-0 text-ink/40">— Excel</span>
      </div>
      {/* Formula bar */}
      <div className="flex shrink-0 items-center gap-2 border-b border-ink/10 bg-[#F3EFE3] px-2.5 py-1 text-[9px] text-ink/60 md:px-3 md:text-[10px]">
        <span>F12</span>
        <span className="text-ink/30">|</span>
        <span className="text-ink/55">fx</span>
        <span className="truncate text-ink/40">=IFERROR(VLOOKUP(A12,prices!A:C,3,FALSE),"#REF!")</span>
      </div>

      {/* Column header row (static, never scrolls) */}
      <div
        className="grid shrink-0 border-b border-ink/15 bg-[#EDE8D9] text-[9px] uppercase tracking-[0.14em] text-ink/55 md:text-[10px]"
        style={{ gridTemplateColumns: COL_GRID }}
      >
        {["A · SKU", "B · Item", "C · Client", "D", "E · Qty", "F · Value"].map((h, i) => (
          <div key={h} className={`truncate px-2 py-1 ${i < 5 ? "border-r border-ink/15" : ""}`}>
            {h}
          </div>
        ))}
      </div>

      {/* Scrolling rows section — fills remaining vertical space. */}
      <div className="relative flex-1 overflow-hidden">
        <div
          className="absolute inset-0 overflow-hidden"
          style={{
            maskImage:
              "linear-gradient(to bottom, transparent 0, #000 22px, #000 calc(100% - 22px), transparent 100%)",
            WebkitMaskImage:
              "linear-gradient(to bottom, transparent 0, #000 22px, #000 calc(100% - 22px), transparent 100%)",
          }}
        >
          <div
            className="will-change-transform"
            style={{
              animation: playing
                ? "kz-sheet-scroll 32s linear infinite"
                : undefined,
            }}
          >
            {rows.map((r, i) => {
              const real = i % SHEET_ROWS.length;
              const highlight =
                real === 4
                  ? "rgba(252,211,77,0.35)"
                  : real === 8
                    ? "rgba(134,239,172,0.30)"
                    : real === 16
                      ? "rgba(186,210,250,0.30)"
                      : undefined;
              return (
                <div
                  key={i}
                  className={`grid border-b border-ink/10 ${
                    real % 2 ? "bg-[#FBFAF6]" : "bg-[#F6F2E6]"
                  }`}
                  style={{ gridTemplateColumns: COL_GRID }}
                >
                  {r.map((cell, ci) => {
                    const isErr = cell === "#REF!" || cell === "#ERROR!" || cell === "#N/A";
                    const isTbd = cell === "TBD" || cell === "??";
                    return (
                      <div
                        key={ci}
                        className={`truncate px-2 py-[3px] tabular-nums ${ci < 5 ? "border-r border-ink/10" : ""}`}
                        style={{
                          color: isErr ? "#B5321A" : isTbd ? "#A85B12" : undefined,
                          background: isErr
                            ? "rgba(181,50,26,0.10)"
                            : ci === 0
                              ? undefined
                              : highlight,
                          fontWeight: isErr ? 600 : undefined,
                        }}
                        title={cell}
                      >
                        {cell}
                      </div>
                    );
                  })}
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* Bottom tabs strip — sits flush at the bottom of the flex column. */}
      <div className="flex shrink-0 items-center gap-2.5 overflow-hidden border-t border-ink/15 bg-[#E8E4D8] px-2.5 py-1 text-[8px] uppercase tracking-[0.12em] text-ink/55 md:gap-3 md:px-3 md:text-[9px] md:tracking-[0.14em]">
        <span>Sheet1</span>
        <span>Sheet2</span>
        <span className="hidden sm:inline">backup_2025-12-03 (3)</span>
        <span className="hidden md:inline">do_not_delete</span>
        <span className="hidden sm:inline">prices_OLD</span>
        <span className="ml-auto shrink-0 whitespace-nowrap text-[#B5321A]">3 errors · unsaved 14m</span>
      </div>
    </div>
  );
};

interface Dispatch {
  route: string;
  customer: string;
  vehicle: string;
  status: string;
  eta: string;
  address: string;
  contact: string;
  notes: string;
}

const DISPATCHES: Dispatch[] = [
  { route: "RT-2841", customer: "Northgate Distribution", vehicle: "TRK-12", status: "in transit", eta: "ETA 2:14p",
    address: "418 Industrial Pkwy, Hamilton ON",   contact: "Marcus Hale · (905) 555-0142",  notes: "Dock B · ring bell twice. Avoid 10–11a shift change." },
  { route: "RT-2842", customer: "Merrick Medical",        vehicle: "TRK-04", status: "scheduled",  eta: "Depart 1:30p",
    address: "1212 Centre St, Mississauga ON",     contact: "Priya Shah · (905) 555-0188",   notes: "Gate code 4471 · loading bay 3 only." },
  { route: "RT-2843", customer: "Ironwood Mfg",           vehicle: "TRK-09", status: "delivered",  eta: "Done 12:48p",
    address: "78 Foundry Rd, Brantford ON",        contact: "Linda Park · (519) 555-0103",   notes: "Signed for by L. Park · photo on file." },
  { route: "RT-2844", customer: "Atrium Partners",        vehicle: "TRK-15", status: "in transit", eta: "ETA 3:02p",
    address: "4400 Yonge St, Toronto ON",          contact: "Reza Tehrani · (416) 555-0271", notes: "Concierge: leave with desk." },
  { route: "RT-2845", customer: "Harlan Foods",           vehicle: "TRK-21", status: "scheduled",  eta: "Depart 2:45p",
    address: "9 Logistics Way, Vaughan ON",        contact: "T. Adeyemi · (905) 555-0244",   notes: "Refrigerated unit required." },
  { route: "RT-2846", customer: "Fairlane Freight",       vehicle: "TRK-07", status: "delivered",  eta: "Done 11:32a",
    address: "61 Cargo Blvd, Etobicoke ON",        contact: "Sam Cho · (416) 555-0322",      notes: "Multi-pallet — driver to confirm count." },
];

const statusColor = (s: string) => {
  if (s === "in transit") return { color: "#F5803E", bg: "rgba(245,128,62,0.10)" };
  if (s === "delivered") return { color: "#3E8F5A", bg: "rgba(62,143,90,0.10)" };
  return { color: "rgba(15,15,18,0.55)", bg: "rgba(15,15,18,0.05)" };
};

/**
 * Cursor cycle — moves between row positions and the CTA, clicking each.
 * Indices: 0–5 → rows. 6 → "Mark delivered" CTA. Loops.
 */
const CURSOR_STEPS: number[] = [1, 3, 0, 6, 2, 4, 0, 6];

const KozaiConsole = ({ playing }: { playing: boolean }) => {
  const [selectedIdx, setSelectedIdx] = useState(0);
  const [stepI, setStepI] = useState(0);
  const [clicking, setClicking] = useState(false);
  const [ctaPulse, setCtaPulse] = useState(false);
  const [cursorXY, setCursorXY] = useState<{ x: number; y: number } | null>(null);

  const rootRef = useRef<HTMLDivElement>(null);
  const rowRefs = useRef<(HTMLLIElement | null)[]>([]);
  const ctaRef = useRef<HTMLButtonElement>(null);

  const selected = DISPATCHES[selectedIdx];
  const currentTarget = CURSOR_STEPS[stepI % CURSOR_STEPS.length];

  // Cycle the cursor through steps
  useEffect(() => {
    if (!playing) return;
    const reduced = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    if (reduced) return;
    let cancelled = false;
    let i = stepI;
    const tick = () => {
      if (cancelled) return;
      setStepI(i);
      const target = CURSOR_STEPS[i % CURSOR_STEPS.length];
      // After the cursor visibly arrives (~0.9s), trigger the "click".
      window.setTimeout(() => {
        if (cancelled) return;
        setClicking(true);
        if (target === 6) {
          setCtaPulse(true);
          window.setTimeout(() => setCtaPulse(false), 700);
        } else {
          setSelectedIdx(target);
        }
        window.setTimeout(() => setClicking(false), 320);
      }, 900);
      i++;
      window.setTimeout(tick, 2400);
    };
    const initial = window.setTimeout(tick, 700);
    return () => {
      cancelled = true;
      window.clearTimeout(initial);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [playing]);

  // Recompute cursor position from the DOM whenever target changes
  // (also after resize so it stays accurate).
  useEffect(() => {
    if (!playing) return;
    const compute = () => {
      const root = rootRef.current;
      if (!root) return;
      const rootRect = root.getBoundingClientRect();
      let targetRect: DOMRect | undefined;
      let offsetX = 0;
      if (currentTarget === 6) {
        const el = ctaRef.current;
        if (!el) return;
        targetRect = el.getBoundingClientRect();
        // tip lands ~65% across the button
        offsetX = targetRect.width * 0.55;
      } else {
        const el = rowRefs.current[currentTarget];
        if (!el) return;
        targetRect = el.getBoundingClientRect();
        // tip lands ~30% across the row (over the customer name)
        offsetX = Math.min(140, targetRect.width * 0.3);
      }
      const x = targetRect.left - rootRect.left + offsetX;
      const y = targetRect.top - rootRect.top + targetRect.height * 0.5;
      setCursorXY({ x, y });
    };
    compute();
    window.addEventListener("resize", compute);
    return () => window.removeEventListener("resize", compute);
  }, [currentTarget, selectedIdx, playing]);

  return (
    <div ref={rootRef} className="relative h-full w-full overflow-hidden bg-paper">
      {/* Top bar */}
      <div className="flex items-center gap-3 border-b border-ink/12 px-3 py-2 md:gap-4 md:px-5 md:py-2.5">
        <div className="font-mono text-[9px] uppercase tracking-[0.24em] text-mute md:text-[10px] md:tracking-[0.28em]">
          DISPATCH · v2.1
        </div>
        <div className="relative mx-auto hidden w-[200px] sm:block md:w-[280px]">
          <input
            disabled
            value="Search route, customer, vehicle…"
            className="w-full border border-ink/12 bg-paper px-3 py-1.5 font-mono text-[10px] text-mute/70 md:text-[11px]"
            readOnly
          />
        </div>
        <div className="ml-auto flex items-center gap-2">
          <div className="flex h-6 w-6 items-center justify-center border border-ink/15 bg-ink/5 font-mono text-[10px] text-ink/70">
            JR
          </div>
          <div className="hidden leading-tight md:block">
            <div className="text-[11px] text-ink">Jen Reyes</div>
            <div className="font-mono text-[9px] uppercase tracking-[0.2em] text-mute">Lead dispatcher</div>
          </div>
        </div>
      </div>

      {/* Toolbar */}
      <div className="flex flex-wrap items-center gap-1.5 border-b border-ink/10 px-3 py-2 md:gap-2 md:px-5">
        <div className="flex flex-wrap items-center gap-1.5 font-mono text-[9px] uppercase tracking-[0.16em] md:gap-2 md:text-[10px] md:tracking-[0.18em]">
          <span className="border border-ink/20 px-2 py-1 text-ink">Today · 6</span>
          <span className="hidden border border-ink/15 px-2 py-1 text-mute sm:inline">Region · GTA-W</span>
          <span className="hidden border border-ink/15 px-2 py-1 text-mute md:inline">Status · Active</span>
        </div>
        <button
          type="button"
          className={`ml-auto px-2.5 py-1.5 text-[11px] font-medium transition-transform md:text-[12px] ${
            ctaPulse ? "scale-[0.96]" : "scale-100"
          }`}
          style={{ background: "#F5803E", color: "#F1EEE5" }}
        >
          + New dispatch
        </button>
      </div>

      {/* Body — desktop is row-list (60) + detail (40); mobile stacks vertical. */}
      <div
        className="flex flex-col md:flex-row"
        style={{ height: "calc(100% - 84px - 28px)" }}
      >
        {/* Rows list */}
        <div className="relative w-full overflow-hidden border-b border-ink/10 md:w-[60%] md:border-b-0 md:border-r">
          <ul>
            {DISPATCHES.map((d, i) => {
              const isActive = i === selectedIdx;
              const sc = statusColor(d.status);
              return (
                <li
                  key={d.route}
                  ref={(el) => { rowRefs.current[i] = el; }}
                  className="relative flex items-center gap-2 px-3 py-2 md:gap-3 md:px-4 md:py-2.5"
                  style={{
                    borderBottom: "1px solid rgba(15,15,18,0.08)",
                    borderLeft: isActive ? "2px solid #F5803E" : "2px solid transparent",
                    background: isActive ? "rgba(245,128,62,0.05)" : undefined,
                    transition: "background 220ms cubic-bezier(0.16,1,0.3,1), border-color 220ms",
                  }}
                >
                  {/* Cursor + ripple now live at the console root (see below) for accurate positioning. */}
                  <div className="w-[58px] shrink-0 font-mono text-[10px] text-ink/85 md:text-[11px]">{d.route}</div>
                  <div className="flex-1 truncate text-[12px] text-ink md:text-[13px]">{d.customer}</div>
                  <div className="hidden w-[54px] shrink-0 font-mono text-[10px] text-mute sm:block">{d.vehicle}</div>
                  <div
                    className="font-mono text-[8.5px] uppercase tracking-[0.14em] px-1.5 py-0.5 md:text-[9px] md:tracking-[0.18em] md:px-2"
                    style={{ color: sc.color, background: sc.bg }}
                  >
                    {d.status}
                  </div>
                  <div className="hidden w-[76px] shrink-0 text-right font-mono text-[10px] text-ink/70 sm:block">{d.eta}</div>
                </li>
              );
            })}
          </ul>

        </div>

        {/* Detail panel */}
        <div className="relative flex w-full flex-col md:w-[40%]">
          <div className="border-b border-ink/10 px-3 py-2 md:px-4 md:py-2.5">
            <div className="font-mono text-[9px] uppercase tracking-[0.22em] text-mute">Selected · {selected.route}</div>
            <div
              key={selected.route}
              className="mt-0.5 text-[13px] font-semibold text-ink md:text-[14px]"
              style={{ animation: "kz-detail-in 0.32s cubic-bezier(0.16,1,0.3,1)" }}
            >
              {selected.customer}
            </div>
          </div>
          <div
            key={selected.route + "-body"}
            className="space-y-2 px-3 py-2.5 text-[11px] leading-[1.5] text-ink/80 md:space-y-2.5 md:px-4 md:py-3 md:text-[12px]"
            style={{ animation: "kz-detail-in 0.42s cubic-bezier(0.16,1,0.3,1)" }}
          >
            <div>
              <div className="mb-0.5 font-mono text-[9px] uppercase tracking-[0.22em] text-mute">Address</div>
              <div>{selected.address}</div>
            </div>
            <div>
              <div className="mb-0.5 font-mono text-[9px] uppercase tracking-[0.22em] text-mute">Contact</div>
              <div>{selected.contact}</div>
            </div>
            <div>
              <div className="mb-0.5 font-mono text-[9px] uppercase tracking-[0.22em] text-mute">Notes</div>
              <div className="text-ink/70">{selected.notes}</div>
            </div>
          </div>

          {/* Tiny route map */}
          <div className="mx-3 mb-3 border border-ink/10 bg-ink/[0.02] p-2 md:mx-4">
            <div className="mb-1 font-mono text-[9px] uppercase tracking-[0.22em] text-mute">Route</div>
            <svg viewBox="0 0 200 50" width="100%" height="36" aria-hidden>
              <path
                d="M 8 35 Q 40 10, 70 30 T 130 22 T 192 14"
                fill="none"
                stroke="#0F0F12"
                strokeOpacity="0.65"
                strokeWidth="1.2"
                strokeLinecap="round"
                strokeDasharray="220"
                strokeDashoffset={playing ? 0 : 220}
                style={{ transition: "stroke-dashoffset 1.4s cubic-bezier(0.16,1,0.3,1)" }}
              />
              <circle cx="8" cy="35" r="2.5" fill="#0F0F12" />
              <circle cx="70" cy="30" r="2.5" fill="#0F0F12" />
              <circle cx="130" cy="22" r="2.5" fill="#0F0F12" />
              <circle cx="192" cy="14" r="3" fill="#F5803E" />
            </svg>
          </div>

          <button
            ref={ctaRef}
            type="button"
            className={`relative mx-3 mb-3 mt-auto border border-ink px-3 py-2 text-[12px] font-medium md:mx-4 md:mb-4 ${
              ctaPulse ? "scale-[0.97]" : "scale-100"
            }`}
            style={{
              background: "#0F0F12",
              color: "#F1EEE5",
              transition: "transform 220ms cubic-bezier(0.16,1,0.3,1)",
            }}
          >
            Mark delivered
            {/* Click ripple on CTA */}
            {ctaPulse && (
              <span
                aria-hidden
                className="pointer-events-none absolute inset-0"
                style={{
                  border: "1.5px solid #F5803E",
                  animation: "kz-cta-ripple 0.6s cubic-bezier(0.16,1,0.3,1)",
                }}
              />
            )}
          </button>
        </div>
      </div>

      {/* Bottom status */}
      <div className="absolute bottom-0 left-0 right-0 flex items-center justify-between border-t border-ink/12 bg-paper px-3 py-1.5 font-mono text-[9px] uppercase tracking-[0.18em] text-mute md:px-5 md:py-2 md:text-[10px] md:tracking-[0.22em]">
        <span>[ ✦ — synced 12s ago ]</span>
        <span className="hidden text-ink/70 sm:inline">6 active · 2 delivered today</span>
      </div>

      {/* Fake cursor (desktop only) — positioned at the console root so it
          can travel between the rows list and the detail-panel CTA. Coords
          come from real refs, so the tip lines up with the click target. */}
      {playing && cursorXY && (
        <>
          <div
            aria-hidden
            className="pointer-events-none absolute z-20 hidden md:block"
            style={{
              left: cursorXY.x - 2,
              top: cursorXY.y - 1,
              transform: clicking ? "scale(0.9)" : "scale(1)",
              transformOrigin: "2px 1px",
              transition:
                "left 0.65s cubic-bezier(0.16,1,0.3,1), top 0.65s cubic-bezier(0.16,1,0.3,1), transform 0.2s cubic-bezier(0.16,1,0.3,1)",
            }}
          >
            <svg width="18" height="22" viewBox="0 0 18 22" fill="none">
              <path
                d="M2 1 L2 17 L6 13 L9 20 L11.5 19 L8.5 12 L14 12 Z"
                fill="#0F0F12"
                stroke="#F1EEE5"
                strokeWidth="1.1"
                strokeLinejoin="round"
              />
            </svg>
          </div>
          {clicking && (
            <span
              key={`${stepI}-ripple`}
              aria-hidden
              className="pointer-events-none absolute z-20 hidden md:block"
              style={{
                left: cursorXY.x - 11,
                top: cursorXY.y - 11,
                width: 22,
                height: 22,
                borderRadius: "50%",
                border: "1.5px solid #F5803E",
                animation: "kz-cursor-ripple-fixed 0.42s cubic-bezier(0.16,1,0.3,1)",
              }}
            />
          )}
        </>
      )}
    </div>
  );
};

const BeforeAfter = () => {
  const wrapRef = useRef<HTMLDivElement>(null);
  const sectionRef = useRef<HTMLElement>(null);
  const [pos, setPos] = useState(50); // percent — width of LEFT (Before) pane
  const [dragging, setDragging] = useState(false);
  const [isMobile, setIsMobile] = useState(false);
  const [mobileTab, setMobileTab] = useState<"before" | "after">("before");
  const [inView, setInView] = useState(true);

  useEffect(() => {
    const mq = window.matchMedia("(max-width: 767px)");
    const update = () => setIsMobile(mq.matches);
    update();
    mq.addEventListener("change", update);
    return () => mq.removeEventListener("change", update);
  }, []);

  // Pause animations when off-screen
  useEffect(() => {
    const el = sectionRef.current;
    if (!el) return;
    const obs = new IntersectionObserver(
      (entries) => {
        for (const e of entries) setInView(e.isIntersecting);
      },
      { threshold: 0.05 },
    );
    obs.observe(el);
    return () => obs.disconnect();
  }, []);

  const setFromClientX = useCallback((clientX: number) => {
    const el = wrapRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    const p = ((clientX - rect.left) / rect.width) * 100;
    setPos(Math.max(2, Math.min(98, p)));
  }, []);

  useEffect(() => {
    if (!dragging) return;
    const onMove = (e: PointerEvent) => setFromClientX(e.clientX);
    const onUp = () => setDragging(false);
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
    return () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
    };
  }, [dragging, setFromClientX]);

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "ArrowLeft") setPos((p) => Math.max(2, p - 4));
    if (e.key === "ArrowRight") setPos((p) => Math.min(98, p + 4));
  };

  // Mobile = stacked tall composition; desktop = side-by-side w/ drag handle.
  const containerHeight = isMobile ? "min(78vh, 700px)" : "560px";

  return (
    <section
      ref={sectionRef}
      id="before-after"
      data-snap
      className="relative px-6 py-20 md:px-10 md:py-28"
    >
      <div className="container-wide">
        <Reveal>
          <div className="mb-10 grid grid-cols-1 gap-6 md:mb-12 md:grid-cols-12 md:items-end md:gap-12">
            <div className="md:col-span-3">
              <div className="label">[ ✦ — Before / After ]</div>
            </div>
            <div className="md:col-span-9">
              <h2
                className="display max-w-[26ch] text-ink"
                style={{ fontSize: "clamp(1.8rem, 5.2vw, 4rem)" }}
              >
                The spreadsheet you have.
                <span className="text-mute"> The screen you could have.</span>
              </h2>
            </div>
          </div>
        </Reveal>

        {isMobile ? (
          <div>
            <div className="mb-3 inline-flex border border-hairline/20">
              {(["before", "after"] as const).map((t) => (
                <button
                  key={t}
                  type="button"
                  onClick={() => setMobileTab(t)}
                  className={`px-4 py-2 font-mono text-[11px] uppercase tracking-[0.22em] ${
                    mobileTab === t ? "bg-ink text-paper" : "text-mute"
                  }`}
                >
                  {t}
                </button>
              ))}
            </div>
            <div
              className="relative border border-hairline/15"
              style={{ height: containerHeight }}
            >
              {mobileTab === "before" ? (
                <Spreadsheet playing={inView} />
              ) : (
                <KozaiConsole playing={inView} />
              )}
            </div>
          </div>
        ) : (
          <div
            ref={wrapRef}
            className="relative select-none overflow-hidden border border-hairline/15"
            style={{ height: containerHeight }}
            onPointerDown={(e) => {
              setDragging(true);
              setFromClientX(e.clientX);
            }}
          >
            {/* After (right, full background) */}
            <div className="absolute inset-0">
              <KozaiConsole playing={inView} />
            </div>
            {/* Before (clipped from left) — z-10 so it sits above the
                After-side fake cursor in the compare view. */}
            <div
              className="absolute inset-0 z-10"
              style={{ clipPath: `inset(0 ${100 - pos}% 0 0)` }}
            >
              <Spreadsheet playing={inView} />
            </div>

            {/* Handle */}
            <div
              role="slider"
              tabIndex={0}
              aria-label="Before / after compare slider"
              aria-valuemin={0}
              aria-valuemax={100}
              aria-valuenow={Math.round(pos)}
              onKeyDown={onKeyDown}
              className="absolute top-0 z-30 h-full w-px bg-signal"
              style={{ left: `calc(${pos}% - 0.5px)`, cursor: "ew-resize" }}
            >
              <span
                className="absolute left-1/2 top-1/2 flex h-10 w-10 -translate-x-1/2 -translate-y-1/2 items-center justify-center border border-signal bg-paper font-mono text-[12px] text-signal"
                aria-hidden
              >
                ⇆
              </span>
              <span className="absolute right-3 top-3 -translate-x-full whitespace-nowrap font-mono text-[10px] uppercase tracking-[0.22em] text-signal">
                before
              </span>
              <span className="absolute left-3 top-3 whitespace-nowrap font-mono text-[10px] uppercase tracking-[0.22em] text-signal">
                after
              </span>
            </div>
          </div>
        )}

        <div className="mt-4 flex items-center justify-between font-mono text-[10px] uppercase tracking-[0.22em] text-mute">
          <span>[ ✦ — Before / After ] · {isMobile ? "Toggle to compare." : "Drag to compare."}</span>
          <span className="hidden md:inline">fig.05 — same data, two surfaces</span>
        </div>
      </div>
    </section>
  );
};

export default BeforeAfter;
