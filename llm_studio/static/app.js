/* LocalLLM Studio 프론트엔드 — 외부 라이브러리 없이 동작 */
"use strict";

const $ = (id) => document.getElementById(id);
const messagesEl = $("messages");
const inputEl = $("input");
const sendBtn = $("sendBtn");
const taskModeBtn = $("taskModeBtn");

let currentConvId = null;
let currentProjectId = null;   // 활성 프로젝트 (null = "기본" 공간). 대화·메모리·프롬프트 격리 기준.
let projectNames = {};   // { id: name } — 새 대화 화면의 활성 공간 배지에 쓴다(loadProjects가 갱신).
let attachments = [];   // {id, name, chars}
let sending = false;
let abortController = null;

/* 활성 프로젝트를 대화·메모리 API 요청에 실어보내는 쿼리 조각. 기본 공간이면 빈 문자열. */
function projQuery(prefix = "?") {
  return currentProjectId ? `${prefix}project_id=${encodeURIComponent(currentProjectId)}` : "";
}
let lastStatus = null;  // 마지막 /api/status 응답 (셋업 바·서빙 버튼 판단용)

// 처리 모드 3단: auto(라우터가 자동 분류) → task(작업 강제) → chat(채팅 강제) → auto…
// 클릭으로 순환한다. auto가 기본 — 요청을 보고 계획-실행/단일응답을 알아서 고른다.
let chatMode = "auto";
const MODE_UI = {
  auto: { icon: "🧭", title: "라우팅: 자동 — 요청을 보고 작업/채팅을 자동 판단 (클릭: 작업 강제)" },
  task: { icon: "🧭", title: "라우팅: 작업 강제 — 항상 계획을 세워 단계별 실행 (클릭: 채팅 강제)" },
  chat: { icon: "💬", title: "라우팅: 채팅 강제 — 항상 단일 응답 (클릭: 자동으로)" },
};
function applyModeBtn() {
  const u = MODE_UI[chatMode];
  taskModeBtn.textContent = u.icon;
  taskModeBtn.title = u.title;
  taskModeBtn.classList.toggle("active", chatMode === "task");
  taskModeBtn.classList.toggle("mode-chat", chatMode === "chat");
}
taskModeBtn.addEventListener("click", () => {
  chatMode = chatMode === "auto" ? "task" : chatMode === "task" ? "chat" : "auto";
  applyModeBtn();
});
applyModeBtn();

// 코드블록 복사 — 버튼은 md() 가 매 렌더마다 새로 그리므로(스트리밍) 개별 리스너
// 대신 messagesEl 한 곳에 위임한다. 코드 원문은 형제 <code>의 textContent(escape 해제)로 읽는다.
messagesEl.addEventListener("click", async (e) => {
  const btn = e.target.closest(".copy-btn");
  if (!btn) return;
  const code = btn.parentElement.querySelector("code");
  if (!code) return;
  try {
    await navigator.clipboard.writeText(code.textContent);
    btn.textContent = "복사됨";
    btn.classList.add("copied");
  } catch {
    btn.textContent = "복사 실패";
  }
  setTimeout(() => { btn.textContent = "복사"; btn.classList.remove("copied"); }, 1500);
});

/* 외부 LLM 프로바이더 — 선택값/키는 전부 서버(config.json)에 저장된다.
   회사 보안 정책상 브라우저 저장소(localStorage 등)는 일절 쓰지 않는다. */
let providersSig = "";        // 헤더 드롭다운 재생성 판단용
let editingProviders = [];    // 설정 모달에서 편집 중인 목록

const PROVIDER_PRESETS = {
  custom:    { name: "", base_url: "", api_key: "", model: "" },
  openai:    { name: "OpenAI", base_url: "https://api.openai.com/v1", api_key: "", model: "gpt-4o-mini" },
  anthropic: { name: "Claude", base_url: "https://api.anthropic.com/v1", api_key: "", model: "claude-sonnet-5" },
  gemini:    { name: "Gemini", base_url: "https://generativelanguage.googleapis.com/v1beta/openai", api_key: "", model: "gemini-2.5-flash" },
};

/* MCP 서버 — 폼으로 편집한다(JSON 직접 입력 대신). 저장 시 mcpServers 규격으로
   직렬화해 /api/mcp/config에 넘긴다. 모델이 못 다루는 필드는 _extra에 보존한다. */
let editingMcpServers = [];   // {name, transport:"url"|"stdio", url, command, args, env, disabled, _extra}

const MCP_PRESETS = {
  url:   { name: "", transport: "url",   url: "", command: "", args: "", env: "", disabled: false },
  stdio: { name: "", transport: "stdio", url: "", command: "", args: "", env: "", disabled: false },
};

/* ==================== 마크다운 렌더러 (최소 구현) ==================== */

function escapeHtml(s) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
          .replace(/"/g, "&quot;");
}

// 목록 항목·문단·표 셀 등의 인라인 마크업. src는 이미 escapeHtml 된 상태로 들어온다.
function mdInline(s) {
  // 백틱 인라인 코드 구간과 그 외를 나눠, 코드 내부에는 굵게/이탤릭/취소선을
  // 적용하지 않는다 (split 으로 격리 — 자리표시자가 없어 텍스트와 충돌하지 않는다).
  return s.split(/(`[^`]+`)/).map((seg) => {
    if (seg.length > 1 && seg[0] === "`" && seg[seg.length - 1] === "`")
      return `<code>${seg.slice(1, -1)}</code>`;
    return seg
      // 이미지 ![alt](url) — 링크보다 먼저 처리해야 앞의 '!'가 남지 않는다.
      // 폐쇄망에선 외부 이미지는 안 뜨고 alt가 대신 보인다(정상). data:image 도 허용.
      .replace(/!\[([^\]]*)\]\((https?:\/\/[^)\s]+|data:image\/[^)\s]+)\)/g,
        '<img src="$2" alt="$1" class="md-img">')
      .replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g,                 // 링크
        '<a href="$2" target="_blank" rel="noopener">$1</a>')
      .replace(/\*\*\*([^*]+)\*\*\*/g, "<strong><em>$1</em></strong>")  // 굵은 이탤릭
      .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")               // 굵게
      .replace(/(^|[^*])\*([^*\s][^*]*)\*/g, "$1<em>$2</em>")           // 이탤릭
      .replace(/~~([^~]+)~~/g, "<del>$1</del>")                         // 취소선
      // 맨 URL 자동 링크 — 위에서 만든 <a>/<img> 속성값(따옴표·괄호 뒤)은 건드리지
      // 않도록 앞 문자를 제한한다.
      .replace(/(^|[^"(>=/])(https?:\/\/[^\s<)]+)/g,
        '$1<a href="$2" target="_blank" rel="noopener">$2</a>');
  }).join("");
}

const LIST_RE = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/;

// 들여쓰기 깊이로 중첩 목록을 재귀 구성한다. items[idx]부터 base 이상 깊이의
// 형제 항목을 한 <ul>/<ol>로 묶고, 더 깊은 항목은 직전 <li> 안으로 내려보낸다.
function renderListItems(items, idx, base) {
  const first = items[idx];
  const tag = first.ordered ? "ol" : "ul";
  const startAttr = first.ordered && first.start !== 1 ? ` start="${first.start}"` : "";
  let html = `<${tag}${startAttr}>`;
  while (idx < items.length && items[idx].indent >= base) {
    if (items[idx].indent > base) { idx++; continue; }   // 방어적: 정상 흐름엔 안 옴
    const it = items[idx];
    let body = it.checked === null
      ? mdInline(it.content)
      : `<label class="task"><input type="checkbox" disabled${it.checked ? " checked" : ""}> ${mdInline(it.content)}</label>`;
    idx++;
    if (idx < items.length && items[idx].indent > it.indent) {   // 하위 목록
      const sub = renderListItems(items, idx, items[idx].indent);
      body += sub.html;
      idx = sub.idx;
    }
    html += `<li${it.checked === null ? "" : ' class="task-item"'}>${body}</li>`;
  }
  return { html: html + `</${tag}>`, idx };
}

function md(src) {
  // ── 수식 보호 ──────────────────────────────────────────────
  // $..$ / $$..$$ / \(..\) / \[..\] 안의 *, _, \ 를 마크다운 인라인 처리가
  // 건드리지 않도록 먼저 placeholder(@@MATHn@@)로 빼둔다. KaTeX 렌더는
  // 여기서 하지 않고(스트리밍 매 토큰 호출 회피) 완성된 메시지에서 enhanceContent()가
  // 한 번에 한다. 코드펜스/인라인코드 안의 $ (셸 변수 등)는 수식으로 보지 않는다.
  const math = [];
  const stash = (tex, display) => `@@MATH${math.push({ tex, display }) - 1}@@`;
  const protectMath = (t) => t
    .replace(/\$\$([\s\S]+?)\$\$/g, (_, x) => stash(x, true))                 // 블록 $$..$$
    .replace(/\\\[([\s\S]+?)\\\]/g, (_, x) => stash(x, true))                 // 블록 \[..\]
    .replace(/(^|[^\\$])\$(?!\s)([^\n$]+?)(?<!\s)\$/g, (_, p, x) => p + stash(x, false))  // 인라인 $..$
    .replace(/\\\(([\s\S]+?)\\\)/g, (_, x) => stash(x, false));              // 인라인 \(..\)
  // 코드(펜스·인라인)를 홀수 인덱스로 분리해 그 안에서는 수식 추출을 건너뛴다.
  src = src.split(/(```[\s\S]*?```|~~~[\s\S]*?~~~|`[^`\n]*`)/g)
           .map((seg, idx) => (idx % 2 ? seg : protectMath(seg))).join("");

  const lines = escapeHtml(src).split("\n");
  const out = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (line.startsWith("```") || line.startsWith("~~~")) {  // 코드 블록
      const fence = line.slice(0, 3);
      const lang = line.slice(3).trim().replace(/[^\w+.#-]/g, "");
      const buf = [];
      i++;
      while (i < lines.length && !lines[i].startsWith(fence)) buf.push(lines[i++]);
      i++;                                                // 닫는 펜스 소비
      const attr = lang ? ` data-lang="${lang}" class="lang-${lang}"` : "";
      // language-<lang> 은 highlight.js 가 문법을 고르는 클래스. 없으면 자동 감지.
      const codeClass = lang ? ` class="language-${lang}"` : "";
      // 복사 버튼은 절대 위치라 흐름 밖(코드 텍스트에 안 섞임). 클릭은 messagesEl 위임.
      out.push(`<pre${attr}><button class="copy-btn" type="button" aria-label="코드 복사">복사</button><code${codeClass}>${buf.join("\n")}</code></pre>`);
      continue;
    } else if (/^#{1,6}\s/.test(line)) {                // 제목 (h1~h6)
      const level = line.match(/^#+/)[0].length;
      out.push(`<h${level}>${mdInline(line.replace(/^#+\s*/, ""))}</h${level}>`);
      i++;  continue;
    } else if (/^\s*(\*{3,}|-{3,}|_{3,})\s*$/.test(line)) {  // 구분선
      out.push("<hr>"); i++;  continue;
    } else if (LIST_RE.test(line)) {                    // 목록 (중첩·체크박스 지원)
      const items = [];
      while (i < lines.length && LIST_RE.test(lines[i])) {
        const m = lines[i].match(LIST_RE);
        const ordered = /\d/.test(m[2]);
        let content = m[3], checked = null;
        const tm = content.match(/^\[([ xX])\]\s+(.*)$/);   // - [ ] / - [x]
        if (tm) { checked = /[xX]/.test(tm[1]); content = tm[2]; }
        items.push({
          indent: m[1].replace(/\t/g, "    ").length,
          ordered,
          start: ordered ? parseInt(m[2], 10) : 1,
          content, checked,
        });
        i++;
      }
      out.push(renderListItems(items, 0, items[0].indent).html);
      continue;
    } else if (line.startsWith("&gt;")) {               // 인용
      const buf = [];
      while (i < lines.length && lines[i].startsWith("&gt;")) {
        buf.push(mdInline(lines[i].replace(/^&gt;\s?/, "")));
        i++;
      }
      out.push(`<blockquote>${buf.join("<br>")}</blockquote>`);
      continue;
    } else if (line.includes("|") && i + 1 < lines.length
               && /^\s*\|?[\s:|-]+\|[\s:|-]*$/.test(lines[i + 1])) {  // 표
      const cells = (row) => row.split("|").map(c => c.trim()).filter((c, idx, a) =>
        !(c === "" && (idx === 0 || idx === a.length - 1)));
      // 구분줄(:---:)에서 열 정렬을 읽어 셀에 style 로 적용한다.
      const aligns = cells(lines[i + 1]).map((sep) => {
        const l = sep.startsWith(":"), r = sep.endsWith(":");
        return l && r ? "center" : r ? "right" : l ? "left" : "";
      });
      const al = (idx) => aligns[idx] ? ` style="text-align:${aligns[idx]}"` : "";
      const head = cells(line).map((c, idx) => `<th${al(idx)}>${mdInline(c)}</th>`).join("");
      i += 2;
      const rows = [];
      while (i < lines.length && lines[i].includes("|")) {
        rows.push(`<tr>${cells(lines[i]).map((c, idx) => `<td${al(idx)}>${mdInline(c)}</td>`).join("")}</tr>`);
        i++;
      }
      out.push(`<table><thead><tr>${head}</tr></thead><tbody>${rows.join("")}</tbody></table>`);
      continue;
    } else if (line.trim() === "") {
      i++;  continue;
    } else {                                            // 문단
      const buf = [line];
      i++;
      while (i < lines.length && lines[i].trim() !== ""
             && !/^(#{1,6}\s|```|~~~|\s*([-*+]|\d+[.)])\s|&gt;)/.test(lines[i])) {
        buf.push(lines[i]); i++;
      }
      out.push(`<p>${buf.map(mdInline).join("<br>")}</p>`);
      continue;
    }
  }
  let html = out.join("\n");
  // 빼뒀던 수식을 KaTeX 가 읽을 span 으로 복원한다. tex 는 escapeHtml 로 넣어
  // 두고, enhanceContent()가 el.textContent(escape 해제된 원문)로 렌더한다.
  if (math.length)
    html = html.replace(/@@MATH(\d+)@@/g, (_, n) => {
      const m = math[+n];
      return `<span class="katex-src" data-display="${m.display ? 1 : 0}">${escapeHtml(m.tex)}</span>`;
    });
  return html;
}

/* 완성된(스트리밍 종료·이력 로드) 메시지에만 코드 하이라이트·수식 렌더를 한 번 적용한다.
   스트리밍 중 매 토큰 재렌더에는 돌리지 않는다(비용·깜빡임). vendor 라이브러리가
   로드 안 됐으면 조용히 건너뛴다(우아한 저하 — 코드/수식은 원문 그대로 보인다). */
function enhanceContent(root) {
  if (window.hljs) {
    root.querySelectorAll("pre code:not([data-highlighted])").forEach((el) => {
      try { window.hljs.highlightElement(el); } catch { /* 하이라이트 실패는 무시 */ }
    });
  }
  if (window.katex) {
    root.querySelectorAll(".katex-src:not([data-rendered])").forEach((el) => {
      const tex = el.textContent;
      el.dataset.rendered = "1";   // 재실행 시 렌더된 내용을 tex 로 오인하지 않게
      try {
        window.katex.render(tex, el, {
          displayMode: el.dataset.display === "1",
          throwOnError: false,
        });
      } catch { /* 잘못된 수식은 원문 유지 */ }
    });
  }
}

/* 카드(도구 결과·생각·단계) 본문에도 답변 말풍선과 같은 마크다운을 입힌다.
   스트리밍으로 토큰이 붙는 카드는 el._mdRaw 에 원문을 누적해 매 토큰 다시 렌더하고
   (답변 본문과 같은 방식 — line 902), 완료 시 enhanceContent 로 코드 하이라이트·
   수식을 한 번만 입힌다. md()/enhanceContent 는 라이브러리 미로드 시 평문으로 저하한다. */
function mdAppend(el, text) {
  el._mdRaw = (el._mdRaw || "") + text;
  el.innerHTML = md(el._mdRaw);
}

function mdSet(el, text) {
  el._mdRaw = text || "";
  el.innerHTML = md(el._mdRaw);
  enhanceContent(el);
}

/* ==================== 메시지 렌더링 ==================== */

function hideEmptyHint() {
  const hint = $("emptyHint");
  if (hint) hint.remove();
}

function addUserMessage(text) {
  hideEmptyHint();
  const div = document.createElement("div");
  div.className = "msg user";
  const bubble = document.createElement("div");
  bubble.className = "bubble";
  bubble.textContent = text;
  div.appendChild(bubble);
  messagesEl.appendChild(div);
  scrollBottom();
}

function addAssistantContainer() {
  hideEmptyHint();
  const div = document.createElement("div");
  div.className = "msg assistant";
  messagesEl.appendChild(div);
  return div;
}

/* 도구/승인 블록 공용: 인자 프리뷰 요소. 표시 개선(정리·마스킹 등)은 여기 한 곳만
   고치면 도구 블록과 승인 카드 양쪽에 함께 적용된다. */
function makeArgsPre(args) {
  const pre = document.createElement("pre");
  pre.textContent = `인자: ${args || "{}"}`;
  return pre;
}

/* 저장된 tool 메시지가 승인 거절/시간초과로 실행되지 않은 결과인지 판별.
   agent.py의 DENIED_RESULT/TIMEOUT_RESULT 접두사와 맞춰야 한다. */
function isDeniedToolResult(content) {
  const c = String(content || "");
  return c.startsWith("[사용자가 이 도구 실행을 거부했습니다") ||
         c.startsWith("[사용자 승인 대기가 시간 초과");
}

function addToolBlock(container, name, args) {
  const details = document.createElement("details");
  details.className = "tool";
  const summary = document.createElement("summary");
  summary.innerHTML = `🔧 <b>${escapeHtml(name)}</b> <span class="tool-spinner">실행 중</span>`;
  details.append(summary, makeArgsPre(args));
  container.appendChild(details);
  scrollBottom();
  return details;
}

function finishToolBlock(details, result, executed = true) {
  const spinner = details.querySelector(".tool-spinner");
  if (spinner) {
    spinner.classList.remove("tool-spinner");
    if (executed) {
      spinner.textContent = "완료";
    } else {
      // 거절/시간초과 — '완료'로 보이면 실행된 것으로 오인하므로 명확히 구분한다
      spinner.textContent = "🚫 실행 안 됨(거절/시간초과)";
      spinner.classList.add("tool-denied");
    }
  }
  // 도구 결과에도 마크다운을 입힌다 — office/excel 서버 등이 돌려주는 표·제목이
  // 답변 본문처럼 렌더된다. (승인 카드의 인자 프리뷰는 JSON이라 그대로 <pre> 유지.)
  const body = document.createElement("div");
  body.className = "content tool-result";
  mdSet(body, result);
  details.appendChild(body);
}

/* 위험 도구 승인 카드 — 모델이 confirm 도구를 실행하려 할 때 서버가 멈추고
   approval_request를 보낸다. 승인/거절을 POST /api/chat/approve 로 답하면
   서버가 approval_result 이벤트로 결정을 확정하고 카드를 마감한다. */
function addApprovalBlock(container, id, name, args) {
  const div = document.createElement("div");
  div.className = "approval";
  const head = document.createElement("div");
  head.className = "approval-head";
  head.innerHTML = `🔐 <b>승인 필요</b> — 모델이 <code>${escapeHtml(name)}</code> 실행 허가를 요청합니다`;
  const pre = makeArgsPre(args);
  const row = document.createElement("div");
  row.className = "approval-btns";
  const ok = document.createElement("button");
  ok.className = "btn primary";
  ok.textContent = "승인하고 실행";
  const no = document.createElement("button");
  no.className = "btn danger";
  no.textContent = "거절";
  const decide = async (approved) => {
    ok.disabled = no.disabled = true;
    try {
      const res = await fetch("/api/chat/approve", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id, approved }),
      });
      if (res.status === 404) {
        // 이미 처리됐거나 만료된 요청 — 버튼을 되살리지 않고 카드를 종결한다
        // (서버의 approval_result가 나중에 오면 그 결정으로 표시가 대체된다)
        expireApprovalBlock(div);
        return;
      }
      if (res.status === 403) {
        // 원격 접속 — 승인은 서버 PC(127.0.0.1)에서만 가능하다
        expireApprovalBlock(div, "⚠ 승인은 서버 PC(127.0.0.1)에서만 할 수 있습니다");
        return;
      }
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
    } catch (e) {
      // 네트워크 등 일시 오류 — 다시 시도할 수 있게 버튼만 되살린다
      ok.disabled = no.disabled = false;
    }
  };
  ok.onclick = () => decide(true);
  no.onclick = () => decide(false);
  row.append(ok, no);
  div.append(head, pre, row);
  container.appendChild(div);
  scrollBottom();
  return div;
}

function finishApprovalBlock(div, approved, timedOut) {
  if (!div) return;
  const row = div.querySelector(".approval-btns");
  if (row) row.remove();
  // 만료 표시가 먼저 붙었더라도 서버의 approval_result가 최종 결정이다 — 대체한다
  const old = div.querySelector(".approval-note");
  if (old) old.remove();
  const note = document.createElement("div");
  note.className = "approval-note";
  note.textContent = timedOut ? "⏱ 시간 초과 — 실행하지 않았습니다"
    : approved ? "✅ 승인됨 — 실행합니다" : "🚫 거절됨 — 실행하지 않았습니다";
  div.appendChild(note);
  div.classList.remove("approved", "denied");
  div.classList.add(approved ? "approved" : "denied");
  scrollBottom();
}

/* 결정이 불가능해진 승인 카드를 종결 상태로 마감한다 — 404(이미 처리/만료),
   스트림 중단 등. 살아있는 버튼이 남아 사용자를 홀리는 일이 없게 한다. */
function expireApprovalBlock(div, text) {
  if (!div) return;
  const row = div.querySelector(".approval-btns");
  if (row) row.remove();
  if (div.querySelector(".approval-note")) return;  // 서버 결정이 이미 표시됨
  const note = document.createElement("div");
  note.className = "approval-note";
  note.textContent = text || "⚠ 만료된 요청 — 이미 처리됐거나 시간이 초과되었습니다";
  div.appendChild(note);
  div.classList.add("denied");
}

/* 모델 생각(추론) 블록 — 추론형 모델이 reasoning_content를 흘릴 때만 나타난다.
   답변 위(맨 앞)에 접을 수 있는 형태로 붙이고, 답변이 시작되면 자동으로 접는다. */
function addReasoningBlock(container) {
  const details = document.createElement("details");
  details.className = "reasoning";
  details.open = true;
  const summary = document.createElement("summary");
  summary.innerHTML = '💭 <b>생각 과정</b> <span class="tool-spinner">진행 중</span>';
  const body = document.createElement("div");
  body.className = "reasoning-body content";  // content: 마크다운 요소 스타일 공유
  details.append(summary, body);
  container.insertBefore(details, container.firstChild);  // 답변보다 위에
  scrollBottom();
  return details;
}

function appendReasoning(details, text) {
  mdAppend(details.querySelector(".reasoning-body"), text);
  scrollBottom();
}

function finishReasoning(details) {
  const sp = details.querySelector(".tool-spinner");
  if (sp) { sp.classList.remove("tool-spinner"); sp.textContent = "완료"; }
  enhanceContent(details.querySelector(".reasoning-body"));  // 코드·수식 마감 렌더
  details.open = false;   // 끝난 생각은 접어 답변에 집중
}

function addErrorNote(container, message) {
  const note = document.createElement("div");
  note.className = "error-note";
  note.textContent = `⚠ ${message}`;
  container.appendChild(note);
}

/* 계획-실행(작업 모드) 렌더 */
function addPlanNote(container, text) {
  const div = document.createElement("div");
  div.className = "plan-note";
  div.textContent = text;
  container.appendChild(div);
  scrollBottom();
}

function addStepBlock(container, index, text) {
  const details = document.createElement("details");
  details.className = "tool step";
  details.open = true;
  const summary = document.createElement("summary");
  summary.innerHTML = `▶ <b>단계 ${index + 1}</b> ${escapeHtml(text)} <span class="tool-spinner">진행 중</span>`;
  const body = document.createElement("div");
  body.className = "step-body content";  // content: 마크다운 요소 스타일 공유
  details.append(summary, body);
  container.appendChild(details);
  scrollBottom();
  return details;
}

function appendStepToken(details, text) {
  if (!details) return;
  mdAppend(details.querySelector(".step-body"), text);
  scrollBottom();
}

function finishStepBlock(details, ok, result) {
  if (!details) return;
  const spinner = details.querySelector(".tool-spinner");
  if (spinner) {
    spinner.classList.remove("tool-spinner");
    spinner.textContent = ok ? "완료" : "실패";
    spinner.classList.add("tool-spinner", ok ? "ok" : "fail");
  }
  enhanceContent(details.querySelector(".step-body"));  // 코드·수식 마감 렌더
  details.open = false;   // 끝난 단계는 접어 답변에 집중
}

/* 의도 판단(라우터) 노드 — 코크핏의 0번 노드. 이번 요청을 작업/채팅 중 무엇으로
   처리하는지와 그 이유를 보여준다. 오분류를 사람이 알아채고 🧭로 강제 전환하게 한다. */
function addRouteNode(container, useTask, reason) {
  const div = document.createElement("div");
  div.className = "route-node " + (useTask ? "task" : "chat");
  div.innerHTML = `🧭 <b>의도 판단</b> — ${useTask ? "작업(계획-실행)" : "채팅(단일 응답)"}` +
    (reason ? ` <span class="route-reason">${escapeHtml(reason)}</span>` : "");
  container.appendChild(div);
  scrollBottom();
}

/* 조건 분기 노드 (Layer 3) — 조건을 규칙/LLM으로 판정해 이후 단계를 건너뛴다.
   branch_start에서 '판정 중' 노드를 붙이고, branch가 오면 참/거짓과 건너뛴 단계를 채운다. */
function addBranchNode(container, index, cond, mode) {
  const div = document.createElement("div");
  div.className = "branch-node pending";
  div.innerHTML = `🔀 <b>분기 ${index + 1}</b> ${escapeHtml(cond || "")} ` +
    `<span class="branch-mode">${mode === "rule" ? "규칙" : "LLM"} 판정 중…</span>`;
  container.appendChild(div);
  scrollBottom();
  return div;
}

function finishBranchNode(div, ev) {
  if (!div) return;
  div.classList.remove("pending");
  div.classList.add(ev.result ? "yes" : "no");
  const skipped = (ev.skipped || []);
  const label = ev.result
    ? "참 → 계속"
    : (skipped.length ? `거짓 → 단계 ${skipped.join(", ")} 건너뜀` : "거짓 → 건너뛸 단계 없음");
  const modeTxt = ev.mode === "rule" ? "규칙" : "LLM";
  const note = ev.note ? ` <span class="branch-mode">(${escapeHtml(ev.note)})</span>` : "";
  div.innerHTML = `🔀 <b>분기</b> ${escapeHtml(ev.cond || "")} ` +
    `<span class="branch-result">${label}</span> ` +
    `<span class="branch-mode">(${modeTxt} 판정)</span>${note}`;
  scrollBottom();
}

function expireBranchNode(div) {
  if (!div || !div.classList.contains("pending")) return;
  div.classList.remove("pending");
  div.querySelector(".branch-mode").textContent = "판정 중단됨";
}

/* ==================== 실행 코크핏 (Layer 4) ====================
   작업 모드의 계획·단계·분기를 흩어진 카드가 아니라 하나의 세로 플로차트로 묶어
   보여준다. 각 노드는 살아있는 상태(대기/실행중/완료/실패/거절/건너뜀)를 색으로
   반영하고, 의존 태그([←N])는 배지와 hover 강조로 데이터 흐름을 드러낸다.
   무거운 그래프 라이브러리 없이 순수 CSS/JS로 그린다(선형+분기만 다루므로 충분). */

const NODE_STATUS_LABEL = {
  pending: "대기", running: "실행중", done: "완료", fail: "실패",
  denied: "거절", skip: "건너뜀", yes: "참", no: "거짓", deciding: "판정중",
};

// _step_display가 붙인 앞머리 [태그]들을 본문에서 떼어낸다(구조는 배지로 따로 보여준다).
function stripStepTags(text) {
  return (text || "").replace(/^(?:\s*\[[^\]]*\]\s*)+/, "").trim() || (text || "");
}

function setNodeStatus(node, status, label) {
  if (!node) return;
  const pill = node.querySelector(".pill");
  if (pill) {
    pill.dataset.status = status;
    pill.textContent = label || NODE_STATUS_LABEL[status] || status;
  }
  node.dataset.status = status;
}

function cockpitNodeLog(node, { open = false } = {}) {
  if (!node) return null;
  const log = node.querySelector(".node-log");
  if (log && open) log.classList.remove("hidden");
  return log;
}

// hover 시 이 노드가 참조하는(←) / 이 노드를 참조하는(→) 단계를 함께 강조한다.
function highlightLinks(flow, node, on) {
  const rel = [];
  (node.dataset.deps || "").split(",").filter(Boolean).forEach(n => rel.push(+n - 1));
  (node.dataset.consumers || "").split(",").filter(Boolean).forEach(n => rel.push(+n - 1));
  rel.forEach(i => {
    const t = flow.querySelector(`.node[data-index="${i}"]`);
    if (t) t.classList.toggle("linked", on);
  });
  node.classList.toggle("linked-src", on);
}

/* 계획(및 재계획/편집)에 맞춰 코크핏 플로차트를 세운다. prevStatus를 주면 이미
   끝난 앞 단계의 상태를 복원한다(재계획은 tail만 바뀌므로 앞 단계는 그대로 유지). */
function buildCockpit(container, steps, meta, prevStatus) {
  const old = container.querySelector(".cockpit");
  // 재계획/편집으로 코크핏을 다시 그려도 사람이 접어둔 상태는 유지한다.
  const wasCollapsed = !!(old && old.classList.contains("collapsed"));
  if (old) old.remove();
  const panel = document.createElement("div");
  panel.className = "cockpit";
  const title = document.createElement("div");
  title.className = "cockpit-title";
  title.innerHTML = "<span class=\"caret\">▾</span>🧭 <b>실행 코크핏</b> " +
    "<span class=\"hint\">제목 클릭 시 접기 · 노드 클릭 시 결과</span>";
  title.title = "클릭하면 접거나 펼칩니다";
  title.onclick = () => panel.classList.toggle("collapsed");
  const flow = document.createElement("div");
  flow.className = "flow";
  const nodes = [];
  steps.forEach((text, i) => {
    const m = (meta && meta[i]) || {};
    const isBranch = !!m.branch;
    const node = document.createElement("div");
    node.className = "node" + (isBranch ? " is-branch" : "");
    node.dataset.index = i;
    if (m.deps) node.dataset.deps = m.deps.join(",");
    if (m.consumers) node.dataset.consumers = m.consumers.join(",");
    const badges = [];
    if (m.deps) badges.push(`<span class="dep in" title="이 단계가 참조하는 단계">← ${m.deps.join(",")}</span>`);
    if (m.consumers) badges.push(`<span class="dep out" title="이 단계 결과를 쓰는 단계">→ ${m.consumers.join(",")}</span>`);
    if (m.scope && m.scope.length) badges.push(`<span class="scope">${escapeHtml(m.scope.join(", "))}</span>`);
    if (isBranch) {
      const mode = m.branch.mode === "rule" ? "규칙" : "LLM";
      badges.push(`<span class="scope">?→${m.branch.target} · ${mode}</span>`);
    }
    const num = isBranch ? "🔀" : String(i + 1);
    node.innerHTML =
      `<div class="node-head">` +
        `<span class="pill" data-status="pending">대기</span>` +
        `<span class="node-num">${num}</span>` +
        `<span class="node-body">${escapeHtml(stripStepTags(text))}</span>` +
        `<span class="node-badges">${badges.join("")}</span>` +
      `</div>` +
      `<pre class="node-log hidden"></pre>`;
    node.querySelector(".node-head").onclick = () => {
      const log = node.querySelector(".node-log");
      if (log && log.textContent.trim()) log.classList.toggle("hidden");
    };
    node.addEventListener("mouseenter", () => highlightLinks(flow, node, true));
    node.addEventListener("mouseleave", () => highlightLinks(flow, node, false));
    flow.append(node);
    nodes.push(node);
  });
  panel.append(title, flow);
  if (wasCollapsed) panel.classList.add("collapsed");
  container.appendChild(panel);
  // 이미 끝난 앞 단계의 상태를 복원한다. 단, 재계획은 실패 지점부터 tail을 갈아끼우므로
  // '본문이 그대로인 노드'만 복원한다 — 자리(index)가 같아도 내용이 바뀐 새 단계에
  // 옛 완료/실패 상태가 잘못 묻어나지 않게 한다.
  const TRANSIENT = new Set(["pending", "running", "deciding"]);
  if (prevStatus) {
    prevStatus.forEach((s, i) => {
      if (s && nodes[i] && !TRANSIENT.has(s.status) && s.body === stripStepTags(steps[i])) {
        setNodeStatus(nodes[i], s.status, s.label);
        // 완료 단계의 결과 로그도 되살린다(접힌 채) — 재계획 후 클릭 시 결과가 비지 않게.
        if (s.log) {
          const log = nodes[i].querySelector(".node-log");
          if (log) log.textContent = s.log;
        }
      }
    });
  }
  scrollBottom();
  return nodes;
}

// 재계획/편집 전, 현재 노드 상태+본문을 갈무리해 복원 판단에 쓴다.
function snapshotCockpit(nodes) {
  if (!nodes) return null;
  return nodes.map(n => {
    const pill = n.querySelector(".pill");
    const body = n.querySelector(".node-body");
    const log = n.querySelector(".node-log");
    return pill ? { status: pill.dataset.status, label: pill.textContent,
                    body: body ? body.textContent : "",
                    log: log ? log.textContent : "" } : null;
  });
}

/* 조종 게이트 카드 (스텝 실패) — 승인 카드와 같은 방식으로 서버가 멈추고
   steer_request를 보내면, 결정을 POST /api/chat/steer 로 답한다. 서버는 그 결정에 따라
   흐름을 이어간다(step_start 이벤트가 뒤따른다). onResolved는 대기 상태 해제용. */
const STEER_LABELS = {
  edit: "✏ 편집한 단계로 재시도합니다",
  abort: "🚫 작업을 중단했습니다", retry: "↻ 이 단계를 재시도합니다",
  skip: "⏭ 이 단계를 건너뜁니다", replan: "🧭 계획을 다시 세웁니다",
};

function mkBtn(label, cls, onclick) {
  const b = document.createElement("button");
  b.className = "btn" + (cls ? " " + cls : "");
  b.textContent = label;
  b.onclick = onclick;
  return b;
}

function addSteerBlock(container, ev, onResolved) {
  const div = document.createElement("div");
  div.className = "steer";
  const decide = async (payload) => {
    div.querySelectorAll("button, textarea, input").forEach(b => b.disabled = true);
    try {
      const res = await fetch("/api/chat/steer", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ id: ev.id, ...payload }),
      });
      if (res.status === 404) { expireSteerBlock(div); return; }
      if (res.status === 403) {
        expireSteerBlock(div, "⚠ 조종은 서버 PC(127.0.0.1)에서만 할 수 있습니다");
        return;
      }
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      finishSteerBlock(div, payload.action);
      if (onResolved) onResolved();
    } catch {
      // 일시 오류 — 다시 시도할 수 있게 컨트롤을 되살린다
      div.querySelectorAll("button, textarea, input").forEach(b => b.disabled = false);
    }
  };
  buildFailGate(div, ev, decide);
  container.appendChild(div);
  scrollBottom();
  return div;
}

// 스텝 실패 게이트: 실패한 단계와 결과를 보여주고 재시도/건너뛰기/재계획/편집/중단.
function buildFailGate(div, ev, decide) {
  const head = document.createElement("div");
  head.className = "steer-head";
  head.innerHTML = `⚠ <b>단계 ${(ev.index ?? 0) + 1} 실패</b> — 어떻게 할까요?`;
  const info = document.createElement("pre");
  info.className = "steer-info";
  info.textContent = `단계: ${ev.step || ""}\n결과: ${ev.result || ""}`;
  const editRow = document.createElement("div");
  editRow.className = "steer-edit hidden";
  const inp = document.createElement("input");
  inp.type = "text";
  inp.value = ev.step || "";
  editRow.append(inp, mkBtn("편집 후 재시도", "primary",
    () => decide({ action: "edit", step: inp.value })));
  const row = document.createElement("div");
  row.className = "steer-btns";
  row.append(
    mkBtn("재시도", "primary", () => decide({ action: "retry" })),
    mkBtn("건너뛰기", "", () => decide({ action: "skip" })),
    mkBtn("재계획", "", () => decide({ action: "replan" })),
    mkBtn("편집", "", () => editRow.classList.toggle("hidden")),
    mkBtn("중단", "danger", () => decide({ action: "abort" })),
  );
  div.append(head, info, row, editRow);
}

function finishSteerBlock(div, action) {
  div.querySelectorAll(".steer-btns, .steer-edit").forEach(e => e.remove());
  const ta = div.querySelector(".steer-plan");
  if (ta) ta.disabled = true;
  const note = document.createElement("div");
  note.className = "steer-note";
  note.textContent = STEER_LABELS[action] || `결정: ${action}`;
  div.appendChild(note);
  div.classList.add("resolved");
  scrollBottom();
}

/* 결정이 불가능해진 조종 카드를 종결한다 — 404(이미 처리/만료)·403(원격)·스트림 중단. */
function expireSteerBlock(div, text) {
  if (!div) return;
  div.querySelectorAll(".steer-btns, .steer-edit").forEach(e => e.remove());
  const ta = div.querySelector(".steer-plan");
  if (ta) ta.disabled = true;
  if (div.querySelector(".steer-note")) return;
  const note = document.createElement("div");
  note.className = "steer-note";
  note.textContent = text || "⚠ 만료된 요청 — 이미 처리됐거나 시간이 초과되었습니다";
  div.appendChild(note);
  div.classList.add("resolved");
}

function scrollBottom() {
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

/* 저장된 대화 이력을 렌더링 */
function renderHistory(messages) {
  messagesEl.innerHTML = "";
  let assistantDiv = null;
  for (const m of messages) {
    if (m.role === "user") {
      addUserMessage(m.content);
      assistantDiv = null;
    } else if (m.role === "assistant") {
      if (!assistantDiv) assistantDiv = addAssistantContainer();
      if (m.tool_calls) {
        for (const tc of m.tool_calls) {
          const d = addToolBlock(assistantDiv, tc.function.name, tc.function.arguments);
          d.dataset.callId = tc.id;
        }
      }
      if (m.content) {
        const content = document.createElement("div");
        content.className = "content";
        content.innerHTML = md(m.content);
        assistantDiv.appendChild(content);
      }
    } else if (m.role === "tool") {
      const block = assistantDiv &&
        assistantDiv.querySelector(`details[data-call-id="${CSS.escape(m.tool_call_id || "")}"]`);
      // 승인 카드는 저장되지 않으므로, 다시 연 대화에서도 거절된 호출이
      // '완료'로 보이지 않게 저장된 거절 문구로 판별해 표시한다
      if (block) finishToolBlock(block, m.content, !isDeniedToolResult(m.content));
    }
  }
  enhanceContent(messagesEl);   // 불러온 이력 전체에 코드 하이라이트·수식 렌더
  scrollBottom();
}

/* ==================== 채팅 전송 (SSE) ==================== */

async function send() {
  const text = inputEl.value.trim();
  if (!text || sending) return;

  const attachNote = attachments.length
    ? "\n📎 " + attachments.map(a => a.name).join(", ") : "";
  addUserMessage(text + attachNote);

  const body = {
    message: text,
    conversation_id: currentConvId,
    project_id: currentProjectId,
    attachments: attachments.map(a => a.id),
    provider: $("providerSelect").value || null,
    mode: chatMode,
  };
  inputEl.value = "";
  autoResize();
  attachments = [];
  renderChips();
  setSending(true);

  const container = addAssistantContainer();
  let contentEl = document.createElement("div");
  contentEl.className = "content";
  container.appendChild(contentEl);
  let buffer = "";
  let currentTool = null;
  let currentApproval = null;  // 대기 중인 위험 도구 승인 카드
  let currentSteer = null;  // 대기 중인 조종 게이트 카드 (계획/실패)
  let currentStep = null;   // 현재 진행 중인 단계 블록 (작업 모드·코크핏 없을 때 폴백)
  let currentBranch = null; // 판정 중인 조건 분기 노드 (작업 모드·폴백)
  let cockpitNodes = null;  // 실행 코크핏 플로차트의 노드 엘리먼트 배열 (Layer 4)
  let nodeDenied = false;   // 현재 단계에서 위험 도구 승인이 거절됐는지 (거절≠실패 구분)
  let reasoningEl = null;   // 모델 생각(추론) 블록 (추론형 모델일 때만)
  // 계획/단계 블록이 붙은 뒤에는 다음 답변 토큰을 맨 아래 새 영역에서 시작한다.
  let answerFresh = false;

  const status = startGenStatus();   // 동작/멈춤을 알려주는 생성 상태 표시
  abortController = new AbortController();
  try {
    const res = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
      signal: abortController.signal,
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let pending = "";
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      pending += decoder.decode(value, { stream: true });
      const events = pending.split("\n\n");
      pending = events.pop();
      for (const raw of events) {
        const line = raw.split("\n").find(l => l.startsWith("data: "));
        if (!line) continue;
        const ev = JSON.parse(line.slice(6));
        // 무슨 이벤트든 도착 = 살아있다는 신호. 토큰류면 진행량으로 센다.
        status.activity(ev.type === "token" || ev.type === "step_token" || ev.type === "reasoning");
        if (ev.type === "meta") {
          currentConvId = ev.conversation_id;
          $("convTitle").textContent = ev.title;
        } else if (ev.type === "reasoning") {
          if (!reasoningEl) reasoningEl = addReasoningBlock(container);
          appendReasoning(reasoningEl, ev.text);
        } else if (ev.type === "token") {
          // 답변이 시작되면 생각 블록은 접는다 (생각 → 답변 전환)
          if (reasoningEl) { finishReasoning(reasoningEl); reasoningEl = null; }
          // 계획/단계 블록 뒤의 첫 답변 토큰이면 맨 아래에 새 답변 영역을 연다
          if (answerFresh) {
            if (!buffer) contentEl.remove();
            buffer = "";
            contentEl = document.createElement("div");
            contentEl.className = "content";
            container.appendChild(contentEl);
            answerFresh = false;
          }
          buffer += ev.text;
          contentEl.innerHTML = md(buffer);
          scrollBottom();
        } else if (ev.type === "route") {
          addRouteNode(container, ev.use_task, ev.reason);
          answerFresh = true;
        } else if (ev.type === "steer_request") {
          if (!buffer) contentEl.remove();
          currentSteer = addSteerBlock(container, ev, () => status.waiting(false));
          status.waiting(true);
          answerFresh = true;
        } else if (ev.type === "plan") {
          // 코크핏 재구성(재계획/편집이면 앞 단계 상태 복원). 계획 요약 카드도 함께 남긴다.
          cockpitNodes = buildCockpit(container, ev.steps || [], ev.meta,
            snapshotCockpit(cockpitNodes));
          if (ev.replan || ev.edited) {
            addPlanNote(container, ev.edited ? "🧭 계획을 편집했습니다"
              : `🧭 계획을 다시 세웠습니다 (#${ev.replan})`);
          }
          answerFresh = true;
        } else if (ev.type === "step_start") {
          const node = cockpitNodes && cockpitNodes[ev.index];
          if (node) {
            nodeDenied = false;
            setNodeStatus(node, "running");
            node.scrollIntoView({ block: "nearest" });
            cockpitNodeLog(node, { open: true }).textContent = "";
          } else {
            currentStep = addStepBlock(container, ev.index, ev.text);  // 폴백
          }
          answerFresh = true;
        } else if (ev.type === "step_token") {
          if (ev.index === -1) { addPlanNote(container, ev.text); }
          else if (cockpitNodes && cockpitNodes[ev.index]) {
            const log = cockpitNodeLog(cockpitNodes[ev.index], { open: true });
            log.textContent += ev.text;
            scrollBottom();
          } else { appendStepToken(currentStep, ev.text); }
        } else if (ev.type === "step_done") {
          const node = cockpitNodes && cockpitNodes[ev.index];
          if (node) {
            setNodeStatus(node, ev.ok ? "done" : (nodeDenied ? "denied" : "fail"));
            const log = cockpitNodeLog(node);
            if (log && ev.result && !log.textContent.trim()) log.textContent = ev.result;
            if (log) log.classList.add("hidden");  // 끝난 단계는 접어 흐름에 집중
            nodeDenied = false;
          } else { finishStepBlock(currentStep, ev.ok, ev.result); currentStep = null; }
          answerFresh = true;
        } else if (ev.type === "branch_start") {
          const node = cockpitNodes && cockpitNodes[ev.index];
          if (node) setNodeStatus(node, "deciding");
          else currentBranch = addBranchNode(container, ev.index, ev.cond, ev.mode);
          answerFresh = true;
        } else if (ev.type === "branch") {
          const node = cockpitNodes && cockpitNodes[ev.index];
          if (node) {
            setNodeStatus(node, ev.result ? "yes" : "no");
            const modeTxt = ev.mode === "rule" ? "규칙" : "LLM";
            const skipTxt = (ev.skipped && ev.skipped.length)
              ? ` → 단계 ${ev.skipped.join(", ")} 건너뜀` : "";
            const log = cockpitNodeLog(node);
            if (log) log.textContent = `${modeTxt} 판정: ${ev.result ? "참" : "거짓"}${skipTxt}`
              + (ev.note ? `\n(${ev.note})` : "");
            (ev.skipped || []).forEach(n => setNodeStatus(cockpitNodes[n - 1], "skip"));
          } else { finishBranchNode(currentBranch, ev); currentBranch = null; }
          answerFresh = true;
        } else if (ev.type === "approval_request") {
          if (!buffer) contentEl.remove();
          currentApproval = addApprovalBlock(container, ev.id, ev.name, ev.arguments);
          status.waiting(true);
        } else if (ev.type === "approval_result") {
          finishApprovalBlock(currentApproval, ev.approved, ev.timeout);
          currentApproval = null;
          status.waiting(false);
          // 거절이면 현재 단계는 '실패'가 아니라 '거절'로 표시한다(코크핏 노드 상태).
          if (!ev.approved) nodeDenied = true;
          // 승인 카드 다음 응답을 위해 새 content 영역을 연다
          buffer = "";
          contentEl = document.createElement("div");
          contentEl.className = "content";
          container.appendChild(contentEl);
        } else if (ev.type === "tool_call") {
          // 도구 호출 전 텍스트가 없으면 빈 content를 정리하고, 도구 블록을 붙인다
          if (!buffer) contentEl.remove();
          currentTool = addToolBlock(container, ev.name, ev.arguments);
        } else if (ev.type === "tool_result") {
          if (currentTool) finishToolBlock(currentTool, ev.result, ev.executed !== false);
          currentTool = null;
          // 도구 결과 이후의 응답은 새 content 요소에 이어서 렌더링한다
          buffer = "";
          contentEl = document.createElement("div");
          contentEl.className = "content";
          container.appendChild(contentEl);
        } else if (ev.type === "error") {
          addErrorNote(container, ev.message);
        }
      }
    }
  } catch (e) {
    if (e.name !== "AbortError") addErrorNote(container, e.message);
  } finally {
    status.stop();
    // 스트림이 approval_result 없이 끝났다면(중단·오류) 카드가 산 채로 남지 않게 마감
    if (currentApproval) {
      expireApprovalBlock(currentApproval, "⚠ 중단됨 — 이 요청은 더 이상 승인할 수 없습니다");
      currentApproval = null;
    }
    if (currentSteer) {
      expireSteerBlock(currentSteer, "⚠ 중단됨 — 이 요청은 더 이상 조종할 수 없습니다");
      currentSteer = null;
    }
    if (currentBranch) { expireBranchNode(currentBranch); currentBranch = null; }
    // 코크핏에 실행중/판정중 상태로 멈춘 노드가 있으면 '중단됨'으로 마감한다.
    if (cockpitNodes) {
      cockpitNodes.forEach(n => {
        const s = n.dataset.status;
        if (s === "running" || s === "deciding") setNodeStatus(n, "fail", "중단됨");
        else if (s === "pending") setNodeStatus(n, "skip", "미실행");  // 중단·예산소진으로 못 돈 단계
      });
    }
    if (reasoningEl) finishReasoning(reasoningEl);
    enhanceContent(container);   // 완성된 답변에 코드 하이라이트·수식 렌더 (한 번)
    setSending(false);
    loadConversations();
  }
}

/* 생성 상태 표시 — "돌아가는 중인지 멈춘 건지" 구분용.
   프리필/계획 단계는 원래 이벤트가 없어(비스트리밍) 오래 걸려도 정상이므로 멈춤으로
   보지 않는다. 빨간 멈춤 경고는 '토큰이 흐르다 20초 넘게 끊긴' 진짜 정지에만 띄운다. */
function startGenStatus() {
  const el = $("genStatus");
  el.classList.remove("hidden", "stall");
  const start = Date.now();
  let last = start, chunks = 0, output = false, stopped = false, waiting = false;
  const fmt = (ms) => {
    const s = Math.max(0, Math.floor(ms / 1000));
    return `${Math.floor(s / 60)}:${String(s % 60).padStart(2, "0")}`;
  };
  const tick = () => {
    if (stopped) return;
    const now = Date.now(), idle = now - last, elapsed = fmt(now - start);
    if (waiting) {
      // 위험 도구 승인 대기 — 모델이 멈춘 게 아니라 사용자의 결정을 기다리는 중
      el.classList.remove("stall");
      el.textContent = `🔐 위험 도구 실행 승인을 기다리는 중… (경과 ${elapsed}) — 위의 [승인/거절] 버튼을 눌러 주세요`;
      return;
    }
    if (output && idle > 20000) {
      // 응답이 흐르다 20초 넘게 끊김 = 진짜 멈춤 의심
      el.classList.add("stall");
      el.textContent = `⚠️ 응답이 ${Math.floor(idle / 1000)}초째 멈춰 있어요 (경과 ${elapsed}) — 느리거나 멈춘 것 같으면 [중단]하세요`;
    } else if (!output) {
      // 아직 첫 응답 전(프리필·계획). 컨텍스트가 크면 원래 오래 걸린다 — 멈춤 아님.
      el.classList.remove("stall");
      const hint = (now - start) > 60000 ? " · 컨텍스트가 크면 오래 걸립니다" : "";
      el.textContent = `⏳ 모델이 처리 중… ${elapsed} (첫 응답 준비${hint})`;
    } else {
      el.classList.remove("stall");
      el.textContent = `▍ 생성 중 · ${elapsed} · ${chunks}조각`;
    }
  };
  const timer = setInterval(tick, 500);
  tick();
  return {
    activity(isOutput) { last = Date.now(); if (isOutput) { output = true; chunks++; } },
    waiting(on) { waiting = on; last = Date.now(); tick(); },
    stop() { stopped = true; clearInterval(timer); el.classList.add("hidden"); el.classList.remove("stall"); },
  };
}

function setSending(on) {
  sending = on;
  sendBtn.textContent = on ? "중단" : "전송";
  sendBtn.classList.toggle("stop", on);
  sendBtn.classList.toggle("primary", !on);
}

sendBtn.addEventListener("click", () => {
  if (sending && abortController) abortController.abort();
  else send();
});

inputEl.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey && !e.isComposing) {
    e.preventDefault();
    send();
  }
});

function autoResize() {
  inputEl.style.height = "auto";
  inputEl.style.height = Math.min(inputEl.scrollHeight, 180) + "px";
}
inputEl.addEventListener("input", autoResize);

/* ==================== 대화 목록 ==================== */

/* 대화 목록 항목 <li>를 만든다. pid는 이 대화가 속한 공간(null=기본 공간).
   삭제·이름변경·열기가 전부 자기 공간(pid)을 대상으로 하도록 pid를 실어 부른다.
   기본 공간 리스트(#convList)와 프로젝트 폴더의 중첩 리스트가 함께 재사용한다. */
function makeConvLi(meta, pid) {
  const li = document.createElement("li");
  // 활성 하이라이트는 "같은 공간의 같은 대화"일 때만 — 다른 공간의 동명 대화가 켜지지 않게.
  if (meta.id === currentConvId && (pid || null) === currentProjectId) li.classList.add("active");
  const q = pid ? `?project_id=${encodeURIComponent(pid)}` : "";
  const title = document.createElement("span");
  title.className = "title";
  title.textContent = meta.title;
  const del = document.createElement("button");
  del.className = "del";
  del.textContent = "✕";
  del.title = "삭제";
  del.addEventListener("click", async (e) => {
    e.stopPropagation();
    if (!confirm(`"${meta.title}" 대화를 삭제할까요?`)) return;
    await fetch(`/api/conversations/${meta.id}${q}`, { method: "DELETE" });
    if (meta.id === currentConvId && (pid || null) === currentProjectId) newChat();
    loadConversations();
  });
  li.append(title, del);
  li.addEventListener("click", () => openConversation(meta.id, pid || null));
  li.addEventListener("dblclick", async () => {
    const name = prompt("새 제목:", meta.title);
    if (!name) return;
    await fetch(`/api/conversations/${meta.id}/rename${q}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ title: name }),
    });
    loadConversations();
    if (meta.id === currentConvId && (pid || null) === currentProjectId) $("convTitle").textContent = name;
  });
  return li;
}

/* 사이드바 갱신. 프로젝트 폴더(loadProjects)를 먼저 그리고, 기본 공간 대화는
   항상 #convList에 그린다(프로젝트 대화는 각 폴더 안에서 따로 채운다). */
async function loadConversations() {
  await loadProjects();
  let list = [];
  try { list = await (await fetch("/api/conversations")).json(); } catch (e) { list = []; }
  const ul = $("convList");
  ul.innerHTML = "";
  if (!list.length) {
    const li = document.createElement("li");
    li.className = "tree-empty";
    li.textContent = "대화 없음";
    ul.appendChild(li);
    return;
  }
  for (const meta of list) ul.appendChild(makeConvLi(meta, null));
}

async function openConversation(id, pid = null) {
  if (sending) return;
  currentProjectId = pid || null;   // 대화를 열면 그 대화의 공간이 활성 공간이 된다
  const res = await fetch(`/api/conversations/${id}${projQuery()}`);
  if (!res.ok) return;
  const conv = await res.json();
  currentConvId = conv.id;
  $("convTitle").textContent = conv.title;
  renderHistory(conv.messages);
  if (conv.summary) prependSummaryBanner(conv.summary);
  loadConversations();
}

/* 압축된 대화 상단에, 사라진 앞부분을 대신하는 요약을 접이식 배너로 보여준다.
   (요약은 conv.summary에 있고 messages에는 최근 원문만 남아 있다.) */
function prependSummaryBanner(summary) {
  const banner = document.createElement("details");
  banner.className = "summary-banner";
  const sm = document.createElement("summary");
  sm.textContent = "🗜 이전 대화 요약 (압축됨)";
  const body = document.createElement("div");
  body.className = "content";
  body.innerHTML = md(summary);
  banner.append(sm, body);
  messagesEl.insertBefore(banner, messagesEl.firstChild);
  enhanceContent(banner);
}

function newChat() {
  if (sending && abortController) abortController.abort();
  currentConvId = null;
  $("convTitle").textContent = "새 대화";
  // 활성 공간 배지 — 이 새 대화가 "일반 대화"에 담길지 어떤 프로젝트에 담길지 미리 보여준다.
  const space = currentProjectId
    ? `📁 ${escapeHtml(projectNames[currentProjectId] || "프로젝트")}`
    : "💬 일반 대화";
  messagesEl.innerHTML =
    `<div id="emptyHint"><div class="space-badge">${space}</div>` +
    "<h2>무엇을 도와드릴까요?</h2><p>메시지를 입력하면 대화가 시작됩니다.</p></div>";
  loadConversations();
}
// 상단 버튼은 항상 "일반 대화"(기본 공간)로 새 대화를 연다 — 활성 프로젝트가 있어도 벗어난다.
// 특정 프로젝트에서 시작하려면 그 프로젝트 폴더 헤더의 ＋ 버튼을 쓴다(makeProjectFolder).
$("newChatBtn").addEventListener("click", () => { currentProjectId = null; newChat(); });

/* ==================== 프로젝트 (프롬프트·기억 격리) ==================== */
/* 사이드바는 인라인 아코디언 폴더 트리다: 프로젝트 폴더를 클릭하면 그 자리에서 펼쳐지며
   해당 프로젝트의 대화가 중첩 리스트로 들여쓰기 되어 보인다. "일반 대화"(기본 공간)는
   맨 아래 별도 섹션(#convList)이라, 프로젝트와 일반 대화를 한 화면에서 구분해 본다. */

const expandedProjects = new Set();   // 펼쳐진 프로젝트 id (재렌더 사이 유지)

/* #projectList를 프로젝트 폴더들로 다시 그린다. 활성 프로젝트는 자동으로 펼친다. */
async function loadProjects() {
  let projects = [];
  try { projects = await (await fetch("/api/projects")).json(); }
  catch (e) { projects = []; }
  projectNames = Object.fromEntries(projects.map(p => [p.id, p.name]));
  // 활성/펼침 상태가 삭제된 프로젝트를 가리키면 정리한다.
  if (currentProjectId && !projects.some(p => p.id === currentProjectId)) currentProjectId = null;
  const ids = new Set(projects.map(p => p.id));
  for (const id of [...expandedProjects]) if (!ids.has(id)) expandedProjects.delete(id);

  const box = $("projectList");
  box.innerHTML = "";
  if (!projects.length) {
    const empty = document.createElement("div");
    empty.className = "tree-empty";
    empty.textContent = "아직 프로젝트가 없습니다. ＋로 만드세요.";
    box.appendChild(empty);
    return;
  }
  for (const p of projects) box.appendChild(makeProjectFolder(p));
}

/* 프로젝트 폴더 하나(헤더 + 중첩 대화 리스트). 헤더 클릭 = 그 프로젝트로 전환/펼치기,
   이미 활성이면 접기/펴기 토글. ⚙ = 설정 모달. */
function makeProjectFolder(p) {
  const isActive = p.id === currentProjectId;
  if (isActive) expandedProjects.add(p.id);   // 활성 프로젝트는 항상 펼쳐 둔다
  const expanded = expandedProjects.has(p.id);

  const wrap = document.createElement("div");
  wrap.className = "project-folder";

  const head = document.createElement("div");
  head.className = "folder-head" + (isActive ? " active" : "");
  const caret = document.createElement("span");
  caret.className = "caret";
  caret.textContent = expanded ? "▾" : "▸";
  const name = document.createElement("span");
  name.className = "folder-name";
  name.textContent = p.name;
  if (p.has_prompt) name.title = "프로젝트 프롬프트 있음";
  const count = document.createElement("span");
  count.className = "folder-count";
  if (p.conversation_count) count.textContent = p.conversation_count;
  const add = document.createElement("button");
  add.className = "gear";
  add.textContent = "＋";
  add.title = "이 프로젝트에서 새 대화 시작";
  add.addEventListener("click", (e) => {
    e.stopPropagation();
    currentProjectId = p.id;   // 활성 공간을 이 프로젝트로 바꾸고 빈 새 대화를 연다
    expandedProjects.add(p.id);
    newChat();                 // → loadConversations()로 사이드바가 이 프로젝트를 활성·펼침으로 다시 그린다
  });
  const gear = document.createElement("button");
  gear.className = "gear";
  gear.textContent = "⚙";
  gear.title = "프로젝트 설정 (이름·프롬프트·기억)";
  gear.addEventListener("click", (e) => { e.stopPropagation(); openProjectSettings(p.id); });
  head.append(caret, name, count, add, gear);

  const kids = document.createElement("ul");
  kids.className = "folder-convs";
  if (!expanded) kids.classList.add("hidden");

  head.addEventListener("click", async () => {
    if (p.id === currentProjectId) {
      // 이미 활성 → 접기/펴기만 토글
      if (expandedProjects.has(p.id)) expandedProjects.delete(p.id);
      else expandedProjects.add(p.id);
    } else {
      // 다른 프로젝트로 전환: 활성 공간을 바꾸고 현재 대화를 비운다
      currentProjectId = p.id;
      expandedProjects.add(p.id);
      newChat();
    }
    await loadConversations();
  });

  wrap.append(head, kids);
  if (expanded) fillProjectConvs(kids, p.id);
  return wrap;
}

/* 한 프로젝트의 대화들을 폴더 중첩 리스트에 채운다(지연 로드). */
async function fillProjectConvs(ul, pid) {
  ul.innerHTML = '<li class="tree-empty">불러오는 중…</li>';
  let list = [];
  try {
    const q = `?project_id=${encodeURIComponent(pid)}`;
    list = await (await fetch(`/api/conversations${q}`)).json();
  } catch (e) { list = []; }
  ul.innerHTML = "";
  if (!list.length) {
    ul.innerHTML = '<li class="tree-empty">대화 없음</li>';
    return;
  }
  for (const meta of list) ul.appendChild(makeConvLi(meta, pid));
}

/* ＋ 새 프로젝트: 이름을 받아 만들고, 바로 그 프로젝트로 전환한 뒤 설정 모달을 연다. */
$("newProjectBtn").addEventListener("click", async () => {
  const name = prompt("새 프로젝트 이름:");
  if (!name || !name.trim()) return;
  try {
    const res = await fetch("/api/projects", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: name.trim() }),
    });
    const proj = await res.json();
    if (!res.ok) { showToast(proj.detail || "생성 실패", "bad"); return; }
    currentProjectId = proj.id;
    expandedProjects.add(proj.id);
    newChat();
    await loadConversations();
    openProjectSettings(proj.id);   // 방금 만든 프로젝트의 프롬프트를 바로 채우게 연다
  } catch (err) { showToast("프로젝트 생성 오류", "bad"); }
});

/* ----- 프로젝트 설정 모달 (이름·프롬프트·이 프로젝트 기억) -----
   모달은 활성 공간과 무관하게 editingProjectId가 가리키는 프로젝트를 편집한다
   (⚙는 활성 프로젝트를 바꾸지 않고 그 프로젝트 설정만 연다). */
const projectModal = $("projectModal");
let editingProjectId = null;

async function openProjectSettings(pid) {
  if (!pid) return;
  editingProjectId = pid;
  let meta = {};
  try { meta = await (await fetch(`/api/projects/${pid}`)).json(); }
  catch (e) { showToast("프로젝트를 불러오지 못했습니다.", "bad"); return; }
  $("projectModalTitle").textContent = `프로젝트: ${meta.name || ""}`;
  $("projectName").value = meta.name || "";
  $("projectPrompt").value = meta.prompt || "";
  $("projectModalMsg").textContent = "";
  renderMemoryList(pid, $("projectMemList"), $("projectMemCount"));
  projectModal.classList.remove("hidden");
}
function closeProjectModal() { projectModal.classList.add("hidden"); }

$("projectModalClose").addEventListener("click", closeProjectModal);
projectModal.addEventListener("click", (e) => { if (e.target === projectModal) closeProjectModal(); });

$("projectSaveBtn").addEventListener("click", async () => {
  if (!editingProjectId) return;
  const name = $("projectName").value.trim();
  const promptText = $("projectPrompt").value;
  try {
    const res = await fetch(`/api/projects/${editingProjectId}`, {
      method: "PATCH", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: name || null, prompt: promptText }),
    });
    const meta = await res.json().catch(() => ({}));
    if (!res.ok) { $("projectModalMsg").textContent = meta.detail || "저장 실패"; return; }
    showToast("프로젝트를 저장했습니다.");
    await loadProjects();
    closeProjectModal();
  } catch (e) { $("projectModalMsg").textContent = "저장 중 오류가 발생했습니다."; }
});

$("projectClearMemBtn").addEventListener("click", async () => {
  if (!editingProjectId) return;
  if (!confirm("이 프로젝트의 장기 기억을 모두 지울까요?\n다른 프로젝트·기본 공간 기억은 그대로입니다. 되돌릴 수 없습니다.")) return;
  try {
    const q = encodeURIComponent(editingProjectId);
    const res = await fetch(`/api/memory?project_id=${q}`, { method: "DELETE" });
    const r = await res.json().catch(() => ({}));
    if (!res.ok) { showToast(r.detail || "비우기 실패", "bad"); return; }
    showToast(`이 프로젝트 기억 ${r.deleted}건을 비웠습니다.`);
    renderMemoryList(editingProjectId, $("projectMemList"), $("projectMemCount"));
  } catch (e) { showToast("기억 비우기 오류", "bad"); }
});

$("projectDeleteBtn").addEventListener("click", async () => {
  if (!editingProjectId) return;
  if (!confirm("이 프로젝트를 삭제할까요?\n이 프로젝트의 모든 대화와 기억이 함께 삭제됩니다. 되돌릴 수 없습니다.")) return;
  try {
    const res = await fetch(`/api/projects/${editingProjectId}`, { method: "DELETE" });
    const r = await res.json().catch(() => ({}));
    if (!res.ok) { showToast(r.detail || "삭제 실패", "bad"); return; }
    showToast("프로젝트를 삭제했습니다.");
    if (editingProjectId === currentProjectId) { currentProjectId = null; newChat(); }
    expandedProjects.delete(editingProjectId);
    editingProjectId = null;
    closeProjectModal();
    await loadConversations();
  } catch (e) { showToast("삭제 중 오류", "bad"); }
});

/* ----- 장기 기억 뷰어 (프로젝트 모달·설정 탭 공용) -----
   저장된 사실을 그대로 보여주고 개별(✕) 삭제한다. pid=null이면 기본 공간. */
async function renderMemoryList(pid, listEl, countEl) {
  const q = pid ? `?project_id=${encodeURIComponent(pid)}` : "";
  let data = { count: 0, items: [] };
  try { data = await (await fetch(`/api/memory${q}`)).json(); }
  catch (e) { if (countEl) countEl.textContent = "기억을 불러오지 못했습니다."; return; }
  const scope = pid ? "이 프로젝트 기억" : "기억";
  if (countEl) countEl.textContent = `${scope} ${data.count}건`;
  if (!listEl) return;
  listEl.innerHTML = "";
  const items = data.items || [];
  if (!items.length) {
    const li = document.createElement("li");
    li.className = "empty";
    li.textContent = "아직 저장된 기억이 없습니다. 대화하면서 자동으로 쌓입니다.";
    listEl.appendChild(li);
    return;
  }
  for (const it of items) listEl.appendChild(makeMemLi(it, pid, listEl, countEl));
}

function makeMemLi(it, pid, listEl, countEl) {
  const li = document.createElement("li");
  const body = document.createElement("span");
  body.className = "mem-text";
  body.textContent = it.content || "";
  if (it.kind && it.kind !== "fact") body.title = it.kind;
  const del = document.createElement("button");
  del.className = "del";
  del.textContent = "✕";
  del.title = "이 기억 삭제";
  del.addEventListener("click", async () => {
    if (!confirm("이 기억을 삭제할까요?\n되돌릴 수 없습니다.")) return;
    const q = pid ? `?project_id=${encodeURIComponent(pid)}` : "";
    const res = await fetch(`/api/memory/${it.id}${q}`, { method: "DELETE" });
    if (!res.ok) { showToast("삭제 실패", "bad"); return; }
    renderMemoryList(pid, listEl, countEl);   // 개수·목록 갱신
  });
  li.append(body, del);
  return li;
}

/* 대화 압축(파괴적) — 앞부분을 요약으로 치환해 컨텍스트를 줄인다. 최근 몇 턴만 원문으로
   남고 그 이전 메시지는 사라진다. 서버가 요약에 실패하면 원본은 그대로다. */
$("compactBtn").addEventListener("click", async () => {
  if (!currentConvId) { showToast("먼저 대화를 시작하세요.", "warn"); return; }
  if (sending) { showToast("응답 생성 중에는 압축할 수 없습니다.", "warn"); return; }
  if (!confirm(
    "이 대화의 앞부분을 요약으로 압축할까요?\n" +
    "최근 몇 턴만 원문으로 남고 그 이전 메시지는 사라집니다 (되돌릴 수 없음)."
  )) return;
  const btn = $("compactBtn");
  const prev = btn.textContent;
  btn.disabled = true; btn.textContent = "⏳";
  try {
    const res = await fetch(`/api/conversations/${currentConvId}/compact${projQuery()}`, { method: "POST" });
    const result = await res.json().catch(() => ({}));
    if (!res.ok) { showToast(result.detail || "압축 실패", "bad"); return; }
    showToast(`${result.dropped}개 메시지를 요약으로 압축했습니다.`);
    await openConversation(currentConvId);   // 줄어든 이력 + 요약 배너로 다시 그린다
  } catch (e) {
    showToast("압축 중 오류가 발생했습니다.", "bad");
  } finally {
    btn.disabled = false; btn.textContent = prev;
  }
});

/* 대화 비우기(파괴적) — 현재 대화를 통째로 삭제하고 이 공간의 장기 기억을 모두 비운다.
   압축이 "앞부분만 요약으로 줄이는" 것과 달리, 이건 대화·기억을 함께 초기화한다.
   기억 비우기는 로컬 접속 전용(_require_local)이므로 원격에선 거절될 수 있다. */
$("clearBtn").addEventListener("click", async () => {
  if (sending) { showToast("응답 생성 중에는 비울 수 없습니다.", "warn"); return; }
  const memScope = currentProjectId ? "이 프로젝트의 장기 기억" : "기본 공간의 장기 기억";
  if (!confirm(
    "이 대화를 삭제하고 " + memScope + "을 모두 비울까요?\n" +
    "대화 내용과 기억이 함께 사라집니다 (되돌릴 수 없음).\n" +
    "(다른 프로젝트·기본 공간의 기억은 그대로입니다.)"
  )) return;
  const btn = $("clearBtn");
  const prev = btn.textContent;
  btn.disabled = true; btn.textContent = "⏳";
  try {
    // 1) 현재 대화 삭제 (있을 때만)
    if (currentConvId) {
      const res = await fetch(`/api/conversations/${currentConvId}${projQuery()}`, { method: "DELETE" });
      if (!res.ok) {
        const r = await res.json().catch(() => ({}));
        showToast(r.detail || "대화 삭제 실패", "bad");
        return;
      }
    }
    // 2) 이 공간의 장기 기억 비우기
    const memRes = await fetch(`/api/memory${projQuery()}`, { method: "DELETE" });
    const memResult = await memRes.json().catch(() => ({}));
    if (!memRes.ok) {
      showToast(memResult.detail || "기억 비우기 실패", "bad");
      return;   // 대화는 이미 지워졌으니 UI는 아래 finally 뒤 새 대화로 정리된다
    }
    showToast(`대화를 삭제하고 기억 ${memResult.deleted}건을 비웠습니다.`);
    refreshMemoryCount();
    newChat();   // 성공 시에만 빈 새 대화 화면으로 되돌린다 (실패 시 현재 화면 유지)
  } catch (e) {
    showToast("비우는 중 오류가 발생했습니다.", "bad");
  } finally {
    btn.disabled = false; btn.textContent = prev;
  }
});

/* ==================== 파일 첨부 (📁 OS 네이티브 대화상자) ==================== */
/* 서버가 로컬에서 도는 점을 이용해 OS 기본 '열기' 대화상자를 띄우고, 고른 원본
   경로를 복사 없이 등록한다. office COM이 원본을 제자리에서 열어(복호화) 읽는다.
   원격(비-localhost) 접속이면 서버가 그 PC의 대화상자를 못 띄우므로(403), 브라우저
   업로드로 자동 폴백한다 — 버튼은 📁 하나로 두되 상황에 맞게 동작만 바뀐다. */

$("pickFileBtn").addEventListener("click", async () => {
  const btn = $("pickFileBtn");
  btn.disabled = true;
  try {
    const res = await fetch("/api/fs/dialog", { method: "POST" });
    if (res.status === 403) {
      // 원격 접속 — 네이티브 대화상자 불가. 브라우저 파일 선택으로 폴백(업로드=사본).
      $("fallbackFile").click();
      return;
    }
    const meta = await res.json();
    if (!res.ok) { alert(`파일 선택 실패: ${meta.detail || res.status}`); return; }
    if (meta.cancelled) return;   // 사용자가 대화상자를 취소
    attachments.push(meta);
    renderChips();
  } catch (err) {
    alert(`파일 선택 실패: ${err.message}`);
  } finally {
    btn.disabled = false;
  }
});

/* 원격 접속용 브라우저 업로드 폴백 — 선택 파일을 /api/upload로 보낸다(서버가 사본을
   만들어 확장자에 맞게 추출/경로 처리). DRM 원본 제자리 읽기는 로컬 📁에서만 가능. */
$("fallbackFile").addEventListener("change", async (e) => {
  for (const file of e.target.files) {
    const form = new FormData();
    form.append("file", file);
    try {
      const res = await fetch("/api/upload", { method: "POST", body: form });
      if (!res.ok) throw new Error((await res.json()).detail || `HTTP ${res.status}`);
      attachments.push(await res.json());
    } catch (err) {
      alert(`첨부 실패 (${file.name}): ${err.message}`);
    }
  }
  e.target.value = "";
  renderChips();
});

function renderChips() {
  const box = $("attachChips");
  box.innerHTML = "";
  attachments.forEach((a, idx) => {
    const chip = document.createElement("span");
    chip.className = "chip";
    // path 모드는 서버가 텍스트를 안 뽑고 경로만 넘긴다 — 도구로 읽는다.
    // ref(📁)는 복사 없이 고른 원본 경로, 그 외 path는 업로드 사본.
    const icon = a.ref ? "📁" : "📄";
    const label = a.mode === "path" ? (a.ref ? "원본 경로" : "도구로 읽음")
                                    : `${(a.chars || 0).toLocaleString()}자`;
    chip.textContent = `${icon} ${a.name} (${label})`;
    const x = document.createElement("button");
    x.textContent = "✕";
    x.addEventListener("click", () => { attachments.splice(idx, 1); renderChips(); });
    chip.appendChild(x);
    box.appendChild(chip);
  });
}

/* ==================== 상태 표시 ==================== */

async function loadStatus() {
  try {
    const st = await (await fetch("/api/status")).json();
    lastStatus = st;
    updateProviderSelect(st);
    const dot = $("statusDot");
    const active = $("providerSelect").value || st.active_provider;
    if (active !== "local") {
      const p = st.providers.find(x => x.name === active);
      dot.className = p && p.has_key ? "dot ok" : "dot mock";
      $("statusText").textContent = `외부 API: ${active}${p && p.model ? " · " + p.model : ""}`;
      $("statusText").title = p && p.has_key ? "" : "API 키가 등록되지 않았습니다.";
    } else if (st.mock) {
      dot.className = "dot mock";
      $("statusText").textContent = "목 모드 (모델 미연결)";
      $("statusText").title = st.mock_reason || "";
    } else if (st.llama.ready) {
      dot.className = "dot ok";
      $("statusText").textContent = st.llama.model || st.llama.alias;
      $("statusText").title = "";
    } else if (st.llama.running) {
      // 프로세스는 떴지만 헬스체크 전 = 모델 로딩 중
      dot.className = "dot mock";
      $("statusText").textContent = "모델 로딩 중… (몇 분 걸릴 수 있습니다)";
      $("statusText").title = "";
    } else {
      dot.className = "dot mock";
      $("statusText").textContent = "로컬 모델 미실행 — 시작 필요";
      $("statusText").title = "상단 [서빙 시작] 또는 설정에서 모델을 골라 시작하세요.";
    }
    // 로컬을 골랐고 프로세스가 아예 안 떠 있을 때만 셋업 바를 띄운다
    // (로딩 중이면 숨김 — 이미 시작했으므로).
    const showSetup = active === "local" && !st.mock && !st.llama.running;
    $("localSetup").classList.toggle("hidden", !showSetup);
    if (showSetup) renderLocalSetup(st);
    const names = Object.entries(st.mcp);
    const toolCount = names.reduce((n, [, s]) => n + (s.tools ? s.tools.length : 0), 0);
    $("mcpText").textContent = names.length
      ? `MCP ${names.filter(([, s]) => s.connected).length}/${names.length} 연결 · 도구 ${toolCount}개`
      : "MCP 서버 없음";
  } catch {
    $("statusDot").className = "dot bad";
    $("statusText").textContent = "서버 연결 끊김";
  }
}

/* 헤더의 모델 선택 드롭다운: 로컬 + 등록된 외부 프로바이더 */
function updateProviderSelect(st) {
  const sel = $("providerSelect");
  const sig = JSON.stringify([st.active_provider, st.providers, st.llama.alias,
                              st.llama.running, st.llama.ready, st.mock]);
  if (sig === providersSig) return;  // 변화 없으면 사용자의 선택을 건드리지 않는다
  providersSig = sig;
  sel.innerHTML = "";
  const local = document.createElement("option");
  local.value = "local";
  const localLabel = st.mock ? "목 모드"
    : st.llama.ready ? (st.llama.model || st.llama.alias || "모델")
    : st.llama.running ? "로딩 중…"
    : "미실행";
  local.textContent = `🖥 로컬: ${localLabel}`;
  sel.appendChild(local);
  for (const p of st.providers) {
    const opt = document.createElement("option");
    opt.value = p.name;
    opt.textContent = `☁ ${p.name}${p.model ? ": " + p.model : ""}${p.has_key ? "" : " (키 없음)"}`;
    sel.appendChild(opt);
  }
  sel.value = st.active_provider || "local";
  if (!sel.value) sel.value = "local";
}

$("providerSelect").addEventListener("change", async () => {
  // 선택을 서버 config.json에 저장 — 새로 열거나 다른 PC에서 접속해도 유지된다
  await fetch("/api/settings", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ active_provider: $("providerSelect").value }),
  });
  loadStatus();
});

/* ==================== 설정 모달 ==================== */

const modal = $("settingsModal");
$("settingsBtn").addEventListener("click", openSettings);
$("settingsClose").addEventListener("click", () => modal.classList.add("hidden"));
modal.addEventListener("click", (e) => { if (e.target === modal) modal.classList.add("hidden"); });

/* 설정 모달 탭 전환 — 패널은 전부 DOM에 남기고 display만 토글한다(숨은 입력값도
   collectSettings가 그대로 읽어야 하므로 언마운트하지 않는다). */
function switchSettingsTab(name) {
  document.querySelectorAll(".settings-tab").forEach(b =>
    b.classList.toggle("active", b.dataset.tab === name));
  document.querySelectorAll(".tab-panel").forEach(p =>
    p.classList.toggle("active", p.dataset.panel === name));
}
$("settingsTabs").addEventListener("click", (e) => {
  const btn = e.target.closest(".settings-tab");
  if (btn) switchSettingsTab(btn.dataset.tab);
});

/* 'LLM 연결' 탭 안의 서브탭(로컬 서버 ↔ 외부 API) 전환. 상위 탭과 마찬가지로
   서브패널은 모두 DOM에 남기고 display만 토글한다(숨은 입력값을 collectSettings가 읽음). */
function switchLlmSubtab(name) {
  document.querySelectorAll("#llmSubtabs .subtab").forEach(b =>
    b.classList.toggle("active", b.dataset.subtab === name));
  document.querySelectorAll('[data-panel="llm"] .subtab-panel').forEach(p =>
    p.classList.toggle("active", p.dataset.subpanel === name));
}
$("llmSubtabs").addEventListener("click", (e) => {
  const btn = e.target.closest(".subtab");
  if (btn) switchLlmSubtab(btn.dataset.subtab);
});

/* 완료 안내 토스트 — 저장·서빙 시작·MCP 재연결 등 성공 동작 뒤에 잠깐 띄운다.
   kind: "ok"(기본)·"warn"·"bad". 상세/오류는 여전히 settingsMsg에 남는다. */
let toastTimer = null;
function showToast(text, kind = "ok") {
  const el = $("toast");
  el.className = "toast";               // 이전 상태(hidden/kind) 초기화
  if (kind !== "ok") el.classList.add(kind);
  el.textContent = text;
  void el.offsetWidth;                  // 리플로우 강제 → 재호출에도 전환 애니메이션 재생
  el.classList.add("show");
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    el.classList.remove("show");
    setTimeout(() => el.classList.add("hidden"), 220);   // 페이드아웃 후 레이아웃에서 제거
  }, 2600);
}

/* 설정의 '장기 기억' 개수·내용 목록을 갱신한다(활성 공간 기준). 실패해도 조용히 물러선다.
   내용 렌더링·개별 삭제는 프로젝트 모달과 같은 공용 renderMemoryList를 쓴다. */
async function refreshMemoryCount() {
  renderMemoryList(currentProjectId, $("memList"), $("memCount"));
}

/* 장기 기억 전체 비우기 (파괴적, 로컬 전용). memory.db의 모든 사실을 지운다. */
$("clearMemoryBtn").addEventListener("click", async () => {
  const scopeMsg = currentProjectId
    ? "이 프로젝트의 모든 장기 기억을 지울까요?\n(다른 프로젝트·기본 공간 기억은 그대로입니다.)"
    : "기본 공간의 모든 장기 기억을 지울까요?\n(프로젝트별 기억은 그대로입니다.)";
  if (!confirm(scopeMsg + "\n되돌릴 수 없습니다. (진행 중인 대화는 그대로 남습니다.)")) return;
  try {
    const res = await fetch(`/api/memory${projQuery()}`, { method: "DELETE" });
    const result = await res.json().catch(() => ({}));
    if (!res.ok) { showToast(result.detail || "비우기 실패", "bad"); return; }
    showToast(`장기 기억 ${result.deleted}건을 비웠습니다.`);
    refreshMemoryCount();
  } catch (e) {
    showToast("기억을 비우는 중 오류가 발생했습니다.", "bad");
  }
});

async function openSettings() {
  switchSettingsTab("chat");   // 열 때마다 첫 탭으로
  const cfg = await (await fetch("/api/settings")).json();
  $("setSystemPrompt").value = cfg.system_prompt;
  $("setTemperature").value = cfg.temperature;
  $("setTopP").value = cfg.top_p;
  $("setTopK").value = cfg.top_k;
  $("setMaxTokens").value = cfg.max_tokens;
  $("setMaxToolRounds").value = cfg.max_tool_rounds;
  $("setApprovalEnabled").checked = cfg.approval_enabled !== false;
  $("setCtx").value = cfg.ctx;
  $("setKvQuant").checked = cfg.kv_quant;
  $("setGpuLayers").value = cfg.gpu_layers;
  $("setModelAlias").value = cfg.model_alias || "";
  $("setModelPath").value = cfg.model_path || "";
  $("setExternalUrl").value = cfg.llama_external_url || "";
  $("setAutostart").checked = !!cfg.autostart_local;
  $("setLlamaMode").value = cfg.llama_external_url ? "external" : "managed";
  applyLlamaMode();
  await refreshSettingsModels(cfg.model_path || "");
  renderServeStatus();
  editingProviders = (cfg.providers || []).map(p => ({ ...p }));
  renderProviderRows();
  $("settingsMsg").textContent = "";
  refreshMemoryCount();
  const mcpCfg = await (await fetch("/api/mcp/config")).json();
  editingMcpServers = parseMcpConfig(mcpCfg.content);
  renderMcpRows();
  renderMcpStatus(await (await fetch("/api/mcp")).json());
  modal.classList.remove("hidden");
}

/* mcp_servers.json 문자열 → 편집용 행 배열. 깨진 JSON이면 빈 목록으로 물러선다. */
function parseMcpConfig(content) {
  let servers = {};
  try {
    servers = (JSON.parse(content || "{}").mcpServers) || {};
  } catch { servers = {}; }
  return Object.entries(servers).map(([name, spec]) => {
    const { url, command, args, env, disabled, ...extra } = spec || {};
    return {
      name,
      transport: url ? "url" : "stdio",
      url: url || "",
      command: command || "",
      args: (args || []).join("\n"),
      env: Object.entries(env || {}).map(([k, v]) => `${k}=${v}`).join("\n"),
      disabled: !!disabled,
      _extra: extra,   // 폼이 모르는 필드(headers 등)를 잃지 않게 보존
    };
  });
}

/* 편집용 행 배열 → mcpServers 규격 JSON 문자열. 빈 이름 행은 버린다. */
function serializeMcpConfig() {
  const mcpServers = {};
  for (const s of editingMcpServers) {
    const name = (s.name || "").trim();
    if (!name) continue;
    const spec = { ...(s._extra || {}) };
    if (s.transport === "url") {
      const url = (s.url || "").trim();
      if (url) spec.url = url;
      delete spec.command; delete spec.args; delete spec.env;
    } else {
      const command = (s.command || "").trim();
      if (command) spec.command = command;
      const args = (s.args || "").split("\n").map(x => x.trim()).filter(Boolean);
      if (args.length) spec.args = args;
      const env = {};
      (s.env || "").split("\n").map(l => l.trim()).filter(Boolean).forEach(line => {
        const i = line.indexOf("=");
        if (i > 0) env[line.slice(0, i).trim()] = line.slice(i + 1).trim();
      });
      if (Object.keys(env).length) spec.env = env;
      delete spec.url;
    }
    if (s.disabled) spec.disabled = true;
    mcpServers[name] = spec;
  }
  return JSON.stringify({ mcpServers }, null, 2);
}

let _lastMcpStatus = null;   // 토글 실패 시 서버 진실값으로 UI를 되돌리기 위한 캐시

function renderMcpStatus(status) {
  _lastMcpStatus = status;
  const el = $("mcpStatus");
  const entries = Object.entries(status);
  const builtins = entries.filter(([, s]) => s.builtin);
  const externals = entries.filter(([, s]) => !s.builtin);
  el.innerHTML = "";

  // 내장 도구 — 설정 없이 앱에 들어있는 서버들. 체크박스로 켜고 끈다(즉시 재적재).
  if (builtins.length) {
    const box = document.createElement("div");
    box.className = "builtin-block";
    box.innerHTML = '<div class="builtin-title">내장 도구 <span class="hint">(설정 없이 바로 사용 — 켜고 끄기)</span></div>';
    for (const [name, s] of builtins) {
      const row = document.createElement("label");
      row.className = "builtin-row";
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = s.enabled !== false;
      cb.addEventListener("change", () => toggleBuiltin(name, cb.checked, cb));
      const dot = s.enabled === false ? "" :
        `<span class="${s.connected ? "ok" : "bad"}">●</span> `;
      const detail = s.enabled === false
        ? '<span class="hint">꺼짐</span>'
        : (s.connected ? `도구 ${s.tools.length}개` : `<span class="hint">${escapeHtml(s.error || "로드 안 됨")}</span>`);
      const label = document.createElement("span");
      label.className = "builtin-label";
      label.innerHTML = `${dot}<b>${escapeHtml(name)}</b>` +
        (s.desc ? ` <span class="hint">${escapeHtml(s.desc)}</span>` : "") + ` — ${detail}`;
      row.append(cb, label);
      box.appendChild(row);
    }
    el.appendChild(box);
  }

  // 외부(mcp_servers.json) 서버.
  const ext = document.createElement("div");
  ext.className = "external-block";
  if (externals.length) {
    ext.innerHTML = externals.map(([name, s]) => s.connected
      ? `<div><span class="ok">●</span> ${escapeHtml(name)} — 도구 ${s.tools.length}개: ${escapeHtml(s.tools.join(", "))}</div>`
      : `<div><span class="bad">●</span> ${escapeHtml(name)} — ${escapeHtml(s.error || "연결 안 됨")}</div>`
    ).join("");
  } else if (!builtins.length) {
    ext.innerHTML = "<div>등록된 MCP 서버가 없습니다.</div>";
  }
  el.appendChild(ext);
}

/* 내장 서버 켜고 끄기 — config.builtin_disabled를 갱신하고 MCP를 즉시 재적재한다. */
async function toggleBuiltin(name, enabled, cb) {
  cb.disabled = true;
  try {
    const res = await fetch("/api/mcp/builtin", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name, enabled }),
    });
    const result = await res.json();
    if (!res.ok) { showToast(result.detail || "변경 실패", "bad"); revertMcpStatus(cb, enabled); return; }
    renderMcpStatus(result.status);
    showToast(`내장 '${name}' ${enabled ? "켬" : "끔"}`, "ok");
  } catch (e) {
    showToast("내장 도구 변경 오류", "bad"); revertMcpStatus(cb, enabled);
  } finally {
    cb.disabled = false;
  }
}

/* 토글 실패 시 UI를 서버 상태로 되돌린다. 서버는 변경되지 않았으므로 마지막으로 렌더한
   상태를 다시 그리면 체크박스·연결 점이 모두 일관되게 복구된다(캐시 없으면 체크박스만 되돌림). */
function revertMcpStatus(cb, enabled) {
  if (_lastMcpStatus) renderMcpStatus(_lastMcpStatus);
  else cb.checked = !enabled;
}

/* MCP 서버 편집기 — 주소/명령을 폼으로 입력한다(JSON 직접 편집 대신) */
function renderMcpRows() {
  const box = $("mcpRows");
  box.innerHTML = "";
  if (!editingMcpServers.length) {
    box.innerHTML = '<p class="hint">등록된 MCP 서버가 없습니다. 아래에서 추가하세요.</p>';
  }
  editingMcpServers.forEach((s, idx) => {
    const row = document.createElement("div");
    row.className = "mcp-row";
    const bind = (el, key) => {
      el.value = s[key] || "";
      el.addEventListener("input", () => { s[key] = el.value; });
      return el;
    };
    const input = (key, placeholder) => {
      const el = document.createElement("input");
      el.type = "text"; el.placeholder = placeholder;
      return bind(el, key);
    };
    const area = (key, placeholder) => {
      const el = document.createElement("textarea");
      el.rows = 2; el.spellcheck = false; el.placeholder = placeholder;
      return bind(el, key);
    };

    // 머리줄: 이름 · 연결 방식 · 삭제
    const head = document.createElement("div");
    head.className = "mcp-row-head";
    const name = input("name", "이름 (도구 앞에 붙음, 예: office)");
    const kind = document.createElement("select");
    kind.innerHTML = '<option value="url">HTTP/SSE 주소</option>'
                   + '<option value="stdio">로컬 명령</option>';
    kind.value = s.transport;
    kind.addEventListener("change", () => { s.transport = kind.value; renderMcpRows(); });
    const del = document.createElement("button");
    del.className = "btn icon"; del.textContent = "✕"; del.title = "이 서버 삭제";
    del.addEventListener("click", () => { editingMcpServers.splice(idx, 1); renderMcpRows(); });
    head.append(name, kind, del);

    // 방식별 입력란
    const fields = document.createElement("div");
    fields.className = "mcp-row-fields";
    if (s.transport === "url") {
      const url = input("url", "http://127.0.0.1:8087/mcp/");
      fields.append(labeled("서버 주소", url));
    } else {
      const cmd = input("command", "실행 명령 (예: python)");
      const args = area("args", "인자 — 한 줄에 하나씩\noutlook_server.py");
      const env = area("env", "환경변수 — KEY=VALUE, 한 줄에 하나씩 (선택)");
      fields.append(labeled("실행 명령", cmd), labeled("인자", args), labeled("환경변수", env));
    }

    // 비활성 토글
    const foot = document.createElement("label");
    foot.className = "mcp-disabled";
    const cb = document.createElement("input");
    cb.type = "checkbox"; cb.checked = s.disabled;
    cb.addEventListener("change", () => { s.disabled = cb.checked; });
    foot.append(cb, document.createTextNode(" 이 서버 비활성(연결 시도 안 함)"));

    row.append(head, fields, foot);
    box.appendChild(row);
  });
  $("mcpConfig").value = serializeMcpConfig();  // 미리보기 갱신
}

/* 라벨 + 입력란을 세로로 묶는다 */
function labeled(text, el) {
  const wrap = document.createElement("label");
  wrap.className = "mcp-field";
  const span = document.createElement("span");
  span.className = "mcp-field-label";
  span.textContent = text;
  wrap.append(span, el);
  return wrap;
}

$("addMcpBtn").addEventListener("click", () => {
  const preset = MCP_PRESETS[$("mcpPreset").value] || MCP_PRESETS.url;
  editingMcpServers.push({ ...preset, _extra: {} });
  renderMcpRows();
});

/* 외부 LLM 프로바이더 편집기 */
function renderProviderRows() {
  const box = $("providerRows");
  box.innerHTML = "";
  if (!editingProviders.length) {
    box.innerHTML = '<p class="hint">등록된 외부 연결이 없습니다. 아래에서 추가하세요.</p>';
    return;
  }
  editingProviders.forEach((p, idx) => {
    const row = document.createElement("div");
    row.className = "provider-row";
    const mk = (key, placeholder, type = "text") => {
      const input = document.createElement("input");
      input.type = type;
      input.value = p[key] || "";
      input.placeholder = placeholder;
      input.addEventListener("input", () => { p[key] = input.value; });
      return input;
    };
    const del = document.createElement("button");
    del.className = "btn icon";
    del.textContent = "✕";
    del.title = "이 연결 삭제";
    del.addEventListener("click", () => { editingProviders.splice(idx, 1); renderProviderRows(); });
    const name = mk("name", "이름 (예: OpenAI)");
    const model = mk("model", "모델 ID (예: gpt-4o-mini)");
    const url = mk("base_url", "https://api.example.com/v1");
    const key = mk("api_key", "API 키", "password");
    url.classList.add("span2");
    key.classList.add("span2");
    row.append(name, model, del, url, key);
    box.appendChild(row);
  });
}

$("addProviderBtn").addEventListener("click", () => {
  const preset = PROVIDER_PRESETS[$("providerPreset").value] || PROVIDER_PRESETS.custom;
  editingProviders.push({ ...preset });
  renderProviderRows();
});

/* 설정 모달의 '전체' 값을 config 규격 본문으로 모은다 (생성 파라미터 + 서버 + 프로바이더).
   [설정만 저장]뿐 아니라 [서빙 시작]·[재시작]도 이걸 저장한다 — 모달에서 고친 시스템
   프롬프트·생성 파라미터가 서빙만 시작하면 사라지는 일이 없게 한다.
   주의: 설정 모달의 입력란이 채워져 있는 상태(openSettings 이후)에서만 부른다.
   셋업 바(모달 미개봉)에서는 setupServeBody를 쓴다 — 빈 입력란으로 덮어쓰면 안 되므로. */
function collectSettings() {
  const providers = editingProviders
    .map(p => ({
      name: (p.name || "").trim(),
      base_url: (p.base_url || "").trim(),
      api_key: (p.api_key || "").trim(),
      model: (p.model || "").trim(),
    }))
    .filter(p => p.name || p.base_url || p.api_key || p.model);  // 완전히 빈 행은 버린다
  const body = {
    system_prompt: $("setSystemPrompt").value,
    temperature: numOr("setTemperature"),
    top_p: numOr("setTopP"),
    top_k: numOr("setTopK", true),
    max_tokens: numOr("setMaxTokens", true),
    max_tool_rounds: numOr("setMaxToolRounds", true),
    approval_enabled: $("setApprovalEnabled").checked,
    ...serverSettingsBody(),
    providers,
  };
  // 선택돼 있던 프로바이더가 삭제됐으면 로컬로 되돌린다
  const current = $("providerSelect").value;
  if (current !== "local" && !providers.some(p => p.name === current)) {
    body.active_provider = "local";
  }
  return body;
}

$("saveSettingsBtn").addEventListener("click", async () => {
  const { ok, result } = await putSettings(collectSettings());
  $("settingsMsg").textContent = ok
    ? (result.restart_required ? "저장됨 — 서버 설정 변경은 [서버 재시작] 후 적용됩니다." : "저장됨.")
    : `저장 실패: ${result.detail}`;
  if (ok) {
    showToast(result.restart_required ? "✅ 설정 저장됨 — 일부는 서버 재시작 후 적용" : "✅ 설정이 저장되었습니다");
    providersSig = ""; loadStatus();
  } else {
    showToast("설정 저장 실패", "bad");
  }
});

/* ==================== 로컬 LLM 서버 (선택·시작·중지·재시작) ==================== */
/* 서빙은 앱 시작과 분리돼 있다. 여기서 GGUF/옵션을 고르고 서빙을 켜고 끈다.
   서버 파라미터는 먼저 PUT /api/settings로 저장한 뒤 /api/server/* 를 호출한다. */

function fmtSize(bytes) {
  if (!bytes) return "?";
  const gb = bytes / 1073741824;
  return gb >= 1 ? gb.toFixed(1) + "GB" : Math.round(bytes / 1048576) + "MB";
}

async function fetchModels() {
  try { return (await (await fetch("/api/models")).json()).models || []; }
  catch { return []; }
}

function fillModelSelect(sel, models, withManual) {
  sel.innerHTML = "";
  sel.appendChild(new Option("자동 (models 폴더 최근 파일)", ""));
  for (const m of models) sel.appendChild(new Option(`${m.name} (${fmtSize(m.size)})`, m.path));
  if (withManual) sel.appendChild(new Option("직접 경로 입력…", "__manual__"));
}

/* 설정 모달의 모델 드롭다운을 채운다. preferPath가 목록에 없고 비어있지 않으면
   '직접 경로 입력'으로 두고 텍스트 입력란을 노출한다(폴더 밖 경로 대응). */
async function refreshSettingsModels(preferPath) {
  const sel = $("setModelSelect");
  fillModelSelect(sel, await fetchModels(), true);
  if (!preferPath) sel.value = "";
  else if ([...sel.options].some(o => o.value === preferPath)) sel.value = preferPath;
  else sel.value = "__manual__";
  applyModelSelectManual();
}

/* 셋업 바(메인)의 모델 드롭다운을 채운다. 이전 선택은 최대한 유지한다. */
async function refreshSetupModels() {
  const sel = $("setupModel");
  const prev = sel.value;
  fillModelSelect(sel, await fetchModels(), false);
  if (prev && [...sel.options].some(o => o.value === prev)) sel.value = prev;
}

function applyModelSelectManual() {
  $("modelPathManual").classList.toggle("hidden", $("setModelSelect").value !== "__manual__");
}

function applyLlamaMode() {
  const external = $("setLlamaMode").value === "external";
  $("managedFields").classList.toggle("hidden", external);
  $("externalFields").classList.toggle("hidden", !external);
}

/* 숫자 입력란을 안전하게 읽는다. 비어 있거나 숫자가 아니면 undefined를 돌려주고,
   호출부는 그 키를 본문에서 빼(JSON.stringify가 undefined 키를 생략) 저장된 설정을
   덮어쓰지 않게 한다 — 입력란을 비웠다고 config가 null로 오염돼 서빙이 깨지는 것을 막는다. */
function numOr(id, isInt) {
  const raw = ($(id).value || "").trim();
  const v = isInt ? parseInt(raw, 10) : parseFloat(raw);
  return Number.isFinite(v) ? v : undefined;
}

/* 설정 모달의 서버 관련 값을 config 규격 dict로 모은다 (PUT /api/settings 본문에 포함). */
function serverSettingsBody() {
  const body = {
    model_alias: $("setModelAlias").value.trim(),
    ctx: numOr("setCtx", true),
    kv_quant: $("setKvQuant").checked,
    gpu_layers: numOr("setGpuLayers", true),
    autostart_local: $("setAutostart").checked,
  };
  if ($("setLlamaMode").value === "external") {
    body.llama_external_url = $("setExternalUrl").value.trim();
    body.model_path = "";                       // external은 모델 파일 불필요
  } else {
    body.llama_external_url = "";
    const sel = $("setModelSelect").value;
    body.model_path = sel === "__manual__" ? $("setModelPath").value.trim() : sel;
  }
  return body;
}

async function putSettings(body) {
  const res = await fetch("/api/settings", {
    method: "PUT", headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return { ok: res.ok, result: await res.json() };
}

/* 현재 서빙 상태를 설정 모달 상단에 표시하고 시작/중지/재시작 버튼을 켜고 끈다. */
async function renderServeStatus() {
  const el = $("serveStatus");
  let st = lastStatus;
  try { st = await (await fetch("/api/status")).json(); lastStatus = st; } catch { /* 캐시 사용 */ }
  if (!st) { el.textContent = ""; return; }
  const l = st.llama;
  if (st.mock) el.innerHTML = '<span class="bad">●</span> 목(--mock) 모드 — 서빙 불가';
  else if (l.ready) el.innerHTML = `<span class="ok">●</span> 서빙 중 — ${escapeHtml(l.model || l.alias || "")}${l.external ? " (외부 서버)" : ""}`;
  else if (l.running) el.innerHTML = '<span class="idle">●</span> 모델 로딩 중… (몇 분 걸릴 수 있습니다)';
  else el.innerHTML = '<span class="idle">●</span> 미실행 (유휴) — 아래에서 시작하세요';
  const running = l.running, mock = st.mock;
  $("startServeBtn").disabled = running || mock;
  $("stopServeBtn").disabled = !running || mock;
  $("restartBtn").disabled = !running || mock;
}

/* 서빙 시작: 설정 저장 → /api/server/start. fromSettings면 설정 모달(전체 설정 저장),
   아니면 셋업 바(모델 경로만). 진행 중엔 시작 버튼을 잠가 중복 요청을 막는다. */
async function startServe(fromSettings) {
  const msgEl = fromSettings ? $("settingsMsg") : $("setupMsg");
  const startBtn = fromSettings ? $("startServeBtn") : $("setupStartBtn");
  const body = fromSettings ? collectSettings() : setupServeBody();
  startBtn.disabled = true;
  msgEl.textContent = "설정 저장 중…";
  try {
    const put = await putSettings(body);
    if (!put.ok) { msgEl.textContent = `저장 실패: ${put.result.detail}`; return; }
    msgEl.textContent = "서빙 시작 중… (모델 로드에 몇 분 걸릴 수 있습니다)";
    const res = await fetch("/api/server/start", { method: "POST" });
    const result = await res.json().catch(() => ({}));
    msgEl.textContent = res.ok ? "서빙이 시작되었습니다." : `시작 실패: ${result.detail || res.status}`;
    showToast(res.ok ? "✅ 로컬 서빙을 시작했습니다 (모델 로드는 잠시 걸릴 수 있습니다)"
                     : `서빙 시작 실패: ${result.detail || res.status}`, res.ok ? "ok" : "bad");
    providersSig = "";                    // 헤더 로컬 라벨 즉시 갱신
  } finally {
    startBtn.disabled = false;            // renderServeStatus/loadStatus가 상태 기반으로 재설정
    if (fromSettings) renderServeStatus();
    loadStatus();
  }
}

/* 셋업 바(메인)에서 시작할 때의 설정 본문. 외부 서버가 설정돼 있으면 그대로 붙는다. */
function setupServeBody() {
  const ext = lastStatus && lastStatus.llama_config && lastStatus.llama_config.external_url;
  if (ext) return {};                   // 외부: 저장된 설정 유지, 붙기만
  return { model_path: $("setupModel").value, llama_external_url: "" };
}

async function stopServe() {
  $("settingsMsg").textContent = "중지 중…";
  const res = await fetch("/api/server/stop", { method: "POST" });
  $("settingsMsg").textContent = res.ok ? "서빙을 중지했습니다." : "중지 실패";
  showToast(res.ok ? "로컬 서빙을 중지했습니다" : "중지 실패", res.ok ? "warn" : "bad");
  providersSig = "";
  renderServeStatus();
  loadStatus();
}

$("setLlamaMode").addEventListener("change", applyLlamaMode);
$("setModelSelect").addEventListener("change", applyModelSelectManual);
$("refreshModelsBtn").addEventListener("click", () => {
  const cur = $("setModelSelect").value === "__manual__" ? $("setModelPath").value.trim() : $("setModelSelect").value;
  refreshSettingsModels(cur);
});
$("startServeBtn").addEventListener("click", () => startServe(true));
$("stopServeBtn").addEventListener("click", stopServe);

/* 셋업 바 버튼 (startServe가 버튼 잠금/해제를 처리한다) */
$("setupStartBtn").addEventListener("click", () => startServe(false));
$("setupRefreshBtn").addEventListener("click", refreshSetupModels);
$("setupSettingsBtn").addEventListener("click", openSettings);

/* 셋업 바 표시 시 호출 — 외부 서버 설정이면 '연결' 버튼으로, 아니면 모델 선택으로. */
function renderLocalSetup(st) {
  const ext = st.llama_config && st.llama_config.external_url;
  if (ext) {
    $("setupText").textContent = `외부 LLM 서버가 설정돼 있습니다 (${ext}). 연결하세요.`;
    $("setupModel").classList.add("hidden");
    $("setupRefreshBtn").classList.add("hidden");
    $("setupStartBtn").textContent = "연결";
  } else {
    $("setupText").textContent = "로컬 모델이 실행되고 있지 않습니다. 모델을 골라 시작하세요.";
    $("setupModel").classList.remove("hidden");
    $("setupRefreshBtn").classList.remove("hidden");
    $("setupStartBtn").textContent = "서빙 시작";
    if ($("setupModel").options.length === 0) refreshSetupModels();
  }
}

$("restartBtn").addEventListener("click", async () => {
  if (!confirm("llama-server를 재시작할까요? 진행 중인 응답이 끊깁니다.")) return;
  const put = await putSettings(collectSettings());  // 모달의 모든 편집(프롬프트·파라미터 포함)을 저장
  if (!put.ok) { $("settingsMsg").textContent = `저장 실패: ${put.result.detail}`; return; }
  $("settingsMsg").textContent = "재시작 중… (모델 로드에 몇 분 걸릴 수 있습니다)";
  const res = await fetch("/api/server/restart", { method: "POST" });
  const result = await res.json().catch(() => ({}));
  $("settingsMsg").textContent = res.ok ? "재시작 완료." : `실패: ${result.detail || res.status}`;
  showToast(res.ok ? "✅ 서버를 재시작했습니다" : `재시작 실패: ${result.detail || res.status}`,
            res.ok ? "ok" : "bad");
  providersSig = "";
  renderServeStatus();
  loadStatus();
});

$("saveMcpBtn").addEventListener("click", async () => {
  // 이름 중복은 서버 이름이 도구 접두사라 조용히 덮어써진다 — 미리 막는다.
  const names = editingMcpServers.map(s => (s.name || "").trim()).filter(Boolean);
  const dup = names.find((n, i) => names.indexOf(n) !== i);
  if (dup) { $("settingsMsg").textContent = `이름이 중복됩니다: ${dup}`; return; }
  const content = serializeMcpConfig();
  $("mcpConfig").value = content;
  $("settingsMsg").textContent = "MCP 재연결 중…";
  const res = await fetch("/api/mcp/config", {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ content }),
  });
  const result = await res.json();
  if (res.ok) {
    $("settingsMsg").textContent = "MCP 저장 + 재연결 완료.";
    const connected = Object.values(result.status || {}).filter(s => s.connected).length;
    const total = Object.keys(result.status || {}).length;
    showToast(`✅ MCP 저장 완료 — ${connected}/${total} 서버 연결됨`);
    renderMcpStatus(result.status);
  } else {
    $("settingsMsg").textContent = `실패: ${result.detail}`;
    showToast(`MCP 저장 실패: ${result.detail}`, "bad");
  }
  loadStatus();
});

/* ==================== 호스팅 종료 ==================== */

let shuttingDown = false;
const statusTimer = setInterval(() => { if (!shuttingDown) loadStatus(); }, 15000);

$("shutdownBtn").addEventListener("click", async () => {
  if (!confirm("호스팅을 종료할까요?\n웹 UI와 로컬 LLM 서버가 모두 내려갑니다. (따로 띄운 외부 LLM 서버는 유지됩니다.)")) return;
  shuttingDown = true;
  clearInterval(statusTimer);
  if (sending && abortController) abortController.abort();  // 진행 중 응답 중단
  try {
    await fetch("/api/shutdown", { method: "POST" });
  } catch { /* 서버가 바로 내려가 응답이 끊겨도 정상 — 오버레이를 띄운다 */ }
  $("shutdownOverlay").classList.remove("hidden");
});

/* ==================== 초기화 ==================== */

loadStatus();
loadConversations();   // 내부에서 loadProjects()도 호출해 사이드바 전체를 그린다
inputEl.focus();
