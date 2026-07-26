/* LocalLLM Studio 프론트엔드 — 외부 라이브러리 없이 동작 */
"use strict";

const $ = (id) => document.getElementById(id);
const messagesEl = $("messages");
const inputEl = $("input");
const sendBtn = $("sendBtn");
const taskModeBtn = $("taskModeBtn");

let currentConvId = null;
let attachments = [];   // {id, name, chars}
let sending = false;
let abortController = null;
let taskMode = false;   // 작업 모드(계획-실행) 토글 상태
let lastStatus = null;  // 마지막 /api/status 응답 (셋업 바·서빙 버튼 판단용)

taskModeBtn.addEventListener("click", () => {
  taskMode = !taskMode;
  taskModeBtn.classList.toggle("active", taskMode);
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

function mdInline(s) {
  return s
    .replace(/`([^`]+)`/g, (_, c) => `<code>${c}</code>`)
    .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
    .replace(/(^|[^*])\*([^*\s][^*]*)\*/g, "$1<em>$2</em>")
    .replace(/\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)/g,
      '<a href="$2" target="_blank" rel="noopener">$1</a>');
}

function md(src) {
  const lines = escapeHtml(src).split("\n");
  const out = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    if (line.startsWith("```")) {                       // 코드 블록
      const buf = [];
      i++;
      while (i < lines.length && !lines[i].startsWith("```")) buf.push(lines[i++]);
      i++;
      out.push(`<pre><code>${buf.join("\n")}</code></pre>`);
    } else if (/^#{1,3}\s/.test(line)) {                // 제목
      const level = line.match(/^#+/)[0].length;
      out.push(`<h${level}>${mdInline(line.replace(/^#+\s*/, ""))}</h${level}>`);
    } else if (/^(\*{3,}|-{3,})\s*$/.test(line)) {      // 구분선
      out.push("<hr>"); i++;  continue;
    } else if (/^\s*([-*]|\d+\.)\s/.test(line)) {       // 목록
      const ordered = /^\s*\d+\./.test(line);
      const tag = ordered ? "ol" : "ul";
      const items = [];
      while (i < lines.length && /^\s*([-*]|\d+\.)\s/.test(lines[i])) {
        items.push(`<li>${mdInline(lines[i].replace(/^\s*([-*]|\d+\.)\s/, ""))}</li>`);
        i++;
      }
      out.push(`<${tag}>${items.join("")}</${tag}>`);
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
      const head = cells(line).map(c => `<th>${mdInline(c)}</th>`).join("");
      i += 2;
      const rows = [];
      while (i < lines.length && lines[i].includes("|")) {
        rows.push(`<tr>${cells(lines[i]).map(c => `<td>${mdInline(c)}</td>`).join("")}</tr>`);
        i++;
      }
      out.push(`<table><thead><tr>${head}</tr></thead><tbody>${rows.join("")}</tbody></table>`);
      continue;
    } else if (line.trim() === "") {
      i++;  continue;
    } else {                                            // 문단
      const buf = [line];
      i++;
      while (i < lines.length && lines[i].trim() !== "" && !/^(#|```|[-*]\s|\d+\.\s|&gt;)/.test(lines[i])) {
        buf.push(lines[i]); i++;
      }
      out.push(`<p>${buf.map(mdInline).join("<br>")}</p>`);
      continue;
    }
    i++;
  }
  return out.join("\n");
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
  const pre = document.createElement("pre");
  pre.textContent = result;
  details.appendChild(pre);
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
  body.className = "reasoning-body";
  details.append(summary, body);
  container.insertBefore(details, container.firstChild);  // 답변보다 위에
  scrollBottom();
  return details;
}

function appendReasoning(details, text) {
  details.querySelector(".reasoning-body").textContent += text;
  scrollBottom();
}

function finishReasoning(details) {
  const sp = details.querySelector(".tool-spinner");
  if (sp) { sp.classList.remove("tool-spinner"); sp.textContent = "완료"; }
  details.open = false;   // 끝난 생각은 접어 답변에 집중
}

function addErrorNote(container, message) {
  const note = document.createElement("div");
  note.className = "error-note";
  note.textContent = `⚠ ${message}`;
  container.appendChild(note);
}

/* 계획-실행(작업 모드) 렌더 */
function addPlanBlock(container, steps, replan) {
  const div = document.createElement("div");
  div.className = "plan-block";
  const title = replan ? `🧭 재계획 #${replan}` : "🧭 계획";
  div.innerHTML = `<div class="plan-title">${title}</div><ol>` +
    steps.map(s => `<li>${escapeHtml(s)}</li>`).join("") + "</ol>";
  container.appendChild(div);
  scrollBottom();
}

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
  const pre = document.createElement("pre");
  pre.className = "step-body";
  details.append(summary, pre);
  container.appendChild(details);
  scrollBottom();
  return details;
}

function appendStepToken(details, text) {
  if (!details) return;
  details.querySelector(".step-body").textContent += text;
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
  details.open = false;   // 끝난 단계는 접어 답변에 집중
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
    attachments: attachments.map(a => a.id),
    provider: $("providerSelect").value || null,
    task_mode: taskMode,
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
  let currentStep = null;   // 현재 진행 중인 단계 블록 (작업 모드)
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
        } else if (ev.type === "plan") {
          addPlanBlock(container, ev.steps || [], ev.replan);
          answerFresh = true;
        } else if (ev.type === "step_start") {
          currentStep = addStepBlock(container, ev.index, ev.text);
          answerFresh = true;
        } else if (ev.type === "step_token") {
          if (ev.index === -1) addPlanNote(container, ev.text);
          else appendStepToken(currentStep, ev.text);
        } else if (ev.type === "step_done") {
          finishStepBlock(currentStep, ev.ok, ev.result);
          currentStep = null;
          answerFresh = true;
        } else if (ev.type === "approval_request") {
          if (!buffer) contentEl.remove();
          currentApproval = addApprovalBlock(container, ev.id, ev.name, ev.arguments);
          status.waiting(true);
        } else if (ev.type === "approval_result") {
          finishApprovalBlock(currentApproval, ev.approved, ev.timeout);
          currentApproval = null;
          status.waiting(false);
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
    if (reasoningEl) finishReasoning(reasoningEl);
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

async function loadConversations() {
  const list = await (await fetch("/api/conversations")).json();
  const ul = $("convList");
  ul.innerHTML = "";
  for (const meta of list) {
    const li = document.createElement("li");
    if (meta.id === currentConvId) li.classList.add("active");
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
      await fetch(`/api/conversations/${meta.id}`, { method: "DELETE" });
      if (meta.id === currentConvId) newChat();
      loadConversations();
    });
    li.append(title, del);
    li.addEventListener("click", () => openConversation(meta.id));
    li.addEventListener("dblclick", async () => {
      const name = prompt("새 제목:", meta.title);
      if (!name) return;
      await fetch(`/api/conversations/${meta.id}/rename`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ title: name }),
      });
      loadConversations();
      if (meta.id === currentConvId) $("convTitle").textContent = name;
    });
    ul.appendChild(li);
  }
}

async function openConversation(id) {
  if (sending) return;
  const res = await fetch(`/api/conversations/${id}`);
  if (!res.ok) return;
  const conv = await res.json();
  currentConvId = conv.id;
  $("convTitle").textContent = conv.title;
  renderHistory(conv.messages);
  loadConversations();
}

function newChat() {
  if (sending && abortController) abortController.abort();
  currentConvId = null;
  $("convTitle").textContent = "새 대화";
  messagesEl.innerHTML =
    '<div id="emptyHint"><h2>무엇을 도와드릴까요?</h2><p>메시지를 입력하면 대화가 시작됩니다.</p></div>';
  loadConversations();
}
$("newChatBtn").addEventListener("click", newChat);

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

function renderMcpStatus(status) {
  const el = $("mcpStatus");
  const entries = Object.entries(status);
  el.innerHTML = entries.length
    ? entries.map(([name, s]) => s.connected
        ? `<div><span class="ok">●</span> ${escapeHtml(name)} — 도구 ${s.tools.length}개: ${escapeHtml(s.tools.join(", "))}</div>`
        : `<div><span class="bad">●</span> ${escapeHtml(name)} — ${escapeHtml(s.error || "연결 안 됨")}</div>`
      ).join("")
    : "<div>등록된 MCP 서버가 없습니다.</div>";
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
loadConversations();
inputEl.focus();
