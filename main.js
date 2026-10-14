/**
 * main.js — Electron 메인 프로세스
 *
 * 역할:
 *   Electron 앱의 Node.js 백엔드. 창 생성, 파일 시스템 접근,
 *   AI API 호출, PPTX 파일 처리 등 렌더러(브라우저)가 직접 할 수 없는
 *   작업을 모두 담당한다.
 *
 * 통신 방식:
 *   렌더러(index.html, screen.html)에서 ipcRenderer.invoke/send 로 요청하면
 *   여기서 ipcMain.handle/on 으로 받아 처리 후 결과를 돌려준다.
 *
 * 주요 기능:
 *   1. 설정 파일(settings.json) 읽기/쓰기
 *   2. AI API(Gemini/OpenAI/Claude)로 가사 자동 분리 및 번역
 *   3. 메인 창 + 외부 송출 창 관리
 *   4. PPTX 내보내기 / 가져오기
 */

// ── 외부 라이브러리 ──
const { app, BrowserWindow, ipcMain, Menu, screen, dialog, shell } = require("electron");
const fs   = require("fs");    // 파일 읽기/쓰기
const path = require("path");  // 경로 조합
const { fileURLToPath } = require("url");
const { GoogleGenAI } = require("@google/genai");  // Gemini AI SDK
const { DOMParser } = require("@xmldom/xmldom");
const DEFAULT_MAX_LINES_PER_SLIDE = 1;

// PPTX 관련 라이브러리는 선택적으로 로드 (없어도 앱 실행 가능)
let AdmZip = null;  // ZIP 파일 파싱 / PPTX 직접 XML 생성
try { AdmZip = require("adm-zip"); } catch(e) { console.warn("adm-zip 미설치:", e.message); }

// PPTX 내보내기 전용 순수 모듈 (Electron 비의존 — 단독 테스트 가능)
let buildPptxBuffer = null;
try { ({ buildPptxBuffer } = require("./pptxExport")); } catch(e) { console.warn("pptxExport 모듈 로드 실패:", e.message); }

// ── 전역 창 참조 ──
let mainWin   = null;  // 메인 조작창 (index.html)
let secondWin = null;  // 외부 송출 화면 (screen.html)

// 현재 열려있는 프로젝트(.wsproj) 파일의 절대 경로. 저장/열기 시 갱신된다.
// "덮어쓰기 저장"은 이 경로에 dialog 없이 바로 기록한다.
let currentProjectPath = null;
// 메인 창 종료를 실제로 진행할지 여부 (렌더러의 저장 확인을 거친 뒤 true)
let appIsQuitting = false;

// ──────────────────────────────────────────────────────────
// 설정 파일 관리
// 저장 위치: OS별 userData 디렉터리 (예: ~/Library/Application Support/Cheon Myeong Slide Pro/settings.json)
// ──────────────────────────────────────────────────────────

/** 설정 파일이 저장되는 절대 경로 */
const settingsPath = path.join(app.getPath("userData"), "settings.json");

/**
 * 설정 파일을 읽어 객체로 반환한다.
 * 파일이 없거나 JSON 파싱 실패 시 빈 객체를 반환한다.
 *
 * @returns {Object} 설정 객체 (aiProvider, models, geminiApiKey, aiApiKeys 등 포함)
 */
function loadSettings() {
  try {
    if (!fs.existsSync(settingsPath)) {
      return {};
    }
    const raw = fs.readFileSync(settingsPath, "utf-8");
    return JSON.parse(raw);
  } catch (error) {
    console.error("설정 불러오기 실패:", error);
    return {};
  }
}

const AI_PROVIDERS = ["gemini", "openai", "claude"];
const DEFAULT_MODELS = {
  gemini: "gemini-2.5-flash",
  openai: "gpt-4o-mini",
  claude: "claude-3-5-haiku-latest"
};

const PROVIDER_LABELS = {
  gemini: "Gemini",
  openai: "GPT / OpenAI",
  claude: "Claude / Anthropic"
};

const PROVIDER_KEY_HINTS = {
  gemini: 'Gemini 키는 보통 "AIza"로 시작합니다. (aistudio.google.com/apikey 에서 발급)',
  openai: 'OpenAI 키는 "sk-"로 시작합니다. (platform.openai.com/api-keys 에서 발급)',
  claude: 'Claude 키는 "sk-ant-"로 시작합니다. (console.anthropic.com 에서 발급)'
};

/**
 * 붙여넣은 API 키를 정규화한다.
 * 메신저·메일로 키를 주고받으면 줄바꿈, 폭 없는 공백(U+200B 등), 따옴표,
 * "Bearer " 접두사가 함께 따라붙는 경우가 많다. 이런 문자가 남아 있으면
 * HTTP 헤더를 만드는 단계에서 바로 실패해 원인을 알 수 없는 오류로 보인다.
 *
 * @param {string} raw - 사용자가 입력/저장한 키 원문
 * @returns {string} 사용 가능한 형태로 정리된 키
 */
function sanitizeApiKey(raw) {
  let key = String(raw == null ? "" : raw);
  key = key.replace(/[\u200B-\u200D\u2060\uFEFF]/g, "");  // 폭 없는 공백류
  key = key.replace(/[\u0000-\u001F\u007F]/g, "");         // 제어문자
  key = key.replace(/[\s\u00A0]+/g, "");                    // 중간 공백·줄바꿈까지 제거
  key = key.replace(/^["'`]+|["'`]+$/g, "");                // 감싸인 따옴표
  key = key.replace(/^Bearer/i, "");
  return key.trim();
}

/**
 * 키 접두사로 어느 제공자의 키인지 추정한다.
 * @returns {string|null} 'gemini' | 'openai' | 'claude' | null(판단 불가)
 */
function detectKeyProvider(key) {
  const k = String(key || "");
  if (/^sk-ant-/i.test(k)) return "claude";
  if (/^AIza/.test(k))     return "gemini";
  if (/^sk-/i.test(k))     return "openai";
  return null;
}

/** 오류 객체에서 HTTP 상태코드를 최대한 뽑아낸다. */
function extractStatusCode(err) {
  const direct = Number(err?.status ?? err?.statusCode ?? err?.code);
  if (Number.isInteger(direct) && direct >= 400 && direct <= 599) return direct;
  const m = String(err?.message || err || "").match(/\b(400|401|403|404|408|409|413|422|429|500|502|503|504|529)\b/);
  return m ? Number(m[1]) : 0;
}

/** 특정 제공자의 키가 저장돼 있지 않을 때 쓰는 안내 오류 */
function missingKeyError(provider) {
  const label = PROVIDER_LABELS[provider] || provider;
  return new Error(
    `설정에 선택된 AI 제공자는 "${label}"인데 ${label} API 키가 저장되어 있지 않습니다.\n` +
    `설정 창에서 AI 제공자를 ${label}로 둔 상태로 키를 입력해 저장해 주세요.\n${PROVIDER_KEY_HINTS[provider] || ""}`
  );
}

/**
 * 제공자 API 오류를 "무엇이 잘못됐고 무엇을 바꿔야 하는지" 한국어로 풀어 준다.
 * 사용자마다 키·모델·결제 상태가 다르기 때문에 영문 원본 오류만 보여 주면
 * 키 문제인지 모델 이름 문제인지 결제/할당량 문제인지 구분할 수 없다.
 *
 * @param {string} provider - 'gemini' | 'openai' | 'claude'
 * @param {string} model    - 호출에 사용한 모델 이름
 * @param {Error}  err      - 원본 오류
 * @returns {Error} 해설이 붙은 새 오류
 */
function describeAIError(provider, model, err) {
  if (err && err.__aiDescribed) return err;

  const label  = PROVIDER_LABELS[provider] || provider;
  const hint   = PROVIDER_KEY_HINTS[provider] || "";
  const raw    = String(err?.message || err || "").trim();
  const lower  = raw.toLowerCase();
  const status = extractStatusCode(err);

  let head = "";

  if (/invalid header|header value|err_invalid_http_token|err_invalid_char/.test(lower)) {
    head = `${label} API 키에 사용할 수 없는 문자가 섞여 있습니다.\n` +
           `설정에서 키 칸을 비우고 키를 다시 복사해 붙여넣어 주세요.\n${hint}`;

  } else if (status === 401 ||
             /api key not valid|api_key_invalid|invalid api key|incorrect api key|invalid x-api-key|authentication_error|unauthorized|missing credentials|api key expired/.test(lower)) {
    head = `${label} API 키가 올바르지 않습니다.\n` +
           `· 설정의 "AI 제공자"가 ${label}인지 확인해 주세요. 다른 제공자의 키는 쓸 수 없습니다.\n` +
           `· ${hint}\n` +
           `· 키를 복사할 때 앞뒤가 잘리지 않았는지, 만료·삭제된 키가 아닌지 확인해 주세요.`;

  } else if (/not supported in your (country|region)|location is not supported|user location/.test(lower)) {
    head = `${label}를 현재 네트워크 위치에서 사용할 수 없습니다.\nVPN을 끄거나 다른 인터넷 회선에서 다시 시도해 주세요.`;

  } else if (status === 403 ||
             /has not been used in project|service_disabled|api is not enabled|permission_denied|forbidden|not authorized|api key not valid for/.test(lower)) {
    head = `${label} 키에 사용 권한이 없습니다.\n` +
           `· 키를 발급한 프로젝트/계정에서 해당 API가 사용 설정되어 있는지 확인해 주세요.\n` +
           `· 키에 IP·앱 사용 제한(restriction)을 걸어 두었다면 해제해 주세요.\n` +
           `· 조직(팀) 계정 키라면 관리자에게 API 사용 허용을 요청해 주세요.`;

  } else if (/insufficient_quota|credit balance|billing|payment/.test(lower)) {
    head = `${label} 계정에 사용 가능한 크레딧(결제)이 없습니다.\n` +
           `키를 새로 만들어도 결제 수단이나 크레딧이 없으면 호출이 거부됩니다.\n` +
           `· 해당 사이트에서 크레딧을 충전한 뒤 다시 시도하거나,\n` +
           `· 설정에서 AI 제공자를 무료 사용량이 있는 Gemini로 바꿔 보세요.`;

  } else if (status === 429 || /quota|rate limit|resource_exhausted|too many requests/.test(lower)) {
    head = `${label} 사용량(할당량) 한도를 초과했습니다.\n` +
           `· 1~2분 뒤 다시 시도해 주세요.\n` +
           `· 무료 등급은 분당·하루 요청 수가 적습니다. 한 번에 넣는 가사 양을 줄이거나 결제를 등록해 주세요.`;

  } else if (status === 404 ||
             /model not found|was not found|unknown model|does not exist|model_not_found|unsupported model|invalid model/.test(lower)) {
    head = `"${model || "(기본)"}" 모델을 ${label}에서 찾을 수 없습니다.\n` +
           `· 설정의 "모델 이름"을 비워 두면 기본 모델(${DEFAULT_MODELS[provider] || "-"})을 사용합니다.\n` +
           `· AI 제공자를 바꿀 때 이전 제공자의 모델 이름이 그대로 남아 있으면 이 오류가 납니다.`;

  } else if (status === 400) {
    head = `${label}가 요청을 거부했습니다(400).\n` +
           `· 설정의 "모델 이름"을 비우고(기본 모델) 다시 시도해 주세요.\n` +
           `· 입력한 키가 ${label} 키인지 확인해 주세요. ${hint}`;

  } else if (status >= 500 || /overloaded|unavailable|high demand|internal error/.test(lower)) {
    head = `${label} 서버가 일시적으로 불안정합니다. 잠시 뒤 다시 시도해 주세요.`;

  } else if (/fetch failed|enotfound|econnrefused|econnreset|etimedout|timeout|network|getaddrinfo|certificate|self signed|proxy/.test(lower)) {
    head = `${label} 서버에 연결할 수 없습니다.\n` +
           `인터넷 연결 상태와, 회사·학교·교회 네트워크의 방화벽 차단 여부를 확인해 주세요.`;

  } else {
    head = `${label} 호출에 실패했습니다.`;
  }

  const wrapped = new Error(
    `${head}\n\n[제공자: ${label} / 모델: ${model || "기본"}${status ? " / 상태: " + status : ""}]\n` +
    `[원본 오류: ${raw.slice(0, 400) || "알 수 없음"}]`
  );
  wrapped.__aiDescribed = true;
  wrapped.status   = status || undefined;
  wrapped.provider = provider;
  return wrapped;
}

// ── 제공자별 단일 호출 래퍼 ──
// 호출 코드가 여러 곳에 흩어져 있어 오류 처리가 제각각이었다.
// 아래 세 함수로 모아 두고, 실패는 항상 describeAIError()를 거쳐 나가게 한다.

/** Gemini 텍스트 생성 1회 */
async function callGeminiText(apiKey, model, prompt) {
  try {
    const ai = new GoogleGenAI({ apiKey });
    const resp = await ai.models.generateContent({ model, contents: prompt });
    return resp.text || "";
  } catch (err) {
    throw describeAIError("gemini", model, err);
  }
}

/**
 * OpenAI chat/completions 1회.
 * 일부 신형 모델은 temperature를 받지 않으므로, 그 때문에 400이 나면
 * temperature 없이 한 번 더 시도한다(사람마다 적어 넣는 모델 이름이 다르다).
 */
async function callOpenAIChat(apiKey, model, messages, options = {}) {
  const temperature = options.temperature === undefined ? 0.2 : options.temperature;

  const request = async (withTemp) => {
    const body = { model, messages };
    if (withTemp && temperature != null) body.temperature = temperature;

    const resp = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
      body: JSON.stringify(body)
    });

    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      const e = new Error(data?.error?.message || `HTTP ${resp.status}`);
      e.status = resp.status;
      e.param  = data?.error?.param || "";
      throw e;
    }
    return data?.choices?.[0]?.message?.content || "";
  };

  try {
    return await request(true);
  } catch (err) {
    const msg = String(err?.message || "").toLowerCase();
    if (err?.status === 400 && (err?.param === "temperature" || msg.includes("temperature"))) {
      try {
        return await request(false);
      } catch (retryErr) {
        throw describeAIError("openai", model, retryErr);
      }
    }
    throw describeAIError("openai", model, err);
  }
}

/** Claude messages 1회 */
async function callClaudeMessages(apiKey, model, prompt, maxTokens = 2048) {
  try {
    const resp = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01"
      },
      body: JSON.stringify({
        model,
        max_tokens: maxTokens,
        temperature: 0.2,
        messages: [{ role: "user", content: prompt }]
      })
    });

    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      const e = new Error(data?.error?.message || `HTTP ${resp.status}`);
      e.status = resp.status;
      throw e;
    }
    return Array.isArray(data?.content) ? data.content.map(part => part.text || "").join("\n") : "";
  } catch (err) {
    throw describeAIError("claude", model, err);
  }
}

function ensureSettingsShape(raw = {}) {
  const settings = (raw && typeof raw === "object") ? raw : {};
  settings.aiProvider = AI_PROVIDERS.includes(settings.aiProvider) ? settings.aiProvider : "gemini";
  settings.aiApiKeys = (settings.aiApiKeys && typeof settings.aiApiKeys === "object") ? settings.aiApiKeys : {};
  settings.models = (settings.models && typeof settings.models === "object") ? settings.models : {};

  for (const provider of AI_PROVIDERS) {
    settings.models[provider] = String(settings.models[provider] || "").trim();

    // 예전에 공백·줄바꿈이 섞인 채로 저장된 키도 여기서 한 번 정리한다
    if (settings.aiApiKeys[provider]) {
      settings.aiApiKeys[provider] = sanitizeApiKey(settings.aiApiKeys[provider]);
      if (!settings.aiApiKeys[provider]) delete settings.aiApiKeys[provider];
    }
  }

  if (settings.geminiApiKey) {
    settings.geminiApiKey = sanitizeApiKey(settings.geminiApiKey);
  }

  const mls = parseInt(settings.maxLinesPerSlide, 10);
  settings.maxLinesPerSlide = (Number.isInteger(mls) && mls >= 1 && mls <= 5) ? mls : DEFAULT_MAX_LINES_PER_SLIDE;

  // 구버전 단일 aiModel 값이 있으면 현재 제공자 슬롯으로 1회 마이그레이션
  if (settings.aiModel && !settings.models[settings.aiProvider]) {
    settings.models[settings.aiProvider] = String(settings.aiModel).trim();
  }
  delete settings.aiModel;

  return settings;
}

/**
 * 설정 객체를 JSON 파일로 저장한다.
 *
 * @param {Object} settings - 저장할 설정 객체
 * @returns {boolean} 저장 성공 여부
 */
function saveSettings(settings) {
  try {
    fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 2), "utf-8");
    return true;
  } catch (error) {
    console.error("설정 저장 실패:", error);
    return false;
  }
}

// ──────────────────────────────────────────────────────────
// IPC: 설정 조회/저장
// ──────────────────────────────────────────────────────────

/**
 * 현재 설정을 렌더러에 전달한다.
 * 보안상 API 키 원문은 보내지 않고 "저장 여부"만 전달한다.
 *
 * 반환값: { hasGeminiApiKey, aiProvider, aiModel, hasAIKey }
 */
ipcMain.handle("settings:get", async () => {
  const settings  = ensureSettingsShape(loadSettings());
  const aiProvider = settings.aiProvider || "gemini";

  // 설정 화면이 제공자를 바꿀 때마다 IPC를 다시 부르지 않도록
  // 제공자별 "키 저장 여부"와 모델 이름을 한꺼번에 내려 준다.
  const hasKeys = {};
  const models  = {};
  for (const provider of AI_PROVIDERS) {
    hasKeys[provider] = !!getAIKey(settings, provider);
    models[provider]  = settings.models?.[provider] || "";
  }

  return {
    hasGeminiApiKey: !!settings.geminiApiKey,
    aiProvider,
    aiModel: settings.models?.[aiProvider] || "",
    hasAIKey: hasKeys[aiProvider],
    hasKeys,
    models,
    defaultModels: { ...DEFAULT_MODELS },
    maxLinesPerSlide: settings.maxLinesPerSlide || DEFAULT_MAX_LINES_PER_SLIDE
  };
});

/**
 * 앱 버전(package.json의 version)을 렌더러에 전달한다.
 * 설정 모달 하단에 표시하는 용도.
 */
ipcMain.handle("app:get-version", async () => {
  try {
    return app.getVersion();
  } catch (e) {
    return "";
  }
});

/**
 * AI 제공자·모델·API 키를 통합 저장한다.
 *
 * @param {Object} data
 *   data.provider - AI 제공자 ('gemini' | 'openai' | 'claude')
 *   data.model    - 모델명 (빈 문자열이면 제공자별 기본 모델 사용)
 *   data.apiKey   - API 키 (빈 문자열이면 기존 키 유지)
 */
async function saveAISettingsCore(data) {
  try {
    const provider = String(data?.provider || "gemini").trim();
    const model    = String(data?.model    || "").trim();
    const rawKey   = String(data?.apiKey   || "");
    const apiKey   = sanitizeApiKey(rawKey);

    if (!AI_PROVIDERS.includes(provider)) {
      return { success: false, error: "지원하지 않는 AI 제공자입니다." };
    }

    if (rawKey.trim() && !apiKey) {
      return {
        success: false,
        error: "API 키에서 사용할 수 있는 문자를 찾지 못했습니다. 키를 다시 복사해 붙여넣어 주세요."
      };
    }

    // 다른 제공자의 키를 그대로 저장해 두면 호출할 때마다 인증 오류가 나는데,
    // 설정 화면에는 "키가 저장되어 있다"고만 보여서 원인을 찾기 어렵다.
    // 그래서 접두사로 제공자가 분명히 다른 키는 저장 단계에서 막는다.
    const detected = apiKey ? detectKeyProvider(apiKey) : null;
    if (detected && detected !== provider) {
      return {
        success: false,
        error: `입력한 키는 ${PROVIDER_LABELS[detected]} 키 형식입니다. 지금 선택된 AI 제공자는 ${PROVIDER_LABELS[provider]}입니다.\n\n` +
               `· ${PROVIDER_LABELS[detected]}를 쓰려면 "AI 제공자"를 ${PROVIDER_LABELS[detected]}로 바꾼 뒤 저장해 주세요.\n` +
               `· ${PROVIDER_LABELS[provider]}를 쓰려면 ${PROVIDER_KEY_HINTS[provider]}`
      };
    }

    const settings = ensureSettingsShape(loadSettings());
    settings.aiProvider = provider;
    settings.models[provider] = model;

    if (apiKey) {
      settings.aiApiKeys[provider] = apiKey;
      if (provider === "gemini") settings.geminiApiKey = apiKey;
    } else if (provider === "gemini" && settings.aiApiKeys.gemini) {
      // 구버전 geminiApiKey 필드와 aiApiKeys 맵이 어긋나 있으면 맞춰 둔다
      settings.geminiApiKey = settings.aiApiKeys.gemini;
    }

    const saved = saveSettings(settings);
    if (!saved) return { success: false, error: "AI 설정 파일 저장에 실패했습니다." };

    const hasKey = !!getAIKey(settings, provider);

    return {
      success: true,
      hasKey,
      warning: hasKey
        ? ""
        : `${PROVIDER_LABELS[provider]} API 키가 아직 저장되지 않았습니다. AI 기능을 쓰려면 ${PROVIDER_LABELS[provider]} 키를 입력해 주세요.\n${PROVIDER_KEY_HINTS[provider]}`
    };
  } catch (error) {
    console.error("AI 설정 저장 실패:", error);
    return { success: false, error: "AI 설정 저장 중 오류가 발생했습니다." };
  }
}

ipcMain.handle("settings:set-ai-settings", async (event, data) => {
  return await saveAISettingsCore(data);
});

/**
 * 입력(또는 저장)된 API 키로 실제 호출을 한 번 해 보고 결과를 알려 준다.
 * 가사 분리까지 가 보지 않아도 "키가 문제인지, 모델 이름이 문제인지,
 * 결제/할당량이 문제인지"를 설정 화면에서 바로 확인할 수 있게 한다.
 *
 * @param {Object} data
 *   data.provider - 'gemini' | 'openai' | 'claude'
 *   data.model    - 모델 이름 (빈 값이면 저장값 → 기본 모델)
 *   data.apiKey   - 테스트할 키 (빈 값이면 저장된 키 사용)
 */
async function testAIKeyCore(data) {
  const provider = String(data?.provider || "gemini").trim();
  if (!AI_PROVIDERS.includes(provider)) {
    return { success: false, error: "지원하지 않는 AI 제공자입니다." };
  }

  const label    = PROVIDER_LABELS[provider];
  const settings = ensureSettingsShape(loadSettings());
  const typedKey = sanitizeApiKey(data?.apiKey);
  const apiKey   = typedKey || getAIKey(settings, provider);

  if (!apiKey) {
    return {
      success: false,
      error: `테스트할 ${label} API 키가 없습니다. 키를 입력한 뒤 다시 눌러 주세요.\n${PROVIDER_KEY_HINTS[provider]}`
    };
  }

  const detected = detectKeyProvider(apiKey);
  if (detected && detected !== provider) {
    return {
      success: false,
      error: `입력한 키는 ${PROVIDER_LABELS[detected]} 키 형식입니다.\n` +
             `"AI 제공자"를 ${PROVIDER_LABELS[detected]}로 바꾸거나, ${label} 키를 입력해 주세요.\n${PROVIDER_KEY_HINTS[provider]}`
    };
  }

  const model = String(data?.model || "").trim() || settings.models[provider] || DEFAULT_MODELS[provider];

  try {
    if (provider === "gemini") {
      await callGeminiText(apiKey, model, "ping");
    } else if (provider === "openai") {
      await callOpenAIChat(apiKey, model, [{ role: "user", content: "ping" }], { temperature: null });
    } else {
      await callClaudeMessages(apiKey, model, "ping", 16);
    }
    return { success: true, message: `${label} 연결 성공 ✓ (모델: ${model})` };
  } catch (error) {
    return { success: false, error: String(error?.message || error) };
  }
}

ipcMain.handle("settings:test-ai-key", async (event, data) => {
  return await testAIKeyCore(data);
});

// ──────────────────────────────────────────────────────────
// 시스템 폰트 목록 조회
// Windows: PowerShell의 System.Drawing.FontFamily로 설치된 폰트 이름 열거
// ──────────────────────────────────────────────────────────
ipcMain.handle("fonts:list-system", async () => {
  try {
    const { execSync } = require("child_process");
    if (process.platform === "win32") {
      const ps = `[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | Out-Null; [System.Drawing.FontFamily]::Families | Select-Object -ExpandProperty Name`;
      const result = execSync(`powershell -NoProfile -Command "${ps}"`, { encoding: "utf8", timeout: 15000 });
      const fonts = result.split("\n").map(f => f.trim()).filter(f => f.length > 0).sort((a, b) => a.localeCompare(b, "ko"));
      return { ok: true, fonts };
    }
    if (process.platform === "darwin") {
      const result = execSync(`fc-list : family --format="%{family}\n" 2>/dev/null || system_profiler SPFontsDataType 2>/dev/null | grep "Full Name:" | awk -F": " '{print $2}'`, { encoding: "utf8", timeout: 15000, shell: "/bin/bash" });
      const fonts = result.split("\n").map(f => f.trim()).filter(f => f.length > 0).sort((a, b) => a.localeCompare(b, "ko"));
      return { ok: true, fonts };
    }
    return { ok: true, fonts: [] };
  } catch (e) {
    console.error("시스템 폰트 목록 조회 실패:", e.message);
    return { ok: false, fonts: [] };
  }
});

// ──────────────────────────────────────────────────────────
// AI 가사 분리 / 번역 공통 유틸
// ──────────────────────────────────────────────────────────

/**
 * AI에게 전달할 가사 분리 프롬프트를 생성한다.
 * translateTo에 언어 코드('en', 'ja')를 포함하면 번역도 함께 요청한다.
 *
 * 핵심 규칙:
 *   - 한 슬라이드의 lines 배열은 최대 maxLinesPerSlide개 (설정 가능)
 *   - 원본 가사를 절대 생략/수정/재창작하지 않음
 *   - 후렴/절 표기, 제목 등도 포함
 *   - 번역 요청 시 translations.en, translations.ja 배열로 반환
 *
 * @param {string}   lyrics      - 원본 가사 텍스트
 * @param {string[]} translateTo - 번역 대상 언어 코드 배열 (예: ['en', 'ja'])
 * @returns {string} 완성된 프롬프트 문자열
 */
function getLyricsSplitPrompt(lyrics, translateTo = [], maxLines = 1) {
  const translationTargets   = Array.isArray(translateTo) ? translateTo : [];
  const translationLanguages = translationTargets
    .map(lang => lang === "en" ? "영어(en)" : lang === "ja" ? "일본어(ja)" : lang)
    .join(", ");

  const safeMax = (Number.isInteger(maxLines) && maxLines >= 1) ? maxLines : 1;

  const linesRule = safeMax === 1
    ? `- 각 슬라이드의 lines 배열에는 원소를 반드시 1개만 넣어라. 절대 2개 이상 넣지 마라.
- 한 문장(한 줄)이 하나의 슬라이드가 된다.
- 여러 문장을 하나의 슬라이드로 합치지 마라. 반드시 문장마다 별도 슬라이드로 분리해라.`
    : `- 각 슬라이드의 lines 배열에는 원소를 최대 ${safeMax}개까지 넣어라.
- 의미상 자연스럽게 이어지는 문장들을 최대 ${safeMax}줄씩 하나의 슬라이드로 묶어라.
- 한 줄로도 되고, 어색하지 않으면 최대 ${safeMax}줄까지 함께 넣어도 된다.`;

  const translationLinesRule = safeMax === 1
    ? `각 배열에 원소를 반드시 1개만 넣어라`
    : `각 배열에 원소를 최대 ${safeMax}개까지 넣어라`;

  let prompt = `너는 예배용 슬라이드 송출 앱의 가사 편집 도우미야.

아래 가사를 예배 송출 화면에 넣기 좋게 슬라이드 단위로 나눠줘.

규칙:
- 반드시 JSON 배열만 반환해.
- 설명문, 마크다운 코드블록, 주석 없이 JSON 배열만 반환한다.
- 입력된 내용을 빠뜨리지 말고 순서를 유지해서 JSON 배열로 반환한다.
${linesRule}
- 의미가 어색하게 끊기지 않게 나눠.
- 빈 줄은 무시해.
- 가사 맨 앞에 <제목> 형태로 꺽쇠(<>)로 감싼 텍스트가 있으면, 그것은 노래 제목이다. 반드시 JSON 배열의 첫 번째 슬라이드로 넣어라. 꺽쇠 기호(<>)는 제거하고 안의 텍스트만 lines에 넣어라.
- 사용자가 붙여넣은 내용 중 제목, 곡명, 절 표시, 후렴, 브릿지, 엔딩, 반복 표시, 괄호 안 문구도 의미 있는 내용이면 절대 생략하지 않는다.
- 사용자가 입력한 문장을 요약하거나 삭제하지 않는다.
- 찬양 가사가 아닌 것처럼 보여도 사용자가 붙여넣은 텍스트는 가능한 한 모두 슬라이드에 포함한다.
- 오타를 임의로 고치거나 가사를 새로 창작하지 않는다.
- 중복으로 보이는 가사도 사용자가 입력했다면 임의로 제거하지 않는다.
- 후렴, 1절, 2절 같은 표기는 가사 내용에서 삭제하지 말고, 사용자가 입력한 순서 그대로 하나의 줄 또는 슬라이드로 포함한다.
- 너의 역할은 입력된 가사를 슬라이드 단위로 "나누는 것"뿐이다. 글자, 띄어쓰기, 문장부호를 절대 추가·삭제·변경하지 말고 원문을 한 글자도 바꾸지 않고 그대로 사용한다.
- 특히 각 줄의 끝이나 마지막 슬라이드 끝에 마침표(.)나 쉼표(,) 같은 문장부호를 임의로 붙이지 마라. 원문에 없으면 절대 넣지 않는다.
`;

  // 번역 요청 시 각 슬라이드 객체에 translations 필드 추가
  if (translationLanguages) {
    prompt += `
번역 추가 규칙:
- 각 슬라이드 객체에 translations 객체를 반드시 포함한다.
- 선택된 언어만 translations 안에 넣는다: ${translationLanguages}
- 영어는 translations.en 배열에 넣는다.
- 일본어는 translations.ja 배열에 넣는다.
- 번역도 원본 lines와 같은 의미를 유지하고 ${translationLinesRule}.
- 번역문의 문장부호는 원문(한국어)과 똑같이 맞춘다. 원문 줄 끝에 마침표(.)나 쉼표(,)가 없으면 영어·일본어 번역 줄 끝에도 절대 붙이지 마라.
- 영어의 마침표(.), 쉼표(,), 일본어의 구점(。), 독점(、), 물음표(?/？), 느낌표(!/！) 모두 해당한다. 원문에 없으면 번역에도 넣지 않는다.
- 반환 형식은 반드시 아래와 같아.

[
  {
    "lines": ["원본 줄 1"],
    "translations": {
      "en": ["English line 1"],
      "ja": ["日本語 1行目"]
    }
  }
]
`;
  } else {
    prompt += `
반환 형식은 반드시 아래와 같아.

[
  {
    "lines": ["슬라이드 내용 1"]
  },
  {
    "lines": ["다음 슬라이드 내용"]
  }
]
`;
  }

  prompt += `
사용자가 붙여넣은 원문:

${lyrics}
`;

  return prompt;
}

/**
 * AI 응답에서 마크다운 코드 블록(```json ... ```)을 제거한다.
 * 일부 AI가 JSON을 코드 블록으로 감싸서 반환하는 경우를 처리한다.
 *
 * @param {string} text - AI 응답 원문
 * @returns {string} 순수 JSON 텍스트
 */
function cleanAIJsonText(text) {
  return String(text || "")
    .replace(/```json/g, "")
    .replace(/```/g, "")
    .trim();
}

function normalizeLines(lines, maxLines = DEFAULT_MAX_LINES_PER_SLIDE) {
  return (Array.isArray(lines) ? lines : [])
    .map(line => String(line || "").trim().replace(/^<(.+)>$/, "$1"))
    .filter(Boolean)
    .slice(0, maxLines);
}

/**
 * AI가 돌려준 슬라이드의 텍스트를 사용자가 입력한 원문 그대로 복원한다.
 *
 * AI에게는 가사를 슬라이드로 "나누기만" 하도록 지시하지만, 모델이 임의로
 * 맞춤법·띄어쓰기·문장부호(특히 마지막 줄 끝의 마침표)를 바꾸는 경우가 있다.
 * 분리된 줄 수가 원문 줄 수와 같고, 공백·문장부호를 뺀 내용이 순서대로 모두 일치하면
 * 각 줄을 원문 텍스트로 되돌려, 가사 자체는 절대 수정되지 않도록 보장한다.
 *
 * 줄 수가 다르거나(병합/분할/누락) 내용이 어긋나면 안전하게 AI 결과를 그대로 둔다.
 *
 * @param {Array}  slides       - 파싱된 슬라이드 배열 (입력 배열을 직접 수정한다)
 * @param {string} originalText - 사용자가 입력한 원문 (없으면 복원하지 않음)
 * @param {string} key          - 복원할 필드명 ('lines' | 'kr' | 'en' | 'ja')
 * @returns {Array} 원문이 복원된 슬라이드 배열
 */
function restoreOriginalLines(slides, originalText, key) {
  if (!Array.isArray(slides) || !originalText || !String(originalText).trim()) {
    return slides;
  }

  // 원문을 정규화된 줄 목록으로 변환 (normalizeLines와 동일 규칙: trim, <제목> 꺽쇠 제거, 빈 줄 제외)
  const originalLines = String(originalText)
    .split(/\r?\n/)
    .map(line => String(line || "").trim().replace(/^<(.+)>$/, "$1"))
    .filter(Boolean);

  // 슬라이드의 해당 필드 줄들을 순서대로 펼친다
  const refs = [];
  for (const slide of slides) {
    const arr = slide && Array.isArray(slide[key]) ? slide[key] : null;
    if (!arr) continue;
    for (let i = 0; i < arr.length; i++) refs.push({ arr, i });
  }

  // 줄 수가 다르면 AI가 줄을 합치거나 나눈 것이므로 그대로 둔다
  if (refs.length !== originalLines.length) return slides;

  // 공백과 문장부호·기호를 제외하고 내용이 순서대로 모두 일치하는지 확인한다
  const stripForCompare = s => String(s).replace(/[\s\p{P}\p{S}]/gu, "");
  const allMatch = refs.every((ref, idx) =>
    stripForCompare(ref.arr[ref.i]) === stripForCompare(originalLines[idx]));
  if (!allMatch) return slides;

  // 모두 일치하면 원문 텍스트로 복원한다 (모델이 덧붙인 마침표 등 모든 변형 제거)
  refs.forEach((ref, idx) => { ref.arr[ref.i] = originalLines[idx]; });
  return slides;
}

// 줄 끝에 붙는 문장부호(마침표·쉼표·물음표·느낌표 등). 한글/영어/일본어 표기를 모두 포함한다.
// 예: "Praise the Lord." / "主をほめたたえよ。" / "주님을 찬양해!"
const TRAILING_PUNCT_RE = /[.。．…‥,，、;；:：!！?？]+[\s"'”’」』)\]]*$/u;

/** 줄 끝에 문장부호가 붙어 있는지 확인한다. */
function hasTrailingPunct(line) {
  return TRAILING_PUNCT_RE.test(String(line || "").trim());
}

/** 줄 끝의 문장부호를 제거한다. 닫는 따옴표·괄호는 그대로 남긴다. */
function stripTrailingPunct(line) {
  const text = String(line || "").trim();
  return text
    .replace(TRAILING_PUNCT_RE, (matched) => matched.replace(/[.。．…‥,，、;；:：!！?？\s]/gu, ""))
    .trim();
}

/**
 * 번역 줄의 끝 문장부호를 원문 줄에 맞춘다.
 *
 * 원문(한국어) 가사에 마침표가 없는데 AI가 번역문 끝에만 마침표(.)나 「。」를
 * 붙여 오는 경우가 많다. 프롬프트로도 막지만 모델이 지키지 않을 수 있으므로
 * 응답을 받은 뒤 코드에서 한 번 더 정리해, 어떤 AI 제공자를 쓰더라도 결과가 같도록 한다.
 *
 * 규칙:
 *   - 원문 줄 끝에 문장부호가 없으면 → 번역 줄 끝의 문장부호를 제거한다.
 *   - 원문 줄 끝에 문장부호가 있으면 → 번역문은 손대지 않는다.
 *   - 원문과 번역의 줄 수가 다르면(병합/분할) 슬라이드 안의 원문 줄 중 하나라도
 *     문장부호가 있을 때만 번역을 그대로 둔다 (안전한 쪽으로 판단).
 *
 * @param {string[]} sourceLines - 기준이 되는 원문 줄 배열
 * @param {string[]} targetLines - 정리할 번역 줄 배열
 * @returns {string[]} 문장부호가 정리된 번역 줄 배열
 */
function matchPunctuationToSource(sourceLines, targetLines) {
  const source = Array.isArray(sourceLines) ? sourceLines : [];
  const target = Array.isArray(targetLines) ? targetLines : [];
  if (!source.length || !target.length) return target;

  const perLine        = source.length === target.length;
  const slideHasPunct  = source.some(hasTrailingPunct);

  return target.map((line, idx) => {
    const sourceHasPunct = perLine ? hasTrailingPunct(source[idx]) : slideHasPunct;
    if (sourceHasPunct) return line;

    const stripped = stripTrailingPunct(line);
    // 부호만으로 이루어진 줄이면 빈 줄이 되어 슬라이드가 어긋나므로 원본을 유지한다
    return stripped || line;
  });
}

/**
 * 슬라이드 배열의 번역 필드를 원문 필드의 문장부호에 맞춰 정리한다.
 *
 * @param {Array}    slides     - 슬라이드 배열 (입력 배열을 직접 수정한다)
 * @param {Function} getSource  - 슬라이드에서 원문 줄 배열을 꺼내는 함수
 * @param {Function} getTarget  - 슬라이드에서 번역 줄 배열을 꺼내는 함수
 * @param {Function} setTarget  - 정리된 번역 줄 배열을 슬라이드에 다시 넣는 함수
 * @returns {Array} 정리된 슬라이드 배열
 */
function applyPunctuationMatch(slides, getSource, getTarget, setTarget) {
  if (!Array.isArray(slides)) return slides;

  for (const slide of slides) {
    if (!slide) continue;
    const target = getTarget(slide);
    if (!Array.isArray(target) || !target.length) continue;
    setTarget(slide, matchPunctuationToSource(getSource(slide), target));
  }

  return slides;
}

/**
 * AI 응답 텍스트를 슬라이드 배열로 파싱한다.
 * 잘못된 형식의 항목은 필터링하고 최대 2줄까지만 허용한다.
 *
 * @param {string}   text         - AI 응답 텍스트
 * @param {string}   providerName - 오류 메시지에 표시할 AI 제공자명
 * @param {string[]} translateTo  - 번역 언어 코드 배열
 * @returns {Array}  슬라이드 배열. 각 항목: { lines: string[], translations: {} }
 * @throws {Error} JSON 파싱 실패 또는 빈 결과 시
 */
function parseLyricsSlidesFromAI(text, providerName = "AI", translateTo = [], maxLines = DEFAULT_MAX_LINES_PER_SLIDE, originalLyrics = "") {
  const cleaned = cleanAIJsonText(text);
  let parsed;

  try {
    parsed = JSON.parse(cleaned);
  } catch (parseError) {
    console.error(providerName + " 응답 JSON 파싱 실패:", cleaned);
    throw new Error(providerName + " 응답을 슬라이드 형식으로 읽지 못했습니다. 다시 시도해 주세요.");
  }

  if (!Array.isArray(parsed)) {
    throw new Error(providerName + " 응답이 배열 형식이 아닙니다.");
  }

  const targets = Array.isArray(translateTo) ? translateTo : [];

  const validSlides = parsed
    .filter(item => item && Array.isArray(item.lines))
    .map(item => {
      const lines = normalizeLines(item.lines, maxLines);

      const translations = {};

      if (targets.includes("en") && Array.isArray(item.translations?.en)) {
        translations.en = normalizeLines(item.translations.en, maxLines);
      }

      if (targets.includes("ja") && Array.isArray(item.translations?.ja)) {
        translations.ja = normalizeLines(item.translations.ja, maxLines);
      }

      return { lines, translations };
    })
    .filter(item => item.lines.length > 0);

  if (validSlides.length === 0) {
    throw new Error(providerName + "가 사용할 수 있는 슬라이드 결과를 만들지 못했습니다.");
  }

  // 분리된 가사를 원문 그대로 복원한다 (모델이 끝에 붙인 마침표 등 임의 수정 제거)
  restoreOriginalLines(validSlides, originalLyrics, "lines");

  // 번역문 끝의 문장부호를 원문(한국어)에 맞춘다.
  // 원문에 마침표가 없으면 영어·일본어 번역에서도 제거한다.
  for (const lang of ["en", "ja"]) {
    if (!targets.includes(lang)) continue;
    applyPunctuationMatch(
      validSlides,
      slide => slide.lines,
      slide => slide.translations?.[lang],
      (slide, lines) => { slide.translations[lang] = lines; }
    );
  }

  return validSlides;
}

/**
 * 설정에서 특정 AI 제공자의 API 키를 꺼낸다.
 * aiApiKeys 맵 → geminiApiKey(Gemini 전용 구버전 필드) 순서로 탐색한다.
 *
 * @param {Object} settings  - loadSettings() 결과
 * @param {string} provider  - 'gemini' | 'openai' | 'claude'
 * @returns {string} API 키 문자열 (없으면 빈 문자열)
 */
function getAIKey(settings, provider) {
  if (settings.aiApiKeys && settings.aiApiKeys[provider]) {
    return sanitizeApiKey(settings.aiApiKeys[provider]);
  }
  if (provider === "gemini") {
    return sanitizeApiKey(settings.geminiApiKey || "");
  }
  return "";
}

function getModelForProvider(settings, provider) {
  const normalized = ensureSettingsShape(settings);
  return normalized.models[provider] || DEFAULT_MODELS[provider] || "";
}

// ── AI 제공자별 가사 분리 함수 ──

/**
 * Gemini API로 가사를 슬라이드 배열로 분리한다.
 *
 * @param {string}   lyrics      - 원본 가사
 * @param {Object}   settings    - 설정 객체
 * @param {string[]} translateTo - 번역 대상 언어 코드 배열
 * @returns {Promise<Array>} 슬라이드 배열
 */
async function splitLyricsWithGemini(lyrics, settings = loadSettings(), translateTo = []) {
  const inputLyrics = String(lyrics || "").trim();
  if (!inputLyrics) throw new Error("가사 내용이 비어 있습니다.");

  const geminiApiKey = getAIKey(settings, "gemini");
  if (!geminiApiKey) throw missingKeyError("gemini");

  const maxLines = ensureSettingsShape(settings).maxLinesPerSlide;
  const model    = getModelForProvider(settings, "gemini");
  const text     = await callGeminiText(geminiApiKey, model, getLyricsSplitPrompt(inputLyrics, translateTo, maxLines));

  return parseLyricsSlidesFromAI(text, "Gemini", translateTo, maxLines, inputLyrics);
}

/**
 * OpenAI API(ChatGPT)로 가사를 슬라이드 배열로 분리한다.
 *
 * @param {string}   lyrics      - 원본 가사
 * @param {Object}   settings    - 설정 객체
 * @param {string[]} translateTo - 번역 대상 언어 코드 배열
 * @returns {Promise<Array>} 슬라이드 배열
 */
async function splitLyricsWithOpenAI(lyrics, settings = loadSettings(), translateTo = []) {
  const inputLyrics = String(lyrics || "").trim();
  if (!inputLyrics) throw new Error("가사 내용이 비어 있습니다.");

  const apiKey = getAIKey(settings, "openai");
  if (!apiKey) throw missingKeyError("openai");

  const model    = getModelForProvider(settings, "openai");
  const maxLines = ensureSettingsShape(settings).maxLinesPerSlide;

  const text = await callOpenAIChat(apiKey, model, [
    { role: "system", content: "You return only valid JSON. No markdown." },
    { role: "user",   content: getLyricsSplitPrompt(inputLyrics, translateTo, maxLines) }
  ]);

  return parseLyricsSlidesFromAI(text, "OpenAI", translateTo, maxLines, inputLyrics);
}

/**
 * Claude API(Anthropic)로 가사를 슬라이드 배열로 분리한다.
 *
 * @param {string}   lyrics      - 원본 가사
 * @param {Object}   settings    - 설정 객체
 * @param {string[]} translateTo - 번역 대상 언어 코드 배열
 * @returns {Promise<Array>} 슬라이드 배열
 */
async function splitLyricsWithClaude(lyrics, settings = loadSettings(), translateTo = []) {
  const inputLyrics = String(lyrics || "").trim();
  if (!inputLyrics) throw new Error("가사 내용이 비어 있습니다.");

  const apiKey = getAIKey(settings, "claude");
  if (!apiKey) throw missingKeyError("claude");

  const model    = getModelForProvider(settings, "claude");
  const maxLines = ensureSettingsShape(settings).maxLinesPerSlide;

  const text = await callClaudeMessages(
    apiKey,
    model,
    getLyricsSplitPrompt(inputLyrics, translateTo, maxLines),
    2048
  );

  return parseLyricsSlidesFromAI(text, "Claude", translateTo, maxLines, inputLyrics);
}

/**
 * 설정에서 선택된 AI 제공자에게 가사 분리를 요청한다.
 * 제공자 선택 로직을 한 곳에 모아 IPC 핸들러들이 공유한다.
 *
 * @param {string}   lyrics      - 원본 가사
 * @param {Object}   settings    - 설정 객체
 * @param {string[]} translateTo - 번역 대상 언어 코드 배열
 * @returns {Promise<Array>} 슬라이드 배열
 */
async function splitLyricsBySelectedProvider(lyrics, settings, translateTo = [], maxLinesOverride = null) {
  const provider      = settings.aiProvider || "gemini";
  const safeTranslate = Array.isArray(translateTo) ? translateTo : [];

  // maxLinesOverride가 유효한 숫자면 settings 값을 덮어쓴다
  const effectiveSettings = (maxLinesOverride != null && Number.isInteger(maxLinesOverride) && maxLinesOverride >= 1)
    ? { ...settings, maxLinesPerSlide: maxLinesOverride }
    : settings;

  if (provider === "gemini")  return await splitLyricsWithGemini(lyrics, effectiveSettings, safeTranslate);
  if (provider === "openai")  return await splitLyricsWithOpenAI(lyrics, effectiveSettings, safeTranslate);
  if (provider === "claude")  return await splitLyricsWithClaude(lyrics, effectiveSettings, safeTranslate);

  throw new Error("지원하지 않는 AI 제공자입니다: " + provider);
}

// IPC 핸들러: 구버전 호환용 (Gemini 고정)
ipcMain.handle("gemini:split-lyrics", async (event, lyrics) => {
  return await splitLyricsWithGemini(lyrics, loadSettings(), []);
});

// IPC 핸들러: 설정된 AI 제공자로 가사 분리 (번역 없음)
ipcMain.handle("ai:split-lyrics", async (event, lyrics, opts) => {
  const maxLines = opts?.maxLinesPerSlide != null ? parseInt(opts.maxLinesPerSlide, 10) : null;
  return await splitLyricsBySelectedProvider(lyrics, loadSettings(), [], maxLines);
});

// IPC 핸들러: 설정된 AI 제공자로 가사 분리 + 번역
// options.translateTo 에 언어 코드 배열, options.maxLinesPerSlide 에 줄 수를 넘긴다
ipcMain.handle("ai:split-lyrics-with-translations", async (event, lyrics, options) => {
  const translateTo = Array.isArray(options?.translateTo) ? options.translateTo : [];
  const maxLines = options?.maxLinesPerSlide != null ? parseInt(options.maxLinesPerSlide, 10) : null;
  return await splitLyricsBySelectedProvider(lyrics, loadSettings(), translateTo, maxLines);
});

// ──────────────────────────────────────────────────────────
// 다언어 가사 동시 분할
// 한국어/영어/일본어 가사를 동시에 입력받아 같은 슬라이드 수로 분할한다.
// ──────────────────────────────────────────────────────────

/**
 * 다언어 가사를 동시에 분할하는 AI 프롬프트를 생성한다.
 * 모든 언어의 슬라이드 수를 동일하게 맞추도록 지시한다.
 *
 * @param {string}  kr    - 한국어 가사 (필수)
 * @param {string}  en    - 영어 가사 (없으면 genEn으로 생성 여부 결정)
 * @param {string}  ja    - 일본어 가사 (없으면 genJa로 생성 여부 결정)
 * @param {boolean} genEn - 영어 가사가 없을 때 AI가 번역 생성 여부
 * @param {boolean} genJa - 일본어 가사가 없을 때 AI가 번역 생성 여부
 * @returns {string} 완성된 프롬프트
 */
function getMultiLangSplitPrompt(kr, en, ja, genEn=false, genJa=false, maxLines=null) {
  const safeMax = (Number.isInteger(maxLines) && maxLines >= 1) ? maxLines : 1;
  const hasEn  = !!(en && en.trim());
  const hasJa  = !!(ja && ja.trim());
  const needEn = hasEn || genEn;
  const needJa = hasJa || genJa;

  const linesRule = safeMax === 1
    ? `- 각 슬라이드의 배열 원소는 반드시 1개만. 절대 2개 이상 넣지 마라.
- 한 문장(한 줄)이 하나의 슬라이드가 된다. 여러 문장을 합치지 마라.
- 반드시 문장마다 별도 슬라이드로 분리해라.`
    : `- 각 슬라이드의 배열 원소는 최대 ${safeMax}개까지 넣어라.
- 의미상 자연스럽게 이어지는 문장들을 최대 ${safeMax}줄씩 하나의 슬라이드로 묶어라.
- 모든 언어의 슬라이드 수를 반드시 동일하게 맞춰.`;

  // AI가 직접 번역문을 만드는 경우에만 문장부호를 한국어 원문에 맞추도록 지시한다
  const aiGeneratesTranslation = (!hasEn && genEn) || (!hasJa && genJa);
  const punctuationRule = aiGeneratesTranslation
    ? `- 새로 번역해서 만드는 언어는 문장부호를 한국어 원문과 똑같이 맞춘다. 한국어 줄 끝에 마침표(.)가 없으면 영어 줄 끝의 마침표(.)도, 일본어 줄 끝의 구점(。)도 절대 붙이지 마라.
- 쉼표(,/、), 물음표(?/？), 느낌표(!/！)도 마찬가지다. 한국어 원문에 없으면 번역에도 넣지 않는다.
`
    : "";

  let prompt = `너는 예배용 슬라이드 앱의 가사 편집 도우미야.

아래 가사를 슬라이드 단위로 나눠줘.

규칙:
- 반드시 JSON 배열만 반환해. 마크다운 코드블록 없음.
${linesRule}
- 빈 줄은 무시해.
- 가사 맨 앞에 <제목> 형태로 꺽쇠(<>)로 감싼 텍스트가 있으면, 그것은 노래 제목이다. 반드시 JSON 배열의 첫 번째 슬라이드로 넣어라. 꺽쇠 기호(<>)는 제거하고 안의 텍스트만 넣어라.
- 내용을 생략, 요약, 수정하지 마.
- 입력된 가사의 글자·띄어쓰기·문장부호를 절대 추가·삭제·변경하지 마라. 원문을 그대로 두고 슬라이드로 "나누기만" 한다.
- 특히 각 줄의 끝이나 마지막 슬라이드 끝에 마침표(.) 등 문장부호를 임의로 붙이지 마라. 원문에 없으면 넣지 않는다.
- 모든 언어의 슬라이드 수를 반드시 동일하게 맞춰.
${punctuationRule}${needEn || needJa ? "- 영어/일본어는 한국어에 의미적으로 대응하는 내용으로 맞춰." : ""}

한국어:
${kr}
`;

  if (hasEn)      prompt += `\n영어 (아래 텍스트를 한국어와 같은 수의 슬라이드로 분할):\n${en}\n`;
  else if (genEn) prompt += `\n영어: (한국어에서 자연스럽게 번역해서 생성)\n`;

  if (hasJa)      prompt += `\n일본어 (아래 텍스트를 한국어와 같은 수의 슬라이드로 분할):\n${ja}\n`;
  else if (genJa) prompt += `\n일본어: (한국어에서 자연스럽게 번역해서 생성)\n`;

  const enField = needEn ? `"en": ["영어 줄 1"], ` : "";
  const jaField = needJa ? `"ja": ["日本語 1行目"]`   : "";

  prompt += `
반환 형식 (JSON 배열):
[
  {"kr": ["한국어 줄 1"], ${enField}${jaField}},
  {"kr": ["다음 슬라이드 줄 1"], ${enField}${jaField}}
]
`;

  return prompt;
}

/**
 * 다언어 AI 응답을 슬라이드 배열로 파싱한다.
 *
 * @param {string}  text   - AI 응답 텍스트
 * @param {boolean} hasEn  - 영어 포함 여부
 * @param {boolean} hasJa  - 일본어 포함 여부
 * @param {boolean} genEn  - 영어 생성 여부
 * @param {boolean} genJa  - 일본어 생성 여부
 * @returns {Array} 슬라이드 배열. 각 항목: { kr: string[], en: string[], ja: string[] }
 */
function parseMultiLangSlides(text, hasEn, hasJa, genEn=false, genJa=false, maxLines=null, krText="", enText="", jaText="") {
  // 사용자가 직접 붙여넣은 가사인지 여부 (재할당 전 원본 값)
  // 직접 입력한 텍스트는 원문이므로 문장부호를 건드리지 않는다.
  const userTypedEn = !!hasEn;
  const userTypedJa = !!hasJa;

  const needEn = hasEn || genEn;
  const needJa = hasJa || genJa;
  hasEn = needEn; hasJa = needJa;

  const safeMax = (Number.isInteger(maxLines) && maxLines >= 1) ? maxLines : DEFAULT_MAX_LINES_PER_SLIDE;

  const cleaned = cleanAIJsonText(text);
  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    throw new Error("AI 응답을 슬라이드 형식으로 읽지 못했습니다. 다시 시도해 주세요.");
  }
  if (!Array.isArray(parsed)) throw new Error("AI 응답이 배열 형식이 아닙니다.");

  const slides = parsed
    .filter(item => item && Array.isArray(item.kr) && item.kr.length)
    .map(item => ({
      kr: normalizeLines(item.kr, safeMax),
      en: hasEn ? normalizeLines(item.en, safeMax) : [],
      ja: hasJa ? normalizeLines(item.ja, safeMax) : []
    }))
    .filter(item => item.kr.length > 0);

  // 사용자가 직접 입력한 가사(한국어 및 붙여넣은 영어/일본어)는 원문 그대로 복원한다.
  // AI가 새로 생성한 번역은 원문이 없으므로(빈 문자열) 복원 대상에서 제외된다.
  restoreOriginalLines(slides, krText, "kr");
  restoreOriginalLines(slides, enText, "en");
  restoreOriginalLines(slides, jaText, "ja");

  // AI가 새로 생성한 번역만 한국어 원문의 문장부호에 맞춘다.
  // (사용자가 직접 붙여넣은 영어·일본어 가사는 원문이므로 그대로 둔다)
  if (needEn && !userTypedEn) {
    applyPunctuationMatch(slides, s => s.kr, s => s.en, (s, lines) => { s.en = lines; });
  }
  if (needJa && !userTypedJa) {
    applyPunctuationMatch(slides, s => s.kr, s => s.ja, (s, lines) => { s.ja = lines; });
  }

  return slides;
}

/**
 * 설정된 AI 제공자로 다언어 가사를 동시 분할한다.
 *
 * @param {string}  kr     - 한국어 가사
 * @param {string}  en     - 영어 가사 (없으면 빈 문자열)
 * @param {string}  ja     - 일본어 가사 (없으면 빈 문자열)
 * @param {Object}  settings - 설정 객체
 * @param {boolean} genEn  - 영어 AI 생성 여부
 * @param {boolean} genJa  - 일본어 AI 생성 여부
 * @returns {Promise<Array>} 다언어 슬라이드 배열
 */
async function splitMultiLangByProvider(kr, en, ja, settings, genEn=false, genJa=false, maxLines=null) {
  const provider = settings.aiProvider || "gemini";
  const prompt   = getMultiLangSplitPrompt(kr, en, ja, genEn, genJa, maxLines);
  const hasEn    = !!(en && en.trim());
  const hasJa    = !!(ja && ja.trim());

  let rawText = "";

  if (provider === "gemini") {
    const geminiApiKey = getAIKey(settings, "gemini");
    if (!geminiApiKey) throw missingKeyError("gemini");
    rawText = await callGeminiText(geminiApiKey, getModelForProvider(settings, "gemini"), prompt);

  } else if (provider === "openai") {
    const apiKey = getAIKey(settings, "openai");
    if (!apiKey) throw missingKeyError("openai");
    rawText = await callOpenAIChat(apiKey, getModelForProvider(settings, "openai"), [
      { role: "system", content: "Return only valid JSON array. No markdown." },
      { role: "user",   content: prompt }
    ]);

  } else if (provider === "claude") {
    const apiKey = getAIKey(settings, "claude");
    if (!apiKey) throw missingKeyError("claude");
    rawText = await callClaudeMessages(apiKey, getModelForProvider(settings, "claude"), prompt, 4096);

  } else {
    throw new Error("지원하지 않는 AI 제공자: " + provider);
  }

  return parseMultiLangSlides(rawText, hasEn, hasJa, genEn, genJa, maxLines, kr, en, ja);
}

/**
 * IPC 핸들러: 한국어/영어/일본어 가사를 동시에 같은 수로 분할한다.
 * genEn/genJa가 true이면 입력이 없어도 AI가 번역 텍스트를 생성한다.
 */
ipcMain.handle("ai:split-multilang", async (event, kr, en, ja, genEn, genJa, maxLines) => {
  const safeMax = (maxLines != null && Number.isInteger(parseInt(maxLines)) && parseInt(maxLines) >= 1) ? parseInt(maxLines) : null;
  return await splitMultiLangByProvider(
    String(kr || "").trim(),
    String(en || "").trim(),
    String(ja || "").trim(),
    loadSettings(),
    !!genEn,
    !!genJa,
    safeMax
  );
});

// ──────────────────────────────────────────────────────────
// 창 생성 및 관리
// ──────────────────────────────────────────────────────────

/**
 * 메인 조작창(index.html)을 생성한다.
 * 창이 닫히면 외부 송출 화면도 함께 닫힌다.
 */
function createMainWindow() {
  // 작업표시줄을 제외한 실제 사용 가능 영역을 기준으로 초기 창 크기 결정
  const { workArea } = screen.getPrimaryDisplay();
  const initWidth  = Math.min(1440, workArea.width);
  const initHeight = Math.min(900,  workArea.height);

  mainWin = new BrowserWindow({
    width:  initWidth,
    height: initHeight,
    x: workArea.x + Math.round((workArea.width  - initWidth)  / 2),
    y: workArea.y + Math.round((workArea.height - initHeight) / 2),
    minWidth: 1100,   // 너무 좁으면 레이아웃이 깨지므로 최소 너비 제한
    minHeight: 700,
    title: "Cheon Myeong Slide Pro",
    icon: path.join(__dirname, process.platform === "win32" ? "icon.ico" : "icon.png"),
    backgroundColor: "#15151f",  // 로딩 전 배경 (흰색 깜빡임 방지) — index.html 의 --bg 와 동일하게 유지
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,  // 렌더러가 Node.js에 직접 접근 불가 (보안)
      nodeIntegration: false   // 렌더러에서 require() 사용 불가 (보안)
    }
  });

  mainWin.loadFile("index.html");

  // 개발 중 DevTools를 열고 싶으면 아래 주석 해제
  // mainWin.webContents.openDevTools();

  // local-fonts 권한 자동 허용 (queryLocalFonts API 사용을 위해 필요)
  mainWin.webContents.session.setPermissionRequestHandler((webContents, permission, callback) => {
    if (permission === "local-fonts") { callback(true); return; }
    callback(false);
  });

  // 최대화/전체화면 상태를 렌더러에 알림 (UI 업데이트용)
  // 최대화 시 명시적으로 workArea 경계를 설정해 작업표시줄과 겹치지 않도록 보장
  mainWin.on("maximize", () => {
    if (mainWin.isDestroyed()) return;
    const display = screen.getDisplayMatching(mainWin.getBounds());
    mainWin.setBounds(display.workArea);
    mainWin.webContents.send("window-fullscreen-change", true);
  });
  mainWin.on("unmaximize",        () => { if (!mainWin.isDestroyed()) mainWin.webContents.send("window-fullscreen-change", false); });
  mainWin.on("enter-full-screen", () => { if (!mainWin.isDestroyed()) mainWin.webContents.send("window-fullscreen-change", true);  });
  mainWin.on("leave-full-screen", () => { if (!mainWin.isDestroyed()) mainWin.webContents.send("window-fullscreen-change", false); });

  // 창을 닫기 전: 렌더러에게 저장 확인을 요청한다.
  // 렌더러가 저장/종료 여부를 결정한 뒤 app:proceed-close 를 호출하면 appIsQuitting=true 가 되어 실제로 닫힌다.
  mainWin.on("close", (e) => {
    if (appIsQuitting) return;            // 이미 확인을 거쳤으면 그대로 닫음
    e.preventDefault();                    // 일단 닫기를 막고 렌더러에 확인 요청
    if (mainWin && !mainWin.isDestroyed() && mainWin.webContents) {
      mainWin.webContents.send("app-confirm-close");
    } else {
      appIsQuitting = true;                // 렌더러를 쓸 수 없으면 그냥 닫음
      if (mainWin && !mainWin.isDestroyed()) mainWin.close();
    }
  });

  // 메인 창이 닫히면 송출 화면도 강제 종료
  mainWin.on("closed", () => {
    if (secondWin && !secondWin.isDestroyed()) {
      secondWin.close();
    }
    mainWin = null;
  });
}

/**
 * IPC 핸들러: 외부 송출 화면(screen.html)을 열거나 이미 열려 있으면 전면으로 가져온다.
 *
 * 동작:
 *   1. 연결된 외부 모니터가 있으면 그 모니터에 전체화면으로 열기
 *   2. 외부 모니터가 없으면 현재 모니터에 전체화면으로 열기
 *   3. 이미 열려 있으면 포커스만 이동
 *
 * 반환값: { success: boolean, alreadyOpen: boolean }
 */
ipcMain.handle("open-second-screen", () => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.setMenuBarVisibility(false);
    secondWin.setFullScreen(true);
    secondWin.focus();
    return { success: true, alreadyOpen: true };
  }

  // 외부 모니터 탐색 (없으면 현재 모니터 사용)
  const displays       = screen.getAllDisplays();
  const primaryDisplay = screen.getPrimaryDisplay();
  const external       = displays.find(display => display.id !== primaryDisplay.id);
  const targetDisplay  = external || primaryDisplay;

  const { x, y, width, height } = targetDisplay.bounds;

  secondWin = new BrowserWindow({
    x, y, width, height,
    title: "Cheon Myeong Slide Pro Output",
    backgroundColor: "#000000",
    frame: false,           // 타이틀바 없음 (전체화면용)
    fullscreen: true,
    autoHideMenuBar: true,
    skipTaskbar: true,      // 작업 표시줄에 표시하지 않음
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  secondWin.setMenuBarVisibility(false);
  secondWin.setFullScreen(true);
  secondWin.loadFile("screen.html");

  // 송출 창 오디오가 시스템에서 음소거되지 않도록 보장
  secondWin.webContents.on("dom-ready", () => {
    secondWin.webContents.setAudioMuted(false);
  });

  // 송출 화면이 닫히면 조작창에 이벤트 알림
  secondWin.on("closed", () => {
    secondWin = null;
    if (mainWin && !mainWin.isDestroyed()) {
      mainWin.webContents.send("second-screen-closed");
    }
  });

  return { success: true, alreadyOpen: false };
});

/**
 * IPC 핸들러: 슬라이드 데이터를 송출 화면으로 전달한다.
 * index.html → main.js → screen.html 방향으로 중계한다.
 * (렌더러끼리 직접 통신 불가 → 메인 프로세스가 중계)
 */
ipcMain.on("slide-update", (event, data) => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("slide-update", data);
  }
});

/**
 * IPC 핸들러: 블랙아웃(화면 암전)을 켜거나 끈다.
 * on=true이면 검은 오버레이로 화면 전체를 덮는다.
 */
ipcMain.on("blank-screen", (event, on) => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("blank-screen", on);
  }
});

/**
 * IPC 핸들러: 자동재생 진행 바 상태를 송출 화면으로 전달한다.
 * (현재 송출 화면에서는 진행 바를 표시하지 않으므로 사실상 미사용)
 */
ipcMain.on("progress-bar", (event, data) => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("progress-bar", data);
  }
});

/**
 * IPC 핸들러: 송출 화면에서 누른 키를 조작창으로 전달한다.
 * screen.html → main.js → index.html 방향으로 중계.
 *
 * 지원 액션:
 *   'next'       - 다음 슬라이드
 *   'prev'       - 이전 슬라이드
 *   'toggleAuto' - 자동재생 토글
 *   'stopAuto'   - 자동재생 중지
 *   {type:'key'} - 그 외 키 (사용자 정의 단축키 처리)
 */
ipcMain.on("output-control", (event, action) => {
  if (!mainWin || mainWin.isDestroyed()) return;

  if (action === "next" || action === "prev" || action === "toggleAuto" || action === "stopAuto") {
    mainWin.webContents.send("output-control", action);
    return;
  }

  if (action && typeof action === "object" && action.type === "key") {
    mainWin.webContents.send("output-control", action);
  }
});

/**
 * IPC 핸들러: 위치 마커 선 상태를 송출 화면으로 전달한다.
 */
ipcMain.on("marker-line", (event, data) => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("marker-line", data);
  }
});

/**
 * IPC 핸들러: 편집 창에서 보낸 영상 재생 트리거를 송출 화면으로 중계한다.
 * 클릭 시 재생 모드에서 편집 창 클릭 → 송출 화면 영상도 재생.
 */
ipcMain.on("video:play", () => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("video:play");
  }
});

ipcMain.on("video:seek", (event, time) => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("video:seek", time);
  }
});

ipcMain.on("video:pause", () => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("video:pause");
  }
});

/**
 * IPC 핸들러: 송출 화면 영상의 재생 위치/상태를 편집 창으로 중계한다.
 * 송출 영상이 기준(마스터)이고, 편집 미리보기가 이 값을 따라가며 싱크를 맞춘다.
 */
ipcMain.on("video:state", (event, state) => {
  if (mainWin && !mainWin.isDestroyed()) {
    mainWin.webContents.send("video:state", state);
  }
});

/**
 * IPC 핸들러: 송출 시작 시 프로젝트 안의 모든 영상 목록을 송출 화면으로 전달한다.
 * 송출 화면은 이를 받아 미리 디코딩(사전 워밍)해 두므로,
 * 이후 어떤 슬라이드로 이동해도 영상이 버벅이지 않는다.
 */
ipcMain.on("prewarm-media", (event, list) => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.webContents.send("prewarm-media", list);
  }
});

/**
 * IPC 핸들러: 외부 송출 화면을 닫는다.
 */
ipcMain.handle("close-second-screen", () => {
  if (secondWin && !secondWin.isDestroyed()) {
    secondWin.close();
  }
  return { success: true };
});

// ──────────────────────────────────────────────────────────
// 앱 메뉴
// ──────────────────────────────────────────────────────────

/**
 * 네이티브 앱 메뉴를 생성하고 등록한다.
 * macOS의 복사/붙여넣기/전체선택 등 표준 기능이 기본적으로 동작하려면
 * 메뉴에 해당 role이 포함되어야 한다.
 */
function buildMenu() {
  const template = [
    {
      label: "Cheon Myeong Slide Pro",
      submenu: [
        { label: "버전 정보", role: "about" },
        { type: "separator" },
        { label: "종료", role: "quit" }
      ]
    },
    {
      label: "편집",
      submenu: [
        { label: "실행 취소", role: "undo" },
        { label: "다시 실행", role: "redo" },
        { type: "separator" },
        { label: "잘라내기", role: "cut" },
        { label: "복사",     role: "copy" },
        { label: "붙여넣기", role: "paste" },
        { label: "전체 선택", role: "selectAll" }
      ]
    },
    {
      label: "보기",
      submenu: [
        { label: "새로고침",   role: "reload" },
        { label: "전체화면",   role: "togglefullscreen" },
        { type: "separator" },
        { label: "개발자 도구", role: "toggleDevTools" }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

// ── 앱 시작 ──

// 영상 소리 자동재생 허용
app.commandLine.appendSwitch('autoplay-policy', 'no-user-gesture-required');
// 미디어 참여도 기반 자동재생 제한 해제 (사이트 방문 기록 없어도 소리 허용)
app.commandLine.appendSwitch('disable-features', 'PreloadMediaEngagementData,MediaEngagementBypassAutoplayPolicies');

/** Electron이 초기화를 완료하면 메뉴와 메인 창을 생성한다. */
app.whenReady().then(() => {
  buildMenu();
  createMainWindow();

  // macOS: 독에서 앱 아이콘 클릭 시 창이 모두 닫혀 있으면 새 창 생성
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createMainWindow();
    }
  });
});

/**
 * 모든 창이 닫히면 non-macOS 플랫폼에서 앱을 종료한다.
 * macOS에서는 표준 동작에 따라 앱을 계속 실행한 채 Dock에 유지한다.
 */
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

// ──────────────────────────────────────────────────────────
// PPTX 내보내기
// ──────────────────────────────────────────────────────────

/**
 * IPC 핸들러: 슬라이드 그룹을 .pptx 파일로 내보낸다.
 *
 * 각 슬라이드의 render(송출 화면용 데이터)를 숨김 창의 screen.html로 그려 PNG로 캡처하고,
 * pptxExport.js(buildPptxBuffer)가 그 이미지를 슬라이드 전체 그림으로 넣는다.
 * 진행 상황: 'pptx:imageProgress' 이벤트로 { current, total } 전달
 */
ipcMain.handle("pptx:export", async (event, { slides, groupName }) => {
  if (!AdmZip || !buildPptxBuffer) {
    return { ok: false, message: "adm-zip가 설치되지 않았습니다.\n터미널에서 npm install 을 실행해 주세요." };
  }

  // 그룹 이름에 경로 구분자 등 파일명에 쓸 수 없는 문자가 섞여 있으면
  // 저장 대화상자가 예상치 못한 경로로 해석할 수 있어 안전한 문자로 치환한다.
  const safeName = (groupName || "슬라이드").replace(/[\\/:*?"<>|]/g, "_").trim() || "슬라이드";

  const dlg = await dialog.showSaveDialog(mainWin, {
    title: "PPTX로 내보내기",
    defaultPath: safeName + ".pptx",
    filters: [{ name: "PowerPoint 파일", extensions: ["pptx"] }]
  });
  if (dlg.canceled) return { ok: false, message: "cancelled" };

  let capWin = null;
  try {
    // 송출 데이터(render)가 있는 슬라이드는 송출 화면(screen.html)과 똑같이 그린 뒤
    // 이미지로 캡처해 넣는다 → 글씨 크기·폰트·색·그림자가 앱 화면과 동일하게 저장된다.
    let outSlides = slides;
    if (slides.some(s => s && s.render)) {
      capWin = new BrowserWindow({
        width: 1920, height: 1080, useContentSize: true,
        show: false, frame: false, enableLargerThanScreen: true,
        backgroundColor: "#000000",
        webPreferences: {
          preload: path.join(__dirname, "preload.js"),
          contextIsolation: true, nodeIntegration: false,
          backgroundThrottling: false
        }
      });
      // 생성 시 크기는 모니터 작업 영역(작업 표시줄 제외)에 맞춰 줄어들 수 있어 다시 지정한다
      capWin.setContentSize(1920, 1080);
      capWin.webContents.setAudioMuted(true);
      await capWin.loadFile("screen.html");

      outSlides = [];
      for (let i = 0; i < slides.length; i++) {
        event.sender.send("pptx:imageProgress", { current: i + 1, total: slides.length });
        const sl = slides[i];
        if (!sl || !sl.render) { outSlides.push(sl); continue; }
        const rect = await capWin.webContents.executeJavaScript(
          `window.__captureSlide(${JSON.stringify(sl.render)})`
        );
        let img = await capWin.webContents.capturePage(rect);
        // 고배율(DPI) 화면에서는 캡처가 더 크게 나오므로 1920×1080으로 맞춘다
        const sz = img.getSize();
        if (sz.width !== 1920 || sz.height !== 1080) {
          img = img.resize({ width: 1920, height: 1080, quality: "best" });
        }
        outSlides.push({ renderedImage: "data:image/png;base64," + img.toPNG().toString("base64") });
      }
      capWin.destroy();
      capWin = null;
    }

    const buf = buildPptxBuffer(outSlides);
    fs.writeFileSync(dlg.filePath, buf);
    return { ok: true, filePath: dlg.filePath, count: slides.length };
  } catch (err) {
    if (capWin) { try { capWin.destroy(); } catch (e) {} }
    console.error("PPTX 내보내기 실패:", err);
    return { ok: false, message: "내보내기 실패: " + err.message };
  }
});

// ──────────────────────────────────────────────────────────
// PPTX 가져오기 (텍스트/이미지 추출 방식)
// ──────────────────────────────────────────────────────────

/**
 * IPC 핸들러: .pptx 파일을 열어 슬라이드 데이터로 변환한다.
 *
 * PPTX 파일 구조:
 *   .pptx는 ZIP 아카이브로, 내부에 XML 파일들이 있다.
 *   - ppt/slides/slide1.xml, slide2.xml, ... : 슬라이드 내용
 *   - ppt/slides/_rels/slide1.xml.rels       : 이미지 등 외부 파일 참조 관계
 *   - ppt/media/                             : 이미지 파일들
 *   - ppt/theme/theme1.xml                   : 폰트/색상 테마
 *   - ppt/presentation.xml                   : 슬라이드 크기 등 전체 설정
 */
ipcMain.handle("pptx:import", async (event) => {
  if (!AdmZip) {
    return { ok: false, message: "adm-zip가 설치되지 않았습니다.\n터미널에서 npm install 을 실행해 주세요." };
  }

  const result = await dialog.showOpenDialog(mainWin, {
    title: "PPTX 가져오기",
    filters: [
      { name: "PowerPoint 파일", extensions: ["pptx", "ppt"] },
      { name: "모든 파일",       extensions: ["*"] }
    ],
    properties: ["openFile"]
  });
  if (result.canceled || !result.filePaths.length) return { ok: false, message: "cancelled" };

  // 구버전 .ppt 형식은 ZIP 구조가 달라 지원 불가
  const ext = path.extname(result.filePaths[0]).toLowerCase();
  if (ext === ".ppt") {
    return { ok: false, message: ".ppt 형식은 지원하지 않습니다.\nPowerPoint에서 '다른 이름으로 저장 → .pptx'로 변환 후 다시 시도해 주세요." };
  }

  try {
    const zip = new AdmZip(result.filePaths[0]);

    // 슬라이드 크기 읽기 (단위: EMU, 1인치 = 914400 EMU)
    let slideW = 9144000, slideH = 5143500;  // 기본값: 16:9 (10인치 × 5.625인치)
    try {
      const presXml = zip.readAsText("ppt/presentation.xml");
      const cxM = presXml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"/);
      const cyM = presXml.match(/<p:sldSz\b[^>]*\bcy="(\d+)"/);
      if (cxM) slideW = parseInt(cxM[1]);
      if (cyM) slideH = parseInt(cyM[1]);
    } catch(e) {}

    // 슬라이드 XML 파일 목록을 번호순으로 정렬
    const slideEntries = zip.getEntries()
      .map(e => ({ entry: e, norm: e.entryName.replace(/\\/g, "/") }))
      .filter(({ norm }) => /^ppt\/slides\/slide\d+\.xml$/i.test(norm))
      .sort((a, b) => {
        const n = x => { const m = x.norm.match(/slide(\d+)\.xml$/i); return m ? parseInt(m[1]) : 0; };
        return n(a) - n(b);
      });

    if (slideEntries.length === 0) {
      return { ok: false, message: "슬라이드를 찾을 수 없습니다. 올바른 .pptx 파일인지 확인해 주세요." };
    }

    const theme  = extractPptxTheme(zip);
    const slides = [];

    for (const { entry, norm } of slideEntries) {
      const xml = entry.getData().toString("utf8");

      // 관계 파일(.rels): 이미지 참조 ID와 실제 파일 경로 매핑
      const relsPath  = norm.replace(/\/([^/]+)$/, "/_rels/$1.rels");
      const relsEntry = zipGetEntry(zip, relsPath);
      const relsMap   = relsEntry ? parseRelsXml(relsEntry.getData().toString("utf8")) : {};

      const slide = parsePptxSlide(xml, relsMap, zip, slideW, slideH, theme);
      slides.push(slide);
    }

    const fileName = path.basename(result.filePaths[0], ".pptx");
    return { ok: true, slides, fileName };
  } catch (err) {
    console.error("PPTX 가져오기 실패:", err);
    return { ok: false, message: "가져오기 실패: " + err.message };
  }
});

// ──────────────────────────────────────────────────────────
// PPTX 이미지 변환 가져오기
// 각 슬라이드를 실제로 렌더링해 PNG 이미지로 캡처하는 방식.
// 텍스트 추출이 어려운 복잡한 디자인 슬라이드에 적합하다.
// ──────────────────────────────────────────────────────────

/**
 * 슬라이드 데이터를 HTML 문자열로 변환한다.
 * 화면 캡처를 위해 1920×1080 크기의 HTML 페이지를 생성한다.
 *
 * @param {Object} slide - parsePptxSlide()가 반환한 슬라이드 데이터
 * @returns {string} 완성된 HTML 문자열
 */
function generateSlideHtml(slide) {
  const esc = s => String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');

  let bgCss = 'background:#000000;';
  if (slide.mediaURL) {
    // cover: 슬라이드 전체를 채워 여백 없이 표시
    bgCss = `background-image:url('${slide.mediaURL}');background-size:cover;background-position:center;background-repeat:no-repeat;background-color:#000000;`;
  } else if (slide.bgColor) {
    bgCss = `background:${slide.bgColor};`;
  }

  const si = (slide.textShadowEnabled !== false) ? Math.max(0, Math.min(1, (slide.textShadowIntensity ?? 80) / 100)) : 0;
  const shadowVal  = si > 0 ? `0 2px ${Math.round(48 * si)}px rgba(0,0,0,${(0.9 * si).toFixed(2)})` : 'none';
  const outlineCss = slide.textOutlineEnabled ? `-webkit-text-stroke:${slide.textOutlineWidth || 2}px rgba(0,0,0,0.85);paint-order:stroke fill;` : '';

  const fs = slide.fontSize || 36;
  const tx = slide.textX ?? 50;
  const ty = slide.textY ?? 50;
  const tw = slide.textWidth ? slide.textWidth + '%' : '80%';
  const textAnchor = slide.textAnchor || 'ctr';
  const textTransY =
    textAnchor === 't' ? '0' :
    textAnchor === 'b' ? '-100%' :
    '-50%';

  // 메인 가사 HTML
  let textHtml = '';
  if (slide.lyrics && slide.lyrics.trim()) {
    textHtml = `<div style="position:absolute;left:${tx}%;top:${ty}%;width:${tw};transform:translate(-50%,${textTransY});font-size:${fs}px;color:${slide.fontColor || '#ffffff'};font-family:'${slide.fontFamily || 'Noto Serif KR'}',serif;font-weight:700;text-shadow:${shadowVal};${outlineCss}text-align:center;line-height:1.7;white-space:pre-wrap;word-break:keep-all;">${esc(slide.lyrics)}</div>`;
    if (slide.sub && slide.sub.trim()) {
      const subFs = Math.round(fs * 0.45);
      const lineCount = Math.max(1, String(slide.lyrics).split(/\n|<br\s*\/?>/i).length);
      textHtml += `<div style="position:absolute;left:${tx}%;top:calc(${ty}% + ${Math.round(fs * 1.7 * lineCount)}px);width:80%;transform:translate(-50%,-50%);font-size:${subFs}px;color:rgba(255,255,255,0.5);font-family:'Noto Sans KR',sans-serif;font-weight:300;text-align:center;letter-spacing:1px;white-space:pre-wrap;">${esc(slide.sub)}</div>`;
    }
  }

  // 도형 HTML (shapeBoxes) — 함수 스코프에서 독립적으로 계산
  const shapeBoxesHtml = (slide.shapeBoxes || []).filter(s => s).map(s => {
    const shapeType = String(s.shapeType || 'rect');
    const borderRadius =
      shapeType === 'ellipse' || shapeType === 'arc' ? '50%' :
      shapeType === 'roundRect' ? '18px' :
      '0';
    const fillColor = s.fillColor || 'transparent';
    const lineColor = s.lineColor || 'transparent';
    const lineWidth = Number.isFinite(Number(s.lineWidth)) ? Number(s.lineWidth) : 0;
    return `<div style="position:absolute;left:${s.x || 50}%;top:${s.y || 50}%;width:${s.w || 1}%;height:${s.h || 1}%;transform:translate(-50%,-50%);background:${fillColor};border:${lineWidth}px solid ${lineColor};border-radius:${borderRadius};box-sizing:border-box;"></div>`;
  }).join('');

  // 추가 텍스트박스 HTML (textBoxes) — shapeBoxes와 분리하여 각 박스를 올바르게 렌더링
  const boxesHtml = (slide.textBoxes || []).filter(b => b && b.text).map(b => {
    const bSi     = (b.textShadowEnabled !== false) ? Math.max(0, Math.min(1, (b.textShadowIntensity ?? 80) / 100)) : 0;
    const bShadow = bSi > 0 ? `0 2px ${Math.round(18 * bSi)}px rgba(0,0,0,${(0.85 * bSi).toFixed(2)})` : 'none';
    const bOutline = b.textOutlineEnabled ? `-webkit-text-stroke:${b.textOutlineWidth || 2}px rgba(0,0,0,0.85);paint-order:stroke fill;` : '';
    const bAnc    = b.textAnchor || 'ctr';
    const bTransY = bAnc === 't' ? '0' : bAnc === 'b' ? '-100%' : '-50%';
    const bFs     = Math.round((b.fontSize || 24) * 1.4);
    const bAlign  = b.align || 'center';
    return `<div style="position:absolute;left:${b.x || 50}%;top:${b.y || 50}%;width:${b.width || 42}%;transform:translate(-50%,${bTransY});font-size:${bFs}px;color:${b.fontColor || '#ffffff'};font-family:'${b.fontFamily || 'Noto Sans KR'}',sans-serif;font-weight:600;text-align:${bAlign};text-shadow:${bShadow};${bOutline}white-space:pre-wrap;word-break:keep-all;">${esc(b.text)}</div>`;
  }).join('');

  // 완성된 HTML 페이지 (1920×1080, 구글 웹폰트 포함)
  return `<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@700;900&family=Noto+Sans+KR:wght@300;400;700&family=Black+Han+Sans&family=Nanum+Gothic&family=Nanum+Myeongjo&display=swap" rel="stylesheet">
<style>*{margin:0;padding:0;box-sizing:border-box;}html,body{width:1920px;height:1080px;overflow:hidden;}</style>
</head><body style="width:1920px;height:1080px;position:relative;${bgCss}">${shapeBoxesHtml}${textHtml}${boxesHtml}</body></html>`;
}

/**
 * IPC 핸들러: .pptx의 각 슬라이드를 PNG 이미지로 변환해 가져온다.
 *
 * 동작 원리:
 *   1. 각 슬라이드를 generateSlideHtml()로 HTML 파일로 변환
 *   2. 숨겨진 오프스크린 BrowserWindow(1920×1080)에 HTML 로드
 *   3. webContents.capturePage()로 PNG 캡처
 *   4. base64 data URL로 변환해 반환
 *
 * 진행 상황: 'pptx:imageProgress' 이벤트로 { current, total } 전달
 */
ipcMain.handle("pptx:importAsImages", async (event) => {
  if (!AdmZip) {
    return { ok: false, message: "adm-zip가 설치되지 않았습니다.\n터미널에서 npm install 을 실행해 주세요." };
  }

  const pickResult = await dialog.showOpenDialog(mainWin, {
    title: "PPTX 이미지로 가져오기",
    filters: [
      { name: "PowerPoint 파일", extensions: ["pptx", "ppt"] },
      { name: "모든 파일",       extensions: ["*"] }
    ],
    properties: ["openFile"]
  });
  if (pickResult.canceled || !pickResult.filePaths.length) return { ok: false, message: "cancelled" };

  const filePath = pickResult.filePaths[0];
  if (path.extname(filePath).toLowerCase() === ".ppt") {
    return { ok: false, message: ".ppt 형식은 지원하지 않습니다.\nPowerPoint에서 '다른 이름으로 저장 → .pptx'로 변환 후 다시 시도해 주세요." };
  }

  let offWin = null;
  try {
    const zip = new AdmZip(filePath);

    // 슬라이드 크기 파싱
    let slideW = 9144000, slideH = 5143500;
    try {
      const presXml = zip.readAsText("ppt/presentation.xml");
      const cxM = presXml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"/);
      const cyM = presXml.match(/<p:sldSz\b[^>]*\bcy="(\d+)"/);
      if (cxM) slideW = parseInt(cxM[1]);
      if (cyM) slideH = parseInt(cyM[1]);
    } catch(e) {}

    const slideEntries = zip.getEntries()
      .map(e => ({ entry: e, norm: e.entryName.replace(/\\/g, "/") }))
      .filter(({ norm }) => /^ppt\/slides\/slide\d+\.xml$/i.test(norm))
      .sort((a, b) => {
        const n = x => { const m = x.norm.match(/slide(\d+)\.xml$/i); return m ? parseInt(m[1]) : 0; };
        return n(a) - n(b);
      });

    if (!slideEntries.length) return { ok: false, message: "슬라이드를 찾을 수 없습니다. 올바른 .pptx 파일인지 확인해 주세요." };
    const theme        = extractPptxTheme(zip);
    const os = require("os");
    // 오프스크린 창: 화면에 보이지 않는 1920×1080 창으로 HTML을 렌더링해 캡처
    offWin = new BrowserWindow({
      width: 1920, height: 1080,
      show: false, frame: false,
      webPreferences: { nodeIntegration: false, contextIsolation: true }
    });

    for (let i = 0; i < slideEntries.length; i++) {
      const { entry, norm } = slideEntries[i];
      // 진행률 이벤트 전송
      event.sender.send("pptx:imageProgress", { current: i + 1, total: slideEntries.length });

      // 메모리 절약을 위해 순차적으로 파싱 및 렌더링 수행
      const xml = entry.getData().toString("utf8");
      const relsPath = norm.replace(/\/([^/]+)$/, "/_rels/$1.rels");
      const relsEntry = zipGetEntry(zip, relsPath);
      const relsMap = relsEntry ? parseRelsXml(relsEntry.getData().toString("utf8")) : {};
      const parsedSlide = parsePptxSlide(xml, relsMap, zip, slideW, slideH, theme);

      const html    = generateSlideHtml(parsedSlide);
      const tmpFile = path.join(os.tmpdir(), `ws_slide_${Date.now()}_${i}.html`);
      
      fs.writeFileSync(tmpFile, html, "utf8");

      try {
        await offWin.loadFile(tmpFile);
        await offWin.webContents.executeJavaScript(`
          Promise.race([
            document.fonts.ready,
            new Promise(r => setTimeout(r, 3000))
          ])
        `);
        const nativeImg = await offWin.webContents.capturePage();
        const b64       = nativeImg.toJPEG(88).toString("base64");
        const slideChunk = {
          mediaURL: `data:image/jpeg;base64,${b64}`,
          mediaType: "image",
          lyrics: "", sub: "", bgColor: null, textBoxes: [],
          layoutType: "centerLyrics",
          fontSize:   parsedSlide.fontSize  || 36,
          fontColor:  parsedSlide.fontColor  || "#ffffff",
          fontFamily: parsedSlide.fontFamily || "Noto Serif KR",
          textX: 50, textY: 50, textWidth: null,
          textShadowEnabled: false, textShadowIntensity: 80,
          textOutlineEnabled: false, textOutlineWidth: 2,
        };
        // 대용량 이미지 데이터는 청크 이벤트로 즉시 전달해 렌더러 프리징을 줄인다.
        event.sender.send("pptx:imageChunk", {
          index: i,
          total: slideEntries.length,
          slide: slideChunk
        });
      } finally {
        try { fs.unlinkSync(tmpFile); } catch(e) {}  // 임시 파일 정리
      }
    }

    offWin.destroy();
    offWin = null;

    const fileName = path.basename(filePath, ".pptx");
    return { ok: true, fileName, total: slideEntries.length, streamed: true };
  } catch (err) {
    if (offWin) { try { offWin.destroy(); } catch(e) {} }
    console.error("PPTX 이미지 변환 오류:", err);
    return { ok: false, message: "변환 실패: " + err.message };
  }
});

// ──────────────────────────────────────────────────────────
// PPTX → 네이티브 이미지 렌더링 (원본 디자인 유지 모드)
// PowerPoint(macOS AppleScript)로 PNG를 직접 저장하고,
// 실패하면 LibreOffice/PDF → Python CoreGraphics PNG 변환으로 폴백한다.
// PowerPoint도 LibreOffice도 없으면 설치 안내 메시지를 반환한다.
// ──────────────────────────────────────────────────────────

function removePptxTextShapeXml(slideXml) {
  return String(slideXml || "").replace(/<p:sp>([\s\S]*?)<\/p:sp>/g, (full, body) => {
    return body.includes("<p:txBody>") ? "" : full;
  });
}

function createPptxWithoutTextShapes(filePath, tmpDir) {
  if (!AdmZip) throw new Error("adm-zip가 설치되지 않았습니다.");
  const zip = new AdmZip(filePath);
  const slideEntries = zip.getEntries()
    .filter(entry => /^ppt\/slides\/slide\d+\.xml$/i.test(entry.entryName));
  for (const entry of slideEntries) {
    const xml     = entry.getData().toString("utf8");
    const stripped = removePptxTextShapeXml(xml);
    if (stripped === xml) continue;
    const nextData = Buffer.from(stripped, "utf8");
    if (typeof zip.updateFile === "function") {
      zip.updateFile(entry.entryName, nextData);
    } else {
      zip.deleteFile(entry.entryName);
      zip.addFile(entry.entryName, nextData);
    }
  }
  const outPath = path.join(tmpDir, "no_text_shapes.pptx");
  zip.writeZip(outPath);
  return outPath;
}

/**
 * PPTX 슬라이드를 Electron 내부 렌더러로 PNG 이미지 배열로 변환한다.
 * PowerPoint / LibreOffice 없이도 동작하는 폴백 경로.
 */
async function renderPptxSlidesWithElectron(filePath, event) {
  const zip = new AdmZip(filePath);

  let slideW = 9144000, slideH = 5143500;
  try {
    const presXml = zip.readAsText("ppt/presentation.xml");
    const cxM = presXml.match(/<p:sldSz\b[^>]*\bcx="(\d+)"/);
    const cyM = presXml.match(/<p:sldSz\b[^>]*\bcy="(\d+)"/);
    if (cxM) slideW = parseInt(cxM[1]);
    if (cyM) slideH = parseInt(cyM[1]);
  } catch(e) {}

  const slideEntries = zip.getEntries()
    .map(e => ({ entry: e, norm: e.entryName.replace(/\\/g, "/") }))
    .filter(({ norm }) => /^ppt\/slides\/slide\d+\.xml$/i.test(norm))
    .sort((a, b) => {
      const n = x => { const m = x.norm.match(/slide(\d+)\.xml$/i); return m ? parseInt(m[1]) : 0; };
      return n(a) - n(b);
    });

  if (!slideEntries.length) throw new Error("슬라이드를 찾을 수 없습니다.");

  const theme = extractPptxTheme(zip);
  const parsedSlides = [];
  for (const { entry, norm } of slideEntries) {
    const xml      = entry.getData().toString("utf8");
    const relsPath  = norm.replace(/\/([^/]+)$/, "/_rels/$1.rels");
    const relsEntry = zipGetEntry(zip, relsPath);
    const relsMap   = relsEntry ? parseRelsXml(relsEntry.getData().toString("utf8")) : {};
    parsedSlides.push(parsePptxSlide(xml, relsMap, zip, slideW, slideH, theme));
  }

  const os = require("os");
  const offWin = new BrowserWindow({
    width: 1920, height: 1080,
    show: false, frame: false,
    webPreferences: { nodeIntegration: false, contextIsolation: true }
  });

  const resultSlides = [];
  try {
    for (let i = 0; i < parsedSlides.length; i++) {
      event.sender.send("pptx:imageProgress", { current: i + 1, total: parsedSlides.length });
      const html    = generateSlideHtml(parsedSlides[i]);
      const tmpFile = path.join(os.tmpdir(), `ws_slide_${Date.now()}_${i}.html`);
      fs.writeFileSync(tmpFile, html, "utf8");
      try {
        await offWin.loadFile(tmpFile);
        await offWin.webContents.executeJavaScript(
          `Promise.race([document.fonts.ready, new Promise(r => setTimeout(r, 3000))])`
        );
        const nativeImg = await offWin.webContents.capturePage();
        const slide = {
          mediaURL:  `data:image/jpeg;base64,${nativeImg.toJPEG(88).toString("base64")}`,
          mediaType: "image",
          bgColor:   null
        };
        resultSlides.push(slide);
        event.sender.send("pptx:imageChunk", { index: i, total: parsedSlides.length, slide });
      } finally {
        try { fs.unlinkSync(tmpFile); } catch(e) {}
      }
    }
  } finally {
    try { offWin.destroy(); } catch(e) {}
  }

  return resultSlides;
}

ipcMain.handle("shell:open-external", async (event, url) => {
  await shell.openExternal(url);
});

ipcMain.handle("pptx:select-file", async () => {
  if (!AdmZip) {
    return { ok: false, message: "adm-zip가 설치되지 않았습니다.\n터미널에서 npm install 을 실행해 주세요." };
  }

  const pick = await dialog.showOpenDialog(mainWin, {
    title: "PPTX 이미지로 가져오기",
    filters: [{ name: "PowerPoint 파일", extensions: ["pptx"] }],
    properties: ["openFile"]
  });
  if (pick.canceled || !pick.filePaths.length) return { ok: false, message: "cancelled" };

  const filePath = pick.filePaths[0];
  try {
    const zip = new AdmZip(filePath);
    const slideCount = zip.getEntries()
      .map(e => e.entryName.replace(/\\/g, "/"))
      .filter(name => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .length;
    if (!slideCount) return { ok: false, message: "슬라이드를 찾을 수 없습니다. 올바른 .pptx 파일인지 확인해 주세요." };
    return {
      ok: true,
      filePath,
      fileName: path.basename(filePath, path.extname(filePath)),
      slideCount
    };
  } catch (err) {
    return { ok: false, message: "PPTX 파일 확인 실패: " + err.message };
  }
});

// PowerPoint 자동화 권한 사전 확인 (Windows/macOS)
// Windows: PowerPoint COM 클래스 등록 여부를 확인한다 (런처 없이 레지스트리만 조회).
// macOS: AppleScript 자동화 권한을 테스트한다.
ipcMain.handle("pptx:check-ppt-permission", async () => {
  const isMac     = process.platform === "darwin";
  const isWindows = process.platform === "win32";

  // ── Windows: PowerPoint COM 가용 여부 확인 ──
  if (isWindows) {
    const { spawnSync } = require("child_process");
    const checkScript = `try { $t=[System.Type]::GetTypeFromProgID('PowerPoint.Application'); if($t -ne $null){Write-Output 'OK'}else{Write-Output 'NOT_INSTALLED'} } catch { Write-Output 'NOT_INSTALLED' }`;
    try {
      const r = spawnSync(
        "powershell",
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", checkScript],
        { timeout: 10000, encoding: "utf8" }
      );
      const out = (r.stdout || "").trim();
      if (out === "OK") return { ok: true, platform: "windows" };
      return { ok: false, notInstalled: true, platform: "windows",
               message: "Microsoft PowerPoint가 설치되지 않았습니다." };
    } catch (e) {
      return { ok: true, platform: "windows" }; // 확인 자체 실패 시 진행 허용
    }
  }

  if (!isMac) return { ok: true, platform: "other" };

  const { spawnSync } = require("child_process");

  let pptAppPath = null;
  const candidatePaths = [
    "/Applications/Microsoft PowerPoint.app",
    "/Applications/Microsoft Office/Microsoft PowerPoint.app",
  ];
  for (const p of candidatePaths) {
    if (fs.existsSync(p) && fs.existsSync(path.join(p, "Contents/MacOS/Microsoft PowerPoint"))) {
      pptAppPath = p;
      break;
    }
  }
  if (!pptAppPath) {
    try {
      const mdf = spawnSync("mdfind", ["kMDItemCFBundleIdentifier == 'com.microsoft.Powerpoint'"], { timeout: 5000, encoding: "utf8" });
      const found = (mdf.stdout || "").split("\n").find(p => p.trim().endsWith(".app"));
      if (found && fs.existsSync(path.join(found.trim(), "Contents/MacOS/Microsoft PowerPoint"))) {
        pptAppPath = found.trim();
      }
    } catch(e) { /* mdfind 없음 */ }
  }

  if (!pptAppPath) {
    return { ok: false, notInstalled: true, message: "Microsoft PowerPoint가 설치되어 있지 않습니다." };
  }

  // 최소 AppleScript로 자동화 권한을 유발 또는 확인한다.
  // macOS는 첫 실행 시 사용자에게 허용/거부 다이얼로그를 자동으로 표시한다.
  const testScript = `tell application ${JSON.stringify(pptAppPath)}\n  get name\nend tell`;
  try {
    const result = spawnSync("osascript", ["-e", testScript], { timeout: 30000, encoding: "utf8" });
    if (result.status === 0) return { ok: true };
    const errMsg = (result.stderr || result.stdout || "");
    if (errMsg.includes("-1743") || errMsg.includes("Not authorized") || errMsg.includes("not allowed to send")) {
      return { ok: false, permissionDenied: true, message: "PowerPoint 자동화 권한이 없습니다." };
    }
    // 기타 오류(PowerPoint가 설치됐지만 응답 불가 등)는 일단 통과시켜 본 렌더에서 처리
    return { ok: true };
  } catch(e) {
    return { ok: true }; // 확인 자체 실패 시 본 렌더에서 처리
  }
});

// Windows: PowerPoint Trust Center 설정 창을 열도록 PowerPoint를 실행한다.
ipcMain.handle("pptx:open-ppt-trust-settings", async () => {
  if (process.platform !== "win32") return;
  const { spawn } = require("child_process");
  try {
    spawn("powershell", ["-Command", "Start-Process POWERPNT.EXE"], {
      detached: true, stdio: "ignore"
    });
  } catch (e) { /* 실패 무시 */ }
});

// Windows: 지정 PPTX 파일을 PowerPoint에서 눈에 보이게 연다.
// 임시 복사본을 만들어 Zone.Identifier(Protected View 유발 ADS)를 확실히 제거한 뒤 열어
// '보호된 보기' 없이 바로 편집 가능한 상태로 열리도록 한다.
// 임시 경로를 tempFilePath로 반환해 exportFromRunningPpt가 동일한 파일을 참조할 수 있게 한다.
ipcMain.handle("pptx:launch-in-powerpoint", async (event, filePath) => {
  if (process.platform !== "win32") return { ok: false, message: "Windows 전용" };
  const { spawnSync } = require("child_process");
  try {
    const os = require("os");
    // 원본을 임시 복사본으로 복사해 Zone.Identifier 제거 (원본 파일 미수정)
    const tmpPath = path.join(os.tmpdir(), `ws_launch_${Date.now()}${path.extname(filePath)}`);
    fs.copyFileSync(filePath, tmpPath);
    // PowerShell 단일 인용 경로로 Unblock-File 실행 (경로 이스케이프 문제 방지)
    const psEsc = s => s.replace(/'/g, "''");
    spawnSync("powershell", [
      "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
      "-Command", `Unblock-File -Path '${psEsc(tmpPath)}' -ErrorAction SilentlyContinue`
    ], { timeout: 5000 });
    // Electron shell.openPath으로 기본 프로그램(PowerPoint)에서 파일 열기
    // Start-Process 대신 사용 — 경로 이스케이프 없이 안정적으로 동작
    const openErr = await shell.openPath(tmpPath);
    if (openErr) return { ok: false, message: openErr };
    return { ok: true, tempFilePath: tmpPath };
  } catch (e) {
    return { ok: false, message: e.message };
  }
});

// Windows: 이미 열려있는 PowerPoint 인스턴스에 연결해 슬라이드를 PNG로 내보낸다.
// GetActiveObject로 실행 중인 PowerPoint에 붙고, 없으면 새 COM 인스턴스를 생성한다.
ipcMain.handle("pptx:export-from-running-ppt", async (event, options = {}) => {
  if (process.platform !== "win32") return { ok: false, message: "Windows 전용" };

  const { filePath, originalFileName } = options;
  if (!filePath) return { ok: false, message: "filePath가 없습니다." };

  const os = require("os");
  const tmpDir  = path.join(os.tmpdir(), "pptx_export_" + Date.now());
  const pngDir  = path.join(tmpDir, "pngs");
  fs.mkdirSync(pngDir, { recursive: true });

  const psEsc    = s => s.replace(/\\/g, "\\\\").replace(/'/g, "''");
  // originalFileName: launchInPowerPoint가 임시 경로를 사용한 경우 원본 파일명을 유지
  const fileName = originalFileName || path.basename(filePath, path.extname(filePath));

  const psScript = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Continue"

$targetPath = '${psEsc(filePath)}'
$targetName = '${psEsc(fileName)}'
$outDir     = '${psEsc(pngDir)}'

$ppt          = $null
$presentation = $null
$createdNew   = $false
$openedByUs   = $false   # 이 스크립트가 직접 연 프레젠테이션인지 여부

try {
    # ── 1. 실행 중인 PowerPoint에 연결, 없으면 새 숨김 인스턴스 생성 ──────
    try {
        $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
    } catch {
        $ppt               = New-Object -ComObject PowerPoint.Application
        $ppt.Visible       = 0    # 숨김: 화면에 PowerPoint 창을 표시하지 않음
        $ppt.DisplayAlerts = 0    # 모든 보안·보호 대화상자 억제
        $createdNew        = $true
    }
    # 기존 인스턴스에도 대화상자 억제 적용
    try { $ppt.DisplayAlerts = 0 } catch {}

    # ── 2. Protected View 먼저 탈출 (탐색보다 먼저 처리해야 정상 Presentation 목록에 나타남) ──
    if ($ppt.ProtectedViewWindows.Count -gt 0) {
        for ($pi = 1; $pi -le $ppt.ProtectedViewWindows.Count; $pi++) {
            try {
                $pvw     = $ppt.ProtectedViewWindows.Item($pi)
                $pvwName = try { [System.IO.Path]::GetFileNameWithoutExtension($pvw.Presentation.Name) } catch { '' }
                # 이름이 일치하거나 창이 하나뿐이면 편집 모드로 전환
                if ($pvwName -eq $targetName -or $ppt.ProtectedViewWindows.Count -eq 1) {
                    $p2 = $pvw.Edit()
                    if ($p2 -ne $null) { $presentation = $p2; Start-Sleep -Seconds 2; break }
                }
            } catch {}
        }
    }

    # ── 3. 열린 프레젠테이션 목록에서 탐색 ──────────────────────────────
    if ($presentation -eq $null) {
        foreach ($p in $ppt.Presentations) {
            if ($p.FullName -eq $targetPath) { $presentation = $p; break }
        }
    }
    if ($presentation -eq $null) {
        foreach ($p in $ppt.Presentations) {
            if ([System.IO.Path]::GetFileNameWithoutExtension($p.Name) -eq $targetName) {
                $presentation = $p; break
            }
        }
    }
    # 열린 파일이 하나뿐이면 그것을 사용
    if ($presentation -eq $null -and $ppt.Presentations.Count -eq 1) {
        $presentation = $ppt.Presentations.Item(1)
    }

    # ── 4. 찾지 못한 경우 직접 열기 (창 없는 숨김 모드) ─────────────────
    if ($presentation -eq $null) {
        $rnd     = [System.IO.Path]::GetRandomFileName().Replace('.','')
        $tmpCopy = [System.IO.Path]::Combine($env:TEMP, "ws_exp_$rnd.pptx")
        try {
            Copy-Item -Path $targetPath -Destination $tmpCopy -Force -ErrorAction Stop
            Unblock-File -Path $tmpCopy -ErrorAction SilentlyContinue
        } catch { $tmpCopy = $targetPath }
        # WithWindow=0 → 창 없이 열기, ReadOnly=0, Untitled=0
        $presentation = $ppt.Presentations.Open($tmpCopy, 0, 0, 0)
        Start-Sleep -Seconds 3
        $openedByUs = $true
    }

    if ($presentation -eq $null) { throw "프레젠테이션을 찾을 수 없습니다." }

    # ── 5. PNG 내보내기 ───────────────────────────────────────────────
    $slideCount = $presentation.Slides.Count
    Write-Output "PAGES:$slideCount"
    [Console]::Out.Flush()

    $sw      = $presentation.PageSetup.SlideWidth
    $sh      = $presentation.PageSetup.SlideHeight
    $exportH = if ($sh -gt 0) { [int]([Math]::Round(1920 * ($sh / $sw))) } else { 1080 }

    for ($i = 1; $i -le $slideCount; $i++) {
        $outPath = [System.IO.Path]::Combine($outDir, ("slide{0:D4}.png" -f $i))
        $presentation.Slides.Item($i).Export($outPath, "PNG", 1920, $exportH)
        Write-Output ("DONE:" + $i + ":" + $slideCount)
        [Console]::Out.Flush()
    }

    # ── 6. 정리: 이 스크립트가 연 것만 닫기 (사용자 열람 파일 유지) ─────
    if ($openedByUs -or $createdNew) {
        try { $presentation.Close() } catch {}
    }
    if ($createdNew) { try { $ppt.Quit() } catch {} }
    Write-Output "SUCCESS"

} catch {
    Write-Output "ERROR:$($_.Exception.Message)"
    if ($openedByUs -and $presentation -ne $null) { try { $presentation.Close() } catch {} }
    if ($createdNew -and $ppt -ne $null) { try { $ppt.Quit() } catch {} }
    exit 1
} finally {
    if ($presentation -ne $null) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch {}
        $presentation = $null
    }
    if ($ppt -ne $null) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
        $ppt = $null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
`.trim();

  const psFile = path.join(tmpDir, "export_running.ps1");
  fs.writeFileSync(psFile,
    Buffer.concat([Buffer.from([0xEF, 0xBB, 0xBF]), Buffer.from(psScript, "utf8")]));

  try {
    await new Promise((resolve, reject) => {
      const { spawn } = require("child_process");
      const child = spawn("powershell",
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", psFile],
        { stdio: ["ignore", "pipe", "pipe"] });
      let stdout = "", stderr = "", totalSlides = 0;

      const parseLine = line => {
        line = line.trim();
        if (line.startsWith("PAGES:")) {
          totalSlides = parseInt(line.slice(6)) || 0;
          event.sender.send("pptx:imageProgress", { current: 0, total: totalSlides });
        } else if (line.startsWith("DONE:")) {
          const [cur, tot] = line.slice(5).split(":");
          event.sender.send("pptx:imageProgress",
            { current: parseInt(cur) || 0, total: parseInt(tot) || totalSlides });
        }
      };
      child.stdout.on("data", chunk => { const t = chunk.toString("utf8"); stdout += t; t.split("\n").forEach(parseLine); });
      child.stderr.on("data", chunk => { const t = chunk.toString("utf8"); stderr += t; t.split("\n").forEach(parseLine); });
      child.on("error", reject);
      child.on("close", code => {
        if (code === 0 || (stdout + stderr).includes("SUCCESS")) resolve();
        else reject(new Error((stderr || stdout || "PowerShell 실패").slice(0, 500)));
      });
    });

    const pngFiles = fs.existsSync(pngDir)
      ? fs.readdirSync(pngDir).filter(f => /\.png$/i.test(f))
          .sort((a, b) => (parseInt(a.match(/\d+/)?.[0]||0) - parseInt(b.match(/\d+/)?.[0]||0)))
          .map(f => path.join(pngDir, f))
      : [];

    if (pngFiles.length > 0) {
      for (let idx = 0; idx < pngFiles.length; idx++) {
        event.sender.send("pptx:imageProgress", { current: idx + 1, total: pngFiles.length });
        event.sender.send("pptx:imageChunk", {
          index: idx, total: pngFiles.length,
          slide: {
            mediaURL:  `data:image/png;base64,${fs.readFileSync(pngFiles[idx]).toString("base64")}`,
            mediaType: "image", bgColor: null
          }
        });
      }
      fs.rmSync(tmpDir, { recursive: true, force: true });
      return { ok: true, fileName, total: pngFiles.length,
               source: "powerpoint-png", streamed: true };
    }
    fs.rmSync(tmpDir, { recursive: true, force: true });
    return { ok: false, message: "PNG 파일이 생성되지 않았습니다." };

  } catch (e) {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
    return { ok: false, message: e.message };
  }
});

// ──────────────────────────────────────────────────────────
// 실행 중인 PowerPoint에 권한 요청 → 활성 프레젠테이션을 PNG로 변환
// GetActiveObject로 연결, Protected View 자동 탈출, ActivePresentation 사용
// ──────────────────────────────────────────────────────────
ipcMain.handle("pptx:request-ppt-access-and-export", async (event, opts = {}) => {
  if (process.platform !== "win32") return { ok: false, message: "Windows 전용" };

  const { fileName = "슬라이드" } = opts || {};
  const os = require("os");
  const tmpDir = path.join(os.tmpdir(), "pptx_grant_" + Date.now());
  const pngDir = path.join(tmpDir, "pngs");
  fs.mkdirSync(pngDir, { recursive: true });

  // 단일 인용부호만 이스케이프 (PS 단일 인용 문자열에서 역슬래시는 리터럴)
  const psEsc = s => s.replace(/'/g, "''");

  const psScript = `
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Continue"
$outDir = '${psEsc(pngDir)}'

$ppt          = $null
$presentation = $null

try {
    # ── 1. 실행 중인 PowerPoint 연결 대기 (최대 20초 재시도) ────────────
    $waited = 0
    while ($ppt -eq $null -and $waited -lt 20) {
        try {
            $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        } catch {
            Start-Sleep -Seconds 1
            $waited++
        }
    }
    if ($ppt -eq $null) {
        Write-Output "ERROR:PowerPoint가 실행 중이지 않습니다. 파일을 먼저 PowerPoint에서 여세요."
        exit 1
    }
    try { $ppt.DisplayAlerts = 0 } catch {}

    # ── 2. 보호된 보기(Protected View) 탈출 ─────────────────────────────
    if ($ppt.ProtectedViewWindows.Count -gt 0) {
        for ($pi = 1; $pi -le $ppt.ProtectedViewWindows.Count; $pi++) {
            try {
                $pvw = $ppt.ProtectedViewWindows.Item($pi)
                $p2  = $pvw.Edit()
                if ($p2 -ne $null) { $presentation = $p2; Start-Sleep -Seconds 2; break }
            } catch {}
        }
    }

    # ── 3. 활성 프레젠테이션 가져오기 ───────────────────────────────────
    if ($presentation -eq $null) {
        try { $presentation = $ppt.ActivePresentation } catch {}
    }
    if ($presentation -eq $null -and $ppt.Presentations.Count -gt 0) {
        $presentation = $ppt.Presentations.Item(1)
    }
    if ($presentation -eq $null) {
        Write-Output "ERROR:열린 프레젠테이션을 찾을 수 없습니다. PowerPoint에서 파일이 열려 있는지 확인하세요."
        exit 1
    }

    # ── 4. 슬라이드 PNG 내보내기 ────────────────────────────────────────
    $slideCount = $presentation.Slides.Count
    Write-Output "PAGES:$slideCount"
    [Console]::Out.Flush()

    $sw      = $presentation.PageSetup.SlideWidth
    $sh      = $presentation.PageSetup.SlideHeight
    $exportH = if ($sh -gt 0) { [int]([Math]::Round(1920 * ($sh / $sw))) } else { 1080 }

    for ($si = 1; $si -le $slideCount; $si++) {
        $outPath = [System.IO.Path]::Combine($outDir, ("slide{0:D4}.png" -f $si))
        $presentation.Slides.Item($si).Export($outPath, "PNG", 1920, $exportH)
        Write-Output ("DONE:" + $si + ":" + $slideCount)
        [Console]::Out.Flush()
    }

    Write-Output "SUCCESS"

} catch {
    Write-Output "ERROR:$($_.Exception.Message)"
    exit 1
} finally {
    if ($presentation -ne $null) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch {}
        $presentation = $null
    }
    if ($ppt -ne $null) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
        $ppt = $null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
`.trim();

  const psFile = path.join(tmpDir, "grant_export.ps1");
  fs.writeFileSync(psFile,
    Buffer.concat([Buffer.from([0xEF, 0xBB, 0xBF]), Buffer.from(psScript, "utf8")]));

  try {
    await new Promise((resolve, reject) => {
      const { spawn } = require("child_process");
      const child = spawn("powershell",
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", psFile],
        { stdio: ["ignore", "pipe", "pipe"] });
      let stdout = "", stderr = "", totalSlides = 0;

      const parseLine = line => {
        line = line.trim();
        if (line.startsWith("PAGES:")) {
          totalSlides = parseInt(line.slice(6)) || 0;
          event.sender.send("pptx:imageProgress", { current: 0, total: totalSlides });
        } else if (line.startsWith("DONE:")) {
          const parts = line.slice(5).split(":");
          event.sender.send("pptx:imageProgress", {
            current: parseInt(parts[0]) || 0, total: parseInt(parts[1]) || totalSlides
          });
        } else if (line.startsWith("ERROR:")) {
          reject(new Error(line.slice(6)));
        }
      };
      child.stdout.on("data", chunk => {
        const t = chunk.toString("utf8"); stdout += t;
        t.split("\n").forEach(parseLine);
      });
      child.stderr.on("data", chunk => { stderr += chunk.toString("utf8"); });
      child.on("error", reject);
      child.on("close", code => {
        if (code === 0 || stdout.includes("SUCCESS")) resolve();
        else reject(new Error((stderr || stdout || "PowerShell 실패").slice(0, 400)));
      });
    });

    const pngFiles = fs.existsSync(pngDir)
      ? fs.readdirSync(pngDir).filter(f => /\.png$/i.test(f))
          .sort((a, b) => parseInt(a.match(/\d+/)?.[0]||0) - parseInt(b.match(/\d+/)?.[0]||0))
          .map(f => path.join(pngDir, f))
      : [];

    if (pngFiles.length > 0) {
      for (let idx = 0; idx < pngFiles.length; idx++) {
        event.sender.send("pptx:imageProgress", { current: idx + 1, total: pngFiles.length });
        event.sender.send("pptx:imageChunk", {
          index: idx, total: pngFiles.length,
          slide: {
            mediaURL:  `data:image/png;base64,${fs.readFileSync(pngFiles[idx]).toString("base64")}`,
            mediaType: "image", bgColor: null
          }
        });
      }
      fs.rmSync(tmpDir, { recursive: true, force: true });
      return { ok: true, fileName, total: pngFiles.length, source: "powerpoint-grant", streamed: true };
    }
    fs.rmSync(tmpDir, { recursive: true, force: true });
    return { ok: false, message: "PNG 파일이 생성되지 않았습니다." };

  } catch (e) {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch {}
    return { ok: false, message: e.message };
  }
});

/**
 * IPC 핸들러: PPTX를 원본 디자인 그대로 PNG 이미지로 변환한다.
 *
 * 처리 순서:
 *   1. PowerPoint(AppleScript, macOS 전용)로 슬라이드를 PNG로 직접 저장
 *   2. 실패 시 PowerPoint PDF 또는 LibreOffice PDF → PNG 변환으로 폴백
 *   3. PNG 파일을 base64 data URL 배열로 반환
 *
 * options.removeTextShapes: true이면 텍스트 shape을 제거한 PPTX로 렌더링 (텍스트 편집 병행 시 사용)
 * options.aspectMode: "16:9" | "original"
 * options.imageFit: "cover" | "contain"
 */
/**
 * PNG 파일들을 180도 회전시켜 base64 data URL로 반환한다.
 * Electron의 오프스크린 BrowserWindow + canvas를 이용해 외부 라이브러리 없이 처리한다.
 * @param {string[]} pngPaths - PNG 파일 절대 경로 배열
 * @returns {Promise<string[]>} base64 data URL 배열
 */
async function rotatePngs180(pngPaths) {
  if (!pngPaths.length) return [];
  const rotWin = new BrowserWindow({
    width: 1920, height: 1080, show: false, frame: false,
    webPreferences: { nodeIntegration: false, contextIsolation: true }
  });
  try {
    await rotWin.loadURL("data:text/html,<canvas id=c></canvas>");
    const results = [];
    for (const p of pngPaths) {
      const buf = fs.readFileSync(p);
      const src = "data:image/png;base64," + buf.toString("base64");
      const rotated = await rotWin.webContents.executeJavaScript(`
        new Promise(resolve => {
          const img = new Image();
          img.onload = () => {
            const c = document.getElementById('c');
            c.width = img.width; c.height = img.height;
            const ctx = c.getContext('2d');
            ctx.clearRect(0, 0, img.width, img.height);
            ctx.save();
            ctx.translate(img.width / 2, img.height / 2);
            ctx.rotate(Math.PI);
            ctx.scale(-1, 1);
            ctx.drawImage(img, -img.width / 2, -img.height / 2);
            ctx.restore();
            resolve(c.toDataURL('image/png'));
          };
          img.onerror = () => resolve(${JSON.stringify(src)});
          img.src = ${JSON.stringify(src)};
        })
      `);
      results.push(rotated);
    }
    return results;
  } finally {
    rotWin.destroy();
  }
}

ipcMain.handle("pptx:render-to-images", async (event, options = {}) => {
  // 핸들러 전체를 감싸는 안전망 — 예기치 않은 JS 에러가 IPC 예외로 터지는 것을 방지
  try {
  let { filePath } = options || {};

  // filePath가 없으면 파일 선택 다이얼로그 표시
  if (!filePath) {
    const pick = await dialog.showOpenDialog(mainWin, {
      title: "PPTX 원본 디자인으로 가져오기",
      filters: [{ name: "PowerPoint 파일", extensions: ["pptx"] }],
      properties: ["openFile"]
    });
    if (pick.canceled || !pick.filePaths.length) return { ok: false, message: "cancelled" };
    filePath = pick.filePaths[0];
  }

  const aspectMode      = options?.aspectMode === "16:9" ? "16:9" : "original";
  const imageFit        = options?.imageFit === "cover" ? "cover" : "contain";
  const removeTextShapes = options?.removeTextShapes === true;
  const os              = require("os");
  const { execSync, spawnSync } = require("child_process");

  const tmpDir = path.join(os.tmpdir(), "pptx_render_" + Date.now());
  fs.mkdirSync(tmpDir, { recursive: true });
  const collectPngFiles = (rootDir) => {
    const found = [];
    const walk = (dir) => {
      for (const name of fs.readdirSync(dir)) {
        const full = path.join(dir, name);
        const st = fs.statSync(full);
        if (st.isDirectory()) walk(full);
        else if (/\.png$/i.test(name)) found.push(full);
      }
    };
    if (fs.existsSync(rootDir)) walk(rootDir);
    return found.sort((a, b) => {
      const na = parseInt(path.basename(a).match(/\d+/)?.[0] || 0);
      const nb = parseInt(path.basename(b).match(/\d+/)?.[0] || 0);
      return na - nb || a.localeCompare(b);
    });
  };

  let renderFilePath = filePath;
  if (removeTextShapes) {
    try {
      renderFilePath = createPptxWithoutTextShapes(filePath, tmpDir);
    } catch(e) {
      console.warn("텍스트 Shape 제거 실패, 원본으로 폴백:", e.message);
      renderFilePath = filePath;
    }
  }

  const pdfPath = path.join(tmpDir, "presentation.pdf");
  let success   = false;
  let pdfSource = null;
  const renderErrors = [];

  // 1단계: PowerPoint AppleScript로 슬라이드를 PNG 파일로 직접 저장 (macOS 전용)
  const isMac = process.platform === "darwin";
  let permissionDenied = false;

  // PowerPoint 설치 경로 탐색 (기본 경로 + Spotlight 검색)
  let pptAppPath = null;
  if (isMac) {
    const candidatePaths = [
      "/Applications/Microsoft PowerPoint.app",
      "/Applications/Microsoft Office/Microsoft PowerPoint.app",
    ];
    for (const p of candidatePaths) {
      if (fs.existsSync(p) && fs.existsSync(path.join(p, "Contents/MacOS/Microsoft PowerPoint"))) {
        pptAppPath = p;
        break;
      }
    }
    if (!pptAppPath) {
      // Spotlight으로 설치 위치 검색
      try {
        const mdf = spawnSync("mdfind", ["kMDItemCFBundleIdentifier == 'com.microsoft.Powerpoint'"], { timeout: 5000, encoding: "utf8" });
        const found = (mdf.stdout || "").split("\n").find(p => p.trim().endsWith(".app"));
        if (found && fs.existsSync(path.join(found.trim(), "Contents/MacOS/Microsoft PowerPoint"))) {
          pptAppPath = found.trim();
        }
      } catch(e) { /* mdfind 없음 */ }
    }
    if (!pptAppPath) {
      renderErrors.push("Microsoft PowerPoint가 설치되어 있지 않습니다.");
    }
  }

  let isPptLaunchable = false;
  if (pptAppPath) {
    const launchCheck = spawnSync("open", ["-b", "com.microsoft.Powerpoint"], { timeout: 15000, encoding: "utf8" });
    isPptLaunchable = launchCheck.status === 0;
    if (!isPptLaunchable) {
      renderErrors.push("PowerPoint 실행 실패: " + ((launchCheck.stderr || launchCheck.stdout || "").trim() || "PowerPoint를 실행하지 못했습니다."));
    }
  }

  if (isPptLaunchable) {
    const directPngDir = path.join(tmpDir, "powerpoint_pngs");
    fs.mkdirSync(directPngDir, { recursive: true });
    const pngDirForScript = directPngDir.endsWith("/") ? directPngDir : directPngDir + "/";

    // open 후 presentation 객체를 직접 받는 대신 active presentation/front document로 접근
    const script = [
      `tell application ${JSON.stringify(pptAppPath)}`,
      `  activate`,
      `  open POSIX file ${JSON.stringify(renderFilePath)}`,
      `  delay 8`,
      `  try`,
      `    set pres to active presentation`,
      `  on error`,
      `    set pres to presentation 1`,
      `  end try`,
      `  try`,
      `    save pres in POSIX file ${JSON.stringify(pngDirForScript)} as save as PNG`,
      `    delay 5`,
      `    close pres saving no`,
      `  on error e number n`,
      `    try`,
      `      close pres saving no`,
      `    end try`,
      `    error e number n`,
      `  end try`,
      `end tell`,
    ].join("\n");
    const scriptFile = path.join(tmpDir, "export.applescript");
    const scriptBin  = path.join(tmpDir, "export.scpt");
    fs.writeFileSync(scriptFile, script, "utf8");
    try {
      execSync(`osacompile -o "${scriptBin}" "${scriptFile}"`, { timeout: 15000 });
      execSync(`osascript "${scriptBin}"`, { timeout: 300000 });
      const directPngFiles = collectPngFiles(directPngDir);
      if (directPngFiles.length > 0) {
        event.sender.send("pptx:imageProgress", { current: 0, total: directPngFiles.length });
        const rotatedUrls = await rotatePngs180(directPngFiles);
        for (let idx = 0; idx < rotatedUrls.length; idx++) {
          event.sender.send("pptx:imageProgress", { current: idx + 1, total: rotatedUrls.length });
          event.sender.send("pptx:imageChunk", {
            index: idx, total: rotatedUrls.length,
            slide: { mediaURL: rotatedUrls[idx], mediaType: "image", bgColor: null }
          });
        }
        fs.rmSync(tmpDir, { recursive: true, force: true });
        return { ok: true, fileName: path.basename(filePath, path.extname(filePath)), total: rotatedUrls.length, source: "powerpoint-png", streamed: true };
      }
      renderErrors.push("PowerPoint PNG 저장 후 파일이 생성되지 않았습니다.");
    } catch(e) {
      const errMsg = e.message || "";
      if (errMsg.includes("-1743") || errMsg.includes("Not authorized") || errMsg.includes("not allowed to send")) {
        permissionDenied = true;
        renderErrors.push("PowerPoint 자동화 권한 없음(-1743): 시스템 설정에서 이 앱의 Automation 권한을 허용해야 합니다.");
      } else {
        renderErrors.push("PowerPoint PNG 저장 실패: " + errMsg.slice(0, 300));
      }
      console.warn("PowerPoint PNG 저장 실패:", errMsg?.slice(0, 200));
    }
  }

  // 2단계: PowerPoint PDF 폴백 (macOS)
  if (!success && isPptLaunchable) {
    const script = [
      `tell application ${JSON.stringify(pptAppPath)}`,
      `  activate`,
      `  open POSIX file ${JSON.stringify(renderFilePath)}`,
      `  delay 8`,
      `  try`,
      `    set pres to active presentation`,
      `  on error`,
      `    set pres to presentation 1`,
      `  end try`,
      `  try`,
      `    save pres in POSIX file ${JSON.stringify(pdfPath)} as (save as PDF)`,
      `  on error e number n`,
      `    try`,
      `      close pres saving no`,
      `    end try`,
      `    error e number n`,
      `  end try`,
      `  delay 2`,
      `  close pres saving no`,
      `end tell`,
    ].join("\n");
    const scriptFile = path.join(tmpDir, "export_pdf.applescript");
    const scriptBin  = path.join(tmpDir, "export_pdf.scpt");
    fs.writeFileSync(scriptFile, script, "utf8");
    try {
      execSync(`osacompile -o "${scriptBin}" "${scriptFile}"`, { timeout: 15000 });
      execSync(`osascript "${scriptBin}"`, { timeout: 300000 });
      if (fs.existsSync(pdfPath) && fs.statSync(pdfPath).size > 0) {
        success = true;
        pdfSource = "powerpoint-pdf";
      } else {
        renderErrors.push("PowerPoint PDF 저장 후 파일이 생성되지 않았습니다.");
      }
    } catch(e) {
      const errMsg = e.message || "";
      if (errMsg.includes("-1743") || errMsg.includes("Not authorized") || errMsg.includes("not allowed to send")) {
        permissionDenied = true;
        renderErrors.push("PowerPoint 자동화 권한 없음(-1743)");
      } else {
        renderErrors.push("PowerPoint PDF 저장 실패: " + errMsg.slice(0, 300));
      }
      console.warn("PowerPoint PDF 저장 실패:", errMsg?.slice(0, 200));
    }
  }

  // 3단계: Windows PowerPoint COM 자동화 (PowerShell, 비동기)
  const isWindows = process.platform === "win32";
  if (!success && isWindows) {
    const winPngDir = path.join(tmpDir, "powerpoint_pngs");
    fs.mkdirSync(winPngDir, { recursive: true });

    // 경로 이스케이프: 백슬래시를 PowerShell 문자열에 안전하게 삽입
    const psEsc = s => s.replace(/\\/g, "\\\\").replace(/'/g, "''");

    const psScript = `
# UTF-8 출력 설정 (한글 경로 대응)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Continue"

$srcPath = '${psEsc(renderFilePath)}'
$outDir  = '${psEsc(winPngDir)}'

# ── 임시 복사본 생성 ──────────────────────────────────────────
# 1) Protected View: 인터넷 다운로드 파일의 Zone.Identifier 제거
# 2) 한글 경로: COM Open() 이 ASCII 경로를 더 안정적으로 처리
$rnd      = [System.IO.Path]::GetRandomFileName().Replace('.','')
$tempPpt  = [System.IO.Path]::Combine($env:TEMP, "ws_ppt_$rnd.pptx")
$openPath = $srcPath
try {
    Copy-Item -Path $srcPath -Destination $tempPpt -Force -ErrorAction Stop
    # Zone.Identifier ADS 제거 → PowerPoint의 Protected View 방지
    Unblock-File -Path $tempPpt -ErrorAction SilentlyContinue
    $openPath = $tempPpt
} catch {
    $tempPpt = $null  # 복사 실패 시 원본 경로 사용
}

$ppt = $null
$presentation = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible      = -1   # msoTrue = -1 (중요: $true 는 1 이므로 잘못된 값)
    $ppt.DisplayAlerts = 0   # ppAlertsNone — 모든 보안 대화상자 억제

    # 파일 열기 (ReadOnly=0, Untitled=0, WithWindow=-1)
    # 정수를 명시적으로 전달해 COM 타입 혼동 방지
    try {
        $presentation = $ppt.Presentations.Open($openPath, 0, 0, -1)
    } catch {
        $presentation = $null
    }
    Start-Sleep -Seconds 3

    # Protected View 탈출 (Unblock-File 이후에도 남아있을 경우 대비)
    if ($presentation -eq $null -or $ppt.ProtectedViewWindows.Count -gt 0) {
        $pvCount = $ppt.ProtectedViewWindows.Count
        for ($pi = 1; $pi -le $pvCount; $pi++) {
            try {
                $pvw = $ppt.ProtectedViewWindows.Item($pi)
                $p2  = $pvw.Edit()
                if ($p2 -ne $null) {
                    $presentation = $p2
                    Start-Sleep -Seconds 2
                    break
                }
            } catch {}
        }
    }

    # 최종 fallback: 열린 프레젠테이션 목록에서 첫 번째 선택
    if ($presentation -eq $null -and $ppt.Presentations.Count -gt 0) {
        $presentation = $ppt.Presentations.Item(1)
    }

    if ($presentation -eq $null) {
        throw "프레젠테이션을 열 수 없습니다. PowerPoint 보안 설정을 확인해 주세요."
    }

    $slideCount = $presentation.Slides.Count
    Write-Output "PAGES:$slideCount"
    [Console]::Out.Flush()

    $sw = $presentation.PageSetup.SlideWidth
    $sh = $presentation.PageSetup.SlideHeight
    $exportH = if ($sh -gt 0) { [int]([Math]::Round(1920 * ($sh / $sw))) } else { 1080 }

    for ($i = 1; $i -le $slideCount; $i++) {
        $outPath = [System.IO.Path]::Combine($outDir, ("slide{0:D4}.png" -f $i))
        $presentation.Slides.Item($i).Export($outPath, "PNG", 1920, $exportH)
        Write-Output ("DONE:" + $i + ":" + $slideCount)
        [Console]::Out.Flush()
    }

    try { $presentation.Close() } catch {}
    try { $ppt.Quit() } catch {}
    Write-Output "SUCCESS"
} catch {
    $errMsg = $_.Exception.Message
    Write-Output "ERROR:$errMsg"
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    if ($ppt -ne $null) { try { $ppt.Quit() } catch {} }
    exit 1
} finally {
    if ($tempPpt -ne $null) {
        try { Remove-Item -Path $tempPpt -Force -ErrorAction SilentlyContinue } catch {}
    }
    if ($presentation -ne $null) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) | Out-Null } catch {}
        $presentation = $null
    }
    if ($ppt -ne $null) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
        $ppt = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
`.trim();

    const psFile = path.join(tmpDir, "export_ppt.ps1");
    // UTF-8 BOM 추가 → PowerShell 5가 한글 경로를 올바르게 읽도록
    const bom = Buffer.from([0xEF, 0xBB, 0xBF]);
    fs.writeFileSync(psFile, Buffer.concat([bom, Buffer.from(psScript, "utf8")]));

    try {
      // 비동기 spawn → 메인 스레드 블록 방지 + 실시간 진행 이벤트
      await new Promise((resolve, reject) => {
        const { spawn } = require("child_process");
        const child = spawn(
          "powershell",
          ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", psFile],
          { stdio: ["ignore", "pipe", "pipe"] }
        );
        let stdout = "", stderr = "", totalSlides = 0;

        const parseLine = (line) => {
          line = line.trim();
          if (line.startsWith("PAGES:")) {
            totalSlides = parseInt(line.slice(6)) || 0;
            event.sender.send("pptx:imageProgress", { current: 0, total: totalSlides });
          } else if (line.startsWith("DONE:")) {
            const parts = line.slice(5).split(":");
            event.sender.send("pptx:imageProgress", {
              current: parseInt(parts[0]) || 0,
              total:   parseInt(parts[1]) || totalSlides
            });
          }
        };

        child.stdout.on("data", chunk => {
          const text = chunk.toString("utf8");
          stdout += text;
          text.split("\n").forEach(parseLine);
        });
        child.stderr.on("data", chunk => {
          const text = chunk.toString("utf8");
          stderr += text;
          // stderr 에도 PAGES:/DONE:/SUCCESS 가 올 수 있으므로 파싱
          text.split("\n").forEach(parseLine);
        });
        child.on("error", reject);
        child.on("close", code => {
          const combined = stdout + stderr;
          if (code === 0 || combined.includes("SUCCESS")) resolve({ stdout, stderr });
          else reject(new Error((stderr || stdout || "PowerShell 실패").slice(0, 500)));
        });
      });

      const winPngFiles = collectPngFiles(winPngDir);
      if (winPngFiles.length > 0) {
        for (let idx = 0; idx < winPngFiles.length; idx++) {
          event.sender.send("pptx:imageProgress", { current: idx + 1, total: winPngFiles.length });
          event.sender.send("pptx:imageChunk", {
            index: idx, total: winPngFiles.length,
            slide: {
              mediaURL:  `data:image/png;base64,${fs.readFileSync(winPngFiles[idx]).toString("base64")}`,
              mediaType: "image", bgColor: null
            }
          });
        }
        fs.rmSync(tmpDir, { recursive: true, force: true });
        return { ok: true, fileName: path.basename(filePath, path.extname(filePath)), total: winPngFiles.length, source: "powerpoint-png", streamed: true };
      }
      renderErrors.push("Windows PowerPoint PNG 저장 후 파일이 생성되지 않았습니다.");
    } catch(e) {
      const errMsg = (e.message || "").slice(0, 400);
      if (errMsg.includes("0x80040154") || errMsg.includes("retrieving the COM") || errMsg.includes("80080005") || errMsg.includes("PowerPoint.Application") || errMsg.includes("클래스가 등록되지")) {
        renderErrors.push("Microsoft PowerPoint가 설치되지 않았거나 COM 등록이 손상되었습니다.");
      } else if (errMsg.includes("Protected View") || errMsg.includes("보호된 보기") || errMsg.includes("보안 설정")) {
        renderErrors.push("PowerPoint 보안 설정으로 인해 파일을 열 수 없습니다. PowerPoint > 파일 > 옵션 > 보안 센터 > 보호된 보기 설정을 확인해 주세요.");
      } else if (errMsg.includes("0x80010001") || errMsg.includes("RPC_E_CALL_REJECTED")) {
        renderErrors.push("PowerPoint가 응답하지 않습니다. PowerPoint를 완전히 닫은 후 다시 시도해 주세요.");
      } else {
        renderErrors.push("Windows PowerPoint COM 실패: " + errMsg);
      }
      console.warn("Windows PowerPoint COM 실패:", errMsg);
    }
  }

  // 4단계: LibreOffice PPTX → PNG 직접 변환 (크로스플랫폼 폴백)
  // PNG 직접 변환을 먼저 시도하고, 실패 시 PDF → PNG 경로로 폴백
  if (!success) {
    const libreOfficePaths = [
      "/Applications/LibreOffice.app/Contents/MacOS/soffice",
      "/usr/local/bin/soffice", "/usr/bin/soffice",
      "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
      "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ];
    for (const soffice of libreOfficePaths) {
      if (!fs.existsSync(soffice)) continue;

      // 1차 시도: PPTX → PNG 직접 변환 (PDF 중간 단계 없음)
      const loPngDir = path.join(tmpDir, "lo_pngs");
      fs.mkdirSync(loPngDir, { recursive: true });
      try {
        await new Promise((resolve, reject) => {
          const { spawn } = require("child_process");
          const child = spawn(soffice, [
            "--headless", "--convert-to", "png",
            "--outdir", loPngDir, renderFilePath
          ], { stdio: ["ignore", "pipe", "pipe"] });
          let stderr = "";
          child.stderr.on("data", d => { stderr += d.toString(); });
          child.on("error", reject);
          child.on("close", code => {
            if (code === 0) resolve();
            else reject(new Error(stderr || "LibreOffice 실패"));
          });
        });
        const loPngFiles = collectPngFiles(loPngDir);
        if (loPngFiles.length > 0) {
          // macOS LibreOffice는 회전 보정 적용, Windows는 정방향이므로 그대로
          const urls = isMac
            ? await rotatePngs180(loPngFiles)
            : loPngFiles.map(f => `data:image/png;base64,${fs.readFileSync(f).toString("base64")}`);
          for (let idx = 0; idx < urls.length; idx++) {
            event.sender.send("pptx:imageProgress", { current: idx + 1, total: urls.length });
            event.sender.send("pptx:imageChunk", {
              index: idx, total: urls.length,
              slide: { mediaURL: urls[idx], mediaType: "image", bgColor: null }
            });
          }
          fs.rmSync(tmpDir, { recursive: true, force: true });
          return { ok: true, fileName: path.basename(filePath, path.extname(filePath)), total: urls.length, source: "libreoffice-png", streamed: true };
        }
      } catch(e) {
        renderErrors.push("LibreOffice PNG 변환 실패: " + (e.message || "").slice(0, 100));
      }

      // 2차 시도: PPTX → PDF (이후 PDF → PNG 단계에서 처리)
      try {
        execSync(`"${soffice}" --headless --convert-to pdf --outdir "${tmpDir}" "${renderFilePath}"`, { timeout: 120000 });
        const pdfs = fs.readdirSync(tmpDir).filter(f => /\.pdf$/i.test(f));
        if (pdfs.length > 0) {
          fs.renameSync(path.join(tmpDir, pdfs[0]), pdfPath);
          success = true;
          pdfSource = "libreoffice-pdf";
          break;
        }
      } catch(e) {
        renderErrors.push("LibreOffice PDF 변환 실패: " + (e.message || "").slice(0, 100));
        console.warn("LibreOffice PDF 실패:", (e.message || "").slice(0, 100));
      }
    }
  }

  // 5단계: Electron 내장 렌더러 폴백
  // PowerPoint/LibreOffice 없이도 동작 — Windows 포함 모든 플랫폼에서 보장된 경로
  if (!success && !permissionDenied) {
    try {
      const electronSlides = await renderPptxSlidesWithElectron(renderFilePath, event);
      fs.rmSync(tmpDir, { recursive: true, force: true });
      return {
        ok: true,
        fileName: path.basename(filePath, path.extname(filePath)),
        total: electronSlides.length,
        source: "electron",
        streamed: true
      };
    } catch(e) {
      renderErrors.push("Electron 렌더링 폴백 실패: " + (e.message || "").slice(0, 100));
    }
  }

  if (!success) {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    if (permissionDenied) {
      return { ok: false, permissionDenied: true, message: "PowerPoint 자동화 권한이 없습니다." };
    }
    let hint = "";
    if (isWindows) {
      hint = "PPTX 이미지 변환에 실패했습니다.\n" +
        "PowerPoint가 실행 중이라면 완전히 닫은 후 다시 시도해 주세요.";
    } else if (!pptAppPath) {
      hint = "Microsoft PowerPoint가 설치되어 있지 않습니다.\nLibreOffice를 설치하면 대체 변환이 가능합니다.";
    } else {
      hint = "PowerPoint 변환에 실패했습니다.\n자동화(Automation) 권한이 허용되어 있는지 확인해 주세요.";
    }
    const detail = renderErrors.length ? "\n\n상세:\n- " + renderErrors.join("\n- ") : "";
    return { ok: false, message: hint + detail };
  }

  // 6단계: PDF → PNG 변환 (플랫폼별)
  const pngDir = path.join(tmpDir, "pngs");
  fs.mkdirSync(pngDir, { recursive: true });
  let pdfPngSuccess = false;

  if (isMac) {
    // macOS: Python + CoreGraphics 프레임워크
    const pythonScript = `
import ctypes, os, sys

cf = ctypes.CDLL('/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation')
cg = ctypes.CDLL('/System/Library/Frameworks/CoreGraphics.framework/CoreGraphics')

class CGRect(ctypes.Structure):
    _fields_ = [('x', ctypes.c_double), ('y', ctypes.c_double),
                ('width', ctypes.c_double), ('height', ctypes.c_double)]

cf.CFStringCreateWithCString.restype = ctypes.c_void_p
cf.CFStringCreateWithCString.argtypes = [ctypes.c_void_p, ctypes.c_char_p, ctypes.c_uint32]
cf.CFURLCreateWithFileSystemPath.restype = ctypes.c_void_p
cf.CFURLCreateWithFileSystemPath.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_long, ctypes.c_bool]
cf.CFRelease.argtypes = [ctypes.c_void_p]
cg.CGPDFDocumentCreateWithURL.restype = ctypes.c_void_p
cg.CGPDFDocumentCreateWithURL.argtypes = [ctypes.c_void_p]
cg.CGPDFDocumentGetNumberOfPages.restype = ctypes.c_size_t
cg.CGPDFDocumentGetNumberOfPages.argtypes = [ctypes.c_void_p]
cg.CGPDFDocumentGetPage.restype = ctypes.c_void_p
cg.CGPDFDocumentGetPage.argtypes = [ctypes.c_void_p, ctypes.c_size_t]
cg.CGPDFPageGetBoxRect.restype = CGRect
cg.CGPDFPageGetBoxRect.argtypes = [ctypes.c_void_p, ctypes.c_int]
cg.CGColorSpaceCreateDeviceRGB.restype = ctypes.c_void_p
cg.CGColorSpaceRelease.argtypes = [ctypes.c_void_p]
cg.CGBitmapContextCreate.restype = ctypes.c_void_p
cg.CGBitmapContextCreate.argtypes = [ctypes.c_void_p, ctypes.c_size_t, ctypes.c_size_t,
    ctypes.c_size_t, ctypes.c_size_t, ctypes.c_void_p, ctypes.c_uint32]
cg.CGContextRelease.argtypes = [ctypes.c_void_p]
cg.CGContextDrawPDFPage.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
cg.CGContextTranslateCTM.argtypes = [ctypes.c_void_p, ctypes.c_double, ctypes.c_double]
cg.CGContextScaleCTM.argtypes = [ctypes.c_void_p, ctypes.c_double, ctypes.c_double]
cg.CGContextSetRGBFillColor.argtypes = [ctypes.c_void_p, ctypes.c_double, ctypes.c_double, ctypes.c_double, ctypes.c_double]
cg.CGContextFillRect.argtypes = [ctypes.c_void_p, CGRect]

def _save_bgra_png(filename, width, height, data):
    import zlib, struct
    def make_chunk(tag, d):
        return struct.pack(">I", len(d)) + tag + d + struct.pack(">I", zlib.crc32(tag + d) & 0xffffffff)
    sig = b'\\x89PNG\\r\\n\\x1a\\n'
    ihdr = make_chunk(b'IHDR', struct.pack(">IIBBBBB", width, height, 8, 6, 0, 0, 0))
    mv = memoryview(data)
    rgba = bytearray(width * height * 4)
    for i in range(0, len(mv), 4):
        b, g, r, a = mv[i:i+4]
        rgba[i] = r; rgba[i+1] = g; rgba[i+2] = b; rgba[i+3] = a
    scanlines = bytearray()
    for y in range(height):
        scanlines.append(0)
        scanlines.extend(rgba[y * width * 4 : (y + 1) * width * 4])
    idat = make_chunk(b'IDAT', zlib.compress(scanlines))
    iend = make_chunk(b'IEND', b'')
    with open(filename, 'wb') as f:
        f.write(sig + ihdr + idat + iend)

pdf_path, out_dir = sys.argv[1], sys.argv[2]
target_w, target_h = 1920, 1080
aspect_mode = ${JSON.stringify(aspectMode)}
image_fit = ${JSON.stringify(imageFit)}

path_str = cf.CFStringCreateWithCString(None, pdf_path.encode('utf-8'), 134217984)
url = cf.CFURLCreateWithFileSystemPath(None, path_str, 0, False)
cf.CFRelease(path_str)
doc = cg.CGPDFDocumentCreateWithURL(url)
cf.CFRelease(url)
if not doc:
    print("ERROR:PDF open failed"); sys.exit(1)

page_count = cg.CGPDFDocumentGetNumberOfPages(doc)
print(f"PAGES:{page_count}", flush=True)

for i in range(1, page_count + 1):
    page = cg.CGPDFDocumentGetPage(doc, i)
    if not page: continue
    box = cg.CGPDFPageGetBoxRect(page, 0)
    if aspect_mode == "16:9":
        rw, rh = target_w, target_h
        scale = max(target_w / box.width, target_h / box.height) if image_fit == "cover" else min(target_w / box.width, target_h / box.height)
        dx = (target_w - box.width * scale) / 2.0
        dy = (target_h - box.height * scale) / 2.0
    else:
        scale = min(target_w / box.width, target_h / box.height)
        rw, rh = int(box.width * scale), int(box.height * scale)
        dx, dy = 0, 0
    colorspace = cg.CGColorSpaceCreateDeviceRGB()
    bpr = rw * 4
    buf = ctypes.create_string_buffer(bpr * rh)
    ctx = cg.CGBitmapContextCreate(buf, rw, rh, 8, bpr, colorspace, ctypes.c_uint32(0x2002))
    cg.CGColorSpaceRelease(colorspace)
    if not ctx: continue
    cg.CGContextSetRGBFillColor(ctx, 1.0, 1.0, 1.0, 1.0)
    cg.CGContextFillRect(ctx, CGRect(0, 0, rw, rh))
    cg.CGContextTranslateCTM(ctx, dx, rh - dy)
    cg.CGContextScaleCTM(ctx, scale, -scale)
    cg.CGContextTranslateCTM(ctx, -box.x, -box.y)
    cg.CGContextDrawPDFPage(ctx, page)
    cg.CGContextRelease(ctx)
    raw = ctypes.string_at(buf, bpr * rh)
    out_path = os.path.join(out_dir, f'slide{i:04d}.png')
    _save_bgra_png(out_path, rw, rh, raw)
    print(f"DONE:{i}:{page_count}", flush=True)
`;
    const pyFile = path.join(tmpDir, "pdf2png.py");
    fs.writeFileSync(pyFile, pythonScript, "utf8");
    try {
      await new Promise((resolve, reject) => {
        const { spawn } = require("child_process");
        const child = spawn("python3", [pyFile, pdfPath, pngDir], { stdio: ["ignore", "pipe", "pipe"] });
        let stdout = "", stderr = "";
        child.stdout.on("data", chunk => {
          const text = chunk.toString("utf8");
          stdout += text;
          for (const raw of text.split("\n")) {
            const line = raw.trim();
            if (line.startsWith("DONE:")) {
              const [, cur, tot] = line.split(":");
              event.sender.send("pptx:imageProgress", { current: parseInt(cur), total: parseInt(tot) });
            }
            if (line.startsWith("ERROR:")) reject(new Error(line.slice(6)));
          }
        });
        child.stderr.on("data", d => { stderr += d.toString(); });
        child.on("error", reject);
        child.on("close", code => {
          if (code === 0) resolve({ stdout, stderr });
          else reject(new Error(stderr || stdout || "python3 실패"));
        });
      });
      pdfPngSuccess = true;
    } catch(e) {
      renderErrors.push("PDF→PNG(Python) 실패: " + (e.message || "").slice(0, 200));
    }
  }

  if (!pdfPngSuccess && isWindows) {
    // Windows: LibreOffice PDF → PNG 변환
    const libreOfficePaths = [
      "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
      "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ];
    for (const soffice of libreOfficePaths) {
      if (!fs.existsSync(soffice)) continue;
      try {
        await new Promise((resolve, reject) => {
          const { spawn } = require("child_process");
          const child = spawn(soffice, [
            "--headless", "--convert-to", "png",
            "--outdir", pngDir, pdfPath
          ], { stdio: ["ignore", "pipe", "pipe"] });
          let stderr = "";
          child.stderr.on("data", d => { stderr += d.toString(); });
          child.on("error", reject);
          child.on("close", code => {
            if (code === 0) resolve();
            else reject(new Error(stderr || "LibreOffice PDF→PNG 실패"));
          });
        });
        if (collectPngFiles(pngDir).length > 0) { pdfPngSuccess = true; break; }
      } catch(e) {
        renderErrors.push("LibreOffice PDF→PNG 실패: " + (e.message || "").slice(0, 100));
      }
    }
  }

  if (!pdfPngSuccess && isWindows) {
    // Windows: PowerShell + Windows.Data.Pdf (Windows 10/11 내장 API)
    const winPdfPngScript = `
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]

function Await($wrt) {
    $task = $wrt.AsTask()
    $task.Wait()
    return $task.Result
}

$pdfPath = ${JSON.stringify(pdfPath)}
$outDir  = ${JSON.stringify(pngDir)}

try {
    $file = Await([Windows.Storage.StorageFile]::GetFileFromPathAsync($pdfPath))
    $doc  = Await([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file))
    $pageCount = $doc.PageCount
    Write-Host "PAGES:$pageCount"

    for ($i = 0; $i -lt $pageCount; $i++) {
        $page = $doc.GetPage($i)
        $outPath = [System.IO.Path]::Combine($outDir, ("slide{0:D4}.png" -f ($i + 1)))
        $stream  = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
        $opts    = [Windows.Data.Pdf.PdfPageRenderOptions]::new()
        $opts.DestinationWidth = 1920
        Await($page.RenderToStreamAsync($stream, $opts)) | Out-Null
        $bytes = [byte[]]::new($stream.Size)
        $ibr   = $stream.GetInputStreamAt(0)
        $reader = [Windows.Storage.Streams.DataReader]::new($ibr)
        Await($reader.LoadAsync($stream.Size)) | Out-Null
        $reader.ReadBytes($bytes)
        [System.IO.File]::WriteAllBytes($outPath, $bytes)
        $stream.Dispose()
        Write-Host "DONE:$($i+1):$pageCount"
    }
    Write-Host "SUCCESS"
} catch {
    Write-Host "ERROR:$_"
    exit 1
}
`.trim();
    const wpsFile = path.join(tmpDir, "pdf2png_win.ps1");
    fs.writeFileSync(wpsFile, winPdfPngScript, "utf8");
    try {
      await new Promise((resolve, reject) => {
        const { spawn } = require("child_process");
        const child = spawn("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", wpsFile], {
          stdio: ["ignore", "pipe", "pipe"]
        });
        let stdout = "", stderr = "", total = 0;
        child.stdout.on("data", chunk => {
          const text = chunk.toString("utf8");
          stdout += text;
          for (const raw of text.split("\n")) {
            const line = raw.trim();
            if (line.startsWith("PAGES:")) total = parseInt(line.slice(6)) || 0;
            else if (line.startsWith("DONE:")) {
              const parts = line.slice(5).split(":");
              event.sender.send("pptx:imageProgress", { current: parseInt(parts[0]) || 0, total: parseInt(parts[1]) || total });
            }
          }
        });
        child.stderr.on("data", d => { stderr += d.toString(); });
        child.on("error", reject);
        child.on("close", code => {
          if (code === 0 || stdout.includes("SUCCESS")) resolve({ stdout, stderr });
          else reject(new Error((stderr || stdout || "PowerShell PDF→PNG 실패").slice(0, 300)));
        });
      });
      if (collectPngFiles(pngDir).length > 0) pdfPngSuccess = true;
    } catch(e) {
      renderErrors.push("Windows.Data.Pdf PDF→PNG 실패: " + (e.message || "").slice(0, 100));
    }
  }

  if (!pdfPngSuccess) {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    const detail = renderErrors.length ? "\n\n상세:\n- " + renderErrors.join("\n- ") : "";
    return { ok: false, message: "PDF→PNG 변환에 실패했습니다." + detail };
  }

  try {
    const pngFiles = collectPngFiles(pngDir);
    if (!pngFiles.length) {
      fs.rmSync(tmpDir, { recursive: true, force: true });
      return { ok: false, message: "PNG 파일이 생성되지 않았습니다." };
    }

    // macOS는 회전 보정, Windows는 정방향이므로 그대로
    const finalUrls = isMac
      ? await rotatePngs180(pngFiles)
      : pngFiles.map(f => `data:image/png;base64,${fs.readFileSync(f).toString("base64")}`);

    for (let idx = 0; idx < finalUrls.length; idx++) {
      event.sender.send("pptx:imageProgress", { current: idx + 1, total: finalUrls.length });
      event.sender.send("pptx:imageChunk", {
        index: idx,
        total: finalUrls.length,
        slide: { mediaURL: finalUrls[idx], mediaType: "image", bgColor: null }
      });
    }

    fs.rmSync(tmpDir, { recursive: true, force: true });
    return { ok: true, fileName: path.basename(filePath, path.extname(filePath)), total: finalUrls.length, source: pdfSource || "pdf", streamed: true };
  } catch(err) {
    fs.rmSync(tmpDir, { recursive: true, force: true });
    return { ok: false, message: "이미지 읽기 실패: " + err.message };
  }
  } catch (fatalErr) {
    // 예기치 않은 최상위 오류 — IPC 예외 대신 ok:false 로 반환
    console.error("pptx:render-to-images 치명적 오류:", fatalErr);
    return { ok: false, message: "변환 중 오류가 발생했습니다: " + (fatalErr.message || String(fatalErr)) };
  }
});

// ──────────────────────────────────────────────────────────
// PPTX XML 파싱 헬퍼 함수들
// ──────────────────────────────────────────────────────────

/**
 * PPTX 관계 파일(.rels)의 XML을 파싱해 ID → 파일경로 매핑을 반환한다.
 * PPTX에서 이미지나 다른 파일을 참조할 때 ID를 사용하므로
 * 실제 파일 경로를 찾으려면 이 매핑이 필요하다.
 *
 * @param {string} xml - .rels 파일의 XML 텍스트
 * @returns {Object} { rId: targetPath } 형태의 매핑 객체
 */
function parseRelsXml(xml) {
  const map = {};
  const re  = /<Relationship\s[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
  let m;
  // Windows에서 생성된 PPTX는 백슬래시를 사용하므로 슬래시로 정규화
  while ((m = re.exec(xml)) !== null) map[m[1]] = m[2].replace(/\\/g, "/");
  return map;
}

/**
 * PPTX 테마 파일에서 색상 팔레트와 폰트 스킴을 추출한다.
 * 슬라이드의 색상/폰트가 테마를 참조할 때 실제 값을 구하는 데 사용한다.
 *
 * @param {AdmZip} zip - 열린 PPTX ZIP 객체
 * @returns {Object} { colors: {dk1, lt1, ...}, majorFont, minorFont, ... }
 */
function extractPptxTheme(zip) {
  const theme = { colors: {}, majorFont: '', minorFont: '', majorFontEa: '', minorFontEa: '', fillStyles: [] };
  try {
    let themeEntry = zipGetEntry(zip, 'ppt/theme/theme1.xml');
    if (!themeEntry) {
      // theme1.xml이 없으면 다른 번호의 테마 파일 탐색
      themeEntry = zip.getEntries().find(e => /^ppt\/theme\/theme\d+\.xml$/i.test(e.entryName.replace(/\\/g, '/'))) || null;
    }
    if (!themeEntry) return theme;
    const xml = themeEntry.getData().toString('utf8');

    // 색상 스킴 추출 (dk1=어두운색1, lt1=밝은색1, accent1~6=강조색 등)
    for (const name of ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink']) {
      const re2  = new RegExp('<a:' + name + '>([\\s\\S]{0,300}?)<\\/a:' + name + '>');
      const m2   = xml.match(re2);
      if (!m2) continue;
      const srgb = m2[1].match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
      const sys  = m2[1].match(/<a:sysClr\s+[^>]*lastClr="([0-9A-Fa-f]{6})"/);
      if (srgb) theme.colors[name] = srgb[1];
      else if (sys) theme.colors[name] = sys[1];
    }

    // 폰트 스킴 (major=제목용, minor=본문용, Ea=동아시아 폰트)
    const mjLa = xml.match(/<a:majorFont>[\s\S]*?<a:latin\s+typeface="([^"]+)"/);
    const mnLa = xml.match(/<a:minorFont>[\s\S]*?<a:latin\s+typeface="([^"]+)"/);
    const mjEa = xml.match(/<a:majorFont>[\s\S]*?<a:ea\s+typeface="([^"]+)"/);
    const mnEa = xml.match(/<a:minorFont>[\s\S]*?<a:ea\s+typeface="([^"]+)"/);
    if (mjLa) theme.majorFont   = mjLa[1];
    if (mnLa) theme.minorFont   = mnLa[1];
    if (mjEa) theme.majorFontEa = mjEa[1];
    if (mnEa) theme.minorFontEa = mnEa[1];

    // fillStyleLst를 읽어 fillRef/bgRef idx 역참조에 사용
    const fillLstM = xml.match(/<a:fillStyleLst>([\s\S]*?)<\/a:fillStyleLst>/);
    if (fillLstM) {
      const fillChunk = fillLstM[1];
      const fillRe = /<(a:solidFill|a:gradFill|a:pattFill)\b([\s\S]*?)<\/\1>/g;
      let fm;
      while ((fm = fillRe.exec(fillChunk)) !== null) {
        const part = fm[0];
        const srgb = part.match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
        const sys = part.match(/<a:sysClr\s+[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"/);
        const sch = part.match(/<a:schemeClr\s+val="([^"]+)"/);
        if (srgb) theme.fillStyles.push("#" + srgb[1]);
        else if (sys) theme.fillStyles.push("#" + sys[1]);
        else if (sch && theme.colors[sch[1]]) theme.fillStyles.push("#" + theme.colors[sch[1]]);
        else theme.fillStyles.push(null);
      }
    }
  } catch(e) {}
  return theme;
}

/**
 * 대소문자를 구분하지 않고 ZIP 내 항목을 찾는다.
 * macOS APFS 같은 대소문자 구분 파일시스템 환경에서도 올바르게 동작한다.
 *
 * @param {AdmZip} zip       - 열린 ZIP 객체
 * @param {string} entryPath - 찾을 항목의 경로
 * @returns {ZipEntry|null} 찾은 항목, 없으면 null
 */
function zipGetEntry(zip, entryPath) {
  const norm  = entryPath.replace(/\\/g, "/");
  let entry   = zip.getEntry(norm);
  if (entry) return entry;
  const lower = norm.toLowerCase();
  return zip.getEntries().find(e => e.entryName.replace(/\\/g, "/").toLowerCase() === lower) || null;
}

/**
 * XML에서 HTML 특수문자 엔티티를 원래 문자로 복원한다.
 * PPTX XML 텍스트에 &amp; &lt; 등이 포함될 수 있다.
 *
 * @param {string} t - 엔티티가 포함된 텍스트
 * @returns {string} 복원된 텍스트
 */
function xmlDecodeText(t) {
  return String(t || "")
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

/**
 * 관계 ID(rId)에 해당하는 이미지를 찾아 base64 data URL로 반환한다.
 * 경로 규칙이 다양하므로 여러 후보 경로를 순서대로 시도한다.
 *
 * @param {string} rId      - 관계 ID (예: 'rId1')
 * @param {Object} relsMap  - parseRelsXml() 결과
 * @param {AdmZip} zip      - 열린 ZIP 객체
 * @param {string} baseDir  - 기본 디렉터리 (예: 'ppt/')
 * @returns {string|null} base64 data URL, 또는 null
 */
function getImageBase64(rId, relsMap, zip, baseDir) {
  let target = relsMap[rId];
  if (!target) return null;
  target = target.replace(/\\/g, "/");

  // 후보 경로 생성 (PPTX 내 상대 경로 규칙이 파일마다 다를 수 있음)
  const tries = [];
  if (target.startsWith("../")) {
    tries.push(baseDir + target.slice(3));
    tries.push("ppt/" + target.slice(3));
  } else if (target.startsWith("/")) {
    tries.push(target.slice(1));
  } else if (/^https?:\/\//.test(target)) {
    return null;  // 외부 URL은 지원 안 함
  } else {
    tries.push(baseDir + target);
    tries.push("ppt/" + target);
    tries.push(target);
  }
  // 파일명만으로 ppt/media/ 하위 재탐색 (경로가 틀려도 파일명으로 찾기)
  const basename = target.split("/").pop();
  if (basename) tries.push("ppt/media/" + basename);

  let entry = null;
  for (const p of tries) {
    entry = zipGetEntry(zip, p);
    if (entry) break;
  }
  if (!entry) return null;

  // MIME 타입 결정
  const ext   = path.extname(entry.entryName).slice(1).toLowerCase();
  const mimes = { jpg:"image/jpeg", jpeg:"image/jpeg", png:"image/png", gif:"image/gif", bmp:"image/bmp", webp:"image/webp", tiff:"image/tiff", svg:"image/svg+xml", emf:"image/x-emf", wmf:"image/x-wmf" };
  const mime  = mimes[ext] || "image/jpeg";
  return "data:" + mime + ";base64," + entry.getData().toString("base64");
}

function resolveOpenXmlTargetPath(currentDir, targetPath) {
  let t = String(targetPath || "").replace(/\\/g, "/");
  const cur = String(currentDir || "").replace(/\\/g, "/");
  if (!t) return "";
  if (t.startsWith("/")) return t.slice(1);
  if (/^https?:\/\//i.test(t)) return "";
  const baseParts = cur.split("/").filter(Boolean);
  const targetParts = t.split("/").filter(Boolean);
  while (targetParts.length && targetParts[0] === "..") {
    targetParts.shift();
    if (baseParts.length) baseParts.pop();
  }
  while (targetParts.length && targetParts[0] === ".") {
    targetParts.shift();
  }
  return [...baseParts, ...targetParts].join("/");
}

/**
 * PPTX 슬라이드 XML을 파싱해 앱 슬라이드 객체로 변환한다.
 *
 * 파싱 순서:
 *   1. 배경: <p:bg> 배경 이미지/색상 → 슬라이드 레이아웃 → 슬라이드 마스터 순으로 탐색
 *   2. <p:pic> 이미지 요소 중 가장 큰 것을 배경으로 사용
 *   3. <p:sp> 텍스트 도형들을 파싱해 폰트/위치/내용 추출
 *   4. placeholder 타입, 폰트 크기 등을 기준으로 메인 가사 텍스트 식별
 *   5. 나머지 텍스트는 textBoxes 배열에 추가
 *
 * @param {string} xml     - 슬라이드 XML 텍스트
 * @param {Object} relsMap - 관계 ID → 파일경로 매핑
 * @param {AdmZip} zip     - 열린 ZIP 객체
 * @param {number} slideW  - 슬라이드 너비 (EMU)
 * @param {number} slideH  - 슬라이드 높이 (EMU)
 * @param {Object} theme   - extractPptxTheme() 결과
 * @returns {Object} 앱 슬라이드 객체
 */
function parsePptxSlide(xml, relsMap, zip, slideW, slideH, theme = {}) {
  const slide = {
    lyrics: "", sub: "",
    fontSize: 36, fontColor: "#ffffff", fontFamily: "Noto Serif KR",
    textX: 50, textY: 50, textWidth: null,
    textShadowEnabled: true, textShadowIntensity: 80,
    textOutlineEnabled: false, textOutlineWidth: 2,
    layoutType: "centerLyrics",
    mediaURL: null, mediaType: null,
    bgColor: null,
    textBoxes: [],
    shapeBoxes: []
  };

  const baseDir = "ppt/";
  const emuToPct = (value, total) => {
    const num = Number(value);
    const den = Number(total);
    if (!Number.isFinite(num) || !Number.isFinite(den) || den <= 0) return 0;
    return Math.round((num / den) * 100 * 10) / 10;
  };
  const resolveSchemeColor = val => (theme.colors && theme.colors[val]) || null;
  const resolveFillRefColor = (idxRaw) => {
    const idx = parseInt(String(idxRaw || ""), 10);
    if (!Number.isFinite(idx) || !Array.isArray(theme.fillStyles) || !theme.fillStyles.length) return null;
    // PPTX fillRef/bgRef idx는 문서에 따라 1-based 혹은 100x/1000x 계열로 나타날 수 있어 보정
    const candidates = [
      idx,
      idx - 1000,
      idx - 1001,
      idx - 100,
      idx - 1
    ].filter(n => Number.isFinite(n) && n >= 1 && n <= theme.fillStyles.length);
    for (const c of candidates) {
      const color = theme.fillStyles[c - 1];
      if (color) return color;
    }
    return null;
  };
  const readFillColorHex = (chunk = "") => {
    if (!chunk) return null;
    const srgb = chunk.match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
    if (srgb) return "#" + srgb[1];
    const sys = chunk.match(/<a:sysClr\s+[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"/);
    if (sys) return "#" + sys[1];
    const grad = chunk.match(/<a:gradFill[\s\S]*?<a:gs\b[^>]*>[\s\S]*?(?:<a:srgbClr\s+val="([0-9A-Fa-f]{6})"|<a:sysClr\s+[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"|<a:schemeClr\s+val="([^"]+)")/);
    if (grad) {
      if (grad[1]) return "#" + grad[1];
      if (grad[2]) return "#" + grad[2];
      if (grad[3]) {
        const resolved = resolveSchemeColor(grad[3]);
        if (resolved) return "#" + resolved.replace(/^#/, "");
      }
    }
    const sch = chunk.match(/<a:schemeClr\s+val="([^"]+)"/);
    if (sch) {
      const resolved = resolveSchemeColor(sch[1]);
      if (resolved) return "#" + resolved.replace(/^#/, "");
    }
    const fillRef = chunk.match(/<(?:a:fillRef|a:bgRef|p:bgRef)\s+[^>]*\bidx="(\d+)"/);
    if (fillRef) {
      const refColor = resolveFillRefColor(fillRef[1]);
      if (refColor) return refColor;
    }
    return null;
  };

  // ── 1. 배경 처리 ──
  // 우선순위: 슬라이드 직접 지정 > 슬라이드 레이아웃 > 슬라이드 마스터

  // (a) <p:bg> 안의 배경 이미지
  const bgBlipM = xml.match(/<p:bg>[\s\S]*?<a:blip\s[^>]*r:embed="([^"]+)"/);
  if (bgBlipM) {
    const dataUrl = getImageBase64(bgBlipM[1], relsMap, zip, baseDir);
    if (dataUrl) { slide.mediaURL = dataUrl; slide.mediaType = "image"; }
  }

  // (a2) <p:bg> 안의 단색 배경
  if (!slide.mediaURL) {
    const bgChunkM = xml.match(/<p:bg>([\s\S]*?)<\/p:bg>/);
    const bgColor = bgChunkM ? readFillColorHex(bgChunkM[1]) : null;
    if (bgColor) slide.bgColor = bgColor;
  }

  // (a3) 슬라이드 레이아웃 → 슬라이드 마스터 배경 확인 (슬라이드 자체 배경이 없을 때)
  if (!slide.mediaURL && !slide.bgColor) {
    const layoutRel = Object.entries(relsMap).find(([, t]) => t.includes("slideLayout"));
    if (layoutRel) {
      const layoutTarget = layoutRel[1].replace(/\\/g, "/");
      const layoutPath   = resolveOpenXmlTargetPath("ppt/slides", layoutTarget);
      try {
        const layoutEntry = zipGetEntry(zip, layoutPath);
        if (layoutEntry) {
          const lxml      = layoutEntry.getData().toString("utf8");
          const lRelsPath  = layoutPath.replace(/\/([^/]+)$/, "/_rels/$1.rels");
          const lRelsEntry = zipGetEntry(zip, lRelsPath);
          const lRelsMap   = lRelsEntry ? parseRelsXml(lRelsEntry.getData().toString("utf8")) : {};

          const lBgBlip = lxml.match(/<p:bg>[\s\S]*?<a:blip\s[^>]*r:embed="([^"]+)"/);
          if (lBgBlip) {
            const lDataUrl = getImageBase64(lBgBlip[1], lRelsMap, zip, "ppt/");
            if (lDataUrl) { slide.mediaURL = lDataUrl; slide.mediaType = "image"; }
          }
          if (!slide.mediaURL) {
            const lBgChunkM = lxml.match(/<p:bg>([\s\S]*?)<\/p:bg>/);
            const lBgColor = lBgChunkM ? readFillColorHex(lBgChunkM[1]) : null;
            if (lBgColor) slide.bgColor = lBgColor;
          }

          // 레이아웃에도 배경이 없으면 슬라이드 마스터 확인
          if (!slide.mediaURL && !slide.bgColor) {
            const masterRel = Object.entries(lRelsMap).find(([, t]) => t.includes("slideMaster"));
            if (masterRel) {
              const masterTarget = masterRel[1].replace(/\\/g, "/");
              const layoutDir = layoutPath.replace(/\/[^/]+$/, "");
              const masterPath = resolveOpenXmlTargetPath(layoutDir, masterTarget);
              try {
                const masterEntry = zipGetEntry(zip, masterPath);
                if (masterEntry) {
                  const mxml      = masterEntry.getData().toString("utf8");
                  const mRelsPath  = masterPath.replace(/\/([^/]+)$/, "/_rels/$1.rels");
                  const mRelsEntry = zipGetEntry(zip, mRelsPath);
                  const mRelsMap   = mRelsEntry ? parseRelsXml(mRelsEntry.getData().toString("utf8")) : {};

                  const mBgBlip = mxml.match(/<p:bg>[\s\S]*?<a:blip\s[^>]*r:embed="([^"]+)"/);
                  if (mBgBlip) {
                    const mDataUrl = getImageBase64(mBgBlip[1], mRelsMap, zip, "ppt/");
                    if (mDataUrl) { slide.mediaURL = mDataUrl; slide.mediaType = "image"; }
                  }
                  if (!slide.mediaURL) {
                    const mBgChunkM = mxml.match(/<p:bg>([\s\S]*?)<\/p:bg>/);
                    const mBgColor = mBgChunkM ? readFillColorHex(mBgChunkM[1]) : null;
                    if (mBgColor) slide.bgColor = mBgColor;
                  }
                }
              } catch(e) {}
            }
          }
        }
      } catch(e) {}
    }
  }

  // (a4) 배경이 <p:bg>가 아니라 "큰 도형 채우기"로 만든 케이스 보정
  // 텍스트가 없는 도형 중 면적이 큰 사각형의 solidFill을 배경으로 사용한다.
  if (!slide.mediaURL && !slide.bgColor) {
    const shapeRe = /<p:sp>([\s\S]*?)<\/p:sp>/g;
    let sm;
    let bestShape = null;
    while ((sm = shapeRe.exec(xml)) !== null) {
      const sp = sm[1];
      if (sp.includes("<p:txBody>")) continue;
      const offM = sp.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/);
      const extM = sp.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
      if (!offM || !extM) continue;
      const x = parseInt(offM[1]);
      const y = parseInt(offM[2]);
      const cx = parseInt(extM[1]);
      const cy = parseInt(extM[2]);
      if (!Number.isFinite(cx) || !Number.isFinite(cy) || cx <= 0 || cy <= 0) continue;
      const fill = readFillColorHex(sp);
      const blipEmbed = (sp.match(/<a:blip\b[^>]*r:embed="([^"]+)"/) || [])[1] || null;
      const fillImage = blipEmbed ? getImageBase64(blipEmbed, relsMap, zip, baseDir) : null;
      if (!fill && !fillImage) continue;

      const area = cx * cy;
      const coverRatio = area / (slideW * slideH);
      if (coverRatio < 0.45) continue; // 배경 후보는 최소 절반 가까이 덮어야 함

      const centerX = x + cx / 2;
      const centerY = y + cy / 2;
      const centerBias =
        Math.abs(centerX - slideW / 2) / slideW +
        Math.abs(centerY - slideH / 2) / slideH;
      const score = coverRatio - centerBias * 0.25;

      if (!bestShape || score > bestShape.score) {
        bestShape = { fill, fillImage, score };
      }
    }
    if (bestShape) {
      if (bestShape.fillImage) {
        slide.mediaURL = bestShape.fillImage;
        slide.mediaType = "image";
      } else {
        slide.bgColor = bestShape.fill;
      }
    }
  }

  // (b) <p:pic> 이미지 요소 — 가장 큰 이미지를 배경으로 사용
  const allPics = [];
  const picRe2  = /<p:pic>([\s\S]*?)<\/p:pic>/g;
  let picM2;
  while ((picM2 = picRe2.exec(xml)) !== null) {
    const picXml = picM2[1];
    const embedM = picXml.match(/r:embed="([^"]+)"/);
    const offM   = picXml.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/);
    const extM   = picXml.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
    if (!embedM) continue;
    const x  = offM ? parseInt(offM[1]) : 0;
    const y  = offM ? parseInt(offM[2]) : 0;
    const cx = extM ? parseInt(extM[1]) : slideW;
    const cy = extM ? parseInt(extM[2]) : slideH;
    allPics.push({ rId: embedM[1], x, y, cx, cy, area: cx * cy });
  }
  allPics.sort((a, b) => b.area - a.area);  // 면적 큰 순서로 정렬
  if (!slide.mediaURL && allPics.length) {
    const bg      = allPics[0];
    const dataUrl = getImageBase64(bg.rId, relsMap, zip, baseDir);
    if (dataUrl) {
      slide.mediaURL  = dataUrl;
      slide.mediaType = "image";
      // 위치/크기를 % 단위로 저장 (앱의 mOffX/Y/Scale 필드)
      slide.mediaOffX  = Math.round((bg.x + bg.cx / 2) / slideW * 100 * 10) / 10;
      slide.mediaOffY  = Math.round((bg.y + bg.cy / 2) / slideH * 100 * 10) / 10;
      slide.mediaScale = Math.round(Math.max(bg.cx / slideW, bg.cy / slideH) * 100 * 10) / 10;
    }
  }

  // ── 2. 슬라이드 레이아웃에서 placeholder 위치 상속 ──
  const layoutPhs = getPptxLayoutPlaceholders(relsMap, zip);

  // 슬라이드 높이를 pt로 변환 (폰트 크기 비율 계산에 사용)
  const slideH_pt = (slideH / 914400) * 72;  // EMU → 인치 → pt

  const pickThemeFont = (tf, preferEa = false) => {
    const val = String(tf || "");
    if (!val) return "";
    if (val === "+mj-lt" || val === "+mj-ea") return (preferEa && theme.majorFontEa) || theme.majorFont || "";
    if (val === "+mn-lt" || val === "+mn-ea") return (preferEa && theme.minorFontEa) || theme.minorFont || "";
    return (!val.startsWith("+") && !val.startsWith("@")) ? val : "";
  };

  const parseShapeColor = (shapeXml = "", fallback = null) => {
    if (!shapeXml) return fallback;
    const srgb = shapeXml.match(/<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
    if (srgb) return "#" + srgb[1];
    const sys = shapeXml.match(/<a:sysClr\s+[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"/);
    if (sys) return "#" + sys[1];
    const sch = shapeXml.match(/<a:schemeClr\s+val="([^"]+)"/);
    if (sch) {
      const resolved = resolveSchemeColor(sch[1]);
      if (resolved) return "#" + String(resolved).replace(/^#/, "");
    }
    return fallback;
  };

  const extractDecorShapesWithDom = () => {
    try {
      const doc = new DOMParser().parseFromString(xml, "text/xml");
      const spNodes = doc.getElementsByTagName("p:sp");
      const result = [];

      const getFirst = (el, tag) => {
        const list = el.getElementsByTagName(tag);
        return list && list.length ? list[0] : null;
      };
      const toInt = (v, d = 0) => {
        const n = parseInt(String(v ?? ""), 10);
        return Number.isFinite(n) ? n : d;
      };

      for (let i = 0; i < spNodes.length; i++) {
        const spEl = spNodes[i];
        const txBodyEl = getFirst(spEl, "p:txBody");
        if (txBodyEl) continue;

        const xfrm = getFirst(spEl, "a:xfrm");
        const off = xfrm ? getFirst(xfrm, "a:off") : null;
        const ext = xfrm ? getFirst(xfrm, "a:ext") : null;
        if (!off || !ext) continue;

        const sx = toInt(off.getAttribute("x"), 0);
        const sy = toInt(off.getAttribute("y"), 0);
        const sw = toInt(ext.getAttribute("cx"), 0);
        const sh = toInt(ext.getAttribute("cy"), 0);
        if (sw <= 0 || sh <= 0) continue;

        const spXml = String(spEl.toString ? spEl.toString() : "");
        const noFill = /<a:noFill\b/.test(spXml);
        const fillColor = noFill ? "transparent" : (parseShapeColor(spXml, null) || "transparent");
        const lineColor = parseShapeColor((spXml.match(/<a:ln\b[\s\S]*?<\/a:ln>/) || [])[0] || "", null);
        const lineW = (() => {
          const lw = (spXml.match(/<a:ln\b[^>]*\bw="(\d+)"/) || [])[1];
          if (!lw) return 0;
          const pt = parseInt(lw, 10) / 12700;
          return Number.isFinite(pt) ? Math.max(0, Math.min(24, Math.round(pt * 10) / 10)) : 0;
        })();
        const prstGeom = (spXml.match(/<a:prstGeom\b[^>]*\bprst="([^"]+)"/) || [])[1] || "rect";

        // 채우기/선이 전혀 없는 보조 도형은 제외
        if (fillColor === "transparent" && !lineColor && lineW <= 0) continue;

        result.push({
          x: Math.max(0, Math.min(100, emuToPct(sx + sw / 2, slideW))),
          y: Math.max(0, Math.min(100, emuToPct(sy + sh / 2, slideH))),
          w: Math.max(0.2, Math.min(100, emuToPct(sw, slideW))),
          h: Math.max(0.2, Math.min(100, emuToPct(sh, slideH))),
          fillColor,
          lineColor: lineColor || "transparent",
          lineWidth: lineW,
          shapeType: prstGeom
        });
      }
      return result;
    } catch (e) {
      return [];
    }
  };

  const extractShapesWithDom = () => {
    try {
      const doc = new DOMParser().parseFromString(xml, "text/xml");
      const spNodes = doc.getElementsByTagName("p:sp");
      const domShapes = [];

      const getFirst = (el, tag) => {
        const list = el.getElementsByTagName(tag);
        return list && list.length ? list[0] : null;
      };
      const toInt = (v, d = 0) => {
        const n = parseInt(String(v ?? ""), 10);
        return Number.isFinite(n) ? n : d;
      };

      for (let i = 0; i < spNodes.length; i++) {
        const spEl = spNodes[i];
        const txBodyEl = getFirst(spEl, "p:txBody");
        if (!txBodyEl) continue;

        const phEl = getFirst(spEl, "p:ph");
        const phType = phEl ? (phEl.getAttribute("type") || "body") : "";
        const phIdx = phEl ? (phEl.getAttribute("idx") || "1") : "";

        const xfrm = getFirst(spEl, "a:xfrm");
        const off = xfrm ? getFirst(xfrm, "a:off") : null;
        const ext = xfrm ? getFirst(xfrm, "a:ext") : null;
        const lph = layoutPhs[phIdx] || layoutPhs[phType] || null;

        const sx = off ? toInt(off.getAttribute("x"), lph ? lph.x : 0) : (lph ? lph.x : 0);
        const sy = off ? toInt(off.getAttribute("y"), lph ? lph.y : 0) : (lph ? lph.y : 0);
        const sw = ext ? toInt(ext.getAttribute("cx"), lph ? lph.w : slideW) : (lph ? lph.w : slideW);
        const sh = ext ? toInt(ext.getAttribute("cy"), lph ? lph.h : slideH) : (lph ? lph.h : slideH);
        if (sw <= 0 || sh <= 0) continue;

        const cxPct = emuToPct(sx + sw / 2, slideW);
        const wPct = emuToPct(sw, slideW);
        const hPct = emuToPct(sh, slideH);

        const bodyPr = getFirst(txBodyEl, "a:bodyPr");
        const anchor = bodyPr ? (bodyPr.getAttribute("anchor") || "t") : "t";
        const cyPct =
          anchor === "ctr" ? emuToPct(sy + sh / 2, slideH) :
          anchor === "b" ? emuToPct(sy + sh, slideH) :
          emuToPct(sy, slideH);

        const pNodes = txBodyEl.getElementsByTagName("a:p");
        let mainFontSize = 0;
        let mainFontColor = "ffffff";
        let mainFontFace = "";
        let mainAlign = "center";
        let alignFixed = false;
        const lines = [];

        for (let pIdx = 0; pIdx < pNodes.length; pIdx++) {
          const p = pNodes[pIdx];
          if (!alignFixed) {
            const algn = p.getAttribute("algn");
            if (algn) {
              alignFixed = true;
              if (algn === "l") mainAlign = "left";
              else if (algn === "r") mainAlign = "right";
              else mainAlign = "center";
            }
          }

          const rNodes = p.getElementsByTagName("a:r");
          let lineText = "";
          for (let rIdx = 0; rIdx < rNodes.length; rIdx++) {
            const r = rNodes[rIdx];
            const rPr = getFirst(r, "a:rPr");
            if (rPr) {
              const runSz = toInt(rPr.getAttribute("sz"), 0);
              if (runSz > 0) mainFontSize = Math.max(mainFontSize, Math.round(runSz / 100));
              const srgb = getFirst(rPr, "a:srgbClr");
              const sys = getFirst(rPr, "a:sysClr");
              const sch = getFirst(rPr, "a:schemeClr");
              if (srgb && srgb.getAttribute("val")) {
                mainFontColor = srgb.getAttribute("val");
              } else if (sys && (sys.getAttribute("val") || sys.getAttribute("lastClr"))) {
                mainFontColor = sys.getAttribute("val") || sys.getAttribute("lastClr");
              } else if (sch && sch.getAttribute("val")) {
                const resolved = resolveSchemeColor(sch.getAttribute("val"));
                if (resolved) mainFontColor = resolved.replace(/^#/, "");
              }

              const ea = getFirst(rPr, "a:ea");
              const latin = getFirst(rPr, "a:latin");
              const runFont = pickThemeFont(ea ? ea.getAttribute("typeface") : "", true) ||
                pickThemeFont(latin ? latin.getAttribute("typeface") : "", false);
              if (runFont) mainFontFace = runFont;
            }
            const tNodes = r.getElementsByTagName("a:t");
            if (tNodes.length) lineText += xmlDecodeText(tNodes[0].textContent || "");
          }
          if (lineText.trim()) lines.push(lineText.trimEnd());
        }

        if (!lines.length) continue;

        domShapes.push({
          text: lines.join("\n").trim(),
          cxPct, cyPct, wPct, hPct,
          fontSize: mainFontSize || 24,
          fontColor: "#" + String(mainFontColor || "ffffff").replace(/^#/, ""),
          fontFace: mainFontFace,
          align: mainAlign,
          anchor,
          phType,
          area: wPct * hPct
        });
      }

      return domShapes;
    } catch (e) {
      return [];
    }
  };

  // ── 3. 텍스트 도형(<p:sp>) 추출 ──
  const shapes = extractShapesWithDom();
  if (!shapes.length) {
    // TODO: DOM 파서가 일부 shape만 반환해도 non-empty면 여기 fallback이 비활성화되어 누락 shape를 보완하지 못할 수 있다.
    const spRe   = /<p:sp>([\s\S]*?)<\/p:sp>/g;
    let spM;
    while ((spM = spRe.exec(xml)) !== null) {
    const sp = spM[1];
    if (!sp.includes("<p:txBody>")) continue;

    // placeholder 타입과 인덱스 (body=본문, title=제목, ctrTitle=중앙제목 등)
    const phTagM  = sp.match(/<p:ph\b([^>]*)>/);
    const phTypeM = phTagM ? phTagM[1].match(/\btype="([^"]+)"/) : null;
    const phIdxM  = phTagM ? phTagM[1].match(/\bidx="([^"]+)"/)  : null;
    const phType  = phTypeM ? phTypeM[1] : (phTagM ? "body" : "");
    const phIdx   = phIdxM  ? phIdxM[1]  : (phTagM ? "1"   : "");

    // 위치/크기: 명시되어 있으면 사용, 없으면 레이아웃 placeholder에서 상속
    const offM = sp.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/);
    const extM = sp.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
    const lph  = layoutPhs[phIdx] || layoutPhs[phType] || null;

    const sx = offM ? parseInt(offM[1]) : (lph ? lph.x : 0);
    const sy = offM ? parseInt(offM[2]) : (lph ? lph.y : 0);
    const sw = extM ? parseInt(extM[1]) : (lph ? lph.w : slideW);
    const sh = extM ? parseInt(extM[2]) : (lph ? lph.h : slideH);

    // 위치/크기를 % 단위로 변환 (중심점 기준)
    const cxPct = emuToPct(sx + sw / 2, slideW);
    const wPct  = emuToPct(sw, slideW);
    const hPct  = emuToPct(sh, slideH);

    // 수직 앵커 (PPTX 기본: 't'=상단, 'ctr'=중앙, 'b'=하단)
    const bprAncM = sp.match(/<a:bodyPr\b[^>]*\banchor="([^"]+)"/);
    const anchor  = bprAncM ? bprAncM[1] : 't';
    const cyPct   =
      anchor === 'ctr' ? emuToPct(sy + sh / 2, slideH) :
      anchor === 'b'   ? emuToPct(sy + sh, slideH) :
                         emuToPct(sy, slideH);

    // txBody 레벨 기본 서식 (run/단락 서식이 없을 때 폴백으로 사용)
    const bodyDefSzM   = sp.match(/<a:defRPr\b[^>]*\bsz="(\d+)"/);
    const bodyDefSize  = bodyDefSzM ? Math.round(parseInt(bodyDefSzM[1]) / 100) : 0;
    const bodyDefClrM  = sp.match(/<a:defRPr\b[^>]*>[\s\S]*?<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
    const bodyDefSchM  = !bodyDefClrM && sp.match(/<a:defRPr\b[^>]*>[\s\S]*?<a:schemeClr\s+val="([^"]+)"/);
    const bodyDefColor = bodyDefClrM ? bodyDefClrM[1] : (bodyDefSchM ? resolveSchemeColor(bodyDefSchM[1]) : null);

    let mainFontSize  = 0, mainFontColor = "ffffff", mainFontFace = "";
    let mainAlign     = "center";
    let alignFixed    = false;
    const lines       = [];

    // 단락(<a:p>) 순서대로 텍스트 추출
    const paraRe = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
    let paraM;
    while ((paraM = paraRe.exec(sp)) !== null) {
      const para = paraM[1];

      // 첫 번째 명시 정렬만 사용 (l=왼쪽, r=오른쪽, 그 외=중앙)
      if (!alignFixed) {
        const algM = para.match(/\balgn="([^"]+)"/);
        if (algM) {
          alignFixed = true;
          if      (algM[1] === "l") mainAlign = "left";
          else if (algM[1] === "r") mainAlign = "right";
          else                      mainAlign = "center";
        }
      }

      // 단락 기본 폰트 크기/색상 (run 서식 없을 때 폴백)
      const pDefSzM   = para.match(/<a:defRPr\b[^>]*\bsz="(\d+)"/);
      const pDefSize  = pDefSzM ? Math.round(parseInt(pDefSzM[1]) / 100) : bodyDefSize;
      const pDefClrM  = para.match(/<a:defRPr\b[^>]*>[\s\S]*?<a:srgbClr\s+val="([0-9A-Fa-f]{6})"/);
      const pDefSchM  = !pDefClrM && para.match(/<a:defRPr\b[^>]*>[\s\S]*?<a:schemeClr\s+val="([^"]+)"/);
      const pDefColor = pDefClrM ? pDefClrM[1] : (pDefSchM ? resolveSchemeColor(pDefSchM[1]) : bodyDefColor);

      let lineText = "";
      // run(<a:r>): 실제 텍스트 데이터를 담는 최소 단위
      const runRe = /<a:r>([\s\S]*?)<\/a:r>/g;
      let runM;
      while ((runM = runRe.exec(para)) !== null) {
        const run = runM[1];

        // 폰트 크기 (PPTX 단위: 1/100pt → pt로 변환)
        const szM2  = run.match(/\bsz="(\d+)"/);
        const runSz = szM2 ? Math.round(parseInt(szM2[1]) / 100) : pDefSize;
        if (runSz > mainFontSize) mainFontSize = runSz;

        // 폰트 색상 (srgbClr/sysClr 우선, schemeClr 폴백)
        const clrM = run.match(/(?:<a:srgbClr|<a:sysClr)\s+[^>]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"/);
        if (clrM) {
          mainFontColor = clrM[1];
        } else {
          const schM     = run.match(/<a:schemeClr\s+val="([^"]+)"/);
          const resolved = schM ? resolveSchemeColor(schM[1]) : null;
          if (resolved) mainFontColor = resolved;
          else if (pDefColor && mainFontColor === "ffffff") mainFontColor = pDefColor;
        }

        // 폰트 이름 (한글 폰트: <a:ea> 우선, 라틴 폰트: <a:latin> 폴백)
        const eaTf   = (run.match(/<a:ea\b[^>]*\btypeface="([^"]+)"/) || [])[1] || '';
        const laTf   = (run.match(/<a:latin\b[^>]*\btypeface="([^"]+)"/) || [])[1] || '';
        const runFont = pickThemeFont(eaTf, true) || pickThemeFont(laTf, false);
        if (runFont) mainFontFace = runFont;

        // 텍스트 내용
        const tM = run.match(/<a:t>([\s\S]*?)<\/a:t>/);
        if (tM) lineText += xmlDecodeText(tM[1]);
      }
      if (lineText.trim()) lines.push(lineText.trimEnd());
    }

    if (!lines.length) continue;

      shapes.push({
      text: lines.join("\n").trim(),
      cxPct, cyPct, wPct, hPct,
      fontSize:  mainFontSize || bodyDefSize || 24,
      fontColor: "#" + mainFontColor,
      fontFace:  mainFontFace,
      align:     mainAlign,
      anchor,
      phType,
      area: wPct * hPct
      });
    }
  }

  // ── 4. 메인 텍스트 식별 ──
  // placeholder 타입 'body' + 폰트 크기 + 면적을 종합해 메인 가사를 선정
  shapes.sort((a, b) => {
    const scoreA = (a.phType === "body" ? 10000 : 0) + a.fontSize * 100 + a.area;
    const scoreB = (b.phType === "body" ? 10000 : 0) + b.fontSize * 100 + b.area;
    return scoreB - scoreA;
  });

  // PPTX pt → 앱 px 변환 (1080p 기준 비율 보존)
  const ptToAppPx = pt => Math.max(8, Math.min(200, Math.round(pt * 1080 / (slideH_pt * 1.4))));
  slide.importedTextBoxes = shapes.map(s => ({
    text: s.text,
    x: Math.max(0, Math.min(100, s.cxPct)),
    y: Math.max(0, Math.min(100, s.cyPct)),
    w: Math.min(99, s.wPct),
    h: Math.max(4, s.hPct),
    width: Math.min(99, s.wPct),
    height: Math.max(4, s.hPct),
    fontSize: ptToAppPx(s.fontSize),
    color: s.fontColor,
    fontColor: s.fontColor,
    fontFamily: pptxFontToKor(s.fontFace),
    align: s.align,
    textAnchor: s.anchor || "t",
    role: s.phType || "custom"
  }));

  const mainShape = shapes[0];
  if (mainShape) {
    slide.lyrics     = mainShape.text;
    slide.fontSize   = ptToAppPx(mainShape.fontSize);
    slide.fontColor  = mainShape.fontColor;
    slide.fontFamily = pptxFontToKor(mainShape.fontFace);
    slide.textX      = Math.max(0, Math.min(100, mainShape.cxPct));
    slide.textY      = Math.max(0, Math.min(100, mainShape.cyPct));
    slide.textWidth  = Math.min(99, mainShape.wPct);
    slide.textAnchor = mainShape.anchor || 't';
    slide.layoutType = "centerLyrics";
  }

  // title/ctrTitle placeholder → 소제목(sub)으로 설정
  const titleShape = shapes.find(s => s.phType === "title" || s.phType === "ctrTitle");
  if (titleShape && titleShape !== mainShape) {
    slide.sub = titleShape.text.replace(/\n/g, " ").slice(0, 80);
  }

  // 나머지 텍스트 도형들은 추가 텍스트박스로 추가
  const mkId = () => Math.random().toString(36).slice(2, 10);
  for (const s of shapes) {
    if (s === mainShape || s === titleShape) continue;
    if (!s.text.trim()) continue;
    slide.textBoxes.push({
      id:        mkId(),
      label:     s.text.slice(0, 20),
      text:      s.text,
      role:      "custom",
      x:         Math.max(0, Math.min(100, s.cxPct)),
      y:         Math.max(0, Math.min(100, s.cyPct)),
      w:         Math.min(99, s.wPct),
      h:         Math.max(4, s.hPct),
      width:     Math.min(99, s.wPct),
      height:    Math.max(4, s.hPct),
      fontSize:  ptToAppPx(s.fontSize),
      color:     s.fontColor,
      fontColor: s.fontColor,
      fontFamily: pptxFontToKor(s.fontFace),
      align:     s.align,
      textAnchor: s.anchor || 't'
    });
  }

slide.shapeBoxes = extractDecorShapesWithDom();

const bigBgShape = (slide.shapeBoxes || [])
  .filter(s => s && s.fillColor && s.fillColor !== 'transparent')
  .sort((a, b) => ((b.w || 0) * (b.h || 0)) - ((a.w || 0) * (a.h || 0)))[0];

if (!slide.mediaURL && !slide.bgColor && bigBgShape && bigBgShape.w > 80 && bigBgShape.h > 80) {
  slide.bgColor = bigBgShape.fillColor;
}

return slide;
}

/**
 * 슬라이드 레이아웃 XML에서 placeholder의 위치/크기를 추출한다.
 * 슬라이드 자체에 위치가 없을 때 레이아웃에서 상속받는 값을 제공한다.
 *
 * @param {Object} relsMap - 슬라이드의 관계 ID 매핑
 * @param {AdmZip} zip     - 열린 ZIP 객체
 * @returns {Object} { idx: {x, y, w, h}, type: {x, y, w, h} } 형태의 매핑
 */
function getPptxLayoutPlaceholders(relsMap, zip) {
  const result = {};
  try {
    const layoutRel = Object.entries(relsMap).find(([, t]) => t.includes("slideLayout"));
    if (!layoutRel) return result;
    const layoutTarget = layoutRel[1].replace(/\\/g, "/");
    const layoutPath   = resolveOpenXmlTargetPath("ppt/slides", layoutTarget);
    const layoutEntry = zipGetEntry(zip, layoutPath);
    if (!layoutEntry) return result;
    const lxml = layoutEntry.getData().toString("utf8");

    const spRe = /<p:sp>([\s\S]*?)<\/p:sp>/g;
    let spM;
    while ((spM = spRe.exec(lxml)) !== null) {
      const sp     = spM[1];
      const phTagM = sp.match(/<p:ph\b([^>]*)>/);
      if (!phTagM) continue;
      const typeM  = phTagM[1].match(/\btype="([^"]+)"/);
      const idxM   = phTagM[1].match(/\bidx="([^"]+)"/);
      const type   = typeM ? typeM[1] : "body";
      const idx    = idxM  ? idxM[1]  : "1";

      const offM = sp.match(/<a:off\s+x="(-?\d+)"\s+y="(-?\d+)"/);
      const extM = sp.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
      if (offM && extM) {
        const phData = {
          x: parseInt(offM[1]), y: parseInt(offM[2]),
          w: parseInt(extM[1]), h: parseInt(extM[2])
        };
        result[idx]  = phData;  // 인덱스로 조회
        result[type] = phData;  // 타입으로도 조회 가능
      }
    }
  } catch(e) {}
  return result;
}

/**
 * PPTX에서 추출한 폰트 이름을 앱에서 사용 가능한 폰트 이름으로 변환한다.
 * 시스템 한글 폰트를 구글 웹폰트 이름으로 매핑한다.
 *
 * @param {string} fontFace - PPTX에서 추출한 폰트 이름
 * @returns {string} 앱 폰트 이름 (매핑 없으면 원본 그대로)
 */
function pptxFontToKor(fontFace) {
  if (!fontFace) return "Noto Serif KR";
  const map = {
    // 한글 명조 계열 → Noto Serif KR
    "나눔명조": "Noto Serif KR",   "NanumMyeongjo": "Noto Serif KR",
    "바탕":     "Noto Serif KR",   "Batang":        "Noto Serif KR",
    "궁서":     "Noto Serif KR",   "Gungsuh":       "Noto Serif KR",
    "HY신명조": "Noto Serif KR",   "HY헤드라인M":   "Noto Serif KR",
    // 한글 고딕 계열 → Noto Sans KR
    "나눔고딕":            "Nanum Gothic",
    "NanumGothic":         "Nanum Gothic",
    "나눔손글씨 붓":       "Nanum Pen Script",
    "맑은 고딕":           "Noto Sans KR",   "Malgun Gothic": "Noto Sans KR",
    "굴림":                "Noto Sans KR",   "Gulim":         "Noto Sans KR",
    "돋움":                "Noto Sans KR",   "Dotum":         "Noto Sans KR",
    "HY고딕B":             "Noto Sans KR",   "HY강B":         "Noto Sans KR",
    // macOS 기본 한글 폰트
    "Apple SD Gothic Neo": "Noto Sans KR",
    "AppleGothic":         "Noto Sans KR",
    "Apple Myungjo":       "Noto Serif KR",
    // 영문 폰트 → 한글 폰트로 대체
    "Calibri":         "Noto Sans KR",   "Calibri Light":   "Noto Sans KR",
    "Arial":           "Noto Sans KR",   "Arial Narrow":    "Noto Sans KR",
    "Helvetica":       "Noto Sans KR",   "Helvetica Neue":  "Noto Sans KR",
    "Times New Roman": "Noto Serif KR",
    "Georgia":         "Noto Serif KR",
  };
  return map[fontFace] || fontFace;  // 매핑 없으면 원본 유지
}

// ──────────────────────────────────────────────────────────
// 가져온 미디어(영상) 임시파일 캐시
//   영상을 base64 data URL로 메모리에 들고 있으면 슬라이드가 많을 때
//   메모리 폭증·디코딩 비용으로 송출이 버벅인다. 대신 영상 바이너리를
//   세션 임시폴더에 1회 기록하고 file:// 경로로 참조하면 Chromium이
//   디스크에서 스트리밍하므로 메모리/멈춤 문제가 사라진다.
//   (저장 시 writeProjectBundle이 file:// 경로를 인식해 .wsproj 안에 담는다)
// ──────────────────────────────────────────────────────────
const MEDIA_TEMP_DIR = path.join(app.getPath("userData"), "media-temp");

// 앱 시작 시 오래된 임시 미디어만 정리한다.
//   localStorage 자동저장이 file:// 경로를 참조하므로 최근 파일은 지우면 안 된다
//   (지우면 재시작 후 영상이 깨짐). 30일 지난 것만 정리해 디스크 누적을 막는다.
//   저장한 프로젝트는 .wsproj 안에 영상이 들어있으므로 어차피 영향 없다.
const MEDIA_TEMP_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
try {
  fs.mkdirSync(MEDIA_TEMP_DIR, { recursive: true });
  const now = Date.now();
  for (const f of fs.readdirSync(MEDIA_TEMP_DIR)) {
    try {
      const fp = path.join(MEDIA_TEMP_DIR, f);
      if (now - fs.statSync(fp).mtimeMs > MEDIA_TEMP_MAX_AGE_MS) fs.rmSync(fp, { force: true });
    } catch(e) {}
  }
} catch(e) {}

// 절대경로 → file:// URL (restoreMediaUrl과 동일한 형식)
function absPathToFileUrl(absPath) {
  const sep = process.platform === "win32" ? "/" : "";
  return "file://" + sep + absPath.replace(/\\/g, "/");
}

/**
 * 렌더러가 보낸 미디어 바이너리를 임시파일로 기록하고 file:// URL을 돌려준다.
 * 주로 영상 가져오기에서 사용 (base64 data URL 대체).
 * @param {{ data: ArrayBuffer|Uint8Array, ext: string }} payload
 */
ipcMain.handle("media:save-temp", async (event, payload) => {
  try {
    const { data, ext } = payload || {};
    if (!data) return { ok: false, message: "데이터가 없습니다." };
    fs.mkdirSync(MEDIA_TEMP_DIR, { recursive: true });
    const safeExt = String(ext || "mp4").replace(/[^a-z0-9]/gi, "").toLowerCase() || "mp4";
    const name = "media_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8) + "." + safeExt;
    const dest = path.join(MEDIA_TEMP_DIR, name);
    fs.writeFileSync(dest, Buffer.from(data));
    return { ok: true, url: absPathToFileUrl(dest) };
  } catch(e) {
    console.error("미디어 임시파일 저장 실패:", e);
    return { ok: false, message: e.message };
  }
});

// ──────────────────────────────────────────────────────────
// 프로젝트 번들 저장 / 열기  (.wsproj = ZIP + JSON + media/)
// ──────────────────────────────────────────────────────────

const PROJ_TEMP_DIR = path.join(app.getPath("userData"), "project-bundle-temp");

// 프로젝트를 열 때 풀어 놓은 미디어의 보관 기간.
//   불러온 레이아웃은 localStorage에 남아 이 file:// 경로를 계속 참조하므로
//   최근 것은 지우면 안 된다. media-temp와 같은 기준으로 30일 지난 것만 정리한다.
const PROJ_TEMP_MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;

function localPathFromMediaUrl(url) {
  if (!url) return null;
  if (url.startsWith("data:") || url.startsWith("http://") || url.startsWith("https://")) return null;
  try {
    if (url.startsWith("file://")) return fileURLToPath(url);
  } catch(e) {}
  if (path.isAbsolute(url)) return url;
  return null;
}

// 프로젝트 데이터를 .wsproj(ZIP) 번들로 만들어 savePath에 기록한다. (dialog 없는 순수 쓰기)
// project:save-bundle(다른 이름으로 저장)와 project:save-bundle-to-path(덮어쓰기) 양쪽에서 공유한다.
function writeProjectBundle(projectData, savePath) {
  const zip = new AdmZip();
  const pathToEntry = new Map();    // 절대경로 → media/ 내 파일명
  const dataUrlToEntry = new Map(); // data: URL → media/ 내 파일명 (중복 저장 방지)
  let mediaIdx = 1;

  const DATA_URL_EXT_MAP = {
    "image/jpeg": "jpg", "image/png": "png", "image/gif": "gif",
    "image/webp": "webp", "image/bmp": "bmp",
    "video/mp4": "mp4", "video/webm": "webm", "video/ogg": "ogv",
    "video/quicktime": "mov", "video/x-msvideo": "avi"
  };

  function resolveMediaUrl(url) {
    if (!url) return url;

    // data: URL → base64 디코딩 후 binary 파일로 ZIP media/ 폴더에 저장
    if (url.startsWith("data:")) {
      if (dataUrlToEntry.has(url)) return "media/" + dataUrlToEntry.get(url);
      const match = url.match(/^data:([^;]+);base64,(.+)$/s);
      if (!match) return url;
      const mime = match[1];
      const base64Data = match[2];
      const ext = DATA_URL_EXT_MAP[mime] || (mime.startsWith("video/") ? "mp4" : "bin");
      const entryName = "media_" + String(mediaIdx++).padStart(4, "0") + "." + ext;
      dataUrlToEntry.set(url, entryName);
      zip.addFile("media/" + entryName, Buffer.from(base64Data, "base64"));
      return "media/" + entryName;
    }

    if (url.startsWith("http://") || url.startsWith("https://")) return url;

    const filePath = localPathFromMediaUrl(url);
    if (!filePath || !fs.existsSync(filePath)) return url;
    if (pathToEntry.has(filePath)) return "media/" + pathToEntry.get(filePath);
    const ext = path.extname(filePath);
    const entryName = "media_" + String(mediaIdx++).padStart(4, "0") + ext;
    pathToEntry.set(filePath, entryName);
    return "media/" + entryName;
  }

  // slides 처리: data: URL을 binary 파일로 추출해 ZIP에 추가
  const processedSlides = (projectData.slides || []).map(s => {
    const slide = { ...s };
    slide.mediaURL = resolveMediaUrl(s.mediaURL);
    return slide;
  });

  // namedLayouts 처리: 레이아웃에 포함된 배경 이미지/영상도 함께 저장
  const rawNamedLayouts = projectData.namedLayouts || {};
  const processedNamedLayouts = {};
  for (const [name, layout] of Object.entries(rawNamedLayouts)) {
    if (layout && layout.mediaURL) {
      processedNamedLayouts[name] = { ...layout, mediaURL: resolveMediaUrl(layout.mediaURL) };
    } else {
      processedNamedLayouts[name] = layout;
    }
  }

  // 로컬 경로 파일을 ZIP에 추가
  for (const [filePath, entryName] of pathToEntry) {
    try {
      zip.addFile("media/" + entryName, fs.readFileSync(filePath));
    } catch(e) {
      console.warn("미디어 파일 추가 실패:", filePath, e.message);
    }
  }

  // 나머지 프로젝트 데이터를 직렬화 (slides, namedLayouts는 처리된 버전으로 교체)
  const data = { ...projectData, slides: processedSlides, namedLayouts: processedNamedLayouts };
  zip.addFile("project.wsproj.json", Buffer.from(JSON.stringify(data, null, 2), "utf8"));

  zip.writeZip(savePath);
  return { mediaCount: mediaIdx - 1, slidesCount: processedSlides.length };
}

// 다른 이름으로 저장: 저장 위치를 묻고 새 파일로 기록한 뒤 currentProjectPath 갱신
ipcMain.handle("project:save-bundle", async (event, projectData) => {
  if (!AdmZip) return { ok: false, message: "adm-zip 패키지가 없습니다. npm install 을 실행해 주세요." };

  const dateStr = new Date().toISOString().slice(0, 10);
  const dlg = await dialog.showSaveDialog(mainWin, {
    title: "프로젝트 저장",
    defaultPath: currentProjectPath || ("project_" + dateStr + ".wsproj"),
    filters: [{ name: "Cheon Myeong Slide Pro 프로젝트", extensions: ["wsproj"] }]
  });
  if (dlg.canceled) return { ok: false, message: "cancelled" };

  try {
    const savePath = dlg.filePath.endsWith(".wsproj") ? dlg.filePath : dlg.filePath + ".wsproj";
    const result = writeProjectBundle(projectData, savePath);
    currentProjectPath = savePath;
    return { ok: true, ...result, filePath: savePath, fileName: path.basename(savePath) };
  } catch(e) {
    console.error("프로젝트 번들 저장 실패:", e);
    return { ok: false, message: e.message };
  }
});

// 덮어쓰기 저장: 현재 열려있는 파일(currentProjectPath)에 dialog 없이 바로 기록.
// 아직 저장된 파일이 없으면 needPath:true 를 돌려주어 렌더러가 "다른 이름으로 저장"으로 넘어가게 한다.
ipcMain.handle("project:save-bundle-to-path", async (event, projectData) => {
  if (!AdmZip) return { ok: false, message: "adm-zip 패키지가 없습니다. npm install 을 실행해 주세요." };
  if (!currentProjectPath) return { ok: false, needPath: true, message: "저장된 파일이 없습니다." };

  try {
    const result = writeProjectBundle(projectData, currentProjectPath);
    return { ok: true, ...result, filePath: currentProjectPath, fileName: path.basename(currentProjectPath) };
  } catch(e) {
    console.error("프로젝트 덮어쓰기 저장 실패:", e);
    return { ok: false, message: e.message };
  }
});

// 현재 열려있는 프로젝트 파일 경로/이름 조회 (렌더러의 저장 버튼 상태 표시용)
ipcMain.handle("project:get-current-path", () => {
  return {
    filePath: currentProjectPath,
    fileName: currentProjectPath ? path.basename(currentProjectPath) : null
  };
});

// 종료 시 저장 확인 (3버튼 네이티브 대화상자) → 'save' | 'discard' | 'cancel'
ipcMain.handle("app:confirm-save-on-close", () => {
  if (!mainWin || mainWin.isDestroyed()) return "discard";
  const r = dialog.showMessageBoxSync(mainWin, {
    type: "question",
    buttons: ["저장하고 종료", "저장하지 않고 종료", "취소"],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
    title: "프로그램 종료",
    message: "변경 사항을 저장하시겠습니까?",
    detail: "저장하지 않은 변경 사항이 있습니다."
  });
  return r === 0 ? "save" : (r === 1 ? "discard" : "cancel");
});

// 렌더러가 종료 진행을 확정하면 실제로 창을 닫는다.
ipcMain.handle("app:proceed-close", () => {
  appIsQuitting = true;
  if (mainWin && !mainWin.isDestroyed()) mainWin.close();
  return { ok: true };
});

// 프로젝트 열기/새 프로젝트 전에 현재 작업을 저장할지 묻는다 → 'save' | 'discard' | 'cancel'
ipcMain.handle("project:confirm-save-before", (event, actionLabel) => {
  if (!mainWin || mainWin.isDestroyed()) return "discard";
  const label = String(actionLabel || "계속");
  const r = dialog.showMessageBoxSync(mainWin, {
    type: "question",
    buttons: ["저장 후 " + label, "저장하지 않고 " + label, "취소"],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
    title: label,
    message: "현재 프로젝트를 저장하시겠습니까?",
    detail: currentProjectPath
      ? "저장하면 \"" + path.basename(currentProjectPath) + "\" 파일에 덮어씁니다."
      : "아직 파일로 저장되지 않은 프로젝트입니다. 저장 위치를 선택하게 됩니다."
  });
  return r === 0 ? "save" : (r === 1 ? "discard" : "cancel");
});

// 새 프로젝트: 이후 저장이 이전 파일을 덮어쓰지 않도록 현재 파일 경로를 비운다.
ipcMain.handle("project:reset-current-path", () => {
  currentProjectPath = null;
  return { ok: true };
});

// opts.keepCurrentPath: true면 currentProjectPath를 바꾸지 않는다 (프로젝트 삽입용).
ipcMain.handle("project:open-bundle", async (event, opts) => {
  if (!AdmZip) return { ok: false, message: "adm-zip 패키지가 없습니다. npm install 을 실행해 주세요." };

  const pick = await dialog.showOpenDialog(mainWin, {
    title: "프로젝트 열기",
    filters: [
      { name: "Cheon Myeong Slide Pro 프로젝트", extensions: ["wsproj"] },
      { name: "모든 파일", extensions: ["*"] }
    ],
    properties: ["openFile"]
  });
  if (pick.canceled || !pick.filePaths.length) return { ok: false, message: "cancelled" };

  try {
    const zip = new AdmZip(pick.filePaths[0]);
    const jsonEntry = zip.getEntry("project.wsproj.json");
    if (!jsonEntry) return { ok: false, message: "project.wsproj.json 을 찾을 수 없습니다." };

    const data = JSON.parse(jsonEntry.getData().toString("utf8"));
    if (!data || data.type !== "full-project") return { ok: false, message: "올바른 프로젝트 형식이 아닙니다." };

    // 프로젝트마다 별도 폴더에 풀어 둔다.
    //   폴더를 통째로 비우고 media_0001.png 같은 이름을 재사용하면, 먼저 열었던
    //   프로젝트에서 가져온 레이아웃(localStorage에 남는다)의 배경 경로가 나중에 연
    //   다른 프로젝트의 파일을 가리키게 된다. 폴더를 분리해 그 충돌을 막고,
    //   오래된 폴더만 나이로 정리한다.
    fs.mkdirSync(PROJ_TEMP_DIR, { recursive: true });
    try {
      const now = Date.now();
      for (const f of fs.readdirSync(PROJ_TEMP_DIR)) {
        try {
          const fp = path.join(PROJ_TEMP_DIR, f);
          if (now - fs.statSync(fp).mtimeMs > PROJ_TEMP_MAX_AGE_MS) fs.rmSync(fp, { recursive: true, force: true });
        } catch(e) {}
      }
    } catch(e) {}

    const sessionDir = path.join(PROJ_TEMP_DIR, "p_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8));
    fs.mkdirSync(sessionDir, { recursive: true });

    const mediaEntries = zip.getEntries().filter(e => e.entryName.startsWith("media/") && !e.isDirectory);
    for (const entry of mediaEntries) {
      const dest = path.join(sessionDir, path.basename(entry.entryName));
      fs.writeFileSync(dest, entry.getData());
    }

    // 상대경로 media/xxx → 절대 file:// URL 로 변환
    const sep = process.platform === "win32" ? "/" : "";

    function restoreMediaUrl(url) {
      if (!url || !url.startsWith("media/")) return url;
      const absPath = path.join(sessionDir, path.basename(url));
      if (!fs.existsSync(absPath)) return url;
      return "file://" + sep + absPath.replace(/\\/g, "/");
    }

    for (const slide of data.slides) {
      slide.mediaURL = restoreMediaUrl(slide.mediaURL);
    }

    // namedLayouts에 포함된 배경 이미지/영상 경로도 복원
    if (data.namedLayouts && typeof data.namedLayouts === "object") {
      for (const name of Object.keys(data.namedLayouts)) {
        const layout = data.namedLayouts[name];
        if (layout && layout.mediaURL) {
          data.namedLayouts[name] = { ...layout, mediaURL: restoreMediaUrl(layout.mediaURL) };
        }
      }
    }

    // 이후 "덮어쓰기 저장"이 이 파일로 바로 저장되도록 경로 기억
    //   삽입은 현재 프로젝트에 끼워 넣는 것이므로 경로를 바꾸지 않는다.
    //   (바꾸면 다음 덮어쓰기 저장이 삽입한 원본 파일을 덮어쓴다)
    if (!(opts && opts.keepCurrentPath)) currentProjectPath = pick.filePaths[0];

    return { ok: true, data, filePath: pick.filePaths[0], fileName: path.basename(pick.filePaths[0]) };
  } catch(e) {
    console.error("프로젝트 번들 열기 실패:", e);
    return { ok: false, message: e.message };
  }
});
