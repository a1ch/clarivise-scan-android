/* Outlook Email Evaluator — Android/Mobile Add-in (taskpane.js) v1.0
   Built on top of the desktop add-in logic.
   Key differences vs desktop:
   - No DOM scraping (Office.js API only — works in mobile task pane)
   - No dark mode toggle (system dark mode handles it on Android)
   - Tap-friendly accordion findings
   - localStorage fallback since roamingSettings may be unavailable on mobile
*/

/* ── Storage helpers ───────────────────────────────────────── */
function storageGet(key) {
  try { const v = Office.context.roamingSettings.get(key); if (v != null) return v; } catch(e) {}
  try { return localStorage.getItem(key) || ''; } catch(e) {}
  return '';
}
function storageSet(key, value) {
  try { Office.context.roamingSettings.set(key, value); Office.context.roamingSettings.saveAsync(); } catch(e) {}
  try { localStorage.setItem(key, value); } catch(e) {}
}

function escapeHtml(s) {
  if (s == null || s === '') return '';
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

const DEFAULT_PROXY_URL = 'https://pikplhvawbhndijpkdbq.supabase.co/functions/v1/analyze-email';

/* ── Boot ──────────────────────────────────────────────────── */
Office.onReady(() => {
  initUI();
  loadEmail();

  // Re-load when user taps a different email (pinned pane on desktop/OWA)
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      loadEmail();
      resetAnalyzeBtn();
      document.getElementById('result-body').innerHTML = buildPlaceholder();
    });
  } catch(e) {}
});

/* ── UI wiring ─────────────────────────────────────────────── */
function initUI() {
  const settingsBtn = document.getElementById('settings-btn');
  const settingsPanel = document.getElementById('settings-panel');
  const mainPanel = document.getElementById('main-panel');
  const closeBtn = document.getElementById('settings-close-btn');

  settingsBtn.addEventListener('click', () => {
    populateSettings();
    settingsPanel.classList.remove('hidden');
    mainPanel.classList.add('hidden');
  });
  closeBtn.addEventListener('click', () => {
    settingsPanel.classList.add('hidden');
    mainPanel.classList.remove('hidden');
  });

  document.getElementById('save-settings-btn').addEventListener('click', saveSettings);
  document.getElementById('analyze-btn').addEventListener('click', analyzeEmail);

  // Tap-to-expand findings (event delegation)
  document.getElementById('result-body').addEventListener('click', e => {
    const hdr = e.target.closest('.finding-header');
    if (hdr) hdr.parentElement.classList.toggle('open');
  });
}

function buildPlaceholder() {
  return '<div class="placeholder"><div class="placeholder-icon">🔍</div><div class="placeholder-text">Tap <strong>Analyze Email</strong> to check this email for threats</div></div>';
}

/* ── Settings ──────────────────────────────────────────────── */
function populateSettings() {
  document.getElementById('proxy-url-input').value    = storageGet('proxyUrl') || DEFAULT_PROXY_URL;
  document.getElementById('ext-token-input').value    = storageGet('extToken') || '';
  document.getElementById('tenant-domain-input').value = storageGet('tenantDomain') || '';
  document.getElementById('custom-prompt-input').value = storageGet('customPrompt') || '';
}

function saveSettings() {
  storageSet('proxyUrl',     (document.getElementById('proxy-url-input').value || '').trim());
  storageSet('extToken',     (document.getElementById('ext-token-input').value || '').trim());
  storageSet('tenantDomain', (document.getElementById('tenant-domain-input').value || '').trim());
  storageSet('customPrompt', (document.getElementById('custom-prompt-input').value || '').trim());
  const msg = document.getElementById('settings-msg');
  msg.classList.remove('hidden');
  setTimeout(() => msg.classList.add('hidden'), 2000);
}

/* ── Load email metadata ───────────────────────────────────── */
function loadEmail() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  const subject = item.subject || '(No subject)';
  document.getElementById('email-subject').textContent =
    subject.length > 60 ? subject.slice(0, 60) + '…' : subject;

  if (item.from) {
    const sender = item.from.emailAddress || '';
    const name   = item.from.displayName  || '';
    document.getElementById('email-sender').textContent =
      name ? name + ' <' + sender + '>' : sender;
  }
}

/* ── URL decoder (SafeLinks / Proofpoint / Trend Micro / Mimecast) ── */
function decodeWrappedUrl(href) {
  if (!href) return href;
  try {
    href = href.replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"');
    if (href.includes('safelinks.protection.outlook.com')) {
      const u = new URL(href); const d = u.searchParams.get('url'); if (d) return decodeURIComponent(d);
    }
    if (href.includes('trendmicro') || href.includes('imsva') || href.includes('tmase')) {
      const u = new URL(href);
      const d = u.searchParams.get('url') || u.searchParams.get('u') || u.searchParams.get('__u');
      if (d) return decodeURIComponent(d);
      const b = u.searchParams.get('redirectUrl') || u.searchParams.get('r');
      if (b) { try { return atob(b); } catch(e) {} }
    }
    if (href.includes('urldefense') && href.includes('/v2/')) {
      const u = new URL(href); let r = u.searchParams.get('u');
      if (r) { r = r.replace(/-/g,'%').replace(/_/g,'/'); return decodeURIComponent(r); }
    }
    if (href.includes('urldefense') && href.includes('/v3/')) {
      const m = href.match(/\/v3\/__([^_]+)__/); if (m) return decodeURIComponent(m[1]);
    }
    if (href.includes('mimecast.com')) {
      const u = new URL(href); const d = u.searchParams.get('url') || u.searchParams.get('u'); if (d) return decodeURIComponent(d);
    }
    if (href.includes('?')) {
      const u = new URL(href); const d = u.searchParams.get('url') || u.searchParams.get('u');
      if (d && (d.startsWith('http') || d.startsWith('%68%74'))) return decodeURIComponent(d);
    }
  } catch(e) {}
  return href;
}

/* ── Extract links from HTML body ──────────────────────────── */
function extractLinks(html) {
  if (!html) return [];
  const doc = (new DOMParser()).parseFromString(html, 'text/html');
  const seen = new Set(); const links = [];
  doc.querySelectorAll('a[href]').forEach(a => {
    try {
      const displayText = (a.textContent || '').trim();
      let href = a.getAttribute('href') || '';
      href = decodeWrappedUrl(href);
      if (!href || href.startsWith('mailto:') || href.startsWith('#') || href.length < 10) return;
      let hrefDomain = ''; try { hrefDomain = new URL(href).hostname.toLowerCase(); } catch(e) { return; }
      if (seen.has(hrefDomain)) return; seen.add(hrefDomain);
      let displayDomain = '';
      const urlPattern = displayText.match(/(?:https?:\/\/|www\.)([\w.-]+)/i);
      if (urlPattern) {
        try { displayDomain = new URL(displayText.startsWith('http') ? displayText : 'https://' + displayText).hostname.toLowerCase(); }
        catch(e) { displayDomain = urlPattern[1].toLowerCase(); }
      }
      const mismatch = displayDomain && hrefDomain &&
        !hrefDomain.includes(displayDomain.replace(/^www\./,'')) &&
        !displayDomain.includes(hrefDomain.replace(/^www\./,''));
      links.push({ display: displayText.slice(0, 80) || '(no text)', href: hrefDomain, fullUrl: href, mismatch });
    } catch(e) {}
  });
  return links.slice(0, 20);
}

/* ── Main analysis ─────────────────────────────────────────── */
async function analyzeEmail() {
  const proxyUrl = storageGet('proxyUrl') || DEFAULT_PROXY_URL;
  const extToken = storageGet('extToken');

  if (!proxyUrl) { showError('No proxy URL set. Tap ⚙️ to configure.'); return; }
  if (!extToken) { showError('No extension token set. Tap ⚙️ to configure.'); return; }

  setLoading();
  const item = Office.context.mailbox.item;

  /* Read body — HTML first, text fallback */
  const bodyHtml = await new Promise(resolve =>
    item.body.getAsync(Office.CoercionType.Html, r =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '')
    )
  );
  const bodyText = await new Promise(resolve =>
    item.body.getAsync(Office.CoercionType.Text, r =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '')
    )
  );

  const sender       = item.from ? (item.from.displayName + ' <' + item.from.emailAddress + '>') : '(Unknown sender)';
  const subject      = item.subject || '(No subject)';
  const links        = extractLinks(bodyHtml);
  const attachNames  = (item.attachments || []).map(a => (a.name || '').toLowerCase());
  const tenantDomain = storageGet('tenantDomain') || '';
  const customPrompt = storageGet('customPrompt') || '';

  let isOutlookExternal = false;
  try {
    const senderEmail = item.from ? item.from.emailAddress.toLowerCase() : '';
    if (tenantDomain && senderEmail && !senderEmail.endsWith('@' + tenantDomain.toLowerCase())) {
      isOutlookExternal = true;
    }
    /* Try to get internet headers for external flag */
    if (Office.context.requirements && Office.context.requirements.isSetSupported('Mailbox', '1.8')) {
      const headers = await new Promise(resolve =>
        item.getAllInternetHeadersAsync(r =>
          resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '')
        )
      );
      if (headers && headers.toLowerCase().includes('x-ms-exchange-organization-scl')) isOutlookExternal = true;
    }
  } catch(e) {}

  const emailData = {
    subject, sender,
    senderHasEmail: sender.includes('@'),
    body: bodyText.slice(0, 3000),
    links, attachments: attachNames,
    isOutlookExternal,
    clientTimestamp: new Date().toISOString(),
    clientTimezone: Intl.DateTimeFormat().resolvedOptions().timeZone
  };

  try {
    const response = await fetch(proxyUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token: extToken, emailData, customPrompt, tenantDomain })
    });

    if (response.status === 429) { showError('Rate limited — wait a few seconds and try again.'); return; }
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      showError('Proxy error ' + response.status + ': ' + (err.error || response.statusText));
      return;
    }

    const data = await response.json();
    window._oe_lastResult = data.result;
    showResult(data.result, { subject, links });

  } catch(err) { showError('Request failed: ' + err.message); }
}

/* ── Loading state ─────────────────────────────────────────── */
function setLoading() {
  document.getElementById('result-body').innerHTML =
    '<div class="loading"><div class="spinner"></div><div class="loading-label">Analyzing email…</div></div>';
  const btn = document.getElementById('analyze-btn');
  btn.disabled = true;
  btn.textContent = 'Analyzing…';
}

function resetAnalyzeBtn() {
  const btn = document.getElementById('analyze-btn');
  btn.disabled = false;
  btn.textContent = '🔍 Analyze Email';
}

function showError(msg) {
  document.getElementById('result-body').innerHTML =
    '<div class="error-card">⚠️ ' + escapeHtml(msg) + '</div>';
  resetAnalyzeBtn();
}

/* ── Render result ─────────────────────────────────────────── */
function showResult(result, { subject, links }) {
  window._oe_lastResult = result;

  const vcMap = { SAFE:'verdict-safe', SUSPICIOUS:'verdict-suspicious', SPAM:'verdict-spam', PHISHING:'verdict-phishing' };
  const viMap = { SAFE:'✅', SUSPICIOUS:'⚠️', SPAM:'🚫', PHISHING:'🎣' };
  const vc = vcMap[result.verdict] || 'verdict-suspicious';
  const vi = viMap[result.verdict] || '⚠️';

  /* Verdict */
  let html = '<div class="verdict-card ' + vc + '">'
    + '<span class="verdict-icon">' + vi + '</span>'
    + '<span class="verdict-label">' + escapeHtml(result.verdict) + '</span>'
    + '</div>';

  /* Scores */
  html += '<div class="scores">'
    + '<div class="score-item"><span class="score-label">Phishing Risk</span><span class="score-val">' + escapeHtml(String(result.phishing_score)) + '/100</span><div class="score-bar"><div class="score-fill phishing-fill" style="width:' + result.phishing_score + '%"></div></div></div>'
    + '<div class="score-item"><span class="score-label">Spam Score</span><span class="score-val">' + escapeHtml(String(result.spam_score)) + '/100</span><div class="score-bar"><div class="score-fill spam-fill" style="width:' + result.spam_score + '%"></div></div></div>'
    + '</div>';

  /* Attachment warnings */
  const highRisk   = result.highRiskFiles  || [];
  const suspicious = result.suspiciousFiles || [];
  if (highRisk.length > 0) {
    html += '<div class="attach-high-risk">⚠️ HIGH RISK ATTACHMENT: ' + escapeHtml(highRisk.join(', ')) + '<br>Do NOT open. Report to IT immediately.</div>';
  } else if (suspicious.length > 0) {
    html += '<div class="attach-suspicious">⚠️ SUSPICIOUS ATTACHMENT: ' + escapeHtml(suspicious.join(', ')) + '<br>Verify with sender before opening.</div>';
  }

  /* Summary */
  html += '<div class="card"><div class="card-title">Summary</div><p>' + escapeHtml(result.summary) + '</p></div>';

  /* Warning banner for phishing-like content */
  const combined = ((subject || '') + ' ' + (result.summary || '')).toLowerCase();
  const showWarn = ['sign in','verification code','one-time','otp','log in','verify your','reset your password','confirm your','your account','click here to'].some(kw => combined.includes(kw))
    || result.verdict === 'PHISHING' || result.phishing_score >= 60;
  if (showWarn) {
    html += '<div class="warning-banner">⚠️ If you did not request this, do not tap any links — report to your IT security team.</div>';
  }

  /* Findings (accordion) */
  if ((result.findings || []).length > 0) {
    html += '<div class="card"><div class="card-title">🔍 What We Found — tap each to learn more</div>';
    result.findings.forEach(f => {
      html += '<div class="finding">'
        + '<div class="finding-header"><span class="finding-icon">🚩</span><span class="finding-flag">' + escapeHtml(f.flag) + '</span><span class="finding-arrow">▼</span></div>'
        + '<div class="finding-body">'
        + '<div class="finding-section"><div class="finding-section-label">What\'s happening</div><div class="finding-section-text">' + escapeHtml(f.explanation) + '</div></div>'
        + '<div class="finding-section finding-tip"><div class="finding-section-label">💡 How to spot this yourself</div><div class="finding-section-text">' + escapeHtml(f.howToSpotIt) + '</div></div>'
        + '</div></div>';
    });
    html += '</div>';
  }

  /* Links */
  if (links.length > 0) {
    html += '<div class="card"><div class="card-title">🔗 Links (' + links.length + ')</div>';
    links.forEach(l => {
      html += '<div class="link-row' + (l.mismatch ? ' link-mismatch' : '') + '">'
        + '<div class="link-display">' + escapeHtml(l.display) + '</div>'
        + '<div class="link-dest">→ ' + escapeHtml(l.href) + (l.mismatch ? ' <span class="link-mismatch-badge">⚠️ MISMATCH</span>' : '') + '</div>'
        + '</div>';
    });
    html += '</div>';
  }

  /* Lesson */
  if (result.lesson) {
    html += '<div class="lesson-card"><div class="lesson-title">📚 Remember for next time</div><div class="lesson-text">' + escapeHtml(result.lesson) + '</div></div>';
  }

  /* Suggested action */
  html += '<div class="card"><div class="card-title">✅ Suggested Action</div><p>' + escapeHtml(result.suggested_action) + '</p></div>';

  /* Feedback */
  html += '<div class="feedback-card" id="feedback-card">'
    + '<div class="feedback-title">Was this analysis accurate?</div>'
    + '<div class="feedback-buttons">'
    + '<button class="fb-btn fb-fp" id="fb-fp">👎 False Positive</button>'
    + '<button class="fb-btn fb-mt" id="fb-mt">🚨 Missed Threat</button>'
    + '</div></div>';

  document.getElementById('result-body').innerHTML = html;
  resetAnalyzeBtn();

  // Scroll to top so user sees verdict first
  document.getElementById('result-body').scrollTop = 0;

  document.getElementById('fb-fp').addEventListener('click', () => showFeedbackForm('false_positive', result));
  document.getElementById('fb-mt').addEventListener('click', () => showFeedbackForm('missed_threat',  result));
}

/* ── Feedback ──────────────────────────────────────────────── */
function showFeedbackForm(feedbackType, result) {
  const card  = document.getElementById('feedback-card');
  const label = feedbackType === 'false_positive'
    ? 'This email was flagged but is actually safe'
    : 'This email is suspicious but was not caught';
  card.innerHTML =
    '<div class="feedback-title">' + label + '</div>'
    + '<textarea id="fb-comment" class="fb-comment" placeholder="Optional: tell us more…" maxlength="500"></textarea>'
    + '<div class="fb-actions">'
    + '<button class="fb-btn fb-send" id="fb-send">Send Report</button>'
    + '<button class="fb-btn fb-cancel" id="fb-cancel">Cancel</button>'
    + '</div>';
  document.getElementById('fb-send').addEventListener('click',   () => submitFeedback(feedbackType, result, (document.getElementById('fb-comment').value || '').trim()));
  document.getElementById('fb-cancel').addEventListener('click', resetFeedbackCard);
}

async function submitFeedback(feedbackType, result, comment) {
  const card = document.getElementById('feedback-card');
  card.innerHTML = '<div class="feedback-title" style="text-align:center"><div class="spinner" style="margin:0 auto 8px;width:24px;height:24px;border-width:2px"></div>Sending…</div>';

  const proxyUrl = storageGet('proxyUrl') || DEFAULT_PROXY_URL;
  const extToken = storageGet('extToken');
  if (!proxyUrl || !extToken) {
    card.innerHTML = '<div class="feedback-title" style="color:#991b1b">Extension not configured.</div>'; return;
  }
  const feedbackUrl = proxyUrl.replace(/\/analyze-email\/?$/, '/report-feedback');

  try {
    const item = Office.context.mailbox.item;
    const response = await fetch(feedbackUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        token: extToken,
        feedbackType,
        originalVerdict:      result.verdict,
        originalPhishingScore: result.phishing_score,
        originalSpamScore:    result.spam_score,
        emailSubject:  (item.subject || '').slice(0, 200),
        emailSender:   item.from ? item.from.emailAddress.slice(0, 200) : '',
        userComment:   comment
      })
    });
    card.innerHTML = response.ok
      ? '<div class="feedback-title" style="color:#065f46">✅ Report submitted — thank you!</div>'
      : '<div class="feedback-title" style="color:#991b1b">Failed to send. Please try again.</div>';
  } catch(e) {
    card.innerHTML = '<div class="feedback-title" style="color:#991b1b">Failed: ' + escapeHtml(e.message) + '</div>';
  }
}

function resetFeedbackCard() {
  const card = document.getElementById('feedback-card');
  const last = window._oe_lastResult || {};
  card.innerHTML =
    '<div class="feedback-title">Was this analysis accurate?</div>'
    + '<div class="feedback-buttons">'
    + '<button class="fb-btn fb-fp" id="fb-fp">👎 False Positive</button>'
    + '<button class="fb-btn fb-mt" id="fb-mt">🚨 Missed Threat</button>'
    + '</div>';
  document.getElementById('fb-fp').addEventListener('click', () => showFeedbackForm('false_positive', last));
  document.getElementById('fb-mt').addEventListener('click', () => showFeedbackForm('missed_threat',  last));
}
