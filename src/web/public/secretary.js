(function () {
  function initSecretaryAssistant() {
    const bootstrap = window.SECRETARY_BOOTSTRAP || {};
    const threads = Array.isArray(bootstrap.threads) ? bootstrap.threads : [];
    const MAX_TURNS = typeof bootstrap.maxTurns === 'number' ? bootstrap.maxTurns : 0;
    window.SECRETARY_BOOTSTRAP = undefined;
  const messageEl = document.getElementById('secretary-message');
  const detailEl = document.getElementById('secretary-email');
  const fromEl = document.getElementById('secretary-from');
  const summaryEl = document.getElementById('secretary-summary');
  const subjectEl = document.getElementById('secretary-subject');
  const nextEl = document.getElementById('secretary-next');
  const linkEl = document.getElementById('secretary-link');
  const buttonEl = document.getElementById('secretary-button');
  const backButtonEl = document.getElementById('secretary-back-button');
  const chatContainer = document.getElementById('secretary-chat');
  const chatLog = document.getElementById('secretary-chat-log');
  const chatForm = document.getElementById('secretary-chat-form');
  const chatInput = document.getElementById('secretary-chat-input');
  const chatHint = document.getElementById('secretary-chat-hint');
  const chatError = document.getElementById('secretary-chat-error');
  const avatarEl = document.getElementById('secretary-avatar');
  const receivedEl = document.getElementById('secretary-received');
  const chipReplyEl = document.getElementById('secretary-chip-reply');
  const chipCategoryEl = document.getElementById('secretary-chip-category');
  const chipEffortEl = document.getElementById('secretary-chip-effort');
  const previewEl = document.getElementById('secretary-preview');
  const previewToggle = document.getElementById('secretary-preview-toggle');
  const insightSection = document.getElementById('assistant-insight');
  const insightText = document.getElementById('assistant-insight-text');
  const positionEl = document.getElementById('assistant-position');
  const draftPanel = document.getElementById('secretary-draft');
  const draftText = document.getElementById('secretary-draft-text');
  if (!buttonEl || !messageEl || !chatForm || !chatInput) return;

  const chatPlaceholderHtml = '<div class="chat-placeholder-card">Ask anything about the selected email and I&#39;ll break it down for you.</div>';

  if (chatLog) chatLog.innerHTML = chatPlaceholderHtml;

  if (!threads.length) {
    messageEl.textContent = 'Morning! Inbox is clear—nothing for us to review.';
    buttonEl.disabled = true;
    buttonEl.textContent = 'No emails';
    chatInput.disabled = true;
    if (backButtonEl) backButtonEl.disabled = true;
    return;
  }

  let index = -1;
  let activeThreadId = '';
  const chatHistories = new Map();
  let chatAdvanceTimer = 0;
  let assistantTyping = false;
  buttonEl.dataset.state = 'idle';
  setChatError('');
  toggleComposer(false);
  updateProgressIndicator(0);
  if (backButtonEl) backButtonEl.disabled = true;

  if (previewToggle && previewEl) {
    previewToggle.addEventListener('click', () => {
      if (previewToggle.disabled) return;
      const collapsed = previewEl.classList.toggle('collapsed');
      previewToggle.textContent = collapsed ? 'Expand' : 'Collapse';
    });
  }

  chatForm.addEventListener('submit', async (event) => {
    event.preventDefault();
    if (!activeThreadId) {
      setChatError('Select an email before chatting.');
      return;
    }
    const question = chatInput.value.trim();
    if (!question) return;
    const history = ensureHistory(activeThreadId);
    if (handleNextIntent(question, history)) {
      chatInput.value = '';
      return;
    }
    const asked = history.filter(turn => turn.role === 'user').length;
    if (asked >= MAX_TURNS) {
      setChatError('Chat limit reached for this thread.');
      return;
    }
    setChatError('');
    const pending = { role: 'user', content: question };
    history.push(pending);
    renderChat(activeThreadId);
    chatInput.value = '';
    chatInput.disabled = true;
    const submitBtn = chatForm.querySelector('button[type="submit"]');
    if (submitBtn) submitBtn.disabled = true;
    setAssistantTyping(true);
    try {
      const historyPayload = history.slice(0, -1);
      const resp = await fetch('/secretary/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          threadId: activeThreadId,
          question,
          history: historyPayload
        })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        history.pop();
        renderChat(activeThreadId);
        setChatError(data?.error || 'Something went wrong. Please try again.');
        chatInput.value = question;
        return;
      }
      history.push({ role: 'assistant', content: data.reply || 'No response received.' });
      renderChat(activeThreadId);
    } catch (err) {
      history.pop();
      renderChat(activeThreadId);
      setChatError('Failed to reach the assistant. Check your connection.');
      chatInput.value = question;
    } finally {
      setAssistantTyping(false);
      chatInput.disabled = false;
      const submitBtn = chatForm.querySelector('button[type="submit"]');
      if (submitBtn) submitBtn.disabled = false;
      chatInput.focus();
    }
  });

  chatInput.addEventListener('keydown', (event) => {
    if (event.defaultPrevented) return;
    if (event.key !== 'Enter') return;
    if (event.shiftKey || event.altKey || event.metaKey || event.ctrlKey) return;
    if (chatInput.disabled) return;
    event.preventDefault();
    if (typeof chatForm.requestSubmit === 'function') {
      chatForm.requestSubmit();
    } else {
      chatForm.dispatchEvent(new Event('submit', { bubbles: true, cancelable: true }));
    }
  });

  buttonEl.addEventListener('click', () => {
    const state = buttonEl.dataset.state;
    if (state === 'complete') return;
    if (state === 'ready-to-close') {
      buttonEl.dataset.state = 'complete';
      buttonEl.disabled = true;
      if (backButtonEl) backButtonEl.disabled = true;
      messageEl.textContent = 'All caught up—ping me if you want another pass.';
      chatHistories.clear();
      resetThreadView();
      return;
    }
    const targetIndex = index + 1;
    showThreadAt(targetIndex, 'next');
  });

  if (backButtonEl) {
    backButtonEl.addEventListener('click', () => {
      if (backButtonEl.disabled) return;
      const previousIndex = index - 1;
      showThreadAt(previousIndex, 'back');
    });
  }

  // Automatically open the first thread so the chat composer is usable immediately.
  showThreadAt(0, 'next');

  function toggleComposer(enabled) {
    chatInput.disabled = !enabled;
    const submitBtn = chatForm.querySelector('button[type="submit"]');
    if (submitBtn) submitBtn.disabled = !enabled;
  }

  function focusChatComposer() {
    if (!chatInput || chatInput.disabled) return;
    const moveCaretToEnd = () => {
      chatInput.focus();
      if (typeof chatInput.setSelectionRange === 'function') {
        const end = chatInput.value.length;
        chatInput.setSelectionRange(end, end);
      }
    };
    if (typeof window.requestAnimationFrame === 'function') {
      window.requestAnimationFrame(moveCaretToEnd);
    } else {
      window.setTimeout(moveCaretToEnd, 0);
    }
  }

  function ensureHistory(threadId) {
    if (!threadId) return [];
    if (!chatHistories.has(threadId)) {
      chatHistories.set(threadId, []);
    }
    return chatHistories.get(threadId);
  }

  function ensureIntroPrompt(thread) {
    if (!thread) return;
    const history = ensureHistory(thread.threadId);
    if (!history.length) {
      const intro = thread.primer || buildFallbackPrimer(thread);
      history.push({ role: 'assistant', content: intro });
    }
  }

  function buildFallbackPrimer(thread) {
    const subject = thread?.subject ? '**' + thread.subject + '**' : 'this thread';
    const nextText = thread?.nextStep && thread.nextStep.toLowerCase() !== 'no action'
      ? 'Need to move on "' + thread.nextStep + '"?'
      : 'Want a recap or draft reply?';
    return 'Need more detail on ' + subject + '? ' + nextText;
  }

  function renderChat(threadId) {
    if (!chatLog) return;
    const history = ensureHistory(threadId);
    if (!history.length) {
      let view = chatPlaceholderHtml;
      if (assistantTyping && threadId === activeThreadId) {
        view += typingIndicatorHtml();
      }
      chatLog.innerHTML = view;
      updateChatHint(threadId);
      updateDraftPanel(history);
      return;
    }
    let markup = history.map(turn => {
      if (turn.role === 'assistant') {
        return '<div class="chat-message assistant"><div class="chat-card">' + renderMarkdown(turn.content) + '</div></div>';
      }
      return '<div class="chat-message user"><div class="chat-card">' + htmlEscape(turn.content) + '</div></div>';
    }).join('');
    if (assistantTyping && threadId === activeThreadId) {
      markup += typingIndicatorHtml();
    }
    chatLog.innerHTML = markup;
    chatLog.scrollTop = chatLog.scrollHeight;
    updateChatHint(threadId);
    updateDraftPanel(history);
  }

  function updateChatHint(threadId) {
    if (!chatHint) return;
    const history = ensureHistory(threadId);
    const asked = history.filter(turn => turn.role === 'user').length;
    const remaining = Math.max(0, MAX_TURNS - asked);
    chatHint.textContent = remaining
      ? remaining + ' question' + (remaining === 1 ? '' : 's') + ' remaining in this chat.'
      : 'Chat limit reached. Wrap up this thread to move on.';
  }

  function updateDraftPanel(history) {
    if (!draftPanel || !draftText) return;
    const draft = extractDraft(history);
    if (!draft) {
      draftPanel.classList.add('hidden');
      draftText.value = '';
      return;
    }
    draftPanel.classList.remove('hidden');
    draftText.value = draft;
  }

  function extractDraft(history) {
    if (!Array.isArray(history)) return '';
    for (let i = history.length - 1; i >= 0; i--) {
      const entry = history[i];
      if (!entry || entry.role !== 'assistant') continue;
      const draft = maybeStripDraft(entry.content);
      if (draft) return draft;
    }
    return '';
  }

  function maybeStripDraft(content) {
    if (!content) return '';
    const trimmed = content.trim();
    if (!trimmed) return '';
    if (/^draft/i.test(trimmed)) {
      return trimmed.replace(/^draft[^:]*:/i, '').trim();
    }
    if (trimmed.length > 240 && trimmed.split(/\n/).length >= 4 && /\b(dear|hello|hi)\b/i.test(trimmed)) {
      return trimmed;
    }
    return '';
  }

  function setChatError(message) {
    if (!chatError) return;
    if (message) {
      chatError.textContent = message;
      chatError.classList.remove('hidden');
    } else {
      chatError.textContent = '';
      chatError.classList.add('hidden');
    }
  }

  function handleNextIntent(question, history) {
    if (!activeThreadId || !shouldTriggerNextIntent(question)) return false;
    if (!history || !Array.isArray(history)) return false;
    setChatError('');
    history.push({ role: 'user', content: question });
    renderChat(activeThreadId);
    chatInput.value = '';
    chatInput.disabled = true;
    const submitBtn = chatForm.querySelector('button[type="submit"]');
    if (submitBtn) submitBtn.disabled = true;
    const isLastThread = index >= threads.length - 1;
    const response = isLastThread
      ? 'That was the final email in your queue. Tap Done when you are ready.'
      : 'Proceeding to the next email...';
    history.push({ role: 'assistant', content: response });
    renderChat(activeThreadId);
    if (chatAdvanceTimer) window.clearTimeout(chatAdvanceTimer);
    const delay = isLastThread ? 900 : 700;
    chatAdvanceTimer = window.setTimeout(() => {
      chatAdvanceTimer = 0;
      chatInput.disabled = false;
      if (submitBtn) submitBtn.disabled = false;
      if (!isLastThread) {
        showThreadAt(index + 1, 'next');
      }
    }, delay);
    return true;
  }

  function shouldTriggerNextIntent(rawText) {
    if (!rawText) return false;
    const normalized = rawText.toLowerCase();
    const simple = normalized.replace(/[^a-z0-9\s]/g, ' ').replace(/\s+/g, ' ').trim();
    if (!simple) return false;
    const matchable = simple.replace(/(?: thanks?| thank you)+$/, '').trim() || simple;

    const directPatterns = [
      /^next( (email|one|thread|message|item|mail))?( please)?$/,
      /^(?:skip|pass)(?: (?:this|it)(?: one)?)?( please)?$/,
      /^(?:let s|lets|let us|shall we|we can|can we|could we|please) move on(?: (now|then))?$/,
      /^(?:onto|on to) the next( (one|email|thread|message|item))?$/,
      /^time for the next( (one|email|thread|message|item))?$/,
      /^(?:ready|i m ready|im ready|we re ready|were ready|ok|okay) (?:for|to move on to) (?:the )?next( (email|one|thread|message|item))?( please)?$/,
      /^(?:all set|done|im done|i m done|we re done|were done) (?:here|with (?:this|it)(?: one| email| thread| message)?)$/,
      /^that s all (?:for|with) (?:this|it)(?: one| email| thread| message)?$/
    ];
    const segments = matchable.split(/[,;]+/).map(part => part.trim()).filter(Boolean);
    const targets = segments.length ? segments : [matchable];
    if (targets.some(part => directPatterns.some(pattern => pattern.test(part)))) {
      return true;
    }

    if (/\bnext( (email|one|thread|message|item|mail))? please$/.test(matchable)) {
      return true;
    }

    const targetedCombos = [
      'move on to the next email',
      'move on to the next one',
      'move on to the next thread',
      'move on to the next message',
      'move onto the next email',
      'move onto the next one',
      'move onto the next thread',
      'go to the next email',
      'go to the next one',
      'go to the next thread',
      'go on to the next email',
      'go on to the next one',
      'go on to the next thread',
      'proceed to the next email',
      'proceed to the next one',
      'proceed to the next thread',
      'show me the next email',
      'show me the next one',
      'ready for the next email',
      'ready for the next one',
      'ready to move on to the next',
      'done with this one',
      'done with this email',
      'done with this thread',
      'done with this message',
      'all set with this one',
      'all set with this thread',
      'skip this email',
      'skip this thread',
      'skip this one',
      'pass this email',
      'pass this thread',
      'let s move on to the next',
      'lets move on to the next',
      'let us move on to the next',
      'let s move on',
      'lets move on',
      'let us move on',
      'onto the next email',
      'onto the next one',
      'on to the next email',
      'on to the next one',
      'time for the next email',
      'time for the next one'
    ];
    return targetedCombos.some(text => matchable.includes(text));
  }

  function resetThreadView() {
    activeThreadId = '';
    if (detailEl) detailEl.classList.add('hidden');
    if (chatContainer) chatContainer.classList.add('hidden');
    if (chatLog) chatLog.innerHTML = chatPlaceholderHtml;
    toggleComposer(false);
    if (chatHint) chatHint.textContent = 'Ask up to ' + MAX_TURNS + ' questions per thread.';
    setChatError('');
    if (backButtonEl) backButtonEl.disabled = true;
    index = -1;
    updateProgressIndicator(0);
    if (draftPanel) draftPanel.classList.add('hidden');
    if (draftText) draftText.value = '';
    if (previewToggle) {
      previewToggle.disabled = true;
      previewToggle.classList.add('hidden');
    }
    setAssistantTyping(false);
    buttonEl.disabled = false;
    updateNavigationButtons();
  }

  function htmlEscape(value) {
    const div = document.createElement('div');
    div.textContent = value || '';
    return div.innerHTML;
  }

  function renderMarkdown(value) {
    if (!value) return '';
    const escaped = htmlEscape(value);
    const lines = escaped.split(/\r?\n/);
    const blocks = [];
    const listStack = [];
    let paragraphLines = [];

    function flushParagraph() {
      if (!paragraphLines.length) return;
      const text = paragraphLines.join(' ').replace(/\s+/g, ' ').trim();
      if (text) blocks.push('<p>' + applyInlineMarkdown(text) + '</p>');
      paragraphLines = [];
    }

    function closeLastList() {
      if (!listStack.length) return;
      const list = listStack.pop();
      if (!list) return;
      const html = '<' + list.type + '>' + list.items.join('') + '</' + list.type + '>';
      if (listStack.length) {
        listStack[listStack.length - 1].items.push(html);
      } else {
        blocks.push(html);
      }
    }

    function closeAllLists() {
      while (listStack.length) closeLastList();
    }

    function ensureList(type, indent) {
      while (listStack.length) {
        const current = listStack[listStack.length - 1];
        if (current.indent > indent) {
          closeLastList();
          continue;
        }
        if (current.indent === indent && current.type !== type) {
          closeLastList();
          continue;
        }
        break;
      }
      const current = listStack[listStack.length - 1];
      if (!current || current.indent !== indent || current.type !== type) {
        listStack.push({ type, indent, items: [] });
      }
    }

    for (const rawLine of lines) {
      const indentMatch = rawLine.match(/^\s*/);
      const indent = indentMatch ? indentMatch[0].length : 0;
      const trimmed = rawLine.trim();
      if (!trimmed) {
        flushParagraph();
        closeAllLists();
        continue;
      }
      const unorderedMatch = trimmed.match(/^[-*+]\s+(.*)$/);
      if (unorderedMatch) {
        flushParagraph();
        ensureList('ul', indent);
        const current = listStack[listStack.length - 1];
        current.items.push('<li>' + applyInlineMarkdown(unorderedMatch[1].trim()) + '</li>');
        continue;
      }
      const orderedMatch = trimmed.match(/^(\d+)\.\s+(.*)$/);
      if (orderedMatch) {
        flushParagraph();
        ensureList('ol', indent);
        const current = listStack[listStack.length - 1];
        current.items.push('<li>' + applyInlineMarkdown(orderedMatch[2].trim()) + '</li>');
        continue;
      }
      closeAllLists();
      paragraphLines.push(trimmed);
    }

    flushParagraph();
    closeAllLists();

    if (!blocks.length) {
      return applyInlineMarkdown(escaped).replace(/(?:\r\n|\r|\n)/g, '<br>');
    }
    return blocks.join('');
  }

  function applyInlineMarkdown(text) {
    if (!text) return '';
    let result = text;
    result = result.replace(/\`([^\`]+)\`/g, '<code>$1</code>');
    result = result.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noreferrer">$1</a>');
    result = result.replace(/~~([^~]+)~~/g, '<del>$1</del>');
    result = result.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    result = result.replace(/__([^_]+)__/g, '<strong>$1</strong>');
    result = result.replace(/(^|[\s>])\*([^*\n]+)\*(?=[\s<.,!?:;]|$)/g, function (_, prefix, content) {
      return prefix + '<em>' + content + '</em>';
    });
    result = result.replace(/(^|[\s>])_([^_\n]+)_(?=[\s<.,!?:;]|$)/g, function (_, prefix, content) {
      return prefix + '<em>' + content + '</em>';
    });
    return result;
  }

  function showThreadAt(newIndex, direction) {
    if (typeof newIndex !== 'number') return;
    if (newIndex < 0 || newIndex >= threads.length) return;
    index = newIndex;
    const current = threads[index];
    if (!current) return;
    if (detailEl) detailEl.classList.remove('hidden');
    if (chatContainer) chatContainer.classList.remove('hidden');
    activeThreadId = current.threadId;
    ensureIntroPrompt(current);
    renderChat(activeThreadId);
    setChatError('');
    toggleComposer(true);
    focusChatComposer();

    const senderParts = splitSender(current.from);
    if (avatarEl) avatarEl.textContent = senderParts.initials;
    if (fromEl) fromEl.textContent = senderParts.name || senderParts.email || 'Sender unknown';
    if (receivedEl) receivedEl.textContent = formatMetaLine(senderParts.email, current.receivedAt);
    if (subjectEl) subjectEl.textContent = current.subject || '(no subject)';
    if (summaryEl) summaryEl.textContent = current.summary || 'No summary captured.';
    if (nextEl) nextEl.textContent = current.nextStep || 'No next step needed.';
    if (chipCategoryEl) chipCategoryEl.textContent = current.category || 'General';
    if (chipEffortEl) chipEffortEl.textContent = 'Effort: ' + computeEffortLevel(current.nextStep);
    if (chipReplyEl) chipReplyEl.classList.toggle('hidden', !computeNeedsReply(current.nextStep));
    updatePreviewPanel(current);
    updateInsight(current);

    if (linkEl) {
      if (current.link) {
        linkEl.href = current.link;
        linkEl.classList.remove('hidden');
      } else {
        linkEl.removeAttribute('href');
        linkEl.classList.add('hidden');
      }
    }

    updateMessageAfterNavigation(direction);
    updateNavigationButtons();
  }

  function updatePreviewPanel(current) {
    if (!previewEl || !previewToggle) return;
    const preview = (current.convo || '').trim();
    if (!preview) {
      previewEl.textContent = 'No preview captured yet.';
      previewEl.classList.remove('collapsed');
      previewToggle.disabled = true;
      previewToggle.classList.add('hidden');
      return;
    }
    previewEl.textContent = preview;
    const shouldCollapse = preview.split(/\n/).length > 6 || preview.length > 480;
    if (shouldCollapse) {
      previewEl.classList.add('collapsed');
      previewToggle.disabled = false;
      previewToggle.classList.remove('hidden');
      previewToggle.textContent = 'Expand';
    } else {
      previewEl.classList.remove('collapsed');
      previewToggle.disabled = true;
      previewToggle.classList.add('hidden');
    }
  }

  function updateInsight(current) {
    if (!insightSection || !insightText) return;
    const insight = (current.headline || current.nextStep || '').trim();
    if (!insight) {
      insightSection.classList.add('hidden');
      insightText.textContent = '';
      return;
    }
    insightSection.classList.remove('hidden');
    insightText.textContent = insight;
  }

  function computeNeedsReply(text) {
    if (!text) return false;
    const normalized = text.toLowerCase();
    return /reply|respond|let (?:them|me) know|follow up|confirm/.test(normalized);
  }

  function computeEffortLevel(text) {
    if (!text) return 'Low';
    const words = text.split(/\s+/).filter(Boolean).length;
    if (words > 40) return 'High';
    if (words > 18) return 'Medium';
    return 'Low';
  }

  function splitSender(raw) {
    if (!raw) return { name: '', email: '', initials: '–' };
    const emailMatch = raw.match(/<([^>]+)>/);
    const email = emailMatch ? emailMatch[1].trim() : (raw.includes('@') ? raw.trim() : '');
    let name = raw;
    if (emailMatch) {
      name = name.replace(emailMatch[0], '').trim();
    } else if (email) {
      name = name.replace(email, '').replace(/[<>]/g, '').trim();
    }
    const base = name || email;
    const initials = base
      ? base.split(/\s+/).slice(0, 2).map(part => part.charAt(0)).join('').toUpperCase()
      : '–';
    return { name, email, initials: initials || '–' };
  }

  function formatMetaLine(email, isoDate) {
    const parts = [];
    if (email) parts.push(email);
    const stamp = formatReceived(isoDate);
    if (stamp) parts.push(stamp);
    return parts.join(' • ') || 'Details coming soon.';
  }

  function formatReceived(isoDate) {
    if (!isoDate) return '';
    const date = new Date(isoDate);
    if (Number.isNaN(date.getTime())) return '';
    const now = Date.now();
    const diff = now - date.getTime();
    if (diff < 24 * 60 * 60 * 1000) {
      return 'Today · ' + date.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
    }
    return date.toLocaleString([], { weekday: 'short', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
  }

  function updateNavigationButtons() {
    updateProgressIndicator(index + 1);
    if (index === threads.length - 1) {
      buttonEl.textContent = 'Done';
      buttonEl.dataset.state = 'ready-to-close';
    } else if (index >= 0) {
      buttonEl.textContent = 'Next email →';
      buttonEl.dataset.state = 'chatting';
    } else {
      buttonEl.textContent = 'Start review';
      buttonEl.dataset.state = 'idle';
    }
    if (backButtonEl) backButtonEl.disabled = index <= 0;
  }

  function updateProgressIndicator(position) {
    if (!positionEl) return;
    if (!threads.length) {
      positionEl.textContent = '0 of 0';
      return;
    }
    if (!position || position < 1) {
      positionEl.textContent = '0 of ' + threads.length;
      return;
    }
    positionEl.textContent = position + ' of ' + threads.length;
  }

  function updateMessageAfterNavigation(direction) {
    if (!messageEl) return;
    if (threads.length === 1) {
      messageEl.textContent = 'Only one email waiting. Tap Done when you are finished.';
      return;
    }
    const position = index + 1;
    if (direction === 'back') {
      messageEl.textContent = 'Back to email ' + position + ' of ' + threads.length + '.';
      return;
    }
    if (position === 1) {
      messageEl.textContent = "Great, here's email 1 of " + threads.length + '.';
      return;
    }
    if (position === threads.length) {
      messageEl.textContent = 'Last one—email ' + threads.length + ' of ' + threads.length + '.';
      return;
    }
    messageEl.textContent = 'Reviewing email ' + position + ' of ' + threads.length + '.';
  }

  function setAssistantTyping(state) {
    if (assistantTyping === state) return;
    assistantTyping = state;
    if (activeThreadId) {
      renderChat(activeThreadId);
    } else if (!state) {
      renderChat('');
    }
  }

  function typingIndicatorHtml() {
    return '<div class="chat-message assistant typing"><div class="chat-card" role="status"><span class="typing-dots" aria-hidden="true"><span class="typing-dot"></span><span class="typing-dot"></span><span class="typing-dot"></span></span><span class="typing-label">Secretary is preparing a reply…</span></div></div>';
  }

  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initSecretaryAssistant, { once: true });
  } else {
    initSecretaryAssistant();
  }
})();
