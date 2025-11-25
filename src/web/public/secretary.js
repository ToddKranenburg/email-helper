(function () {
  const bootstrap = window.SECRETARY_BOOTSTRAP || {};
  const threads = Array.isArray(bootstrap.threads) ? bootstrap.threads : [];
  const MAX_TURNS = typeof bootstrap.maxTurns === 'number' ? bootstrap.maxTurns : 0;
  const TOTAL_ITEMS = typeof bootstrap.totalItems === 'number' ? bootstrap.totalItems : threads.length;
  window.SECRETARY_BOOTSTRAP = undefined;

  const refs = {
    count: document.getElementById('triage-count'),
    progress: document.getElementById('triage-progress'),
    progressTrack: document.querySelector('.triage-progress'),
    queuePill: document.getElementById('queue-pill-text'),
    emailCard: document.getElementById('email-card'),
    emailEmpty: document.getElementById('email-empty'),
    emailEmptyText: document.querySelector('#email-empty p'),
    avatar: document.getElementById('email-avatar'),
    sender: document.getElementById('email-sender'),
    received: document.getElementById('email-received'),
    subject: document.getElementById('email-subject'),
    position: document.getElementById('email-position'),
    cta: document.getElementById('email-cta'),
    preview: document.getElementById('email-preview'),
    previewToggle: document.getElementById('preview-toggle'),
    chatLog: document.getElementById('assistant-chat-log'),
    chatScroll: document.getElementById('assistant-chat'),
    chatForm: document.getElementById('assistant-form'),
    chatInput: document.getElementById('assistant-input-field'),
    chatError: document.getElementById('assistant-error'),
    chatHint: document.getElementById('assistant-hint'),
    doneBtn: document.getElementById('action-done'),
    skipBtn: document.getElementById('action-skip'),
    mapToggle: document.getElementById('map-toggle'),
    drawer: document.getElementById('inbox-drawer'),
    drawerClose: document.getElementById('drawer-close'),
    needsList: document.getElementById('needs-list'),
    doneList: document.getElementById('done-list'),
    needsCount: document.getElementById('needs-count'),
    doneCount: document.getElementById('done-count'),
    doneDetails: document.getElementById('done-details')
  };

  if (!refs.chatLog || !refs.chatForm || !refs.emailCard || !refs.emailEmpty) {
    return;
  }

  const state = {
    lookup: new Map(),
    positions: new Map(),
    needs: [],
    done: [],
    histories: new Map(),
    activeId: '',
    typing: false,
    headerTotal: TOTAL_ITEMS || threads.length,
    autoAdvanceTimer: 0
  };

  threads.forEach((thread, index) => {
    if (!thread || !thread.threadId) return;
    state.lookup.set(thread.threadId, thread);
    state.positions.set(thread.threadId, index);
    state.needs.push(thread.threadId);
  });

  init();

  function init() {
    updateHeaderCount();
    updateProgress();
    updateQueuePill();
    updateDrawerLists();
    wireEvents();

    if (refs.previewToggle && refs.preview) {
      refs.previewToggle.addEventListener('click', () => {
        if (refs.previewToggle.disabled) return;
        const isHidden = refs.preview.classList.toggle('hidden');
        refs.previewToggle.textContent = isHidden ? 'See email body' : 'Hide email body';
        if (!isHidden) {
          refs.preview.scrollTop = 0;
        }
      });
    }

    if (state.needs.length) {
      setActiveThread(state.needs[0]);
    } else {
      setEmptyState('No emails queued. Tap Sync Gmail to pull fresh ones.');
      toggleComposer(false);
    }
  }

  function wireEvents() {
    refs.chatForm.addEventListener('submit', handleChatSubmit);
    refs.chatInput.addEventListener('keydown', handleChatKeydown);

    refs.doneBtn.addEventListener('click', () => markCurrentDone('button'));
    refs.skipBtn.addEventListener('click', () => skipCurrent('button'));

    if (refs.mapToggle && refs.drawer) {
      refs.mapToggle.addEventListener('click', () => toggleDrawer(true));
      refs.drawer.addEventListener('click', (event) => {
        if (event.target === refs.drawer) toggleDrawer(false);
      });
      if (refs.drawerClose) {
        refs.drawerClose.addEventListener('click', () => toggleDrawer(false));
      }
      document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape' && !refs.drawer.classList.contains('hidden')) {
          toggleDrawer(false);
        }
      });
      if (refs.needsList) refs.needsList.addEventListener('click', handleDrawerClick);
      if (refs.doneList) refs.doneList.addEventListener('click', handleDrawerClick);
    }
  }

  function handleDrawerClick(event) {
    const target = event.target.closest('.drawer-thread');
    if (!target) return;
    const threadId = target.dataset.threadId;
    if (!threadId || !state.lookup.has(threadId)) return;
    toggleDrawer(false);
    setActiveThread(threadId);
  }

  function handleChatKeydown(event) {
    if (event.defaultPrevented) return;
    if (event.key !== 'Enter') return;
    if (event.shiftKey || event.altKey || event.metaKey || event.ctrlKey) return;
    if (refs.chatInput.disabled) return;
    event.preventDefault();
    refs.chatForm.requestSubmit();
  }

  async function handleChatSubmit(event) {
    event.preventDefault();
    if (!state.activeId) return;
    const question = refs.chatInput.value.trim();
    if (!question) return;

    const intent = detectIntent(question);
    if (intent === 'done' || intent === 'skip') {
      refs.chatInput.value = '';
      handleAutoIntent(intent, question);
      return;
    }

    const history = ensureHistory(state.activeId);
    const asked = history.filter(turn => turn.role === 'user').length;
    if (MAX_TURNS > 0 && asked >= MAX_TURNS) {
      setChatError('Chat limit reached for this email.');
      return;
    }
    setChatError('');

    history.push({ role: 'user', content: question });
    renderChat(state.activeId);
    refs.chatInput.value = '';
    toggleComposer(false);
    setAssistantTyping(true);
    const submitBtn = refs.chatForm.querySelector('button[type="submit"]');
    if (submitBtn) submitBtn.disabled = true;

    try {
      const historyPayload = history.slice(0, -1);
      const resp = await fetch('/secretary/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          threadId: state.activeId,
          question,
          history: historyPayload
        })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        history.pop();
        renderChat(state.activeId);
        setChatError(data?.error || 'Something went wrong. Try again.');
        refs.chatInput.value = question;
        return;
      }
      history.push({ role: 'assistant', content: data.reply || 'No response received.' });
      renderChat(state.activeId);
    } catch (err) {
      history.pop();
      renderChat(state.activeId);
      setChatError('Unable to reach the assistant. Check your connection.');
      refs.chatInput.value = question;
    } finally {
      setAssistantTyping(false);
      toggleComposer(true);
      const submitBtn2 = refs.chatForm.querySelector('button[type="submit"]');
      if (submitBtn2) submitBtn2.disabled = false;
      refs.chatInput.focus();
    }
  }

  function setActiveThread(threadId) {
    if (!threadId || !state.lookup.has(threadId)) return;
    clearAutoAdvance();
    state.activeId = threadId;
    const thread = state.lookup.get(threadId);
    if (!thread) return;

    refs.emailCard.classList.remove('hidden');
    refs.emailEmpty.classList.add('hidden');
    if (refs.mapToggle) {
      refs.mapToggle.disabled = false;
      refs.mapToggle.removeAttribute('aria-disabled');
    }

    updateEmailCard(thread);
    ensureHistory(threadId);
    setChatError('');
    renderChat(threadId);
    updateHint(threadId);
    updateDrawerLists();
    updateQueuePill();
    toggleComposer(true);
    refs.chatInput.value = '';
  }

  function updateEmailCard(thread) {
    if (refs.position) {
      const label = formatEmailPosition(thread.threadId);
      refs.position.textContent = label || '';
      refs.position.classList.toggle('hidden', !label);
    }
    if (refs.avatar) refs.avatar.textContent = initialsFromSender(thread.from);
    if (refs.sender) refs.sender.textContent = thread.from || 'Unknown sender';
    if (refs.received) refs.received.textContent = formatTimestamp(thread.receivedAt);
    if (refs.subject) refs.subject.textContent = thread.subject || '(no subject)';
    if (refs.cta) {
      if (thread.link) {
        refs.cta.href = thread.link;
        refs.cta.classList.remove('hidden');
      } else {
        refs.cta.classList.add('hidden');
        refs.cta.removeAttribute('href');
      }
    }
    if (refs.preview && refs.previewToggle) {
      const previewText = (thread.convo || '').trim();
      if (previewText) {
        refs.preview.textContent = previewText;
        refs.preview.classList.add('hidden');
        refs.previewToggle.disabled = false;
        refs.previewToggle.textContent = 'See email body';
      } else {
        refs.preview.textContent = 'Email body is unavailable for this message.';
        refs.preview.classList.remove('hidden');
        refs.previewToggle.disabled = true;
        refs.previewToggle.textContent = 'Email body unavailable';
      }
    }
  }

  function updateHeaderCount() {
    if (!refs.count) return;
    const total = state.headerTotal;
    const label = total
      ? `${total} email${total === 1 ? '' : 's'}`
      : 'No emails';
    refs.count.textContent = `· ${label}`;
  }

  function updateQueuePill() {
    if (!refs.queuePill) return;
    const remaining = state.needs.length;
    refs.queuePill.textContent = remaining
      ? `${remaining} remaining`
      : 'All done';
  }

  function updateProgress() {
    if (!refs.progress) return;
    const total = state.headerTotal || (state.needs.length + state.done.length) || 1;
    const done = state.done.length;
    const pct = Math.min(100, Math.round((done / total) * 100));
    refs.progress.style.width = `${pct}%`;
    if (refs.progressTrack) {
      refs.progressTrack.setAttribute('aria-valuenow', String(pct));
    }
  }

  function updateDrawerLists() {
    if (!refs.needsList || !refs.doneList) return;
    refs.needsList.innerHTML = '';
    refs.doneList.innerHTML = '';

    if (state.needs.length) {
      state.needs.forEach(id => appendDrawerItem(refs.needsList, id));
    } else {
      refs.needsList.innerHTML = '<li class="drawer-empty">Nothing queued up.</li>';
    }

    if (state.done.length) {
      state.done.forEach(id => appendDrawerItem(refs.doneList, id));
    } else {
      refs.doneList.innerHTML = '<li class="drawer-empty">No finished emails yet.</li>';
    }

    if (refs.needsCount) refs.needsCount.textContent = String(state.needs.length);
    if (refs.doneCount) refs.doneCount.textContent = String(state.done.length);
  }

  function appendDrawerItem(listEl, threadId) {
    const thread = state.lookup.get(threadId);
    if (!thread) return;
    const li = document.createElement('li');
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'drawer-thread' + (threadId === state.activeId ? ' active' : '');
    btn.dataset.threadId = threadId;
    btn.innerHTML = `<strong>${htmlEscape(thread.from || 'Unknown sender')}</strong><span>${htmlEscape(thread.subject || '(no subject)')}</span>`;
    li.appendChild(btn);
    listEl.appendChild(li);
  }

  function updateHint(threadId) {
    if (!refs.chatHint) return;
    if (!MAX_TURNS || !threadId) {
      refs.chatHint.textContent = 'Ask anything or tap Done / Skip.';
      return;
    }
    const history = ensureHistory(threadId);
    const asked = history.filter(turn => turn.role === 'user').length;
    const remaining = Math.max(0, MAX_TURNS - asked);
    refs.chatHint.textContent = remaining
      ? `${remaining} question${remaining === 1 ? '' : 's'} left on this email.`
      : 'Chat limit reached here.';
  }

  function ensureHistory(threadId) {
    if (!state.histories.has(threadId)) {
      state.histories.set(threadId, []);
    }
    const history = state.histories.get(threadId);
    if (!history.length) {
      const intro = buildIntroMessage(state.lookup.get(threadId));
      history.push({ role: 'assistant', content: intro });
    }
    return history;
  }

  function buildIntroMessage(thread) {
    if (!thread) return 'Need a quick summary or a draft? I can help.';
    const primer = (thread.primer || '').trim();
    if (primer) return primer;

    const sender = thread.from ? thread.from.split('<')[0].trim() || thread.from : '';
    const subject = (thread.subject || '').trim();
    const summary = (thread.summary || thread.headline || '').split('\n')[0]?.trim() || '';
    const contextParts = [];
    if (sender) contextParts.push(sender);
    if (subject) contextParts.push(`about “${subject}”`);
    let context = contextParts.join(' ');
    if (summary) {
      context = context ? `${context} — ${summary}` : summary;
    }
    context = context || 'an email that needs your call';
    const starter = `Heads up: ${context}.`;
    const normalizedNext = (thread.nextStep || '').trim();
    const hasNext = normalizedNext && normalizedNext.toLowerCase() !== 'no action';
    const followUps = [
      'Want me to draft something and send it for you?',
      'Want me to nudge them so it moves along?',
      'Should I remind you about it later?'
    ];
    const next = hasNext
      ? `Want me to run with "${normalizedNext}"?`
      : followUps[Math.floor(Math.random() * followUps.length)];
    return `${starter} ${next}`;
  }

  function renderChat(threadId) {
    if (!refs.chatLog) return;
    const history = ensureHistory(threadId);
    if (!history.length) {
      let view = chatPlaceholder();
      if (state.typing && threadId === state.activeId) {
        view += typingIndicatorHtml();
      }
      refs.chatLog.innerHTML = view;
      return;
    }
    let markup = history.map(turn => {
      if (turn.role === 'assistant') {
        return `<div class="chat-message assistant"><div class="chat-card">${renderMarkdown(turn.content)}</div></div>`;
      }
      return `<div class="chat-message user"><div class="chat-card">${htmlEscape(turn.content)}</div></div>`;
    }).join('');
    if (state.typing && threadId === state.activeId) {
      markup += typingIndicatorHtml();
    }
    refs.chatLog.innerHTML = markup;
    if (refs.chatScroll) refs.chatScroll.scrollTop = refs.chatScroll.scrollHeight;
  }

  function chatPlaceholder() {
    return '<div class="chat-placeholder-card">I’ll drop the gist and nudge you forward.</div>';
  }

  function setChatError(message) {
    if (!refs.chatError) return;
    if (message) {
      refs.chatError.textContent = message;
      refs.chatError.classList.remove('hidden');
    } else {
      refs.chatError.textContent = '';
      refs.chatError.classList.add('hidden');
    }
  }

  function toggleComposer(enabled) {
    refs.chatInput.disabled = !enabled || !state.activeId;
    refs.doneBtn.disabled = !enabled || !state.activeId;
    refs.skipBtn.disabled = !enabled || !state.activeId;
  }

  function setAssistantTyping(value) {
    state.typing = Boolean(value);
    renderChat(state.activeId);
  }

  function markCurrentDone(source) {
    if (!state.activeId || !state.needs.length) return;
    const threadId = state.activeId;
    const index = state.needs.indexOf(threadId);
    if (index === -1) return;
    state.needs.splice(index, 1);
    state.done.unshift(threadId);
    updateProgress();
    updateDrawerLists();

    if (!state.needs.length) {
      setEmptyState('Inbox clear. Celebrate the tiny win.');
      toggleComposer(false);
      return;
    }
    const nextId = state.needs[index] || state.needs[0];
    setActiveThread(nextId);
  }

  function skipCurrent(source) {
    if (!state.activeId || state.needs.length <= 1) {
      return;
    }
    const threadId = state.activeId;
    const index = state.needs.indexOf(threadId);
    if (index === -1) return;
    state.needs.splice(index, 1);
    state.needs.push(threadId);
    updateDrawerLists();
    const nextId = state.needs[index] || state.needs[0];
    setActiveThread(nextId);
  }

  function handleAutoIntent(intent, userText) {
    if (!state.activeId) return;
    const history = ensureHistory(state.activeId);
    history.push({ role: 'user', content: userText });
    renderChat(state.activeId);

    const acknowledgement = intent === 'done'
      ? 'Cool — marking this handled and moving forward.'
      : 'Skipping for now. It stays in Needs Review.';
    history.push({ role: 'assistant', content: acknowledgement });
    renderChat(state.activeId);
    updateHint(state.activeId);

    clearAutoAdvance();
    const targetId = state.activeId;
    state.autoAdvanceTimer = window.setTimeout(() => {
      if (state.activeId !== targetId) return;
      if (intent === 'done') {
        markCurrentDone('auto');
      } else {
        skipCurrent('auto');
      }
      clearAutoAdvance();
    }, 600);
  }

  function clearAutoAdvance() {
    if (state.autoAdvanceTimer) {
      window.clearTimeout(state.autoAdvanceTimer);
      state.autoAdvanceTimer = 0;
    }
  }

  function setEmptyState(message) {
    state.activeId = '';
    refs.emailCard.classList.add('hidden');
    refs.emailEmpty.classList.remove('hidden');
    if (refs.position) {
      refs.position.textContent = '';
      refs.position.classList.add('hidden');
    }
    if (refs.emailEmptyText) refs.emailEmptyText.textContent = message;
    if (refs.mapToggle) {
      refs.mapToggle.setAttribute('aria-disabled', 'true');
      refs.mapToggle.disabled = true;
    }
    refs.chatLog.innerHTML = chatPlaceholder();
    updateQueuePill();
  }

  function detectIntent(text) {
    const cleaned = text.replace(/[.!?]/g, '').trim().toLowerCase();
    if (!cleaned) return '';
    const donePhrases = ['ok', 'done', 'next', 'cool', 'handled', 'all good', 'all set', 'next one', 'next email'];
    if (donePhrases.includes(cleaned)) return 'done';
    if (cleaned === 'skip' || cleaned === 'skip it') return 'skip';
    return '';
  }

  function toggleDrawer(open) {
    if (!refs.drawer || !refs.mapToggle) return;
    if (open) {
      refs.drawer.classList.remove('hidden');
      refs.drawer.setAttribute('aria-hidden', 'false');
      refs.mapToggle.setAttribute('aria-expanded', 'true');
    } else {
      refs.drawer.classList.add('hidden');
      refs.drawer.setAttribute('aria-hidden', 'true');
      refs.mapToggle.setAttribute('aria-expanded', 'false');
    }
  }

  function initialsFromSender(fromLine) {
    if (!fromLine) return '–';
    const clean = fromLine.replace(/<[^>]*>/g, '').trim();
    const words = clean.split(/\s+/).filter(Boolean);
    const initials = words.slice(0, 2).map(word => word[0]?.toUpperCase() || '').join('');
    return initials || (fromLine[0]?.toUpperCase() || '–');
  }

  function formatTimestamp(iso) {
    if (!iso) return 'Timestamp unavailable';
    const date = new Date(iso);
    if (!Number.isFinite(date.getTime())) return 'Timestamp unavailable';
    return date.toLocaleString(undefined, { dateStyle: 'medium', timeStyle: 'short' });
  }

  function formatEmailPosition(threadId) {
    if (!threadId) return '';
    const total = Math.max(state.headerTotal || 0, state.positions.size || 0);
    if (!total) return '';
    const position = state.positions.get(threadId);
    if (typeof position !== 'number') return '';
    const current = position + 1;
    return `${current} of ${total}`;
  }

  function htmlEscape(value) {
    const div = document.createElement('div');
    div.textContent = value || '';
    return div.innerHTML;
  }

  function renderMarkdown(value) {
    if (!value) return '';
    const escaped = htmlEscape(value);
    return escaped.replace(/(?:\r\n|\r|\n)/g, '<br>');
  }

  function typingIndicatorHtml() {
    return '<div class="chat-message assistant"><div class="typing-dots"><span class="typing-dot"></span><span class="typing-dot"></span><span class="typing-dot"></span></div></div>';
  }
})();
