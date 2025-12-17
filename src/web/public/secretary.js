(function () {
  const bootstrap = window.SECRETARY_BOOTSTRAP || {};
  const threads = Array.isArray(bootstrap.threads) ? bootstrap.threads : [];
  const MAX_TURNS = typeof bootstrap.maxTurns === 'number' ? bootstrap.maxTurns : 0;
  const PAGE_SIZE = typeof bootstrap.pageSize === 'number' ? bootstrap.pageSize : 20;
  const HAS_MORE = Boolean(bootstrap.hasMore);
  const NEXT_PAGE = typeof bootstrap.nextPage === 'number' ? bootstrap.nextPage : (HAS_MORE ? 2 : 0);
  const PRIMER_ENDPOINT = '/secretary/primer';
  const PRIMER_RETRY_DELAY = 1500;
  const MAX_PRIMER_POLLS = 8;
  const DEFAULT_NUDGE = 'Type your next move here.';
  const REVIEW_PROMPT = 'Give me a more detailed but easy-to-digest summary of this email. Highlight the main points, asks, deadlines, and any decisions in quick bullets.';
  window.SECRETARY_BOOTSTRAP = undefined;

  const refs = {
    count: document.getElementById('triage-count'),
    loadMoreHead: document.getElementById('load-more-head'),
    progress: document.getElementById('triage-progress'),
    progressTrack: document.querySelector('.triage-progress'),
    queuePill: document.getElementById('queue-pill-text'),
    emailCard: document.getElementById('email-card'),
    emailEmpty: document.getElementById('email-empty'),
    loadMoreEmpty: document.getElementById('load-more-empty'),
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
    reviewBtn: document.getElementById('action-review'),
    archiveBtn: document.getElementById('action-archive'),
    moreBtn: document.getElementById('action-more'),
    moreMenu: document.getElementById('more-menu'),
    taskPanel: document.getElementById('task-panel'),
    taskPanelHelper: document.getElementById('task-panel-helper'),
    taskTitle: document.getElementById('task-title'),
    taskNotes: document.getElementById('task-notes'),
    taskDue: document.getElementById('task-due'),
    taskError: document.getElementById('task-error'),
    taskSuccess: document.getElementById('task-success'),
    taskSuccessMeta: document.getElementById('task-success-meta'),
    taskCancel: document.getElementById('task-cancel'),
    taskSubmit: document.getElementById('task-submit'),
    taskReset: document.getElementById('task-reset'),
    taskClose: document.getElementById('task-close'),
    mapToggle: document.getElementById('map-toggle'),
    drawer: document.getElementById('inbox-drawer'),
    drawerClose: document.getElementById('drawer-close'),
    needsList: document.getElementById('needs-list'),
    needsCount: document.getElementById('needs-count'),
    reviewList: document.getElementById('review-list')
  };

  if (!refs.chatLog || !refs.chatForm || !refs.emailCard || !refs.emailEmpty) {
    return;
  }

  const state = {
    lookup: new Map(),
    positions: new Map(),
    needs: [],
    histories: new Map(),
    timeline: [],
    activeId: '',
    typing: false,
    totalLoaded: threads.length,
    pageSize: PAGE_SIZE,
    hasMore: HAS_MORE,
    nextPage: NEXT_PAGE,
    loadingMore: false,
    autoAdvanceTimer: 0,
    prepTyping: '',
    pendingCreateThreadId: '',
    pendingArchiveThreadId: '',
    pendingSuggestedActions: new Map()
  };
  const taskState = {
    open: false,
    status: 'idle', // idle | submitting | success | error
    suggested: { title: '', notes: '', due: '' },
    values: { title: '', notes: '', due: '' },
    error: '',
    lastSourceId: ''
  };
  const reviewedIds = new Set();
  const primerStatus = new Map();
  const primerRetryTimers = new Map();
  let composerNudgeTimer = 0;

  const markedLib = resolveMarked();
  const linkify = typeof window.linkifyIt === 'function' ? window.linkifyIt() : null;
  const sanitizeHtml = typeof window.DOMPurify?.sanitize === 'function'
    ? (html) => window.DOMPurify.sanitize(html, { ADD_ATTR: ['target', 'rel'] })
    : (html) => html;

  if (markedLib?.setOptions) {
    markedLib.setOptions({ breaks: true, mangle: false, headerIds: false });
  }

  threads.forEach((thread, index) => {
    if (!thread || !thread.threadId) return;
    thread.primer = typeof thread.primer === 'string' ? thread.primer.trim() : '';
    thread.messageId = typeof thread.messageId === 'string' ? thread.messageId.trim() : '';
    state.lookup.set(thread.threadId, thread);
    state.positions.set(thread.threadId, index);
    state.needs.push(thread.threadId);
    primerStatus.set(thread.threadId, thread.primer ? 'ready' : 'idle');
  });
  state.totalLoaded = state.positions.size;

  init();

  function init() {
    updateHeaderCount();
    updateProgress();
    updateQueuePill();
    updateDrawerLists();
    updateLoadMoreButtons();
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
    refs.chatInput.addEventListener('input', clearComposerNudge);

    if (refs.reviewBtn) {
      refs.reviewBtn.addEventListener('click', () => requestReview());
    }
    if (refs.archiveBtn) {
      refs.archiveBtn.addEventListener('click', () => archiveCurrent('button'));
    }
    if (refs.moreBtn) {
      refs.moreBtn.addEventListener('click', toggleMoreMenu);
    }
    if (refs.moreMenu) {
      refs.moreMenu.addEventListener('click', handleMoreMenuClick);
    }

    if (refs.loadMoreHead) {
      refs.loadMoreHead.addEventListener('click', async () => {
        await fetchNextPage('heading');
        nudgeComposer(DEFAULT_NUDGE, { focus: false });
      });
    }
    if (refs.loadMoreEmpty) {
      refs.loadMoreEmpty.addEventListener('click', async () => {
        await fetchNextPage('empty-card');
        nudgeComposer(DEFAULT_NUDGE, { focus: false });
      });
    }

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
      if (refs.needsList) refs.needsList.addEventListener('click', (event) => handleThreadListClick(event, 'drawer'));
    }
    if (refs.reviewList) refs.reviewList.addEventListener('click', (event) => handleThreadListClick(event, 'queue'));

    document.addEventListener('click', (event) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
      if (refs.moreMenu && !refs.moreMenu.classList.contains('hidden')) {
        const inMenu = target.closest('#more-menu');
        const toggle = target.closest('#action-more');
        if (!inMenu && !toggle) hideMoreMenu();
      }
      const btn = target.closest('button');
      if (!btn) return;
      if (btn.closest('#assistant-form')) return;
      if (!refs.chatInput || refs.chatInput.disabled) return;
      nudgeComposer(DEFAULT_NUDGE, { focus: false });
    });

    if (refs.taskSubmit) {
      refs.taskSubmit.addEventListener('click', async (event) => {
        const result = await submitTask(event);
        if (result?.ok) {
          clearPendingCreate();
          promptArchiveAfterTask(result, { includeSuccess: true });
        }
      });
    }
    if (refs.taskCancel) refs.taskCancel.addEventListener('click', () => closeTaskPanel(true));
    if (refs.taskReset) refs.taskReset.addEventListener('click', resetTaskToSuggested);
    if (refs.taskClose) refs.taskClose.addEventListener('click', () => closeTaskPanel(true));
    if (refs.taskTitle) refs.taskTitle.addEventListener('input', syncTaskValues);
    if (refs.taskNotes) refs.taskNotes.addEventListener('input', syncTaskValues);
    if (refs.taskDue) refs.taskDue.addEventListener('input', syncTaskValues);
  }

  function handleThreadListClick(event, variant) {
    const targetEl = event.target;
    if (!(targetEl instanceof Element)) return;
    const loadMoreBtn = targetEl.closest('.load-more');
    if (loadMoreBtn) {
      fetchNextPage('button');
      return;
    }
    const selector = variant === 'queue' ? '.queue-item' : '.drawer-thread';
    const target = targetEl.closest(selector);
    if (!target) return;
    const threadId = target.dataset.threadId;
    if (!threadId || !state.lookup.has(threadId)) return;
    if (variant === 'drawer') toggleDrawer(false);
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

    const history = ensureHistory(state.activeId);
    const asked = history.filter(turn => turn.role === 'user').length;
    const pendingCreate = isCreateConfirmationPending(state.activeId);
    const pendingArchive = isArchiveConfirmationPending(state.activeId);
    const pendingSuggested = Boolean(getPendingSuggestedAction(state.activeId));
    if (!pendingCreate && !pendingArchive && !pendingSuggested && MAX_TURNS > 0 && asked >= MAX_TURNS) {
      setChatError('Chat limit reached for this email.');
      return;
    }
    setChatError('');

    appendTurn(state.activeId, { role: 'user', content: question });
    renderChat();
    refs.chatInput.value = '';
    toggleComposer(false, { preserveTaskPanel: pendingCreate || taskState.open });
    setAssistantTyping(true);
    const submitBtn = refs.chatForm.querySelector('button[type="submit"]');
    if (submitBtn) submitBtn.disabled = true;

    try {
      if (pendingSuggested) {
        const handled = await handleSuggestedActionResponse(question);
        if (handled) return;
      }
      if (pendingCreate) {
        await handleCreateConfirmationResponse(question);
        return;
      }
      if (pendingArchive) {
        await handleArchiveConfirmationResponse(question);
        return;
      }

      const intent = await detectIntent(question);
      if (intent === 'skip') {
        setAssistantTyping(false);
        toggleComposer(Boolean(state.activeId));
        if (submitBtn) submitBtn.disabled = false;
        handleAutoIntent(intent, question, { alreadyLogged: true });
        return;
      }
      if (intent === 'archive') {
        setAssistantTyping(false);
        toggleComposer(Boolean(state.activeId));
        if (submitBtn) submitBtn.disabled = false;
        await handleArchiveIntent(question, { alreadyLogged: true });
        return;
      }
      if (intent === 'create_task') {
        setAssistantTyping(false);
        toggleComposer(Boolean(state.activeId));
        if (submitBtn) submitBtn.disabled = false;
        handleCreateTaskIntent();
        return;
      }

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
        popLastTurn(state.activeId);
        renderChat();
        setChatError(data?.error || 'Something went wrong. Try again.');
        refs.chatInput.value = question;
        return;
      }
      appendTurn(state.activeId, { role: 'assistant', content: data.reply || 'No response received.' });
      renderChat();
    } catch (err) {
      popLastTurn(state.activeId);
      renderChat();
      setChatError('Unable to reach the assistant. Check your connection.');
      refs.chatInput.value = question;
    } finally {
      setAssistantTyping(false);
      toggleComposer(true);
      const submitBtn2 = refs.chatForm.querySelector('button[type="submit"]');
      if (submitBtn2) submitBtn2.disabled = false;
      nudgeComposer(DEFAULT_NUDGE, { focus: true });
    }
  }

  async function fetchNextPage(reason) {
    if (!state.hasMore || state.loadingMore) return [];
    state.loadingMore = true;
    updateLoadMoreButtons();
    updateDrawerLists();
    const targetPage = state.nextPage || Math.floor((state.totalLoaded || 0) / state.pageSize) + 1;
    try {
      const resp = await fetch(`/api/threads?page=${targetPage}`, {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        credentials: 'same-origin'
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok || !Array.isArray(data?.threads)) {
        throw new Error(data?.error || 'Unable to load more emails.');
      }
      const incoming = data.threads.map(normalizeThread).filter(Boolean);
      const added = appendThreads(incoming);
      const meta = data.meta || {};
      state.hasMore = Boolean(meta.hasMore);
      state.nextPage = state.hasMore
        ? (typeof meta.nextPage === 'number' ? meta.nextPage : targetPage + 1)
        : 0;
      state.totalLoaded = state.positions.size;
      updateHeaderCount();
      updateProgress();
      updateQueuePill();
      updateDrawerLists();
      updateLoadMoreButtons();
      if (state.activeId && state.lookup.has(state.activeId)) {
        updateEmailCard(state.lookup.get(state.activeId));
      }
      if (!state.activeId && state.needs.length) {
        setActiveThread(state.needs[0]);
      }
      return added;
    } catch (err) {
      console.error('Failed to load more threads', err);
      return [];
    } finally {
      state.loadingMore = false;
      updateDrawerLists();
      updateLoadMoreButtons();
      nudgeComposer('Loaded more - type your next move here.', { focus: false });
    }
  }

  function setActiveThread(threadId) {
    if (!threadId || !state.lookup.has(threadId)) return;
    clearAutoAdvance();
    state.activeId = threadId;
    const thread = state.lookup.get(threadId);
    if (!thread) return;
    hideMoreMenu();
    if (taskState.open && taskState.lastSourceId !== threadId) {
      closeTaskPanel(true);
    }
    if (state.pendingCreateThreadId && state.pendingCreateThreadId !== threadId) {
      clearPendingCreate();
    }
    if (state.pendingArchiveThreadId && state.pendingArchiveThreadId !== threadId) {
      clearPendingArchive();
    }

    refs.emailCard.classList.remove('hidden');
    refs.emailEmpty.classList.add('hidden');
    if (refs.mapToggle) {
      refs.mapToggle.disabled = false;
      refs.mapToggle.removeAttribute('aria-disabled');
    }

    updateEmailCard(thread);
    refreshTaskSuggestion(thread);
    ensurePrimerFetch(threadId);
    ensureHistory(threadId);
    refreshPrepTyping(threadId);
    setChatError('');
    renderChat(threadId);
    updateHint(threadId);
    updateDrawerLists();
    updateQueuePill();
    toggleComposer(true);
    refs.chatInput.value = '';
    nudgeComposer(DEFAULT_NUDGE, { focus: true });
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
        refs.preview.innerHTML = renderPlainText(previewText, { preserveLineBreaks: true });
        refs.preview.classList.add('hidden');
        refs.previewToggle.disabled = false;
        refs.previewToggle.textContent = 'See email body';
      } else {
        const fallbackCopy = 'Email body is unavailable for this message.';
        refs.preview.innerHTML = renderPlainText(fallbackCopy);
        refs.preview.classList.remove('hidden');
        refs.previewToggle.disabled = true;
        refs.previewToggle.textContent = 'Email body unavailable';
      }
    }
  }

  function refreshTaskSuggestion(thread) {
    if (!thread) return;
    const suggestion = buildTaskSuggestion(thread);
    const sameSource = taskState.lastSourceId === thread.threadId;
    taskState.suggested = suggestion;
    taskState.lastSourceId = thread.threadId;
    if (!taskState.open || !sameSource) {
      taskState.values = { ...suggestion };
      taskState.status = 'idle';
      taskState.error = '';
    }
    renderTaskPanel();
  }

  function buildTaskSuggestion(thread) {
    const subject = thread.subject || 'Follow up';
    const sender = thread.from || '';
    const action = thread.nextStep || thread.headline || '';
    const baseTitle = action ? `${action} — ${subject}` : `${subject}${sender ? ` — ${sender}` : ''}`;
    const link = buildMessageLink(thread);
    const summaryLines = [];
    if (sender) summaryLines.push(`From: ${sender}`);
    if (thread.summary) summaryLines.push(`Summary: ${thread.summary}`);
    if (thread.nextStep) summaryLines.push(`Next step: ${thread.nextStep}`);
    if (link) {
      summaryLines.push(`Email: ${link}`);
    } else if (thread.messageId) {
      summaryLines.push(`Message ID: ${thread.messageId}`);
    }
    const due = suggestDueDate(thread);
    return {
      title: truncateText(baseTitle.trim(), 140),
      notes: summaryLines.join('\n'),
      due
    };
  }

  function buildMessageLink(thread) {
    if (thread?.messageId) {
      return `https://mail.google.com/mail/u/0/#all/${encodeURIComponent(thread.messageId)}`;
    }
    if (thread?.link) return thread.link;
    return '';
  }

  function suggestDueDate(thread) {
    const parts = [thread.nextStep, thread.summary, thread.headline, thread.subject].filter(Boolean);
    const combined = parts.join(' ');
    return extractDateFromText(combined);
  }

  function extractDateFromText(text) {
    if (!text) return '';
    const lower = text.toLowerCase();
    const inDays = lower.match(/\bin\s+(\d+)\s+days?\b/);
    if (inDays) {
      const days = Number(inDays[1]);
      if (Number.isFinite(days)) return formatDateInput(addDays(new Date(), days));
    }
    if (lower.includes('end of day') || lower.includes('eod')) {
      return formatDateInput(new Date());
    }
    if (lower.includes('end of week') || lower.includes('by end of week') || lower.includes('this week')) {
      return formatDateInput(nextWeekdayDate(5)); // Friday target
    }
    if (lower.includes('tomorrow')) {
      return formatDateInput(addDays(new Date(), 1));
    }
    if (lower.includes('today')) {
      return formatDateInput(new Date());
    }
    const weekday = detectWeekday(lower);
    if (weekday !== null) {
      return formatDateInput(nextWeekdayDate(weekday));
    }
    const isoMatch = text.match(/\b(\d{4}-\d{2}-\d{2})\b/);
    if (isoMatch) {
      const date = new Date(isoMatch[1]);
      return isValidDate(date) ? formatDateInput(date) : '';
    }
    const slash = text.match(/\b(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\b/);
    if (slash) {
      const month = Number(slash[1]) - 1;
      const day = Number(slash[2]);
      const year = slash[3] ? Number(slash[3].length === 2 ? `20${slash[3]}` : slash[3]) : new Date().getFullYear();
      const parsed = new Date(year, month, day);
      return isValidDate(parsed) ? formatDateInput(parsed) : '';
    }
    return '';
  }

  function detectWeekday(text) {
    const days = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'];
    for (let i = 0; i < days.length; i++) {
      const needle = days[i];
      const pattern = new RegExp(`\\b(?:by|on|this|next)?\\s*${needle}\\b`, 'i');
      if (pattern.test(text)) return i;
    }
    return null;
  }

  function nextWeekdayDate(targetDay) {
    const today = new Date();
    const result = new Date(today);
    const delta = (targetDay - today.getDay() + 7) % 7 || 7;
    result.setDate(today.getDate() + delta);
    return result;
  }

  function addDays(date, days) {
    const copy = new Date(date);
    copy.setDate(copy.getDate() + days);
    return copy;
  }

  function isValidDate(date) {
    return date instanceof Date && !Number.isNaN(date.getTime());
  }

  function formatDateInput(date) {
    if (!isValidDate(date)) return '';
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  function truncateText(text, max) {
    if (!max || text.length <= max) return text;
    return `${text.slice(0, max - 1)}…`;
  }

  function renderTaskPanel() {
    if (!refs.taskPanel || !refs.taskTitle || !refs.taskNotes || !refs.taskDue) return;
    const showPanel = taskState.open;
    refs.taskPanel.classList.toggle('hidden', !showPanel);
    refs.taskPanel.classList.toggle('loading', taskState.status === 'submitting');
    refs.taskTitle.value = taskState.values.title || '';
    refs.taskNotes.value = taskState.values.notes || '';
    refs.taskDue.value = taskState.values.due || '';

    const disabled = taskState.status === 'submitting';
    refs.taskTitle.disabled = disabled;
    refs.taskNotes.disabled = disabled;
    refs.taskDue.disabled = disabled;
    if (refs.taskSubmit) {
      refs.taskSubmit.disabled = disabled;
      refs.taskSubmit.textContent = disabled ? 'Creating…' : 'Create task';
    }
    if (refs.taskCancel) refs.taskCancel.disabled = disabled;
    if (refs.taskReset) refs.taskReset.disabled = disabled;

    if (refs.taskError) {
      refs.taskError.textContent = taskState.error || '';
      refs.taskError.classList.toggle('hidden', !taskState.error);
    }
    if (refs.taskSuccess && refs.taskSuccessMeta) {
      refs.taskSuccess.classList.toggle('hidden', taskState.status !== 'success');
      if (taskState.status !== 'success') {
        refs.taskSuccessMeta.textContent = '';
      }
    }
    if (refs.taskPanelHelper) {
      const pending = isCreateConfirmationPending(taskState.lastSourceId);
      refs.taskPanelHelper.textContent = pending
        ? 'Review and confirm in chat to create.'
        : 'Edit anything before saving.';
    }
  }

  function openTaskPanel(options = {}) {
    const opts = options instanceof Event ? {} : options;
    const preserveValues = Boolean(opts.preserveValues);
    if (!state.activeId || !refs.taskPanel) return;
    const thread = state.lookup.get(state.activeId);
    if (!thread) return;
    refreshTaskSuggestion(thread);
    taskState.open = true;
    taskState.status = 'idle';
    taskState.error = '';
    if (!preserveValues) {
      taskState.values = { ...taskState.suggested };
    }
    renderTaskPanel();
    if (refs.taskTitle) refs.taskTitle.focus();
  }

  function closeTaskPanel(resetValues) {
    if (!refs.taskPanel) return;
    taskState.open = false;
    if (resetValues) {
      taskState.status = 'idle';
      taskState.error = '';
      taskState.values = { ...taskState.suggested };
    }
    clearPendingCreate();
    renderTaskPanel();
  }

  function syncTaskValues() {
    if (!refs.taskTitle || !refs.taskNotes || !refs.taskDue) return;
    taskState.values = {
      title: refs.taskTitle.value,
      notes: refs.taskNotes.value,
      due: refs.taskDue.value
    };
    taskState.error = '';
    if (refs.taskError) refs.taskError.classList.add('hidden');
  }

  function resetTaskToSuggested() {
    taskState.values = { ...taskState.suggested };
    taskState.status = 'idle';
    taskState.error = '';
    renderTaskPanel();
  }

  async function submitTask(arg) {
    if (arg instanceof Event && typeof arg.preventDefault === 'function') {
      arg.preventDefault();
    }
    if (!state.activeId || !refs.taskSubmit) return { ok: false, error: 'No email selected.' };
    const thread = state.lookup.get(state.activeId);
    if (!thread) return { ok: false, error: 'Email context missing.' };
    const title = (taskState.values.title || '').trim();
    if (!title) {
      taskState.error = 'Add a task title before saving.';
      taskState.status = 'error';
      renderTaskPanel();
      return { ok: false, error: taskState.error };
    }
    taskState.error = '';
    taskState.status = 'submitting';
    renderTaskPanel();
    let result = { ok: false, error: '' };
    try {
      const resp = await fetch('/api/tasks', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({
          threadId: thread.threadId,
          messageId: thread.messageId || '',
          title,
          notes: taskState.values.notes || '',
          due: taskState.values.due || ''
        })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        throw new Error(data?.error || 'Unable to create that task.');
      }
      const dueRaw = typeof data?.due === 'string' ? data.due : taskState.values.due;
      const friendlyDue = formatFriendlyDate(dueRaw);
      const finalTitle = typeof data?.title === 'string' && data.title.trim() ? data.title : title;
      const taskUrl = typeof data?.taskUrl === 'string' ? data.taskUrl : '';
      taskState.status = 'success';
      taskState.error = '';
      taskState.open = false;
      if (refs.taskSuccessMeta) {
        const bits = [finalTitle];
        if (friendlyDue) bits.push(`Due ${friendlyDue}`);
        refs.taskSuccessMeta.textContent = bits.join(' • ');
      }
      result = { ok: true, title: finalTitle, due: friendlyDue, url: taskUrl };
    } catch (err) {
      taskState.status = 'error';
      taskState.error = err instanceof Error ? err.message : 'Unable to create that task.';
      result = { ok: false, error: taskState.error };
    } finally {
      renderTaskPanel();
    }
    return result;
  }

  function formatFriendlyDate(raw) {
    if (!raw) return '';
    const parsed = parseDateFriendly(raw);
    if (!parsed) return '';
    return parsed.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
  }

  function parseDateFriendly(raw) {
    if (!raw) return null;
    const str = String(raw).trim();
    if (!str) return null;
    const dateOnly = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (dateOnly) {
      const [, y, m, d] = dateOnly;
      const year = Number(y);
      const month = Number(m) - 1;
      const day = Number(d);
      const date = new Date(year, month, day);
      return isValidDate(date) ? date : null;
    }
    const isoWithTime = str.match(/^(\d{4})-(\d{2})-(\d{2})[T\s]/);
    if (isoWithTime) {
      const [, y, m, d] = isoWithTime;
      const year = Number(y);
      const month = Number(m) - 1;
      const day = Number(d);
      const date = new Date(year, month, day);
      return isValidDate(date) ? date : null;
    }
    const parsed = new Date(str);
    return isValidDate(parsed) ? parsed : null;
  }

  function updateHeaderCount() {
    if (!refs.count) return;
    const loaded = getLoadedCount();
    let label = '0 emails under review';
    if (loaded > 0 && state.hasMore) {
      label = `${loaded}+ emails under review`;
    } else if (loaded > 0) {
      label = `${loaded} email${loaded === 1 ? '' : 's'} under review`;
    }
    refs.count.textContent = label;
    updateLoadMoreButtons();
  }

  function updateQueuePill() {
    if (!refs.queuePill) return;
    const remaining = state.needs.length;
    const suffix = state.hasMore ? '+' : '';
    if (remaining) {
      refs.queuePill.textContent = `${remaining}${suffix} remaining`;
      return;
    }
    refs.queuePill.textContent = state.hasMore ? 'Load more to continue' : 'All done';
  }

  function updateLoadMoreButtons() {
    if (refs.loadMoreHead) {
      refs.loadMoreHead.classList.remove('hidden');
      if (state.hasMore) {
        refs.loadMoreHead.disabled = state.loadingMore;
        refs.loadMoreHead.textContent = state.loadingMore ? 'Loading…' : 'Load more';
      } else {
        refs.loadMoreHead.disabled = true;
        refs.loadMoreHead.textContent = 'All emails loaded';
      }
    }
    if (refs.loadMoreEmpty) {
      refs.loadMoreEmpty.classList.remove('hidden');
      if (state.hasMore) {
        refs.loadMoreEmpty.disabled = state.loadingMore;
        refs.loadMoreEmpty.textContent = state.loadingMore ? 'Loading…' : 'Load more';
      } else {
        refs.loadMoreEmpty.disabled = true;
        refs.loadMoreEmpty.textContent = 'All emails loaded';
      }
    }
  }

  function toggleMoreMenu(event) {
    if (event) event.preventDefault();
    if (!refs.moreMenu || !refs.moreBtn || !state.activeId) return;
    const isOpen = !refs.moreMenu.classList.contains('hidden');
    if (isOpen) {
      hideMoreMenu();
      return;
    }
    refs.moreMenu.classList.remove('hidden');
    refs.moreBtn.setAttribute('aria-expanded', 'true');
  }

  function hideMoreMenu() {
    if (!refs.moreMenu || !refs.moreBtn) return;
    refs.moreMenu.classList.add('hidden');
    refs.moreBtn.setAttribute('aria-expanded', 'false');
  }

  function handleMoreMenuClick(event) {
    const target = event.target;
    if (!(target instanceof Element)) return;
    const action = target.closest('button')?.dataset?.action;
    if (!action) return;
    event.preventDefault();
    hideMoreMenu();
    if (action === 'task') {
      clearPendingCreate();
      openTaskPanel();
    }
    if (action === 'skip') {
      skipCurrent('menu');
    }
  }

  function getLoadedCount() {
    return state.totalLoaded || state.positions.size || 0;
  }

  function getReviewedCount() {
    return Math.min(reviewedIds.size, getLoadedCount());
  }

  function markReviewed(threadId) {
    if (threadId) reviewedIds.add(threadId);
  }

  function updateProgress() {
    if (!refs.progress) return;
    const loaded = getLoadedCount();
    const done = getReviewedCount();
    const totalForCalc = loaded || 1;
    const pct = Math.min(100, Math.round((done / totalForCalc) * 100));
    refs.progress.style.width = `${pct}%`;
    if (refs.progressTrack) {
      refs.progressTrack.setAttribute('aria-valuenow', String(pct));
      const labelTotal = state.hasMore ? `${loaded}+` : `${loaded}`;
      refs.progressTrack.setAttribute('aria-valuetext', `${done} reviewed out of ${labelTotal}`);
    }
  }

  function updateDrawerLists() {
    renderThreadList(refs.needsList, 'drawer');
    renderThreadList(refs.reviewList, 'queue');
    if (refs.needsCount) refs.needsCount.textContent = state.hasMore ? `${state.needs.length}+` : String(state.needs.length);
  }

  function renderThreadList(listEl, variant = 'drawer') {
    if (!listEl) return;
    listEl.innerHTML = '';

    if (state.needs.length) {
      state.needs.forEach(id => appendThreadItem(listEl, id, variant));
    } else {
      const li = document.createElement('li');
      li.className = variant === 'queue' ? 'queue-empty drawer-empty' : 'drawer-empty';
      li.textContent = 'Nothing queued up.';
      listEl.appendChild(li);
    }

    if (state.hasMore) {
      const li = document.createElement('li');
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = buildThreadClass(variant, { loadMore: true });
      btn.textContent = state.loadingMore ? 'Loading…' : 'Load more';
      btn.disabled = state.loadingMore;
      li.appendChild(btn);
      listEl.appendChild(li);
    }
  }

  function appendThreadItem(listEl, threadId, variant) {
    const thread = state.lookup.get(threadId);
    if (!thread) return;
    const li = document.createElement('li');
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = buildThreadClass(variant, { active: threadId === state.activeId });
    btn.dataset.threadId = threadId;
    btn.innerHTML = `<strong>${htmlEscape(thread.from || 'Unknown sender')}</strong><span>${htmlEscape(thread.subject || '(no subject)')}</span>`;
    li.appendChild(btn);
    listEl.appendChild(li);
  }

  function buildThreadClass(variant, options = {}) {
    const { loadMore = false, active = false } = options;
    const base = variant === 'queue' ? 'queue-item' : 'drawer-thread';
    const classes = [base];
    if (loadMore) classes.push('load-more');
    if (active) classes.push('active');
    return classes.join(' ');
  }

  function appendThreads(items) {
    const added = [];
    items.forEach(thread => {
      if (!thread || !thread.threadId) return;
      if (state.lookup.has(thread.threadId)) return;
      state.lookup.set(thread.threadId, thread);
      state.positions.set(thread.threadId, state.positions.size);
      state.needs.push(thread.threadId);
      primerStatus.set(thread.threadId, thread.primer ? 'ready' : 'idle');
      added.push(thread.threadId);
    });
    if (added.length) {
      state.totalLoaded = state.positions.size;
    }
    return added;
  }

  function normalizeThread(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const threadId = typeof raw.threadId === 'string' ? raw.threadId.trim() : '';
    if (!threadId) return null;
    return {
      threadId,
      messageId: typeof raw.messageId === 'string' ? raw.messageId.trim() : '',
      headline: typeof raw.headline === 'string' ? raw.headline.trim() : '',
      from: typeof raw.from === 'string' ? raw.from.trim() : '',
      subject: typeof raw.subject === 'string' ? raw.subject.trim() : '(no subject)',
      summary: typeof raw.summary === 'string' ? raw.summary.trim() : '',
      nextStep: typeof raw.nextStep === 'string' ? raw.nextStep.trim() : '',
      link: typeof raw.link === 'string' ? raw.link : '',
      primer: typeof raw.primer === 'string' ? raw.primer.trim() : '',
      suggestedAction: normalizeSuggestedAction(raw.suggestedAction) || guessSuggestedAction(raw),
      category: typeof raw.category === 'string' ? raw.category : '',
      receivedAt: typeof raw.receivedAt === 'string' ? raw.receivedAt : '',
      convo: typeof raw.convo === 'string' ? raw.convo : ''
    };
  }

  function normalizeSuggestedAction(value) {
    const val = typeof value === 'string' ? value.trim().toLowerCase() : '';
    if (val === 'archive' || val === 'more_info' || val === 'create_task' || val === 'skip') return val;
    return '';
  }

  function guessSuggestedAction(thread) {
    const next = normalizeSuggestedAction(actionFromNextStep(thread?.nextStep));
    if (next) return next;
    const summary = `${thread?.summary || thread?.headline || thread?.subject || ''}`.toLowerCase();
    if (summary.includes('deadline') || summary.includes('follow up') || summary.includes('follow-up') || summary.includes('due')) {
      return 'create_task';
    }
    if (summary.includes('fyi') || summary.includes('newsletter')) return 'skip';
    return 'more_info';
  }

  function actionFromNextStep(nextStep) {
    const text = typeof nextStep === 'string' ? nextStep.toLowerCase() : '';
    if (!text) return '';
    if (text.includes('archive')) return 'archive';
    if (text.includes('remind') || text.includes('task') || text.includes('follow up') || text.includes('follow-up')) {
      return 'create_task';
    }
    return '';
  }

  function updateHint(threadId) {
    if (!refs.chatHint) return;
    if (!MAX_TURNS || !threadId) {
      refs.chatHint.textContent = 'Ask anything or tap Archive / More actions.';
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
      insertThreadDivider(threadId);
      const thread = state.lookup.get(threadId);
      const intro = buildIntroMessage(thread);
      if (intro) {
        const turn = { role: 'assistant', content: intro };
        history.push(turn);
        state.timeline.push({ type: 'turn', threadId, turn });
      }
      if (thread) {
        const action = normalizeSuggestedAction(thread.suggestedAction) || guessSuggestedAction(thread);
        if (action) setPendingSuggestedAction(threadId, action);
      }
    }
    return history;
  }

  function appendTurn(threadId, turn) {
    const history = ensureHistory(threadId);
    history.push(turn);
    state.timeline.push({ type: 'turn', threadId, turn });
    return history;
  }

  function popLastTurn(threadId) {
    const history = state.histories.get(threadId);
    if (!history?.length) return null;
    const removed = history.pop();
    if (!removed) return null;
    for (let i = state.timeline.length - 1; i >= 0; i--) {
      const item = state.timeline[i];
      if (item.type === 'turn' && item.threadId === threadId && item.turn === removed) {
        state.timeline.splice(i, 1);
        break;
      }
    }
    return removed;
  }

  function insertThreadDivider(threadId) {
    const thread = state.lookup.get(threadId);
    if (!thread) return;
    const alreadyInserted = state.timeline.some(item => item.type === 'divider' && item.threadId === threadId);
    if (alreadyInserted) return;
    const sender = thread.from ? thread.from.split('<')[0].trim() || thread.from : '';
    const subject = (thread.subject || '').trim() || '(no subject)';
    const labelParts = [];
    if (sender) labelParts.push(sender);
    labelParts.push(subject);
    const label = labelParts.join(' — ');
    state.timeline.push({
      type: 'divider',
      threadId,
      label,
      subject,
      sender,
      receivedAt: thread.receivedAt || '',
      link: thread.link || ''
    });
  }

  function buildIntroMessage(thread) {
    if (!thread) return 'Need a quick summary or a draft? I can help.';
    const primer = (thread.primer || '').trim();
    if (primer) return primer;

    const status = getPrimerStatus(thread.threadId);
    if (status === 'loading' || status === 'pending') return '';
    if (status === 'error') {
      return buildFallbackPrimer(thread);
    }
    return buildFallbackPrimer(thread);
  }

  function buildFallbackPrimer(thread) {
    if (!thread) return 'Heads up: you have an email waiting.';
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
    const action = normalizeSuggestedAction(thread.suggestedAction) || guessSuggestedAction(thread);
    thread.suggestedAction = action;
    const next = formatActionNudge(action, normalizedNext);
    return `${starter} ${next}`;
  }

  function formatActionNudge(action, nextStep) {
    if (action === 'archive') return 'Looks wrapped—want me to archive it?';
    if (action === 'skip') return 'We can park this and move to the next email if you want.';
    if (action === 'create_task') {
      const detail = nextStep ? ` for "${nextStep}"` : '';
      return `I can log a quick task${detail} so it doesn’t slip. Want that?`;
    }
    return 'Want me to pull a tighter rundown or key deadlines?';
  }


  function ensurePrimerFetch(threadId) {
    const thread = state.lookup.get(threadId);
    if (!thread || (thread.primer || '').trim()) return;
    const status = getPrimerStatus(threadId);
    if (status === 'loading' || status === 'pending') return;
    setPrimerStatus(threadId, 'loading');
    state.prepTyping = threadId;
    fetchPrimer(threadId, 0);
  }

  function fetchPrimer(threadId, attempt) {
    const nextAttempt = typeof attempt === 'number' ? attempt : 0;
    fetch(`${PRIMER_ENDPOINT}/${encodeURIComponent(threadId)}`, {
      headers: { 'Accept': 'application/json' }
    }).then(async resp => {
      if (resp.status === 202) {
        if (nextAttempt >= MAX_PRIMER_POLLS) {
          setPrimerStatus(threadId, 'error');
          return;
        }
        setPrimerStatus(threadId, 'pending');
        schedulePrimerRetry(threadId, nextAttempt + 1);
        return;
      }
      if (!resp.ok) throw new Error('Unable to load primer');
      const data = await resp.json().catch(() => ({}));
      const primer = typeof data?.primer === 'string' ? data.primer.trim() : '';
      const suggestedAction = normalizeSuggestedAction(data?.suggestedAction);
      if (!primer) {
        if (nextAttempt >= MAX_PRIMER_POLLS) {
          setPrimerStatus(threadId, 'error');
          return;
        }
        setPrimerStatus(threadId, 'pending');
        schedulePrimerRetry(threadId, nextAttempt + 1);
        return;
      }
      applyPrimerToThread(threadId, primer, suggestedAction);
    }).catch(() => {
      if (nextAttempt >= MAX_PRIMER_POLLS) {
        setPrimerStatus(threadId, 'error');
        return;
      }
      setPrimerStatus(threadId, 'pending');
      schedulePrimerRetry(threadId, nextAttempt + 1);
    });
  }

  function schedulePrimerRetry(threadId, attempt) {
    clearPrimerRetry(threadId);
    const timer = window.setTimeout(() => fetchPrimer(threadId, attempt), PRIMER_RETRY_DELAY);
    primerRetryTimers.set(threadId, timer);
  }

  function clearPrimerRetry(threadId) {
    if (!primerRetryTimers.has(threadId)) return;
    window.clearTimeout(primerRetryTimers.get(threadId));
    primerRetryTimers.delete(threadId);
  }

  function getPrimerStatus(threadId) {
    return primerStatus.get(threadId) || 'idle';
  }

  function setPrimerStatus(threadId, status) {
    primerStatus.set(threadId, status);
    if (status === 'loading' || status === 'pending') {
      state.prepTyping = threadId;
    } else if (state.prepTyping === threadId) {
      state.prepTyping = '';
    }
    const history = state.histories.get(threadId);
    if (!history || !history.length) {
      const thread = state.lookup.get(threadId);
      if (thread && status !== 'loading' && status !== 'pending') {
        const intro = buildIntroMessage(thread);
        if (intro) {
          insertThreadDivider(threadId);
          const turn = { role: 'assistant', content: intro };
          state.histories.set(threadId, [turn]);
          state.timeline.push({ type: 'turn', threadId, turn });
          const action = normalizeSuggestedAction(thread.suggestedAction) || guessSuggestedAction(thread);
          if (action) setPendingSuggestedAction(threadId, action);
        }
      }
      if (threadId === state.activeId) renderChat(threadId);
      return;
    }
    const thread = state.lookup.get(threadId);
    if (!thread || (thread.primer || '').trim()) return;
    if (history[0]?.role === 'assistant') {
      history[0].content = buildIntroMessage(thread);
      if (threadId === state.activeId) {
        renderChat(threadId);
      }
    }
  }

  function refreshPrepTyping(threadId) {
    const status = getPrimerStatus(threadId);
    if (status === 'loading' || status === 'pending') {
      state.prepTyping = threadId;
    } else if (state.prepTyping === threadId) {
      state.prepTyping = '';
    }
  }

  function applyPrimerToThread(threadId, primer, suggestedAction) {
    clearPrimerRetry(threadId);
    const thread = state.lookup.get(threadId);
    if (!thread) return;
    const action = normalizeSuggestedAction(suggestedAction) || thread.suggestedAction || guessSuggestedAction(thread);
    thread.suggestedAction = action;
    setPendingSuggestedAction(threadId, action);
    thread.primer = primer;
    const history = state.histories.get(threadId);
    if (history && history.length && history[0]?.role === 'assistant') {
      history[0].content = primer;
    } else {
      insertThreadDivider(threadId);
      const turn = { role: 'assistant', content: primer };
      state.histories.set(threadId, [turn]);
      state.timeline.push({ type: 'turn', threadId, turn });
    }
    setPrimerStatus(threadId, 'ready');
    if (state.prepTyping === threadId) {
      state.prepTyping = '';
    }
    if (threadId === state.activeId) {
      renderChat(threadId);
    }
  }

  function renderChat(threadId = state.activeId) {
    if (!refs.chatLog) return;
    const timeline = state.timeline;
    if (!timeline.length) {
      const showTyping = (state.typing && threadId === state.activeId) || state.prepTyping === threadId;
      let view = showTyping ? '' : chatPlaceholder();
      if (showTyping) {
        view += typingIndicatorHtml();
      }
      refs.chatLog.innerHTML = view;
      return;
    }
    let markup = timeline.map(entry => {
      if (entry.type === 'divider') {
        const label = htmlEscape(entry.label || 'New email thread');
        const sender = htmlEscape(entry.sender || '');
        const subject = htmlEscape(entry.subject || '');
        const timestamp = entry.receivedAt ? formatTimestamp(entry.receivedAt) : '';
        const meta = htmlEscape([sender, timestamp].filter(Boolean).join(' • '));
        const initials = initialsFromSender(entry.sender || entry.label || '');
        const link = entry.link ? escapeAttribute(entry.link) : '';
        const linkHtml = link
          ? `<a class="chat-divider-link" href="${link}" target="_blank" rel="noopener noreferrer">Open in Gmail ↗</a>`
          : '';
        return `
          <div class="chat-divider">
            <span class="chat-divider-line" aria-hidden="true"></span>
            <div class="chat-divider-card">
              <div class="chat-divider-avatar" aria-hidden="true">${initials}</div>
              <div class="chat-divider-content">
                <p class="chat-divider-meta">${meta}</p>
                <p class="chat-divider-subject">${subject || label}</p>
                ${linkHtml}
              </div>
            </div>
            <span class="chat-divider-line" aria-hidden="true"></span>
          </div>
        `;
      }
      const turn = entry.turn;
      if (turn.role === 'assistant') {
        return `<div class="chat-message assistant"><div class="chat-card">${renderAssistantMarkdown(turn.content)}</div></div>`;
      }
      return `<div class="chat-message user"><div class="chat-card">${renderPlainText(turn.content, { preserveLineBreaks: true })}</div></div>`;
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

  function clearComposerNudge() {
    if (composerNudgeTimer) {
      window.clearTimeout(composerNudgeTimer);
      composerNudgeTimer = 0;
    }
    if (!refs.chatForm) return;
    refs.chatForm.classList.remove('nudged');
    delete refs.chatForm.dataset.nudge;
  }

  function nudgeComposer(message, options = {}) {
    if (!refs.chatForm || !refs.chatInput || refs.chatInput.disabled) return;
    const { focus = true } = options;
    if (focus) refs.chatInput.focus();
    const label = message && message.trim() ? message.trim() : DEFAULT_NUDGE;
    refs.chatForm.dataset.nudge = label;
    refs.chatForm.classList.add('nudged');
    if (composerNudgeTimer) window.clearTimeout(composerNudgeTimer);
    composerNudgeTimer = window.setTimeout(() => {
      clearComposerNudge();
    }, 2600);
  }

  function toggleComposer(enabled, options = {}) {
    const preserveTaskPanel = Boolean(options.preserveTaskPanel);
    refs.chatInput.disabled = !enabled || !state.activeId;
    if (refs.reviewBtn) refs.reviewBtn.disabled = !enabled || !state.activeId;
    if (refs.archiveBtn) refs.archiveBtn.disabled = !enabled || !state.activeId;
    if (refs.moreBtn) refs.moreBtn.disabled = !enabled || !state.activeId;
    if (!enabled) {
      clearComposerNudge();
      hideMoreMenu();
      if (!preserveTaskPanel) {
        closeTaskPanel(true);
      }
    }
  }

  function setPendingSuggestedAction(threadId, action) {
    const normalized = normalizeSuggestedAction(action);
    if (!threadId || !normalized) return;
    state.pendingSuggestedActions.set(threadId, normalized);
  }

  function clearPendingSuggestedAction(threadId = state.activeId) {
    if (!threadId) return;
    state.pendingSuggestedActions.delete(threadId);
  }

  function getPendingSuggestedAction(threadId = state.activeId) {
    if (!threadId) return '';
    return state.pendingSuggestedActions.get(threadId) || '';
  }

  function setPendingCreate(threadId) {
    state.pendingCreateThreadId = threadId || '';
    renderTaskPanel();
  }

  function clearPendingCreate() {
    if (!state.pendingCreateThreadId) return;
    state.pendingCreateThreadId = '';
    renderTaskPanel();
  }

  function isCreateConfirmationPending(threadId = state.activeId) {
    return Boolean(threadId && state.pendingCreateThreadId && state.pendingCreateThreadId === threadId);
  }

  function setPendingArchive(threadId) {
    state.pendingArchiveThreadId = threadId || '';
  }

  function clearPendingArchive() {
    if (!state.pendingArchiveThreadId) return;
    state.pendingArchiveThreadId = '';
  }

  function isArchiveConfirmationPending(threadId = state.activeId) {
    return Boolean(threadId && state.pendingArchiveThreadId && state.pendingArchiveThreadId === threadId);
  }

  function setAssistantTyping(value) {
    state.typing = Boolean(value);
    renderChat(state.activeId);
  }

  function withButtonBusy(btn, label) {
    const original = btn.textContent;
    btn.disabled = true;
    if (label) btn.textContent = label;
    return () => {
      btn.disabled = false;
      if (label) btn.textContent = original;
    };
  }

  async function requestReview() {
    if (!state.activeId || !refs.reviewBtn) return;
    const threadId = state.activeId;
    const history = ensureHistory(threadId);
    const asked = history.filter(turn => turn.role === 'user').length;
    if (MAX_TURNS > 0 && asked >= MAX_TURNS) {
      setChatError('Chat limit reached for this email.');
      return;
    }
    setChatError('');
    appendTurn(threadId, { role: 'user', content: REVIEW_PROMPT });
    renderChat();
    toggleComposer(false);
    setAssistantTyping(true);
    const restore = withButtonBusy(refs.reviewBtn, 'Getting details…');

    try {
      const resp = await fetch('/secretary/review', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({ threadId })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        const last = history[history.length - 1];
        if (last?.role === 'user' && last.content === REVIEW_PROMPT) {
          popLastTurn(threadId);
          renderChat();
        }
        throw new Error(data?.error || 'Unable to review this email.');
      }
      const reply = typeof data?.review === 'string' ? data.review.trim() : '';
      appendTurn(threadId, { role: 'assistant', content: reply || 'Here’s what I could pull together.' });
      renderChat();
      updateHint(threadId);
    } catch (err) {
      console.error('Review request failed', err);
      const message = err instanceof Error ? err.message : 'Unable to review this email.';
      setChatError(message);
    } finally {
      restore();
      setAssistantTyping(false);
      toggleComposer(Boolean(state.activeId));
      nudgeComposer(DEFAULT_NUDGE, { focus: true });
    }
  }

  async function archiveCurrent(source) {
    if (!state.activeId || !refs.archiveBtn) return;
    const threadId = state.activeId;
    const restoreBtn = withButtonBusy(refs.archiveBtn, 'Archiving…');
    toggleComposer(false);
    setChatError('');
    try {
      const resp = await fetch('/api/archive', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({ threadId })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        throw new Error(data?.error || 'Unable to archive this email.');
      }
      removeCurrentFromQueue();
    } catch (err) {
      console.error('Failed to archive thread', err);
      const message = err instanceof Error ? err.message : 'Unable to archive this email.';
      setChatError(message);
    } finally {
      restoreBtn();
      toggleComposer(Boolean(state.activeId));
    }
  }

  function removeCurrentFromQueue() {
    if (!state.activeId || !state.needs.length) return;
    const threadId = state.activeId;
    const index = state.needs.indexOf(threadId);
    if (index === -1) return;
    state.needs.splice(index, 1);
    markReviewed(threadId);
    updateProgress();
    updateDrawerLists();
    updateHeaderCount();
    updateQueuePill();
    closeTaskPanel(true);
    hideMoreMenu();
    clearPendingSuggestedAction(threadId);

    if (!state.needs.length) {
      const message = state.hasMore
        ? 'You reviewed everything loaded. Tap Load more to keep going.'
        : 'All emails reviewed. Nice work.';
      setEmptyState(message);
      toggleComposer(false);
      return;
    }
    const nextId = state.needs[index] || state.needs[0];
    setActiveThread(nextId);
  }

  function skipCurrent(source) {
    if (!state.activeId) return;
    const threadId = state.activeId;
    const index = state.needs.indexOf(threadId);
    const hasRoomToAdvance = state.needs.length > 1;
    markReviewed(threadId);
    updateProgress();
    updateDrawerLists();
    updateHeaderCount();
    updateQueuePill();
    closeTaskPanel(true);
    hideMoreMenu();
    clearPendingSuggestedAction(threadId);
    if (!hasRoomToAdvance) return;
    const nextIndex = index === -1 ? 0 : (index + 1) % state.needs.length;
    const nextId = state.needs[nextIndex] || state.needs[0];
    setActiveThread(nextId);
  }

  function handleAutoIntent(intent, userText, options = {}) {
    if (!state.activeId || intent !== 'skip') return;
    const alreadyLogged = Boolean(options.alreadyLogged);
    const history = ensureHistory(state.activeId);
    if (!alreadyLogged) {
      appendTurn(state.activeId, { role: 'user', content: userText });
      renderChat();
    }

    appendTurn(state.activeId, { role: 'assistant', content: 'Skipping for now. It stays in Needs Review.' });
    renderChat();
    updateHint(state.activeId);

    clearAutoAdvance();
    const targetId = state.activeId;
    state.autoAdvanceTimer = window.setTimeout(() => {
      if (state.activeId !== targetId) return;
      skipCurrent('auto');
      clearAutoAdvance();
    }, 600);
  }

  function clearAutoAdvance() {
    if (state.autoAdvanceTimer) {
      window.clearTimeout(state.autoAdvanceTimer);
      state.autoAdvanceTimer = 0;
    }
  }

  async function handleArchiveIntent(userText, options = {}) {
    if (!state.activeId) return;
    const alreadyLogged = Boolean(options.alreadyLogged);
    const history = ensureHistory(state.activeId);
    if (!alreadyLogged) {
      appendTurn(state.activeId, { role: 'user', content: userText });
      renderChat();
    }

    appendTurn(state.activeId, { role: 'assistant', content: 'Archiving this email in Gmail…' });
    renderChat();
    updateHint(state.activeId);

    setAssistantTyping(true);
    try {
      await archiveCurrent('auto-intent');
    } finally {
      setAssistantTyping(false);
    }
  }

  function handleCreateTaskIntent() {
    if (!state.activeId) return;
    const thread = state.lookup.get(state.activeId);
    if (!thread) return;
    const alreadyPending = isCreateConfirmationPending(state.activeId);
    openTaskPanel({ preserveValues: alreadyPending && taskState.open });
    setPendingCreate(state.activeId);
    taskState.status = 'idle';
    taskState.error = '';
    renderTaskPanel();

    const prompt = alreadyPending ? 'Please confirm: create the task as shown? (yes/no)' : buildTaskConfirmationPrompt();
    appendTurn(state.activeId, { role: 'assistant', content: prompt });
    renderChat();
    updateHint(state.activeId);
  }

  async function handleMoreInfoIntent() {
    await requestReview();
  }

  async function handleSuggestedActionResponse(userText) {
    if (!state.activeId) return false;
    const action = getPendingSuggestedAction(state.activeId);
    if (!action) return false;
    const normalized = (userText || '').trim().toLowerCase();

    if (isAffirmativeResponse(normalized)) {
      clearPendingSuggestedAction(state.activeId);
      if (action === 'archive') {
        await handleArchiveIntent(userText, { alreadyLogged: true });
        return true;
      }
      if (action === 'skip') {
        handleAutoIntent('skip', userText, { alreadyLogged: true });
        return true;
      }
      if (action === 'create_task') {
        handleCreateTaskIntent();
        return true;
      }
      if (action === 'more_info') {
        await handleMoreInfoIntent();
        return true;
      }
    }

    if (isNegativeResponse(normalized)) {
      clearPendingSuggestedAction(state.activeId);
      appendTurn(state.activeId, { role: 'assistant', content: 'Okay, I’ll hold off. Tell me what you’d like me to do instead.' });
      renderChat();
      return true;
    }

    const intent = await detectIntent(userText);
    if (intent === 'archive') {
      clearPendingSuggestedAction(state.activeId);
      await handleArchiveIntent(userText, { alreadyLogged: true });
      return true;
    }
    if (intent === 'skip') {
      clearPendingSuggestedAction(state.activeId);
      handleAutoIntent('skip', userText, { alreadyLogged: true });
      return true;
    }
    if (intent === 'create_task') {
      clearPendingSuggestedAction(state.activeId);
      handleCreateTaskIntent();
      return true;
    }

    // User asked something else — clear pending and let normal flow handle it.
    clearPendingSuggestedAction(state.activeId);
    return false;
  }


  async function handleCreateConfirmationResponse(userText) {
    if (!state.activeId || !isCreateConfirmationPending(state.activeId)) return;
    const normalized = (userText || '').trim().toLowerCase();

    if (isAffirmativeResponse(normalized)) {
      const result = await submitTask();
      if (result?.ok) {
        appendTurn(state.activeId, { role: 'assistant', content: buildTaskCreatedMessage(result) });
        clearPendingCreate();
        promptArchiveAfterTask(result, { includeSuccess: false });
      } else if (result?.error) {
        appendTurn(state.activeId, { role: 'assistant', content: `Couldn't create the task: ${result.error}` });
      }
      renderChat();
      return;
    }

    if (isNegativeResponse(normalized)) {
      clearPendingCreate();
      renderTaskPanel();
      appendTurn(state.activeId, { role: 'assistant', content: 'Okay, I won’t create it. Adjust the fields or ask another action.' });
      renderChat();
      return;
    }

    const intent = await detectIntent(userText);
    if (intent === 'archive') {
      clearPendingCreate();
      closeTaskPanel(true);
      await handleArchiveIntent(userText, { alreadyLogged: true });
      return;
    }
    if (intent === 'skip') {
      clearPendingCreate();
      closeTaskPanel(true);
      handleAutoIntent('skip', userText, { alreadyLogged: true });
      return;
    }

    appendTurn(state.activeId, { role: 'assistant', content: 'Please confirm: create the task as shown? (yes/no)' });
    renderChat();
  }

  function buildTaskConfirmationPrompt() {
    const title = (taskState.values.title || taskState.suggested.title || 'New task').trim();
    const friendlyDue = taskState.values.due ? formatFriendlyDate(taskState.values.due) : '';
    const dueLabel = friendlyDue ? `due ${friendlyDue}` : 'with no due date';
    return `I can create a task: ${title} (${dueLabel}). Create it?`;
  }

  function buildTaskCreatedMessage(result) {
    const bits = ['✅ Task created'];
    if (result?.title) bits.push(result.title);
    const due = result?.due || formatFriendlyDate(taskState.values.due);
    if (due) {
      bits.push(`Due ${due}`);
    } else {
      bits.push('No due date set');
    }
    if (result?.url) {
      bits.push(`[Open in Google Tasks](${result.url})`);
    }
    return bits.join(' — ');
  }

  function isAffirmativeResponse(text) {
    const normalized = (text || '').trim().toLowerCase();
    if (!normalized) return false;
    const affirm = ['yes', 'y', 'yeah', 'yep', 'yup', 'sure', 'sure thing', 'do it', 'create it', 'confirm', 'please do', 'go for it', 'go ahead', 'sounds good', 'ok', 'okay', 'affirmative', 'absolutely'];
    return affirm.some(word => normalized === word || normalized.startsWith(`${word},`) || normalized.startsWith(`${word} `));
  }

  function isNegativeResponse(text) {
    const normalized = (text || '').trim().toLowerCase();
    if (!normalized) return false;
    const negative = ['no', 'n', 'nah', 'nope', 'not now', 'cancel', 'stop', 'hold on', 'wait', 'don’t', "don't", 'do not', 'no thanks', 'no thank you'];
    return negative.some(word => normalized === word || normalized.startsWith(`${word},`) || normalized.startsWith(`${word} `));
  }

  function promptArchiveAfterTask(result, options = {}) {
    if (!state.activeId) return;
    const threadId = state.activeId;
    const includeSuccess = Boolean(options.includeSuccess);
    setPendingArchive(threadId);
    if (includeSuccess) {
      appendTurn(threadId, { role: 'assistant', content: buildTaskCreatedMessage(result) });
    }
    const prompt = 'Archive this email and move on? I can keep it here if you want.';
    appendTurn(threadId, { role: 'assistant', content: prompt });
    renderChat();
    updateHint(threadId);
  }

  async function handleArchiveConfirmationResponse(userText) {
    if (!state.activeId || !isArchiveConfirmationPending(state.activeId)) return;
    const normalized = (userText || '').trim().toLowerCase();

    const intent = await detectIntent(userText);
    if (intent === 'archive') {
      clearPendingArchive();
      await handleArchiveIntent(userText, { alreadyLogged: true });
      return;
    }
    if (intent === 'skip') {
      clearPendingArchive();
      handleAutoIntent('skip', userText, { alreadyLogged: true });
      return;
    }

    if (isAffirmativeResponse(normalized)) {
      clearPendingArchive();
      await handleArchiveIntent(userText, { alreadyLogged: true });
      return;
    }

    if (isNegativeResponse(normalized)) {
      clearPendingArchive();
      appendTurn(state.activeId, { role: 'assistant', content: 'Okay, leaving it in Needs Review. Ask to archive anytime.' });
      renderChat();
      return;
    }

    appendTurn(state.activeId, { role: 'assistant', content: 'Want me to archive this email or keep it here?' });
    renderChat();
  }

  function setEmptyState(message) {
    const fallback = state.hasMore
      ? 'You reviewed everything loaded. Tap Load more to keep going.'
      : 'All emails reviewed. Nice work.';
    const copy = message || fallback;
    closeTaskPanel(true);
    hideMoreMenu();
    state.activeId = '';
    refs.emailCard.classList.add('hidden');
    refs.emailEmpty.classList.remove('hidden');
    if (refs.position) {
      refs.position.textContent = '';
      refs.position.classList.add('hidden');
    }
    if (refs.emailEmptyText) refs.emailEmptyText.textContent = copy;
    if (refs.loadMoreEmpty) {
      refs.loadMoreEmpty.classList.toggle('hidden', !state.hasMore);
    }
    if (refs.mapToggle) {
      refs.mapToggle.setAttribute('aria-disabled', 'true');
      refs.mapToggle.disabled = true;
    }
    renderChat();
    updateQueuePill();
    updateLoadMoreButtons();
  }

  async function detectIntent(text) {
    const cleaned = text.replace(/[.!?]/g, '').trim().toLowerCase();
    if (!cleaned) return '';
    const archivePhrases = ['archive', 'archive this', 'archive it', 'archive email', 'archive message', 'archive thread'];
    if (archivePhrases.includes(cleaned)) return 'archive';
    const skipPhrases = ['skip', 'skip it', 'skip this', 'skip this one', 'skip this email', 'skip this thread'];
    if (skipPhrases.includes(cleaned)) return 'skip';
    const taskPhrases = [
      'create task',
      'create a task',
      'make a task',
      'make this a task',
      'add a task',
      'set a reminder',
      'reminder',
      'add reminder',
      'remind me'
    ];
    if (taskPhrases.includes(cleaned) || cleaned.includes('reminder')) return 'create_task';
    const intent = await evaluateIntent(text);
    return intent === 'archive' || intent === 'skip' || intent === 'create_task' ? intent : '';
  }

  async function evaluateIntent(rawText) {
    try {
      const resp = await fetch('/secretary/intent', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({ text: rawText })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) return '';
      return typeof data?.intent === 'string' ? data.intent : '';
    } catch (err) {
      console.warn('Intent check failed', err);
      return '';
    }
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
    const total = getLoadedCount();
    if (!total) return '';
    const position = state.positions.get(threadId);
    if (typeof position !== 'number') return '';
    const current = position + 1;
    const totalLabel = state.hasMore ? `${total}+` : `${total}`;
    return `${current} of ${totalLabel}`;
  }

  function renderAssistantMarkdown(value) {
    const text = value == null ? '' : String(value);
    if (!text) return '';
    if (!markedLib) {
      return renderPlainText(text, { preserveLineBreaks: true });
    }
    const prepared = linkifyMarkdownSource(text);
    const html = markedLib.parse(prepared);
    return sanitizeHtml(html);
  }

  function linkifyMarkdownSource(text) {
    if (!linkify) return text;
    const matches = linkify.match(text);
    if (!matches || !matches.length) return text;
    let result = '';
    let cursor = 0;
    for (const match of matches) {
      const start = match.index ?? 0;
      const end = match.lastIndex ?? start;
      if (start < cursor) continue;
      if (shouldSkipMarkdownAutolink(text, start)) {
        result += text.slice(cursor, end);
        cursor = end;
        continue;
      }
      result += text.slice(cursor, start);
      const target = match.url || match.raw || match.text || '';
      result += `<${target}>`;
      cursor = end;
    }
    result += text.slice(cursor);
    return result;
  }

  function shouldSkipMarkdownAutolink(text, index) {
    if (index <= 0) return false;
    const prevChar = text[index - 1];
    if (prevChar === '<') return true;
    if (prevChar === '(') {
      for (let i = index - 2; i >= 0; i--) {
        const ch = text[i];
        if (ch === ']') return true;
        if (!/\s/.test(ch)) break;
      }
    }
    return false;
  }

  function renderPlainText(value, options = {}) {
    const text = value == null ? '' : String(value);
    if (!text) return '';
    const preserve = Boolean(options.preserveLineBreaks);
    const html = linkifyPlainString(text, preserve);
    return sanitizeHtml(html);
  }

  function linkifyPlainString(text, preserveLineBreaks) {
    if (!linkify) {
      return escapeTextSegment(text, preserveLineBreaks);
    }
    const matches = linkify.match(text);
    if (!matches || !matches.length) {
      return escapeTextSegment(text, preserveLineBreaks);
    }
    let html = '';
    let cursor = 0;
    for (const match of matches) {
      const start = match.index ?? 0;
      const end = match.lastIndex ?? start;
      if (start < cursor) continue;
      if (start > cursor) {
        html += escapeTextSegment(text.slice(cursor, start), preserveLineBreaks);
      }
      const href = buildHref(match.url || match.raw || match.text || '');
      const label = htmlEscape(match.text || match.raw || match.url || '');
      html += `<a href="${href}" target="_blank" rel="noopener noreferrer">${label}</a>`;
      cursor = end;
    }
    html += escapeTextSegment(text.slice(cursor), preserveLineBreaks);
    return html;
  }

  function escapeTextSegment(segment, preserveLineBreaks) {
    const escaped = htmlEscape(segment || '');
    return preserveLineBreaks
      ? escaped.replace(/(?:\r\n|\r|\n)/g, '<br>')
      : escaped;
  }

  function buildHref(value) {
    const raw = (value || '').trim();
    if (!raw) return '#';
    const normalized = /^https?:\/\//i.test(raw) ? raw : `https://${raw}`;
    return escapeAttribute(normalized);
  }

  function htmlEscape(value) {
    const div = document.createElement('div');
    div.textContent = value || '';
    return div.innerHTML;
  }

  function escapeAttribute(value) {
    return String(value || '').replace(/[&"'<>]/g, ch => ({
      '&': '&amp;',
      '"': '&quot;',
      "'": '&#39;',
      '<': '&lt;',
      '>': '&gt;'
    }[ch] || ch));
  }

  function resolveMarked() {
    if (window.marked && typeof window.marked.marked === 'function') {
      return window.marked.marked;
    }
    if (typeof window.marked === 'function') {
      return window.marked;
    }
    return null;
  }

  function typingIndicatorHtml() {
    return '<div class="chat-message assistant"><div class="typing-dots"><span class="typing-dot"></span><span class="typing-dot"></span><span class="typing-dot"></span></div></div>';
  }
})();
