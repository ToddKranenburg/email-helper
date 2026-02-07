(function () {
  const bootstrap = window.SECRETARY_BOOTSTRAP || {};
  const threads = Array.isArray(bootstrap.threads) ? bootstrap.threads : [];
  const priorityPayload = Array.isArray(bootstrap.priority) ? bootstrap.priority : null;
  const priorityProgress = bootstrap.priorityProgress && typeof bootstrap.priorityProgress === 'object'
    ? bootstrap.priorityProgress
    : null;
  const syncMeta = window.SYNC_META && typeof window.SYNC_META === 'object'
    ? window.SYNC_META
    : null;
  const priorityBatchFinishedAt = typeof bootstrap.priorityBatchFinishedAt === 'string'
    ? bootstrap.priorityBatchFinishedAt
    : null;
  const priorityItems = normalizePriorityItems(priorityPayload);
  const hasServerPriority = Array.isArray(bootstrap.priority);
  const MAX_TURNS = typeof bootstrap.maxTurns === 'number' ? bootstrap.maxTurns : 0;
  const PAGE_SIZE = typeof bootstrap.pageSize === 'number' ? bootstrap.pageSize : 20;
  const TOTAL_ITEMS = typeof bootstrap.totalItems === 'number' ? bootstrap.totalItems : 0;
  const HAS_MORE = Boolean(bootstrap.hasMore);
  const NEXT_PAGE = typeof bootstrap.nextPage === 'number' ? bootstrap.nextPage : (HAS_MORE ? 2 : 0);
  const DEFAULT_NUDGE = 'Type your next move here.';
  const REVIEW_PROMPT = 'Give me a more detailed but easy-to-digest summary of this email. Highlight the main points, asks, deadlines, and any decisions in quick bullets.';
  const REVISIT_TEMPLATES = [
    desc => `This is ${desc}. Want a recap or a next step?`,
    desc => `Quick context: ${desc}. Want me to summarize or suggest a reply?`,
    desc => `Email overview: ${desc}. What should we do next?`,
    desc => `Here’s the email: ${desc}. Want a refresher or a next move?`,
    desc => `Context check: ${desc}. Want me to draft a reply or log a task?`
  ];
  const PRIORITY_LIMIT = 6;
  const PRIORITY_MIN_SCORE = 4;
  const PRIORITY_POLL_INTERVAL_MS = 8000;
  const PRIORITY_LOADING_MIN_MS = 450;
  const PRIORITY_SYNC_STORAGE_KEY = 'prioritySyncStartAt';
  const PRIORITY_SYNC_MAX_AGE_MS = 10 * 60 * 1000;
  const AUTO_THREAD_PAUSE_MS = 1500;
  const PRIORITY_PATTERNS = {
    urgent: /\b(urgent|asap|immediately|time[-\s]?sensitive|deadline|final notice|action required|response required|reply needed|respond by|past due|overdue|expir(?:e|es|ing))\b/i,
    security: /\b(security|verify|verification|password|2fa|unauthorized|suspicious|fraud|breach|locked?|login|sign[-\s]?in|account alert)\b/i,
    payment: /\b(payment|invoice|receipt|billing|charge|charged|refund|past due|overdue|card)\b/i,
    approval: /\b(approve|approval|sign[-\s]?off|contract|legal|compliance|policy)\b/i,
    scheduling: /\b(meeting|call|calendar|schedule|reschedule|availability|zoom|appointment|rsvp|invite)\b/i
  };
  window.SECRETARY_BOOTSTRAP = undefined;
  window.SYNC_META = undefined;
  const debugEnabled = true;
  const logDebug = (...args) => {
    if (!debugEnabled) return;
    // eslint-disable-next-line no-console
    console.info('[secretary]', ...args);
  };

  function parseTimestamp(value) {
    if (!value) return 0;
    const ts = Date.parse(value);
    return Number.isFinite(ts) ? ts : 0;
  }

  function sleep(ms) {
    return new Promise(resolve => window.setTimeout(resolve, ms));
  }

  function normalizePriorityProgress(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const prioritizedCount = Number(raw.prioritizedCount);
    const totalCount = Number(raw.totalCount);
    if (!Number.isFinite(prioritizedCount) || !Number.isFinite(totalCount)) return null;
    return {
      prioritizedCount: Math.max(0, prioritizedCount),
      totalCount: Math.max(0, totalCount)
    };
  }

  function readPrioritySyncStart() {
    const raw = window.localStorage.getItem(PRIORITY_SYNC_STORAGE_KEY);
    if (!raw) return 0;
    const value = Number(raw);
    if (!Number.isFinite(value)) return 0;
    if (Date.now() - value > PRIORITY_SYNC_MAX_AGE_MS) {
      window.localStorage.removeItem(PRIORITY_SYNC_STORAGE_KEY);
      return 0;
    }
    return value;
  }

  function writePrioritySyncStart(ts) {
    if (!ts) return;
    window.localStorage.setItem(PRIORITY_SYNC_STORAGE_KEY, String(ts));
  }

  function clearPrioritySyncStart() {
    window.localStorage.removeItem(PRIORITY_SYNC_STORAGE_KEY);
  }

  const refs = {
    count: document.getElementById('triage-count'),
    loadMoreHead: document.getElementById('load-more-head'),
    progress: document.getElementById('triage-progress'),
    progressTrack: document.querySelector('.triage-progress'),
    priorityPill: document.getElementById('priority-pill-text'),
    queuePill: document.getElementById('queue-pill-text'),
    emailEmpty: document.getElementById('email-empty'),
    loadMoreEmpty: document.getElementById('load-more-empty'),
    emailEmptyText: document.querySelector('#email-empty p'),
    priorityQueue: document.getElementById('priority-queue'),
    priorityToggle: document.getElementById('priority-toggle'),
    reviewQueue: document.getElementById('review-queue'),
    queueToggle: document.getElementById('queue-toggle'),
    subject: document.getElementById('email-subject'),
    chatLog: document.getElementById('assistant-chat-log'),
    chatScroll: document.getElementById('assistant-chat'),
    chatForm: document.getElementById('assistant-form'),
    chatInput: document.getElementById('assistant-input-field'),
    chatError: document.getElementById('assistant-error'),
    chatHint: document.getElementById('assistant-hint'),
    reviewBtn: document.getElementById('action-review'),
    replyBtn: document.getElementById('action-reply'),
    archiveBtn: document.getElementById('action-archive'),
    unsubscribeBtn: document.getElementById('action-unsubscribe'),
    openLinkBtn: document.getElementById('action-open-link'),
    taskBtn: document.getElementById('action-task'),
    skipBtn: document.getElementById('action-skip'),
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
    priorityList: document.getElementById('priority-list'),
    priorityProgress: document.getElementById('priority-progress'),
    replyPanel: document.getElementById('reply-panel'),
    replyPanelHelper: document.getElementById('reply-panel-helper'),
    replyTo: document.getElementById('reply-to'),
    replySubject: document.getElementById('reply-subject'),
    replyBody: document.getElementById('reply-body'),
    replyError: document.getElementById('reply-error'),
    replyCancel: document.getElementById('reply-cancel'),
    replySubmit: document.getElementById('reply-submit'),
    replyClose: document.getElementById('reply-close'),
    reviewList: document.getElementById('review-list'),
    syncForm: document.getElementById('ingest-form'),
    sidebarToggle: document.getElementById('sidebar-expand'),
    sidebarBack: document.getElementById('sidebar-back')
  };
  const loadingOverlay = document.getElementById('loading');
  const loadingText = loadingOverlay ? loadingOverlay.querySelector('.loading-text') : null;
  const OVERLAY_OWNER = 'secretary';

  if (!refs.chatLog || !refs.chatForm || !refs.emailEmpty) {
    return;
  }

  const mobileSidebarMedia = window.matchMedia('(max-width: 1024px)');
  const sharedHeaderCenter = document.querySelector('.shared-header-center');
  const desktopSyncSlot = document.getElementById('desktop-sync-slot');
  const syncFormEl = document.getElementById('ingest-form');
  const syncMetaEl = document.getElementById('sync-meta');

  function moveSyncToDesktop() {
    if (!desktopSyncSlot || !syncFormEl || !syncMetaEl) return;
    desktopSyncSlot.append(syncFormEl, syncMetaEl);
  }

  function moveSyncToShared() {
    if (!sharedHeaderCenter || !syncFormEl || !syncMetaEl) return;
    sharedHeaderCenter.append(syncFormEl, syncMetaEl);
  }

  function updateToggleButton(btn, label, isOpen) {
    if (!btn) return;
    btn.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
    btn.setAttribute('aria-label', label);
    const labelEl = btn.querySelector('.mobile-expand-label');
    if (labelEl) {
      labelEl.textContent = label;
    }
  }

  function setSidebarState({ queueOpen, sidebarCollapsed }) {
    document.body.classList.toggle('queue-open', queueOpen);
    document.body.classList.toggle('sidebar-collapsed', sidebarCollapsed);
    if (mobileSidebarMedia.matches) {
      updateToggleButton(refs.sidebarToggle, 'Inbox', false);
      updateToggleButton(refs.sidebarBack, 'Chat', queueOpen);
      return;
    }
    const label = sidebarCollapsed ? 'Inbox' : 'Chat';
    updateToggleButton(refs.sidebarToggle, label, !sidebarCollapsed);
    updateToggleButton(refs.sidebarBack, label, !sidebarCollapsed);
  }

  function syncSidebarForViewport() {
    if (!mobileSidebarMedia.matches) {
      document.body.classList.remove('sidebar-collapsed');
      moveSyncToDesktop();
      document.body.classList.remove('queue-open');
      setSidebarState({
        queueOpen: false,
        sidebarCollapsed: false
      });
      return;
    }
    moveSyncToShared();
    document.body.classList.remove('sidebar-collapsed');
    setSidebarState({
      queueOpen: document.body.classList.contains('queue-open'),
      sidebarCollapsed: false
    });
  }

  function openQueueView() {
    if (mobileSidebarMedia.matches) {
      setSidebarState({ queueOpen: true, sidebarCollapsed: false });
    } else {
      const nextCollapsed = !document.body.classList.contains('sidebar-collapsed');
      setSidebarState({ queueOpen: false, sidebarCollapsed: nextCollapsed });
    }
  }

  function closeQueueView() {
    if (mobileSidebarMedia.matches) {
      setSidebarState({ queueOpen: false, sidebarCollapsed: false });
    }
  }

  if (refs.sidebarToggle) {
    refs.sidebarToggle.addEventListener('click', openQueueView);
  }
  if (refs.sidebarBack) {
    refs.sidebarBack.addEventListener('click', closeQueueView);
  }

  if (mobileSidebarMedia.addEventListener) {
    mobileSidebarMedia.addEventListener('change', syncSidebarForViewport);
  } else if (mobileSidebarMedia.addListener) {
    mobileSidebarMedia.addListener(syncSidebarForViewport);
  }
  syncSidebarForViewport();

  const DEFAULT_PLACEHOLDER = refs.chatInput?.placeholder || 'Type a message…';
  const SUGGESTED_PLACEHOLDER = 'Press Enter to accept • or type to respond…';

  const state = {
    lookup: new Map(),
    positions: new Map(),
    needs: [],
    priority: [],
    priorityMeta: new Map(),
    prioritySource: hasServerPriority ? 'server' : 'local',
    priorityLoading: false,
    priorityReady: false,
    priorityReadyAnnounced: false,
    priorityWaitingNotified: false,
    prioritySwitchPending: false,
    prioritySwitchInProgress: false,
    userEngagedInbox: false,
    pendingPriorityAnnouncement: '',
    autoSelectBlocked: false,
    prioritySyncStartAt: 0,
    priorityLastBatchAt: parseTimestamp(priorityBatchFinishedAt),
    priorityProgress: normalizePriorityProgress(priorityProgress),
    priorityPolling: false,
    priorityPollTimer: 0,
    histories: new Map(),
    timeline: [],
    timelineMessageIds: new Set(),
    serverTimelines: new Map(),
    actionFlows: new Map(),
    activeId: '',
    typing: false,
    totalLoaded: threads.length,
    totalInboxCount: TOTAL_ITEMS,
    totalInboxCountKnown: Boolean(syncMeta?.lastSyncAt),
    pageSize: PAGE_SIZE,
    hasMore: HAS_MORE,
    nextPage: NEXT_PAGE,
    loadingMore: false,
    autoAdvanceTimer: 0,
    snapThreadId: '',
    pendingCreateThreadId: '',
    pendingArchiveThreadId: '',
    hydrated: new Set(),
    hydrating: new Set(),
    pendingSuggestedActions: new Map(),
    seenThreads: new Set(),
    loadTypingThreads: new Set(),
    mustKnowByThread: new Map(),
    suggestedLinksByThread: new Map(),
    openLinkProgress: new Map()
  };
  const assistantQueues = new Map();
  const pendingTranscripts = new Map();
  let typingSessions = 0;
  const taskState = {
    open: false,
    status: 'idle', // idle | submitting | success | error
    suggested: { title: '', notes: '', due: '' },
    values: { title: '', notes: '', due: '' },
    error: '',
    lastSourceId: ''
  };
  const replyState = {
    open: false,
    status: 'idle', // idle | sending | error
    values: { to: '', subject: '', body: '' },
    error: '',
    lastSourceId: '',
    baseBody: '',
    suggested: '',
    suggesting: false
  };
  const reviewedIds = new Set();
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
    const normalized = normalizeThread(thread);
    if (!normalized) return;
    state.lookup.set(normalized.threadId, normalized);
    state.positions.set(normalized.threadId, index);
    state.needs.push(normalized.threadId);
    state.serverTimelines.set(normalized.threadId, normalized.timeline || []);
    if (normalized.actionFlow) {
      state.actionFlows.set(normalized.threadId, normalized.actionFlow);
    }
    logDebug('bootstrap thread', {
      threadId: normalized.threadId,
      timelineCount: normalized.timeline?.length || 0,
      actionFlow: normalized.actionFlow
    });
  });

  if (hasServerPriority) {
    state.priority = priorityItems.map(item => item.threadId);
    priorityItems.forEach(item => {
      state.priorityMeta.set(item.threadId, {
        score: item.score,
        reason: item.reason,
        reasonWeight: item.reasonWeight
      });
    });
  }
  state.totalLoaded = state.positions.size;

  init();

  function init() {
    state.priorityReady = state.priority.length > 0;
    state.prioritySyncStartAt = readPrioritySyncStart();
    if (state.prioritySyncStartAt && (!state.priorityLastBatchAt || state.priorityLastBatchAt < state.prioritySyncStartAt)) {
      state.priorityLoading = true;
    }
    rebuildPriorityQueue();
    updateHeaderCount();
    updateProgress();
    updatePriorityPill();
    updatePriorityProgress();
    updateQueuePill();
    updateDrawerLists();
    updateLoadMoreButtons();
    wireEvents();
    startPriorityPolling();

    if (state.needs.length) {
      state.autoSelectBlocked = isPriorityWaiting();
      if (state.autoSelectBlocked) {
        setEmptyState('I\'m prioritizing your inbox to find emails that need your attention. You can wait here, or jump into your full inbox and start reviewing anytime.');
        toggleComposer(false);
      } else {
        const startId = getPriorityFirstCandidate() || state.needs[0];
        setActiveThread(startId, { source: 'init' });
        maybeAnnouncePriorityWaiting();
        maybeSwitchToPriority('init');
      }
    } else {
      setEmptyState('No emails queued. Tap Sync Gmail to pull fresh ones.');
      toggleComposer(false);
    }

  }

  function wireEvents() {
    refs.chatForm.addEventListener('submit', handleChatSubmit);
    refs.chatInput.addEventListener('keydown', handleChatKeydown);
    refs.chatInput.addEventListener('input', clearComposerNudge);
    refs.chatLog.addEventListener('click', handleTimelineClick);
    if (refs.syncForm) {
      refs.syncForm.addEventListener('submit', () => {
        const now = Date.now();
        state.prioritySyncStartAt = now;
        writePrioritySyncStart(now);
        setPriorityLoading(true);
        startPriorityPolling();
      });
    }

    if (refs.reviewBtn) {
      refs.reviewBtn.addEventListener('click', () => {
        markUserEngagedIfInbox(state.activeId);
        requestReview();
      });
    }
    if (refs.replyBtn) {
      refs.replyBtn.addEventListener('click', () => {
        markUserEngagedIfInbox(state.activeId);
        openReplyPanel();
      });
    }
    if (refs.archiveBtn) {
      refs.archiveBtn.addEventListener('click', () => {
        markUserEngagedIfInbox(state.activeId);
        archiveCurrent('button');
      });
    }
    if (refs.unsubscribeBtn) {
      refs.unsubscribeBtn.addEventListener('click', () => {
        markUserEngagedIfInbox(state.activeId);
        unsubscribeCurrent('button');
      });
    }
    if (refs.openLinkBtn) {
      refs.openLinkBtn.addEventListener('click', () => {
        markUserEngagedIfInbox(state.activeId);
        openSuggestedLinks('button');
      });
    }
    if (refs.taskBtn) {
      refs.taskBtn.addEventListener('click', () => {
        markUserEngagedIfInbox(state.activeId);
        if (state.activeId) requestDraft(state.activeId, 'generate');
      });
    }
    if (refs.skipBtn) {
      refs.skipBtn.addEventListener('click', () => {
        markUserEngagedIfInbox(state.activeId);
        skipCurrent('button');
      });
    }

    if (refs.loadMoreHead) {
      refs.loadMoreHead.addEventListener('click', () => {
        handleReviewLatestClick();
      });
    }
    if (refs.loadMoreEmpty) {
      refs.loadMoreEmpty.addEventListener('click', () => {
        handleReviewLatestClick();
      });
    }

    if (refs.priorityList) refs.priorityList.addEventListener('click', (event) => handleThreadListClick(event, 'priority'));
    if (refs.reviewList) refs.reviewList.addEventListener('click', (event) => handleThreadListClick(event, 'queue'));
    if (refs.priorityToggle && refs.priorityQueue) {
      refs.priorityToggle.addEventListener('click', () => toggleQueueSection(refs.priorityQueue, refs.priorityToggle));
    }
    if (refs.queueToggle && refs.reviewQueue) {
      refs.queueToggle.addEventListener('click', () => toggleQueueSection(refs.reviewQueue, refs.queueToggle));
    }

    document.addEventListener('click', (event) => {
      const target = event.target;
      if (!(target instanceof Element)) return;
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
    if (refs.replySubmit) {
      refs.replySubmit.addEventListener('click', async (event) => {
        await submitReply(event);
      });
    }
    if (refs.replyCancel) refs.replyCancel.addEventListener('click', () => closeReplyPanel(true));
    if (refs.replyClose) refs.replyClose.addEventListener('click', () => closeReplyPanel(true));
    if (refs.replyBody) refs.replyBody.addEventListener('input', syncReplyValues);
  }

  function handleThreadListClick(event, variant) {
    const targetEl = event.target;
    if (!(targetEl instanceof Element)) return;
    const loadMoreBtn = targetEl.closest('.load-more');
    if (loadMoreBtn) {
      fetchNextPage('button');
      return;
    }
    const selector = variant === 'drawer' ? '.drawer-thread' : '.queue-item';
    const target = targetEl.closest(selector);
    if (!target) return;
    const threadId = target.dataset.threadId;
    if (!threadId || !state.lookup.has(threadId)) return;
    if (variant !== 'priority') {
      markUserEngagedIfInbox(threadId);
    }
    if (variant === 'drawer') toggleDrawer(false);
    setActiveThread(threadId, { source: 'user' });
    if (mobileSidebarMedia.matches && document.body.classList.contains('queue-open')) {
      closeQueueView();
    }
  }

  function handleTimelineClick(event) {
    const target = event.target;
    if (!(target instanceof Element)) return;
    const action = target.getAttribute('data-action');
    if (!action) return;
    const threadId = target.getAttribute('data-thread-id') || state.activeId;
    if (!threadId) return;
    if (action === 'suggested-primary') {
      const actionType = normalizeSuggestedAction(target.getAttribute('data-action-type'));
      handleSuggestedActionClick(threadId, actionType);
      return;
    }
    if (action === 'draft-create') {
      executeActionForThread(threadId, 'create_task');
      return;
    }
    if (action === 'draft-edit') {
      requestDraft(threadId, 'edit');
      return;
    }
    if (action === 'editor-create' || action === 'editor-save') {
      const editor = target.closest('.draft-editor');
      const values = readInlineEditorValues(editor);
      if (!values) return;
      if (action === 'editor-create') {
        executeActionForThread(threadId, 'create_task', values);
      } else {
        requestDraft(threadId, 'save', values);
      }
      return;
    }
  }

  function handleChatKeydown(event) {
    if (event.defaultPrevented) return;
    if (event.key !== 'Enter') return;
    if (event.shiftKey || event.altKey || event.metaKey || event.ctrlKey) return;
    if (refs.chatInput.disabled) return;
    const trimmed = refs.chatInput.value.trim();
    if (!trimmed) {
      if (isCreateConfirmationPending(state.activeId)) {
        event.preventDefault();
        refs.chatInput.value = 'yes';
        refs.chatForm.requestSubmit();
        return;
      }
      if (isArchiveConfirmationPending(state.activeId)) {
        event.preventDefault();
        refs.chatInput.value = 'yes';
        refs.chatForm.requestSubmit();
        return;
      }
      const pendingSuggested = getPendingSuggestedAction(state.activeId);
      if (pendingSuggested) {
        event.preventDefault();
        refs.chatInput.value = 'yes';
        refs.chatForm.requestSubmit();
        return;
      }
    }
    event.preventDefault();
    refs.chatForm.requestSubmit();
  }

  async function handleChatSubmit(event) {
    event.preventDefault();
    if (!state.activeId) return;
    markUserEngagedIfInbox(state.activeId);
    const question = refs.chatInput.value.trim();
    if (!question) return;

    const history = ensureHistory(state.activeId);
    const asked = history.filter(turn => turn.role === 'user').length;
    const pendingCreate = isCreateConfirmationPending(state.activeId);
    const pendingArchive = isArchiveConfirmationPending(state.activeId);
    if (!pendingCreate && !pendingArchive && MAX_TURNS > 0 && asked >= MAX_TURNS) {
      setChatError('Chat limit reached for this email.');
      return;
    }
    setChatError('');

    appendTurn(state.activeId, { role: 'user', content: question });
    renderChat();
    refs.chatInput.value = '';
    toggleComposer(false, { preserveTaskPanel: pendingCreate || taskState.open, preserveReplyPanel: replyState.open });
    startAssistantTyping(state.activeId);
    const submitBtn = refs.chatForm.querySelector('button[type="submit"]');
    if (submitBtn) submitBtn.disabled = true;

    try {
      if (pendingCreate) {
        await handleCreateConfirmationResponse(question);
        return;
      }
      if (pendingArchive) {
        await handleArchiveConfirmationResponse(question);
        return;
      }

      const handledSuggested = await handleSuggestedActionResponse(question);
      if (handledSuggested) {
        stopAssistantTyping(state.activeId);
        toggleComposer(Boolean(state.activeId));
        if (submitBtn) submitBtn.disabled = false;
        return;
      }

      const intent = await detectIntent(question);
      if (intent === 'skip') {
        stopAssistantTyping(state.activeId);
        toggleComposer(Boolean(state.activeId));
        if (submitBtn) submitBtn.disabled = false;
        handleAutoIntent(intent, question, { alreadyLogged: true });
        return;
      }
      if (intent === 'archive') {
        stopAssistantTyping(state.activeId);
        toggleComposer(Boolean(state.activeId));
        if (submitBtn) submitBtn.disabled = false;
        await handleArchiveIntent(question, { alreadyLogged: true });
        return;
      }
      if (intent === 'reply') {
        stopAssistantTyping(state.activeId);
        toggleComposer(Boolean(state.activeId));
        if (submitBtn) submitBtn.disabled = false;
        openReplyPanelFromPrompt(question);
        return;
      }
      if (intent === 'unsubscribe') {
        stopAssistantTyping(state.activeId);
        toggleComposer(Boolean(state.activeId));
        if (submitBtn) submitBtn.disabled = false;
        await executeActionForThread(state.activeId, 'unsubscribe');
        return;
      }
      if (intent === 'open_link') {
        stopAssistantTyping(state.activeId);
        toggleComposer(Boolean(state.activeId));
        if (submitBtn) submitBtn.disabled = false;
        await openSuggestedLinks('intent');
        return;
      }
      if (intent === 'create_task') {
        stopAssistantTyping(state.activeId);
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
      await enqueueAssistantMessage(state.activeId, data.reply || 'No response received.');
    } catch (err) {
      popLastTurn(state.activeId);
      renderChat();
      setChatError('Unable to reach the assistant. Check your connection.');
      refs.chatInput.value = question;
    } finally {
      stopAssistantTyping(state.activeId);
      toggleComposer(true);
      const submitBtn2 = refs.chatForm.querySelector('button[type="submit"]');
      if (submitBtn2) submitBtn2.disabled = false;
      nudgeComposer(DEFAULT_NUDGE, { focus: true });
    }
  }

  async function handleSuggestedActionClick(threadId, actionType) {
    const normalized = normalizeSuggestedAction(actionType);
    if (!normalized) return;
    markUserEngagedIfInbox(threadId);
    clearPendingSuggestedAction(threadId);
    if (normalized === 'reply') {
      if (threadId && threadId !== state.activeId) setActiveThread(threadId, { source: 'user' });
      openReplyPanel();
      return;
    }
    if (normalized === 'create_task') {
      await requestDraft(threadId, 'generate');
      return;
    }
    if (normalized === 'open_link') {
      await openSuggestedLinks('suggested');
      return;
    }
    await executeActionForThread(threadId, normalized);
  }

  async function fetchAutoSummary(threadId) {
    if (!threadId || state.hydrating.has(threadId)) return;
    state.hydrating.add(threadId);
    startThreadLoadTyping(threadId);
    logDebug('fetchAutoSummary:start', threadId);
    try {
      const resp = await fetch('/secretary/auto-summarize', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({ threadId, fresh: true })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) throw new Error(data?.error || 'Unable to prep this email.');
      if (data.flow) updateActionFlow(threadId, data.flow);
      if (Array.isArray(data.timeline)) {
        syncTimelineFromServer(threadId, data.timeline);
      }
      state.hydrated.add(threadId);
      renderChat(threadId);
      logDebug('fetchAutoSummary:done', { threadId, timelineCount: Array.isArray(data.timeline) ? data.timeline.length : 0, flow: data.flow });
    } catch (err) {
      console.error('Auto summarize failed', err);
      if (threadId === state.activeId) {
        setChatError(err instanceof Error ? err.message : 'Unable to load that email right now.');
      }
      state.hydrated.delete(threadId);
    } finally {
      state.hydrating.delete(threadId);
      stopThreadLoadTyping(threadId);
    }
  }

  async function requestDraft(threadId, mode = 'generate', draft) {
    if (!threadId) return;
    try {
      const resp = await fetch('/secretary/action/draft', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({ threadId, mode, draft })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) throw new Error(data?.error || 'Unable to prepare the draft.');
      if (data.flow) updateActionFlow(threadId, data.flow);
      if (Array.isArray(data.timeline)) {
        syncTimelineFromServer(threadId, data.timeline);
      }
      renderChat(threadId);
    } catch (err) {
      setChatError(err instanceof Error ? err.message : 'Unable to prepare that action.');
    }
  }

  async function executeActionForThread(threadId, actionType, draftPayload) {
    if (!threadId || !actionType) return;
    setChatError('');
    try {
      const body = { threadId, actionType };
      if (draftPayload) {
        if (actionType === 'open_link' && Array.isArray(draftPayload.links)) {
          body.links = draftPayload.links;
        } else {
          body.draft = draftPayload;
        }
      }
      const resp = await fetch('/secretary/action/execute', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify(body)
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        if (data.flow) updateActionFlow(threadId, data.flow);
        if (Array.isArray(data.timeline)) {
          syncTimelineFromServer(threadId, data.timeline);
        }
        throw new Error(data?.error || 'Unable to process that action.');
      }
      if (data.flow) updateActionFlow(threadId, data.flow);
      if (Array.isArray(data.timeline)) {
        syncTimelineFromServer(threadId, data.timeline);
      }
      if (actionType === 'archive') {
        await waitForAssistantSettled(threadId);
        removeCurrentFromQueue();
      } else if (actionType === 'unsubscribe') {
        await waitForAssistantSettled(threadId);
        removeCurrentFromQueue();
      } else if (actionType === 'skip') {
        await waitForAssistantSettled(threadId);
        skipCurrent('action');
      } else if (actionType === 'create_task') {
        clearPendingCreate();
        if (data?.status === 'created') {
          await waitForAssistantSettled(threadId);
          promptArchiveAfterTask(null, { includeSuccess: false, threadId });
        }
      }
      clearPendingSuggestedAction(threadId);
      if (threadId === state.activeId) {
        renderChat(threadId);
      } else if (state.activeId) {
        renderChat(state.activeId);
      }
    } catch (err) {
      if (actionType === 'unsubscribe') {
        enqueueAssistantMessage(threadId, pickUnsubscribeUnavailableMessage());
      }
      setChatError(err instanceof Error ? err.message : 'Unable to process that action.');
    }
  }

  function syncTimelineFromServer(threadId, timeline) {
    state.serverTimelines.set(threadId, Array.isArray(timeline) ? timeline : []);
    state.hydrated.add(threadId);
    mergeServerTimeline(threadId, timeline || []);
    logDebug('syncTimelineFromServer', { threadId, timelineCount: Array.isArray(timeline) ? timeline.length : 0 });
  }

  function updateActionFlow(threadId, flow) {
    if (!flow) return;
    state.actionFlows.set(threadId, flow);
    const thread = state.lookup.get(threadId);
    if (thread) thread.actionFlow = flow;
    rebuildPriorityQueue();
    updatePriorityPill();
    updateDrawerLists();
  }

  function readInlineEditorValues(editorEl) {
    if (!editorEl) return null;
    const title = editorEl.querySelector('[data-field=\"title\"]');
    const notes = editorEl.querySelector('[data-field=\"notes\"]');
    const due = editorEl.querySelector('[data-field=\"dueDate\"]');
    return {
      title: title?.value?.trim() || '',
      notes: notes?.value?.trim() || '',
      dueDate: due?.value?.trim() || ''
    };
  }

  async function fetchNextPage(reason) {
    if (!state.hasMore || state.loadingMore) return [];
    const overlayActive = reason === 'auto';
    if (overlayActive) showLoadingOverlay('Loading more emails…');
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
      if (!state.activeId && state.needs.length && reason !== 'auto') {
        const nextId = getNextReviewCandidate() || state.needs[0];
        setActiveThread(nextId, { source: 'user' });
        if (!isPriorityWaiting()) {
          const nextId = getNextReviewCandidate() || state.needs[0];
          setActiveThread(nextId);
        }
      }
      return added;
    } catch (err) {
      console.error('Failed to load more threads', err);
      return [];
    } finally {
      state.loadingMore = false;
      if (overlayActive) hideLoadingOverlay();
      updateDrawerLists();
      updateLoadMoreButtons();
      nudgeComposer('Loaded more - type your next move here.', { focus: false });
    }
  }

  function setActiveThread(threadId, options = {}) {
    if (!threadId || !state.lookup.has(threadId)) return;
    if (state.activeId === threadId) return;
    const wasSeen = state.seenThreads.has(threadId);
    const source = typeof options.source === 'string' ? options.source : 'auto';
    if (state.autoSelectBlocked && source === 'auto' && !isPriorityThread(threadId)) {
      return;
    }
    clearAutoAdvance();
    state.snapThreadId = threadId;
    state.activeId = threadId;
    const thread = state.lookup.get(threadId);
    logDebug('setActiveThread', threadId, {
      source,
      hasTimeline: state.timeline.some(item => item.threadId === threadId),
      serverTimelineCount: (state.serverTimelines.get(threadId) || []).length,
      hydrated: state.hydrated.has(threadId)
    });
    if (!thread) return;
    if (taskState.open && taskState.lastSourceId !== threadId) {
      closeTaskPanel(true);
    }
    if (replyState.open && replyState.lastSourceId !== threadId) {
      closeReplyPanel(true);
    }
    if (state.pendingCreateThreadId && state.pendingCreateThreadId !== threadId) {
      clearPendingCreate();
    }
    if (state.pendingArchiveThreadId && state.pendingArchiveThreadId !== threadId) {
      clearPendingArchive();
    }

    refs.emailEmpty.classList.add('hidden');
    updateEmailCard(thread);
    if (state.pendingPriorityAnnouncement && isPriorityThread(threadId)) {
      appendTurn(threadId, { role: 'assistant', content: state.pendingPriorityAnnouncement });
      state.pendingPriorityAnnouncement = '';
    }
    insertThreadDivider(threadId);
    startThreadLoadTyping(threadId);
    hydrateThreadTimeline(threadId);
    revealPendingTranscripts(threadId);
    ensureHistory(threadId);
    setChatError('');
    renderChat(threadId, { scrollMode: 'divider' });
    updateHint(threadId);
    updateDrawerLists();
    updateQueuePill();
    toggleComposer(true);
    refs.chatInput.value = '';
    nudgeComposer(DEFAULT_NUDGE, { focus: true });
    if (wasSeen) {
      const prompt = buildRevisitPrompt(thread);
      enqueueAssistantMessage(threadId, prompt);
    }
    state.seenThreads.add(threadId);
    maybeAnnouncePriorityWaiting();

    const pendingCount = pendingTranscripts.get(threadId)?.length || 0;
    const serverItems = state.serverTimelines.get(threadId) || [];
    const needsFetch = !state.hydrated.has(threadId) && !serverItems.length;
    const isHydrating = state.hydrating.has(threadId);
    if (!pendingCount && !needsFetch && !isHydrating) {
      stopThreadLoadTyping(threadId);
    }
  }

  function buildRevisitPrompt(thread) {
    if (!thread) return 'Want a recap or next step?';
    const headline = typeof thread.headline === 'string' ? thread.headline.trim() : '';
    const summary = typeof thread.summary === 'string' ? thread.summary.trim() : '';
    const subject = typeof thread.subject === 'string' ? thread.subject.trim() : '';
    const sender = typeof thread.from === 'string' ? thread.from.trim() : '';
    let descriptor = headline || summary || subject || 'this email';
    descriptor = descriptor.replace(/\s+/g, ' ').trim();
    if (sender) descriptor = `${descriptor} from ${sender}`;
    if (!descriptor || descriptor === 'this email') return 'Want a recap or next step?';
    const template = REVISIT_TEMPLATES[Math.floor(Math.random() * REVISIT_TEMPLATES.length)];
    return template(descriptor);
  }

  function updateEmailCard(thread) {
    if (refs.position) {
      const label = formatEmailPosition(thread.threadId);
      refs.position.textContent = label || '';
      refs.position.classList.toggle('hidden', !label);
    }
    const subjectCopy = thread.subject || '(no subject)';
    if (refs.subject) refs.subject.textContent = subjectCopy;
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
    closeReplyPanel(true);
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

  function renderReplyPanel() {
    if (!refs.replyPanel || !refs.replyBody || !refs.replyTo || !refs.replySubject) return;
    const showPanel = replyState.open;
    refs.replyPanel.classList.toggle('hidden', !showPanel);
    refs.replyPanel.classList.toggle('loading', replyState.status === 'sending');
    refs.replyTo.textContent = replyState.values.to || 'Unknown sender';
    refs.replySubject.textContent = replyState.values.subject || '(no subject)';
    refs.replyBody.value = replyState.values.body || '';

    const disabled = replyState.status === 'sending';
    refs.replyBody.disabled = disabled;
    const replyField = refs.replyBody.closest('.reply-field');
    if (replyField) replyField.classList.toggle('drafting', replyState.suggesting);
    if (replyState.suggesting && !(replyState.values.body || '').trim()) {
      refs.replyBody.setAttribute('placeholder', 'Weaving a reply…');
    } else {
      refs.replyBody.setAttribute('placeholder', 'Type your reply…');
    }
    if (refs.replySubmit) {
      refs.replySubmit.disabled = disabled;
      refs.replySubmit.textContent = disabled ? 'Sending…' : 'Send reply';
    }
    if (refs.replyCancel) refs.replyCancel.disabled = disabled;
    if (refs.replyClose) refs.replyClose.disabled = disabled;

    if (refs.replyError) {
      refs.replyError.textContent = replyState.error || '';
      refs.replyError.classList.toggle('hidden', !replyState.error);
    }
    if (refs.replyPanelHelper) {
      if (replyState.suggesting) {
        refs.replyPanelHelper.textContent = 'Drafting a suggested reply...';
      } else if (replyState.suggested) {
        refs.replyPanelHelper.textContent = 'Suggested reply added. Edit anything before sending.';
      } else {
        refs.replyPanelHelper.textContent = 'Send a quick response without leaving the queue.';
      }
    }
  }

  function openReplyPanel(options = {}) {
    const opts = options instanceof Event ? {} : options;
    const preserveValues = Boolean(opts.preserveValues);
    const prefillBody = typeof opts.prefillBody === 'string' ? opts.prefillBody.trim() : '';
    const skipSuggestion = Boolean(opts.skipSuggestion || prefillBody);
    if (!state.activeId || !refs.replyPanel) return;
    const thread = state.lookup.get(state.activeId);
    if (!thread) return;
    closeTaskPanel(true);
    clearPendingSuggestedAction(thread.threadId);
    const sameSource = replyState.lastSourceId === thread.threadId;
    replyState.open = true;
    replyState.status = 'idle';
    replyState.error = '';
    replyState.lastSourceId = thread.threadId;
    if (!preserveValues || !sameSource) {
      const draft = buildReplyDraft(thread);
      const baseBody = draft.body || '';
      if (prefillBody) {
        draft.body = mergeReplyBody(baseBody, prefillBody);
      }
      replyState.values = draft;
      replyState.baseBody = prefillBody ? baseBody : draft.body || '';
      replyState.suggested = '';
      replyState.suggesting = false;
    }
    renderReplyPanel();
    if (refs.replyBody) refs.replyBody.focus();
    if (!preserveValues || !sameSource) {
      if (skipSuggestion) return;
      requestReplyDraft(thread.threadId);
    }
  }

  function closeReplyPanel(resetValues) {
    if (!refs.replyPanel) return;
    replyState.open = false;
    if (resetValues) {
      replyState.status = 'idle';
      replyState.error = '';
      replyState.values = { to: '', subject: '', body: '' };
      replyState.baseBody = '';
      replyState.suggested = '';
      replyState.suggesting = false;
    }
    renderReplyPanel();
  }

  function syncReplyValues() {
    if (!refs.replyBody) return;
    replyState.values.body = refs.replyBody.value;
    replyState.error = '';
    if (refs.replyError) refs.replyError.classList.add('hidden');
  }

  function buildReplyDraft(thread) {
    const fallbackParticipant = Array.isArray(thread.participants) && thread.participants.length
      ? thread.participants[0]
      : '';
    const to = thread.from || fallbackParticipant || '';
    const subject = formatReplySubject(thread.subject || '');
    const greeting = buildReplyGreeting(to);
    return {
      to: to || 'Unknown sender',
      subject,
      body: greeting
    };
  }

  function buildReplyGreeting(fromLine) {
    const name = extractFirstName(fromLine);
    if (!name) return '';
    return `Hi ${name},\n\n`;
  }

  function extractFirstName(fromLine) {
    const raw = (fromLine || '').replace(/<[^>]+>/g, '').replace(/\"/g, '').trim();
    if (!raw || /@/.test(raw)) return '';
    const first = raw.split(/\s+/)[0];
    return first && /@/.test(first) ? '' : first;
  }

  function formatReplySubject(subject) {
    const trimmed = (subject || '').trim();
    if (!trimmed) return 'Re:';
    if (/^re:/i.test(trimmed)) return trimmed;
    return `Re: ${trimmed}`;
  }

  async function requestReplyDraft(threadId) {
    if (!threadId || replyState.suggesting) return;
    replyState.suggesting = true;
    renderReplyPanel();
    try {
      const resp = await fetch('/secretary/reply-draft', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({ threadId })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        throw new Error(data?.error || 'Unable to draft a reply.');
      }
      const suggestion = typeof data?.body === 'string' ? data.body.trim() : '';
      const shouldApply = Boolean(data?.suggested && suggestion);
      if (!shouldApply) return;
      if (!replyState.open || replyState.lastSourceId !== threadId) return;
      const current = (replyState.values.body || '').trim();
      const base = (replyState.baseBody || '').trim();
      if (current && current !== base) return;
      replyState.values.body = mergeReplyBody(replyState.baseBody || '', suggestion);
      replyState.suggested = suggestion;
      renderReplyPanel();
    } catch (err) {
      console.error('Reply draft request failed', err);
    } finally {
      replyState.suggesting = false;
      renderReplyPanel();
    }
  }

  async function requestReplyIntentDraft(threadId, userText) {
    if (!threadId || replyState.suggesting) return;
    const prompt = typeof userText === 'string' ? userText.trim() : '';
    if (!prompt) return;
    replyState.suggesting = true;
    renderReplyPanel();
    let shouldFallback = false;
    try {
      const resp = await fetch('/secretary/reply-intent-draft', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({ threadId, text: prompt })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        if (resp.status === 403 && data?.error) {
          replyState.error = data.error;
          replyState.status = 'error';
          return;
        }
        console.warn('Reply intent draft unavailable', { status: resp.status, error: data?.error });
        shouldFallback = true;
        return;
      }
      const suggestion = typeof data?.body === 'string' ? data.body.trim() : '';
      const shouldApply = Boolean(data?.suggested && suggestion);
      if (!shouldApply) {
        shouldFallback = true;
        return;
      }
      if (!replyState.open || replyState.lastSourceId !== threadId) return;
      const current = (replyState.values.body || '').trim();
      const base = (replyState.baseBody || '').trim();
      if (current && current !== base) return;
      replyState.values.body = mergeReplyBody(replyState.baseBody || '', suggestion);
      replyState.suggested = suggestion;
      renderReplyPanel();
    } catch (err) {
      console.warn('Reply intent draft failed', err);
      shouldFallback = true;
    } finally {
      replyState.suggesting = false;
      renderReplyPanel();
    }
    if (shouldFallback) {
      replyState.error = '';
      replyState.status = 'idle';
      renderReplyPanel();
      await requestReplyDraft(threadId);
    }
  }

  function mergeReplyBody(base, suggestion) {
    const baseText = String(base || '').replace(/\s+$/, '');
    const suggestionText = String(suggestion || '').trim();
    if (!suggestionText) return baseText;
    if (!baseText) return suggestionText;
    return `${baseText}\n\n${suggestionText}`;
  }

  function openReplyPanelFromPrompt(userText) {
    const prompt = typeof userText === 'string' ? userText.trim() : '';
    const useGuided = shouldUseGuidedReply(prompt);
    openReplyPanel({ skipSuggestion: useGuided });
    if (useGuided && state.activeId) {
      requestReplyIntentDraft(state.activeId, prompt);
    }
  }

  function extractReplyBodyFromPrompt(rawText) {
    const raw = typeof rawText === 'string' ? rawText.trim() : '';
    if (!raw) return '';

    const quoted = extractQuotedReply(raw);
    if (quoted) return normalizeReplyPrefill(quoted);

    const patterns = [
      /(?:^|\b)(?:reply|respond|send a reply|send a response|email back|write back)(?:\s+to\s+[^,.:;-]+)?\s*(?:with|saying|and say|and tell|and ask|to say|to tell|to ask|that|:|-|,)\s*(.+)$/i,
      /(?:^|\b)(?:tell|let)\s+(?:him|her|them)\s+(?:know\s+)?(?:that\s+)?(.+)$/i,
      /(?:^|\b)ask\s+(?:him|her|them)?\s*(.+)$/i,
      /(?:^|\b)say\s+(?:that\s+)?(.+)$/i
    ];

    for (const pattern of patterns) {
      const match = raw.match(pattern);
      if (match && match[1]) {
        const candidate = normalizeReplyPrefill(match[1]);
        if (candidate) return candidate;
      }
    }

    const fallback = extractReplyAfterLead(raw);
    return fallback ? normalizeReplyPrefill(fallback) : '';
  }

  function extractQuotedReply(text) {
    const doubleMatch = text.match(/(?:^|\s)"([^"]{2,})"(?:\s|$)/);
    if (doubleMatch?.[1]) return doubleMatch[1].trim();
    const singleMatch = text.match(/(?:^|\s)'([^']{2,})'(?:\s|$)/);
    return singleMatch?.[1]?.trim() || '';
  }

  function extractReplyAfterLead(text) {
    const cleaned = text.replace(/\s+/g, ' ').trim();
    if (!cleaned) return '';
    const match = cleaned.match(/^\s*(?:please\s+)?(?:reply back|reply|respond|send a reply|send a response|email back|write back)\s+(?!to\b)(.+)$/i);
    let remainder = match?.[1]?.trim() || '';
    remainder = remainder.replace(/^back\s+/i, '');
    remainder = remainder.replace(/^and\s+/i, '');
    remainder = remainder.replace(/^(?:say|tell|ask)\s+/i, '');
    return remainder.trim();
  }

  function normalizeReplyPrefill(text) {
    let body = String(text || '').trim();
    if (!body) return '';
    body = body.replace(/^[\s"'`]+|[\s"'`]+$/g, '');
    body = body.replace(/^(?:that|please)\s+/i, '').trim();
    if (!body) return '';
    const lower = body.toLowerCase();
    const discard = new Set(['please', 'pls', 'plz', 'thanks', 'thank you', 'ok', 'okay']);
    if (discard.has(lower)) return '';
    if (/^if\s+/i.test(body) && !/[?]$/.test(body)) {
      const clause = body.replace(/^if\s+/i, '').trim();
      if (clause) return `Could you let me know if ${clause}?`;
    }
    if (!/[.!?]$/.test(body)) {
      body = `${body}.`;
    }
    return body;
  }

  function shouldUseGuidedReply(rawText) {
    const raw = typeof rawText === 'string' ? rawText.trim() : '';
    if (!raw) return false;
    const cleaned = raw.replace(/[.!?]/g, '').trim().toLowerCase();
    const shortCommands = new Set([
      'reply',
      'reply back',
      'reply to',
      'respond',
      'respond to',
      'send a reply',
      'send a response',
      'draft a reply',
      'write back',
      'email back',
      'answer them'
    ]);
    if (shortCommands.has(cleaned)) return false;
    if (extractReplyBodyFromPrompt(raw)) return true;
    if (/["']/.test(raw)) return true;
    if (/(tell|say|ask|let)\s+(him|her|them)\b/i.test(raw)) return true;
    if (/(reply|respond|email back|write back)\b/i.test(raw) && raw.length > 20) return true;
    return false;
  }

  async function submitReply(event) {
    if (event instanceof Event && typeof event.preventDefault === 'function') {
      event.preventDefault();
    }
    if (!state.activeId || !refs.replySubmit) return { ok: false, error: 'No email selected.' };
    const threadId = state.activeId;
    const body = (replyState.values.body || '').trim();
    if (!body) {
      replyState.error = 'Write a reply before sending.';
      replyState.status = 'error';
      renderReplyPanel();
      return { ok: false, error: replyState.error };
    }
    replyState.error = '';
    replyState.status = 'sending';
    renderReplyPanel();
    let result = { ok: false, error: '' };
    try {
      const resp = await fetch('/secretary/action/execute', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        credentials: 'same-origin',
        body: JSON.stringify({
          threadId,
          actionType: 'reply',
          draft: { body }
        })
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        throw new Error(data?.error || 'Unable to send that reply.');
      }
      if (data.flow) updateActionFlow(threadId, data.flow);
      if (Array.isArray(data.timeline)) {
        syncTimelineFromServer(threadId, data.timeline);
      }
      if (replyState.open && replyState.lastSourceId === threadId) {
        closeReplyPanel(true);
      }
      await waitForAssistantSettled(threadId);
      if (state.activeId === threadId) {
        promptArchiveAfterReply({ threadId });
      }
      result = { ok: true };
    } catch (err) {
      replyState.status = 'error';
      replyState.error = err instanceof Error ? err.message : 'Unable to send that reply.';
      result = { ok: false, error: replyState.error };
    } finally {
      if (replyState.open) {
        replyState.status = replyState.status === 'error' ? 'error' : 'idle';
        renderReplyPanel();
      }
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
    if (!state.totalInboxCountKnown) {
      refs.count.textContent = 'Inbox count unavailable';
      updateLoadMoreButtons();
      return;
    }
    const total = state.totalInboxCount || getLoadedCount();
    let label = '0 emails in inbox';
    if (total > 0) {
      label = `${total} email${total === 1 ? '' : 's'} in inbox`;
    }
    refs.count.textContent = label;
    updateLoadMoreButtons();
  }

  function updateQueuePill() {
    if (!refs.queuePill) return;
    if (!state.totalInboxCountKnown) {
      refs.queuePill.textContent = 'Sync to load count';
      return;
    }
    const total = state.totalInboxCount || getLoadedCount();
    if (total > 0) {
      refs.queuePill.textContent = `${total} in inbox`;
      return;
    }
    refs.queuePill.textContent = 'All done';
  }

  function updatePriorityPill() {
    if (!refs.priorityPill) return;
    if (state.priorityLoading) {
      refs.priorityPill.textContent = 'Prioritizing…';
      return;
    }
    const remaining = state.priority.length;
    if (remaining) {
      refs.priorityPill.textContent = `Found ${remaining} emails to address`;
      return;
    }
    refs.priorityPill.textContent = state.needs.length ? 'All clear' : 'No priority mail';
  }

  function updatePriorityProgress() {
    if (!refs.priorityProgress) return;
    const progress = state.priorityProgress;
    if (!progress || !progress.totalCount) {
      refs.priorityProgress.textContent = '';
      return;
    }
    const prioritized = Math.min(progress.prioritizedCount, progress.totalCount);
    if (prioritized >= progress.totalCount) {
      if (state.priorityLoading) {
        setPriorityLoading(false);
      }
      if (state.prioritySyncStartAt) {
        state.prioritySyncStartAt = 0;
        clearPrioritySyncStart();
      }
    }
    const needsSpinner = prioritized < progress.totalCount;
    const spinner = needsSpinner ? '<span class="priority-mini-spinner" aria-hidden="true"></span>' : '';
    refs.priorityProgress.innerHTML = `${spinner}<span>Reviewed ${prioritized} of ${progress.totalCount}</span>`;
  }

  function updateLoadMoreButtons() {
    const canReview = state.needs.length > 0;
    if (refs.loadMoreHead) {
      refs.loadMoreHead.classList.remove('hidden');
      refs.loadMoreHead.disabled = !canReview;
      refs.loadMoreHead.textContent = 'Review latest emails';
    }
    if (refs.loadMoreEmpty) {
      refs.loadMoreEmpty.classList.remove('hidden');
      refs.loadMoreEmpty.disabled = !canReview;
      refs.loadMoreEmpty.textContent = 'Review latest emails';
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

  function markUserEngaged() {
    if (state.prioritySwitchInProgress) return;
    state.userEngagedInbox = true;
    state.autoSelectBlocked = false;
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
    renderThreadList(refs.priorityList, 'priority');
    renderThreadList(refs.reviewList, 'queue');
  }

  function setPriorityLoading(isLoading) {
    if (state.priorityLoading === isLoading) return;
    state.priorityLoading = isLoading;
    if (refs.priorityList) {
      refs.priorityList.classList.toggle('is-loading', isLoading);
    }
    updatePriorityPill();
    renderThreadList(refs.priorityList, 'priority');
  }

  function startPriorityPolling() {
    if (state.priorityPolling) return;
    state.priorityPolling = true;
    schedulePriorityPoll(0);
  }

  function schedulePriorityPoll(delay = PRIORITY_POLL_INTERVAL_MS) {
    if (!state.priorityPolling) return;
    if (state.priorityPollTimer) {
      window.clearTimeout(state.priorityPollTimer);
    }
    state.priorityPollTimer = window.setTimeout(() => {
      void pollPriority();
    }, delay);
  }

  async function pollPriority() {
    try {
      const resp = await fetch('/api/priority', {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        credentials: 'same-origin'
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) {
        throw new Error(data?.error || 'Unable to load priority queue.');
      }

      const batchAt = parseTimestamp(data?.batchFinishedAt);
      const hasNewBatch = batchAt && batchAt > state.priorityLastBatchAt;
      const shouldApply = hasNewBatch ||
        (state.priorityLoading && batchAt && state.prioritySyncStartAt && batchAt >= state.prioritySyncStartAt);

    if (shouldApply) {
      if (hasNewBatch) {
        setPriorityLoading(true);
      }
      window.setTimeout(() => {
        applyPriorityPayload(data);
        if (batchAt) {
          state.priorityLastBatchAt = batchAt;
        }
        if (state.prioritySyncStartAt && batchAt && batchAt >= state.prioritySyncStartAt) {
          state.prioritySyncStartAt = 0;
          clearPrioritySyncStart();
        }
        setPriorityLoading(false);
      }, hasNewBatch ? PRIORITY_LOADING_MIN_MS : 0);
    } else if (data?.progress) {
      const progress = normalizePriorityProgress(data.progress);
      if (progress) {
        state.priorityProgress = progress;
        state.totalInboxCount = progress.totalCount;
        state.totalInboxCountKnown = true;
        updatePriorityProgress();
        updateHeaderCount();
        updateQueuePill();
        if (progress.prioritizedCount >= progress.totalCount) {
          setPriorityLoading(false);
          if (state.prioritySyncStartAt) {
            state.prioritySyncStartAt = 0;
            clearPrioritySyncStart();
          }
        }
      }
    }
    } catch (err) {
      logDebug('priority poll failed', err);
    } finally {
      schedulePriorityPoll();
    }
  }

  function applyPriorityPayload(payload) {
    const priorityItems = normalizePriorityItems(Array.isArray(payload?.priority) ? payload.priority : null);
    if (priorityItems) {
      state.prioritySource = 'server';
      state.priority = priorityItems.map(item => item.threadId);
      state.priorityMeta.clear();
      priorityItems.forEach(item => {
        state.priorityMeta.set(item.threadId, {
          score: item.score,
          reason: item.reason,
          reasonWeight: item.reasonWeight
        });
      });
    }

    const incoming = Array.isArray(payload?.threads)
      ? payload.threads.map(normalizeThread).filter(Boolean)
      : [];
    if (incoming.length) {
      appendThreads(incoming);
    }

    const progress = normalizePriorityProgress(payload?.progress);
    if (progress) {
      state.priorityProgress = progress;
      state.totalInboxCount = progress.totalCount;
      state.totalInboxCountKnown = true;
    }

    updateHeaderCount();
    updateProgress();
    updatePriorityPill();
    updatePriorityProgress();
    updateQueuePill();
    updateDrawerLists();
    syncPriorityReadiness();
  }

  function getThreadListIds(variant) {
    if (variant === 'priority') return state.priority;
    if (variant === 'queue') return sortThreadIdsByReceivedAt(state.needs);
    return state.needs;
  }

  function renderThreadList(listEl, variant = 'drawer') {
    if (!listEl) return;
    listEl.innerHTML = '';

    const ids = getThreadListIds(variant);
    if (variant === 'priority' && state.priorityLoading) {
      const li = document.createElement('li');
      li.className = 'queue-loading priority-loading';
      li.innerHTML = '<span class="priority-spinner" aria-hidden="true"></span><span>Prioritizing inbox…</span>';
      listEl.appendChild(li);
    }

    if (ids.length) {
      ids.forEach(id => appendThreadItem(listEl, id, variant));
    } else if (!state.priorityLoading || variant !== 'priority') {
      const li = document.createElement('li');
      li.className = variant === 'queue' ? 'queue-empty drawer-empty' : 'drawer-empty';
      if (variant === 'priority') {
        li.className = 'queue-empty priority-empty';
        li.textContent = 'No urgent emails right now.';
      } else {
        li.textContent = 'Nothing queued up.';
      }
      listEl.appendChild(li);
    }

    if (variant === 'queue' && state.hasMore) {
      const li = document.createElement('li');
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = buildThreadClass(variant, { loadMore: true });
      btn.textContent = state.loadingMore ? 'Loading…' : 'Review latest emails';
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
    const receivedAt = thread.receivedAt ? formatQueueTimestamp(thread.receivedAt) : '';
    const timeHtml = receivedAt ? `<span class="queue-item-time">${htmlEscape(receivedAt)}</span>` : '';
    const reason = variant === 'priority' ? getPriorityReason(threadId) : '';
    const reasonHtml = reason ? `<span class="priority-reason">${htmlEscape(reason)}</span>` : '';
    btn.innerHTML = `<div class="queue-item-header"><strong>${htmlEscape(thread.from || 'Unknown sender')}</strong>${timeHtml}</div><span>${htmlEscape(thread.subject || '(no subject)')}</span>${reasonHtml}`;
    li.appendChild(btn);
    listEl.appendChild(li);
  }

  function buildThreadClass(variant, options = {}) {
    const { loadMore = false, active = false } = options;
    const base = variant === 'queue' || variant === 'priority' ? 'queue-item' : 'drawer-thread';
    const classes = [base];
    if (variant === 'priority') classes.push('priority-item');
    if (loadMore) classes.push('load-more');
    if (active) classes.push('active');
    return classes.join(' ');
  }

  function rebuildPriorityQueue() {
    if (state.prioritySource === 'server') {
      const needsSet = new Set(state.needs);
      state.priority = state.priority.filter(id => needsSet.has(id) && !reviewedIds.has(id));
      for (const key of state.priorityMeta.keys()) {
        if (!needsSet.has(key) || reviewedIds.has(key)) {
          state.priorityMeta.delete(key);
        }
      }
      syncPriorityReadiness();
      return;
    }
    const scored = [];
    state.priorityMeta.clear();
    state.needs.forEach(threadId => {
      if (reviewedIds.has(threadId)) return;
      const thread = state.lookup.get(threadId);
      if (!thread) return;
      const evaluation = scoreThreadPriority(thread);
      state.priorityMeta.set(threadId, evaluation);
      if (evaluation.score >= PRIORITY_MIN_SCORE) {
        scored.push({
          threadId,
          score: evaluation.score,
          reasonWeight: evaluation.reasonWeight,
          receivedAt: thread.receivedAt || ''
        });
      }
    });
    scored.sort((a, b) => {
      if (b.score !== a.score) return b.score - a.score;
      if (b.reasonWeight !== a.reasonWeight) return b.reasonWeight - a.reasonWeight;
      const aTime = a.receivedAt ? new Date(a.receivedAt).getTime() : 0;
      const bTime = b.receivedAt ? new Date(b.receivedAt).getTime() : 0;
      return bTime - aTime;
    });
    state.priority = scored.slice(0, PRIORITY_LIMIT).map(item => item.threadId);
    syncPriorityReadiness();
  }

  function syncPriorityReadiness() {
    const wasReady = state.priorityReady;
    const nowReady = state.priority.length > 0;
    state.priorityReady = nowReady;
    if (!nowReady) {
      state.prioritySwitchPending = false;
      return;
    }
    state.autoSelectBlocked = false;
    if (!wasReady) {
      maybeSwitchToPriority('ready');
    }
  }

  function isPriorityWaiting() {
    if (state.priority.length > 0) return false;
    const progress = state.priorityProgress;
    if (state.priorityLoading) return true;
    if (progress && progress.totalCount === 0) return true;
    if (progress && progress.prioritizedCount < progress.totalCount) return true;
    return false;
  }

  function isPriorityThread(threadId) {
    if (!threadId) return false;
    return state.priority.includes(threadId);
  }

  function getNextPriorityCandidate() {
    if (!state.priority.length) return '';
    const needsSet = new Set(state.needs);
    for (let i = 0; i < state.priority.length; i += 1) {
      const candidate = state.priority[i];
      if (!needsSet.has(candidate)) continue;
      if (!reviewedIds.has(candidate)) return candidate;
    }
    return '';
  }

  function getPriorityFirstCandidate() {
    const priorityCandidate = getNextPriorityCandidate();
    if (priorityCandidate) return priorityCandidate;
    return getNextReviewCandidate();
  }

  function maybeAnnouncePriorityWaiting() {
    if (!state.activeId) return;
    if (state.priorityWaitingNotified) return;
    if (state.userEngagedInbox) return;
    if (state.priority.length > 0) return;
    const progress = state.priorityProgress;
    const prioritizing = state.priorityLoading || (progress && progress.prioritizedCount < progress.totalCount);
    if (!prioritizing) return;
    state.priorityWaitingNotified = true;
    enqueueAssistantMessage(
      state.activeId,
      'I\'m prioritizing your inbox to find emails that need your attention. You can wait here, or jump into your full inbox and start reviewing anytime.'
    );
  }

  function maybeSwitchToPriority(source) {
    if (!state.priority.length) return;
    if (state.prioritySwitchInProgress) return;
    if (state.activeId && isPriorityThread(state.activeId)) {
      state.prioritySwitchPending = false;
      return;
    }
    if (!state.activeId && !state.userEngagedInbox) {
      const nextPriority = getNextPriorityCandidate();
      if (!nextPriority) return;
      state.prioritySwitchInProgress = true;
      state.pendingPriorityAnnouncement = 'Priority emails are ready to review. Starting with the top priority email.';
      setEmptyState('Priority emails are ready to review. Starting with the top priority email.');
      window.setTimeout(() => {
        setQueueSectionExpanded(refs.priorityQueue, refs.priorityToggle, true);
        setQueueSectionExpanded(refs.reviewQueue, refs.queueToggle, false);
        setActiveThread(nextPriority);
      }, 0);
      state.priorityReadyAnnounced = true;
      state.prioritySwitchPending = false;
      state.priorityWaitingNotified = false;
      state.prioritySwitchInProgress = false;
      return;
    }
    if (state.userEngagedInbox) {
      state.prioritySwitchPending = true;
      return;
    }
    const nextPriority = getNextPriorityCandidate();
    if (!nextPriority) return;
    const announce = state.priorityWaitingNotified && !state.priorityReadyAnnounced;
    const originId = state.activeId;
    state.prioritySwitchInProgress = true;
    const finishSwitch = () => {
      state.priorityReadyAnnounced = true;
      state.prioritySwitchPending = false;
      state.priorityWaitingNotified = false;
      setQueueSectionExpanded(refs.priorityQueue, refs.priorityToggle, true);
      setQueueSectionExpanded(refs.reviewQueue, refs.queueToggle, false);
      setActiveThread(nextPriority);
      state.prioritySwitchInProgress = false;
    };
    if (announce && originId) {
      enqueueAssistantMessage(originId, 'Priority emails are ready to review. Starting with the top priority email.')
        .then(() => {
          if (state.activeId === originId) {
            finishSwitch();
          } else {
            state.prioritySwitchInProgress = false;
          }
        });
    } else {
      finishSwitch();
    }
  }

  function maybeSwitchToPriorityAfterResolve() {
    if (!state.prioritySwitchPending) return false;
    const nextPriority = getNextPriorityCandidate();
    state.prioritySwitchPending = false;
    if (!nextPriority) return false;
    state.prioritySwitchInProgress = true;
    const originId = state.activeId;
    const announce = originId
      ? enqueueAssistantMessage(originId, 'Priority emails are ready to review. Starting with the top priority email.')
      : Promise.resolve();
    announce.then(() => {
      setQueueSectionExpanded(refs.priorityQueue, refs.priorityToggle, true);
      setQueueSectionExpanded(refs.reviewQueue, refs.queueToggle, false);
      setActiveThread(nextPriority);
      state.priorityReadyAnnounced = true;
      state.prioritySwitchInProgress = false;
    });
    return true;
  }

  function markUserEngagedIfInbox(threadId = state.activeId) {
    if (!threadId) return;
    if (isPriorityThread(threadId)) return;
    markUserEngaged();
    if (state.priorityReady && state.priority.length) {
      state.prioritySwitchPending = true;
    }
  }

  function scoreThreadPriority(thread) {
    const signals = [];
    let score = 0;
    const text = buildPriorityText(thread);
    const nextStep = typeof thread.nextStep === 'string' ? thread.nextStep.trim() : '';

    if (requiresAction(nextStep)) {
      score += addPrioritySignal(signals, 'Action needed', 2);
    }

    if (thread?.actionFlow?.actionType === 'open_link') {
      score += addPrioritySignal(signals, 'Action required', 4);
    }

    const dueSignal = applyDueSignal(text, signals);
    score += dueSignal;

    if (PRIORITY_PATTERNS.urgent.test(text)) {
      score += addPrioritySignal(signals, 'Time-sensitive', 2);
    }
    if (PRIORITY_PATTERNS.security.test(text)) {
      score += addPrioritySignal(signals, 'Account alert', 4);
    }
    if (PRIORITY_PATTERNS.payment.test(text)) {
      score += addPrioritySignal(signals, 'Payment issue', 2);
    }
    if (PRIORITY_PATTERNS.approval.test(text)) {
      score += addPrioritySignal(signals, 'Needs approval', 2);
    }
    if (PRIORITY_PATTERNS.scheduling.test(text)) {
      score += addPrioritySignal(signals, 'Scheduling', 1);
    }

    const categoryBoost = categoryPriorityBoost(thread.category);
    score += categoryBoost.score;
    if (categoryBoost.label) {
      addPrioritySignal(signals, categoryBoost.label, categoryBoost.score);
    }

    const normalizedScore = Math.max(0, score);
    const reason = pickPriorityReason(signals);
    return {
      score: normalizedScore,
      reason: reason?.label || 'Needs attention',
      reasonWeight: reason?.weight || 0
    };
  }

  function buildPriorityText(thread) {
    return [thread.nextStep, thread.summary, thread.headline, thread.subject].filter(Boolean).join(' ');
  }

  function requiresAction(nextStep) {
    const text = (nextStep || '').toLowerCase();
    if (!text) return false;
    return !/(no action|fyi|none|no need|no response needed)/i.test(text);
  }

  function addPrioritySignal(signals, label, weight) {
    if (!label || !weight) return 0;
    signals.push({ label, weight });
    return weight;
  }

  function applyDueSignal(text, signals) {
    const raw = extractDateFromText(text);
    const due = parsePriorityDate(raw);
    if (!due) return 0;
    const days = diffCalendarDays(new Date(), due);
    if (days < 0) return addPrioritySignal(signals, 'Overdue', 4);
    if (days === 0) return addPrioritySignal(signals, 'Due today', 3);
    if (days === 1) return addPrioritySignal(signals, 'Due tomorrow', 3);
    if (days <= 3) return addPrioritySignal(signals, 'Due soon', 2);
    if (days <= 7) return addPrioritySignal(signals, 'Due this week', 1);
    return 0;
  }

  function parsePriorityDate(raw) {
    if (!raw) return null;
    const iso = raw.length === 10 ? `${raw}T00:00:00` : raw;
    const parsed = new Date(iso);
    if (Number.isNaN(parsed.getTime())) return null;
    return parsed;
  }

  function diffCalendarDays(from, to) {
    const start = new Date(from.getFullYear(), from.getMonth(), from.getDate());
    const end = new Date(to.getFullYear(), to.getMonth(), to.getDate());
    return Math.round((end.getTime() - start.getTime()) / (24 * 60 * 60 * 1000));
  }

  function categoryPriorityBoost(category) {
    const label = typeof category === 'string' ? category.toLowerCase() : '';
    if (label.includes('billing')) return { score: 2, label: 'Billing' };
    if (label.includes('personal request')) return { score: 2, label: 'Personal request' };
    if (label.includes('introduction')) return { score: 1, label: 'Introduction' };
    if (label.includes('personal event')) return { score: 1, label: 'Event planning' };
    if (label.includes('catch up')) return { score: 0, label: '' };
    if (label.includes('marketing') || label.includes('promotion')) return { score: -3, label: '' };
    if (label.includes('editorial') || label.includes('writing')) return { score: -2, label: '' };
    if (label.includes('fyi')) return { score: -1, label: '' };
    return { score: 0, label: '' };
  }

  function pickPriorityReason(signals) {
    if (!signals.length) return null;
    const sorted = signals.slice().sort((a, b) => b.weight - a.weight);
    return sorted[0] || null;
  }

  function getPriorityReason(threadId) {
    const meta = state.priorityMeta.get(threadId);
    if (!meta) return 'Needs attention';
    return meta.reason || 'Needs attention';
  }

  function appendThreads(items) {
    const added = [];
    items.forEach(thread => {
      const normalized = normalizeThread(thread);
      if (!normalized || !normalized.threadId) return;
      if (state.lookup.has(normalized.threadId)) return;
      state.lookup.set(normalized.threadId, normalized);
      state.positions.set(normalized.threadId, state.positions.size);
      state.needs.push(normalized.threadId);
      state.serverTimelines.set(normalized.threadId, normalized.timeline || []);
      if (normalized.actionFlow) state.actionFlows.set(normalized.threadId, normalized.actionFlow);
      added.push(normalized.threadId);
    });
    if (added.length) {
      state.totalLoaded = state.positions.size;
      rebuildPriorityQueue();
      updatePriorityPill();
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
      convo: typeof raw.convo === 'string' ? raw.convo : '',
      participants: Array.isArray(raw.participants) ? raw.participants.map(p => String(p || '').trim()).filter(Boolean) : [],
      unsubscribe: normalizeUnsubscribe(raw.unsubscribe),
      actionFlow: normalizeActionFlow(raw.actionFlow),
      timeline: Array.isArray(raw.timeline) ? raw.timeline.map(normalizeTimelineMessage).filter(Boolean) : []
    };
  }

  function normalizePriorityItems(raw) {
    if (!Array.isArray(raw)) return [];
    return raw
      .map(item => {
        if (!item || typeof item !== 'object') return null;
        const threadId = typeof item.threadId === 'string' ? item.threadId.trim() : '';
        if (!threadId) return null;
        const reason = typeof item.reason === 'string' ? item.reason.trim() : '';
        const score = typeof item.score === 'number' && Number.isFinite(item.score) ? item.score : 0;
        const reasonWeight = typeof item.reasonWeight === 'number' && Number.isFinite(item.reasonWeight)
          ? item.reasonWeight
          : 0;
        return {
          threadId,
          reason: reason || 'Needs attention',
          score,
          reasonWeight
        };
      })
      .filter(Boolean);
  }

  function normalizeActionFlow(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const allowedStates = new Set(['suggested', 'draft_ready', 'editing', 'executing', 'completed', 'failed']);
    const allowedTypes = new Set(['archive', 'create_task', 'more_info', 'skip', 'open_link', 'external_action', 'reply', 'unsubscribe']);
    const actionType = typeof raw.actionType === 'string' && allowedTypes.has(raw.actionType) ? raw.actionType : '';
    const state = typeof raw.state === 'string' && allowedStates.has(raw.state) ? raw.state : 'suggested';
    return actionType ? { ...raw, actionType, state } : null;
  }

  function normalizeUnsubscribe(raw) {
    if (!raw || typeof raw !== 'object') {
      return { supported: false, oneClick: false, bulk: false };
    }
    const supported = Boolean(raw.supported);
    const oneClick = Boolean(raw.oneClick);
    const bulk = Boolean(raw.bulk);
    return { supported, oneClick, bulk };
  }

  function normalizeTimelineMessage(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const messageType = typeof raw.type === 'string' ? raw.type : '';
    const content = typeof raw.content === 'string' ? raw.content : '';
    if (!raw.threadId) {
      logDebug('normalizeTimelineMessage missing threadId', raw);
    }
    logDebug('normalizeTimelineMessage', { id: raw.id, messageType, contentPreview: content?.slice?.(0, 80) });
    return {
      id: typeof raw.id === 'string' ? raw.id : '',
      threadId: typeof raw.threadId === 'string' ? raw.threadId : '',
      type: 'transcript',
      messageType,
      content,
      payload: raw.payload || {},
      createdAt: typeof raw.createdAt === 'string' ? raw.createdAt : ''
    };
  }

  function normalizeSuggestedAction(value) {
    const val = typeof value === 'string' ? value.trim().toLowerCase() : '';
    if (val === 'archive' || val === 'more_info' || val === 'create_task' || val === 'skip' || val === 'reply' || val === 'unsubscribe' || val === 'open_link' || val === 'external_action') return val;
    return '';
  }

  function normalizeActionType(value) {
    const val = typeof value === 'string' ? value.trim().toLowerCase() : '';
    if (val === 'open link') return 'open_link';
    if (val === 'archive' || val === 'more_info' || val === 'create_task' || val === 'skip' || val === 'open_link' || val === 'external_action' || val === 'reply' || val === 'unsubscribe') {
      return val;
    }
    return '';
  }

  function normalizeSuggestedActionsPayload(payload, fallbackPrompt) {
    const raw = payload?.suggestedActions || payload?.suggested_actions;
    if (Array.isArray(raw)) {
      return raw.map(entry => {
        const actionType = normalizeActionType(entry?.actionType || entry?.action_type);
        if (!actionType) return null;
        return {
          actionType,
          userFacingPrompt: typeof entry?.userFacingPrompt === 'string' ? entry.userFacingPrompt.trim() : '',
          externalAction: entry?.externalAction || entry?.external_action || null
        };
      }).filter(Boolean);
    }
    const actionType = normalizeActionType(payload?.actionType);
    if (!actionType) return [];
    return [{
      actionType,
      userFacingPrompt: typeof fallbackPrompt === 'string' ? fallbackPrompt.trim() : '',
      externalAction: payload?.externalAction || payload?.external_action || null
    }];
  }

  function guessSuggestedAction(thread) {
    const next = normalizeSuggestedAction(actionFromNextStep(thread?.nextStep));
    if (next) return next;
    const summary = `${thread?.summary || thread?.headline || thread?.subject || ''}`.toLowerCase();
    if (thread?.unsubscribe?.supported && (thread?.category || '').toLowerCase().startsWith('marketing')) {
      return 'unsubscribe';
    }
    if (summary.includes('reply') || summary.includes('respond')) {
      return 'reply';
    }
    if (summary.includes('deadline') || summary.includes('follow up') || summary.includes('follow-up') || summary.includes('due')) {
      return 'create_task';
    }
    if (summary.includes('fyi') || summary.includes('newsletter')) return 'skip';
    return 'more_info';
  }

  function normalizeDraftPayload(raw) {
    const draft = raw && typeof raw === 'object' ? raw : {};
    const title = typeof draft.title === 'string' ? draft.title.trim() : '';
    const notes = typeof draft.notes === 'string' ? draft.notes.trim() : '';
    const dueDate = typeof draft.dueDate === 'string' ? draft.dueDate.trim() : '';
    return { title, notes, dueDate };
  }

  function normalizeExternalActionPayload(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const payload = raw;
    const steps = typeof payload.steps === 'string' ? payload.steps.trim() : '';
    const links = Array.isArray(payload.links)
      ? payload.links
        .map(normalizeExternalLink)
        .filter(Boolean)
        .slice(0, 3)
      : [];
    if (!steps) return null;
    return { steps, links };
  }

  function normalizeExternalLink(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const label = typeof raw.label === 'string' ? raw.label.trim() : '';
    const url = typeof raw.url === 'string' ? raw.url.trim() : '';
    if (!url) return null;
    return { label: label || url, url };
  }

  function buildLinksKey(links) {
    if (!Array.isArray(links)) return '';
    return links.map(link => String(link?.url || '').trim()).filter(Boolean).join('|');
  }

  function ordinalLabel(value) {
    const n = Number(value);
    if (!Number.isFinite(n) || n <= 0) return 'next';
    if (n === 1) return 'first';
    if (n === 2) return 'second';
    if (n === 3) return 'third';
    if (n === 4) return 'fourth';
    if (n === 5) return 'fifth';
    return `${n}th`;
  }

  function setSuggestedLinks(threadId, payload) {
    if (!threadId) return;
    const links = payload?.links || [];
    if (links.length) {
      state.suggestedLinksByThread.set(threadId, payload);
      const key = buildLinksKey(links);
      const existing = state.openLinkProgress.get(threadId);
      if (!existing || existing.key !== key) {
        state.openLinkProgress.set(threadId, { key, index: 0 });
      }
    } else {
      state.suggestedLinksByThread.delete(threadId);
      state.openLinkProgress.delete(threadId);
    }
    updateOpenLinkButton();
  }

  function getSuggestedLinks(threadId = state.activeId) {
    if (!threadId) return null;
    return state.suggestedLinksByThread.get(threadId) || null;
  }

  function getRemainingLinkCount(threadId = state.activeId) {
    const payload = getSuggestedLinks(threadId);
    const links = Array.isArray(payload?.links) ? payload.links : [];
    if (!links.length) return 0;
    const key = buildLinksKey(links);
    const progress = state.openLinkProgress.get(threadId);
    const index = progress && progress.key === key ? progress.index : 0;
    return Math.max(0, links.length - index);
  }

  function hasSuggestedLinks(threadId = state.activeId) {
    return getRemainingLinkCount(threadId) > 0;
  }

  function openLinkActionLabel(threadId = state.activeId) {
    const payload = getSuggestedLinks(threadId);
    const links = Array.isArray(payload?.links) ? payload.links : [];
    if (!links.length) return 'Open link';
    if (links.length === 1) return 'Open link';
    const key = buildLinksKey(links);
    const progress = state.openLinkProgress.get(threadId);
    const index = progress && progress.key === key ? progress.index : 0;
    const position = Math.min(index + 1, links.length || 1);
    return `Open ${ordinalLabel(position)} link`;
  }

  function updateOpenLinkButton(enabled = !refs.chatInput?.disabled) {
    if (!refs.openLinkBtn) return;
    const hasLinks = Boolean(state.activeId && hasSuggestedLinks(state.activeId));
    refs.openLinkBtn.disabled = !enabled || !state.activeId || !hasLinks;
    const label = openLinkActionLabel(state.activeId);
    const labelEl = refs.openLinkBtn.querySelector('span:not(.action-icon)');
    if (labelEl) labelEl.textContent = label;
  }

  function suggestedActionLabel(actionType) {
    if (actionType === 'archive') return 'Archive';
    if (actionType === 'create_task') return 'Draft task';
    if (actionType === 'more_info') return 'Tell me more';
    if (actionType === 'reply') return 'Reply';
    if (actionType === 'skip') return 'Skip';
    if (actionType === 'unsubscribe') return 'Unsubscribe';
    if (actionType === 'open_link') return 'Open link';
    if (actionType === 'external_action') return 'Confirm';
    return 'Do it';
  }

  function actionFromNextStep(nextStep) {
    const text = typeof nextStep === 'string' ? nextStep.toLowerCase() : '';
    if (!text) return '';
    if (text.includes('archive')) return 'archive';
    if (text.includes('unsubscribe') || text.includes('opt out')) return 'unsubscribe';
    if (text.includes('reply') || text.includes('respond')) return 'reply';
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

  function updateComposerPlaceholder(threadId = state.activeId) {
    if (!refs.chatInput) return;
    const isActive = threadId && threadId === state.activeId;
    const hasSuggested = isActive
      && (Boolean(getPendingSuggestedAction(threadId)) || isCreateConfirmationPending(threadId));
    refs.chatInput.placeholder = hasSuggested ? SUGGESTED_PLACEHOLDER : DEFAULT_PLACEHOLDER;
  }

  function ensureHistory(threadId) {
    if (!state.histories.has(threadId)) {
      state.histories.set(threadId, []);
    }
    hydrateThreadTimeline(threadId);
    logDebug('ensureHistory', threadId, {
      historyLength: state.histories.get(threadId)?.length || 0
    });
    return state.histories.get(threadId);
  }

  function hydrateThreadTimeline(threadId) {
    if (!threadId) return;
    const hasTimelineEntries = state.timeline.some(item => item.threadId === threadId && item.type === 'transcript');
    const serverItems = state.serverTimelines.get(threadId) || [];
    const alreadyHydrated = state.hydrated.has(threadId);
    const isHydrating = state.hydrating.has(threadId);
    logDebug('hydrateThreadTimeline', threadId, {
      hasTimelineEntries,
      serverCount: serverItems.length,
      alreadyHydrated,
      isHydrating
    });
    if (alreadyHydrated && (hasTimelineEntries || !serverItems.length)) {
      return;
    }
    const thread = state.lookup.get(threadId);
    if (!thread) return;
    if (serverItems.length) {
      mergeServerTimeline(threadId, serverItems);
      state.hydrated.add(threadId);
      return;
    }
    if (!isHydrating) {
      logDebug('hydrateThreadTimeline: fetching auto summary', threadId);
      fetchAutoSummary(threadId);
    }
  }

  function mergeServerTimeline(threadId, messages) {
    if (!Array.isArray(messages) || !messages.length) return;
    const ordered = messages.slice().sort((a, b) => {
      const aTime = new Date(a.createdAt || '').getTime();
      const bTime = new Date(b.createdAt || '').getTime();
      return aTime - bTime;
    });
    const shouldDefer = true;
    ordered.forEach(msg => {
      const id = typeof msg?.id === 'string' ? msg.id : '';
      if (id && state.timelineMessageIds.has(id)) return;
      if (id) state.timelineMessageIds.add(id);
      const resolvedType = typeof msg.messageType === 'string' && msg.messageType
        ? msg.messageType
        : (typeof msg.type === 'string' ? msg.type : '');
      const entry = {
        type: 'transcript',
        id,
        threadId,
        messageType: resolvedType,
        content: typeof msg.content === 'string' ? msg.content : '',
        payload: msg.payload || {},
        createdAt: msg.createdAt || ''
      };
      if (shouldDefer) {
        storePendingTranscript(threadId, entry);
      } else {
        applyTranscriptEntryEffects(threadId, entry);
        state.timeline.push(entry);
      }
    });
    logDebug('mergeServerTimeline', {
      threadId,
      added: ordered.length,
      totalTimeline: state.timeline.length,
      ids: ordered.map(m => m.id || '(no id)'),
      messageTypes: ordered.map(m => m.type)
    });
    if (threadId === state.activeId) {
      revealPendingTranscripts(threadId);
    }
  }

  function appendTurn(threadId, turn) {
    const history = ensureHistory(threadId);
    history.push(turn);
    state.timeline.push({ type: 'turn', threadId, turn });
    logDebug('appendTurn', { threadId, role: turn.role, timelineLength: state.timeline.length });
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

  function getPinnedDividerIndex(threadId, timeline) {
    if (!threadId || !timeline?.length) return -1;
    for (let i = timeline.length - 1; i >= 0; i--) {
      const entry = timeline[i];
      if (entry.type === 'divider' && entry.threadId === threadId) return i;
    }
    return -1;
  }

  function snapToPinnedDivider(threadId) {
    if (!refs.chatScroll || !refs.chatLog) return;
    window.requestAnimationFrame(() => {
      const pinned = refs.chatLog.querySelector('.chat-divider.is-pinned');
      if (!(pinned instanceof HTMLElement)) return;
      const top = pinned.offsetTop;
      refs.chatScroll.scrollTo({ top, behavior: 'smooth' });
    });
  }

  function renderChat(threadId = state.activeId, options = {}) {
    if (!refs.chatLog) return;
    const timeline = state.timeline.slice();
    const pinnedDividerIndex = getPinnedDividerIndex(threadId, timeline);
    logDebug('renderChat', {
      threadId,
      timelineCount: timeline.length,
      items: timeline.map(t => ({
        id: t.id,
        type: t.type,
        messageType: t.messageType,
        createdAt: t.createdAt,
        contentPreview: (t.content || '').slice(0, 80)
      }))
    });
    let markup = timeline.map((entry, index) => renderTimelineEntry(entry, {
      isPinned: index === pinnedDividerIndex
    })).join('');
    logDebug('renderChat markup preview', markup.slice(0, 200));
    if (!markup && !state.typing) {
      markup = chatPlaceholder();
    }
    if (state.typing && threadId === state.activeId) {
      markup += typingIndicatorHtml();
    }
    refs.chatLog.innerHTML = markup;
    const scrollMode = options.scrollMode || (state.snapThreadId === threadId ? 'divider' : 'bottom');
    if (refs.chatScroll) {
      if (scrollMode === 'divider') {
        snapToPinnedDivider(threadId);
      } else if (scrollMode === 'bottom') {
        refs.chatScroll.scrollTop = refs.chatScroll.scrollHeight;
      }
    }
    updateComposerPlaceholder(threadId);
  }

  function renderTimelineEntry(entry, options = {}) {
    if (entry.type === 'divider') {
      const isPinned = Boolean(options.isPinned);
      const label = htmlEscape(entry.label || 'New email thread');
      const sender = htmlEscape(entry.sender || '');
      const subject = htmlEscape(entry.subject || '');
      const timestamp = entry.receivedAt ? formatTimestamp(entry.receivedAt) : '';
      const meta = htmlEscape([sender, timestamp].filter(Boolean).join(' • '));
      const initials = initialsFromSender(entry.sender || entry.label || '');
      const link = entry.link ? escapeAttribute(entry.link) : '';
      const mustKnow = state.mustKnowByThread.get(entry.threadId) || '';
      const linkHtml = link
        ? `<a class="chat-divider-link" href="${link}" target="_blank" rel="noopener noreferrer">Open in Gmail ↗</a>`
        : '';
      const summaryHtml = mustKnow ? renderAssistantMarkdown(mustKnow) : '';
      const isPriority = state.priority.includes(entry.threadId);
      const dividerClass = isPriority ? 'chat-divider-card is-priority' : 'chat-divider-card is-inbox';
      const summaryClass = isPriority ? 'chat-divider-summary is-priority' : 'chat-divider-summary is-inbox';
      const dividerClassName = isPinned ? 'chat-divider is-pinned' : 'chat-divider';
      return `
          <div class="${dividerClassName}" data-thread-id="${escapeAttribute(entry.threadId)}">
            <div class="${dividerClass}">
              <div class="chat-divider-body">
                <div class="chat-divider-avatar" aria-hidden="true">${initials}</div>
                <div class="chat-divider-content">
                  <p class="chat-divider-meta">${meta}</p>
                  <p class="chat-divider-subject">${subject || label}</p>
                  ${linkHtml}
                </div>
              </div>
              ${mustKnow ? `<div class="${summaryClass}"><span class="chat-divider-summary-icon" aria-hidden="true">✨</span><div>${summaryHtml}</div></div>` : ''}
            </div>
          </div>
        `;
    }
    if (entry.type === 'turn') {
      const turn = entry.turn;
      if (turn.role === 'assistant') {
        return `<div class="chat-message assistant"><div class="assistant-avatar" aria-hidden="true">S</div><div class="chat-card">${renderAssistantMarkdown(turn.content)}</div></div>`;
      }
      return `<div class="chat-message user"><div class="chat-card">${renderPlainText(turn.content, { preserveLineBreaks: true })}</div></div>`;
    }
    if (entry.type === 'transcript') {
      return renderTranscriptEntry(entry);
    }
    return '';
  }

  function renderTranscriptEntry(entry) {
    const actionFlow = state.actionFlows.get(entry.threadId);
    if (!entry.messageType) {
      logDebug('renderTranscriptEntry missing messageType', entry);
    }
    if (entry.messageType === 'must_know') {
      return '';
    }
    if (entry.messageType === 'suggested_action') {
      const suggestedActions = normalizeSuggestedActionsPayload(entry?.payload, entry?.content);
      const primaryAction = suggestedActions[0];
      const primaryType = normalizeActionType(primaryAction?.actionType || entry?.payload?.actionType);
      if (primaryType === 'open_link') {
        const externalAction = normalizeExternalActionPayload(
          primaryAction?.externalAction
          || entry?.payload?.externalAction
          || entry?.payload?.external_action
        );
        const steps = externalAction?.steps || entry.content || '';
        const linksMarkup = renderExternalLinks(externalAction?.links || []);
        const hasLinks = Boolean(externalAction?.links?.length);
        const disableAll = actionFlow && actionFlow.state === 'completed' && actionFlow.actionType === 'skip';
        const disabled = disableAll || !hasLinks;
        const label = openLinkActionLabel(entry.threadId);
        return `<div class="chat-message assistant">
          <div class="assistant-avatar" aria-hidden="true">S</div>
          <div class="chat-card suggested-card">
            <p class="suggested-copy">${renderPlainText(steps, { preserveLineBreaks: true })}</p>
            ${linksMarkup}
          </div>
        </div>
        <div class="suggested-action-row" role="group" aria-label="Suggested action">
          <button type="button" class="suggested-btn" data-action="suggested-primary" data-thread-id="${escapeAttribute(entry.threadId)}" data-action-type="open_link" ${disabled ? 'disabled' : ''}>${label}<span class="suggested-enter-hint" aria-hidden="true">⏎</span></button>
        </div>`;
      }
      if (primaryType === 'external_action') {
        const externalAction = normalizeExternalActionPayload(
          primaryAction?.externalAction
          || entry?.payload?.externalAction
          || entry?.payload?.external_action
        );
        const steps = externalAction?.steps || entry.content || '';
        const disableAll = actionFlow && actionFlow.state === 'completed' && actionFlow.actionType === 'skip';
        const disabled = disableAll;
        return `<div class="chat-message assistant">
          <div class="assistant-avatar" aria-hidden="true">S</div>
          <div class="chat-card suggested-card">
            <p class="suggested-copy">${renderPlainText(steps, { preserveLineBreaks: true })}</p>
          </div>
        </div>
        <div class="suggested-action-row" role="group" aria-label="Suggested action">
          <button type="button" class="suggested-btn" data-action="suggested-primary" data-thread-id="${escapeAttribute(entry.threadId)}" data-action-type="external_action" ${disabled ? 'disabled' : ''}>Confirm<span class="suggested-enter-hint" aria-hidden="true">⏎</span></button>
        </div>`;
      }
      const thread = state.lookup.get(entry.threadId);
      const canUnsubscribe = Boolean(thread?.unsubscribe?.supported);
      const disableAll = actionFlow && actionFlow.state === 'completed' && actionFlow.actionType === 'skip';
      const agenticType = normalizeSuggestedAction(primaryAction?.actionType || primaryType) || '';
      const disabled = disableAll
        || (agenticType === 'unsubscribe' && !canUnsubscribe)
        || (agenticType === 'open_link' && !hasSuggestedLinks(entry.threadId));
      const label = agenticType === 'open_link'
        ? openLinkActionLabel(entry.threadId)
        : suggestedActionLabel(agenticType);
      const hint = '<span class="suggested-enter-hint" aria-hidden="true">⏎</span>';
      let messageHtml = '';
      if (suggestedActions.length > 1) {
        const listItems = suggestedActions
          .map(action => {
            const prompt = typeof action?.userFacingPrompt === 'string' ? action.userFacingPrompt.trim() : '';
            if (!prompt) return '';
            return `<li>${renderPlainText(prompt, { preserveLineBreaks: true })}</li>`;
          })
          .filter(Boolean)
          .join('');
        messageHtml = `<p class="suggested-copy">Recommended sequence:</p><ol class="suggested-sequence">${listItems}</ol>`;
      } else {
        messageHtml = `<p class="suggested-copy">${renderPlainText(entry.content, { preserveLineBreaks: true })}</p>`;
      }
      return `<div class="chat-message assistant">
        <div class="assistant-avatar" aria-hidden="true">S</div>
        <div class="chat-card">
          ${messageHtml}
        </div>
      </div>
      <div class="suggested-action-row" role="group" aria-label="Suggested action">
        <button type="button" class="suggested-btn" data-action="suggested-primary" data-thread-id="${escapeAttribute(entry.threadId)}" data-action-type="${escapeAttribute(agenticType)}" ${disabled ? 'disabled' : ''}>${label}${hint}</button>
      </div>`;
    }
    if (entry.messageType === 'draft_details') {
      const draft = normalizeDraftPayload(entry.payload);
      const due = draft.dueDate ? `Due ${formatFriendlyDate(draft.dueDate)}` : 'No due date';
      return `<div class="chat-message assistant">
        <div class="assistant-avatar" aria-hidden="true">S</div>
        <div class="chat-card draft-card" data-thread-id="${escapeAttribute(entry.threadId)}">
          <p><strong>Title:</strong> ${htmlEscape(draft.title || 'New task')}</p>
          <p><strong>Notes:</strong> ${htmlEscape(draft.notes || 'None')}</p>
          <p><strong>Due:</strong> ${htmlEscape(due)}</p>
          <div class="suggested-actions">
            <button type="button" class="enter-hint-btn" data-action="draft-create" data-thread-id="${escapeAttribute(entry.threadId)}">Create task<span class="suggested-enter-hint" aria-hidden="true">⏎</span></button>
            <button type="button" data-action="draft-edit" data-thread-id="${escapeAttribute(entry.threadId)}">Edit</button>
          </div>
        </div>
      </div>`;
    }
    if (entry.messageType === 'inline_editor') {
      const draft = normalizeDraftPayload(entry.payload);
      const editorId = entry.id || `${entry.threadId}-editor`;
      return `<div class="chat-message assistant">
        <div class="assistant-avatar" aria-hidden="true">S</div>
        <div class="chat-card draft-editor" data-editor-id="${escapeAttribute(editorId)}" data-thread-id="${escapeAttribute(entry.threadId)}">
          <div class="task-field"><label>Title</label><input type="text" data-field="title" value="${escapeAttribute(draft.title || '')}" /></div>
          <div class="task-field"><label>Notes</label><textarea data-field="notes">${htmlEscape(draft.notes || '')}</textarea></div>
          <div class="task-field"><label>Due date</label><input type="date" data-field="dueDate" value="${escapeAttribute(draft.dueDate || '')}" /></div>
          <div class="suggested-actions">
            <button type="button" data-action="editor-create" data-editor-id="${escapeAttribute(editorId)}" data-thread-id="${escapeAttribute(entry.threadId)}">Create task</button>
            <button type="button" data-action="editor-save" data-editor-id="${escapeAttribute(editorId)}" data-thread-id="${escapeAttribute(entry.threadId)}">Save draft</button>
          </div>
        </div>
      </div>`;
    }
    if (entry.messageType === 'action_result') {
      return `<div class="chat-message assistant"><div class="assistant-avatar" aria-hidden="true">S</div><div class="chat-card">${renderAssistantMarkdown(entry.content)}</div></div>`;
    }
    return `<div class="chat-message assistant"><div class="assistant-avatar" aria-hidden="true">S</div><div class="chat-card">${renderAssistantMarkdown(entry.content)}</div></div>`;
  }

  function renderExternalLinks(links) {
    if (!links || !links.length) return '';
    const items = links.map(link => {
      const label = htmlEscape(link.label || link.url || 'Open link');
      const href = escapeAttribute(link.url || '');
      const displayUrl = htmlEscape(formatLinkLabel(link.url || ''));
      return `<li><a href="${href}" target="_blank" rel="noopener noreferrer">${label}</a><span class="external-link-url">${displayUrl}</span></li>`;
    }).join('');
    return `<ul class="external-links">${items}</ul>`;
  }

  function formatLinkLabel(url) {
    if (!url) return '';
    try {
      const parsed = new URL(url);
      const host = parsed.host.replace(/^www\./, '');
      const path = parsed.pathname.length > 1 ? parsed.pathname.replace(/\/$/, '') : '';
      return `${host}${path}`;
    } catch {
      return url;
    }
  }

  function chatPlaceholder() {
    return '';
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
    const preserveReplyPanel = Boolean(options.preserveReplyPanel);
    const activeThread = state.activeId ? state.lookup.get(state.activeId) : null;
    const canUnsubscribe = Boolean(activeThread?.unsubscribe?.supported);
    refs.chatInput.disabled = !enabled || !state.activeId;
    if (refs.reviewBtn) refs.reviewBtn.disabled = !enabled || !state.activeId;
    if (refs.replyBtn) refs.replyBtn.disabled = !enabled || !state.activeId;
    if (refs.archiveBtn) refs.archiveBtn.disabled = !enabled || !state.activeId;
    if (refs.unsubscribeBtn) refs.unsubscribeBtn.disabled = !enabled || !state.activeId || !canUnsubscribe;
    updateOpenLinkButton(enabled);
    if (refs.taskBtn) refs.taskBtn.disabled = !enabled || !state.activeId;
    if (refs.skipBtn) refs.skipBtn.disabled = !enabled || !state.activeId;
    if (!enabled) {
      clearComposerNudge();
      if (!preserveTaskPanel) {
        closeTaskPanel(true);
      }
      if (!preserveReplyPanel) {
        closeReplyPanel(true);
      }
    }
  }

  function setPendingSuggestedAction(threadId, actionType) {
    const normalized = normalizeSuggestedAction(actionType);
    if (!threadId || !normalized) return;
    state.pendingSuggestedActions.set(threadId, normalized);
    logDebug('setPendingSuggestedAction', { threadId, actionType: normalized });
    updateComposerPlaceholder(threadId);
  }
  function clearPendingSuggestedAction(threadId = state.activeId) {
    if (!threadId) return;
    state.pendingSuggestedActions.delete(threadId);
    logDebug('clearPendingSuggestedAction', { threadId });
    updateComposerPlaceholder(threadId);
  }
  function getPendingSuggestedAction(threadId = state.activeId) {
    if (!threadId) return '';
    return state.pendingSuggestedActions.get(threadId) || '';
  }

  function setPendingCreate(threadId) {
    state.pendingCreateThreadId = threadId || '';
    renderTaskPanel();
    updateComposerPlaceholder(threadId);
  }

  function clearPendingCreate() {
    if (!state.pendingCreateThreadId) return;
    state.pendingCreateThreadId = '';
    renderTaskPanel();
    updateComposerPlaceholder();
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

  function startThreadLoadTyping(threadId = state.activeId) {
    if (!threadId || threadId !== state.activeId) return;
    if (state.loadTypingThreads.has(threadId)) return;
    state.loadTypingThreads.add(threadId);
    startAssistantTyping(threadId);
  }

  function stopThreadLoadTyping(threadId = state.activeId) {
    if (!threadId || threadId !== state.activeId) return;
    if (!state.loadTypingThreads.has(threadId)) return;
    state.loadTypingThreads.delete(threadId);
    stopAssistantTyping(threadId);
    if (state.snapThreadId === threadId) {
      state.snapThreadId = '';
    }
  }

  function startAssistantTyping(threadId = state.activeId) {
    if (threadId !== state.activeId) return;
    typingSessions += 1;
    if (typingSessions === 1) setAssistantTyping(true);
  }

  function stopAssistantTyping(threadId = state.activeId) {
    if (threadId !== state.activeId) return;
    typingSessions = Math.max(0, typingSessions - 1);
    if (typingSessions === 0) setAssistantTyping(false);
  }

  function nextTypingDelay(message) {
    const base = 220;
    const extra = Math.min(420, Math.max(0, Math.floor((message || '').length * 6)));
    return base + extra;
  }

  function enqueueAssistantMessage(threadId, content, options = {}) {
    if (!threadId) return Promise.resolve();
    const preDelay = Number.isFinite(options.preDelayMs) ? Math.max(0, options.preDelayMs) : 0;
    const bonusDelay = Number.isFinite(options.typingDelayMs) ? Math.max(0, options.typingDelayMs) : 0;
    const delay = options.instant ? 0 : nextTypingDelay(content) + bonusDelay;
    const queue = assistantQueues.get(threadId) || Promise.resolve();
    const next = queue.then(async () => {
      const isActive = threadId === state.activeId;
      if (delay > 0) {
        if (preDelay > 0) await sleep(preDelay);
        if (isActive) startAssistantTyping(threadId);
        await sleep(delay);
      }
      appendTurn(threadId, { role: 'assistant', content });
      renderChat(threadId);
      if (delay > 0 && isActive) stopAssistantTyping(threadId);
    }).catch((err) => {
      console.error('Failed to enqueue assistant message', err);
      stopAssistantTyping(threadId);
    });
    assistantQueues.set(threadId, next);
    return next;
  }

  async function waitForAssistantQueue(threadId) {
    const queue = assistantQueues.get(threadId);
    if (!queue) return;
    try {
      await queue;
    } catch (err) {
      console.error('Failed while waiting for assistant queue', err);
    }
  }

  async function waitForAssistantSettled(threadId) {
    if (!threadId) return;
    revealPendingTranscripts(threadId);
    let current = assistantQueues.get(threadId);
    while (current) {
      try {
        await current;
      } catch (err) {
        console.error('Failed while waiting for assistant to settle', err);
      }
      const next = assistantQueues.get(threadId);
      if (next === current) break;
      current = next;
    }
  }

  function applyTranscriptEntryEffects(threadId, entry) {
    if (entry.messageType === 'must_know') {
      const mustKnow = typeof entry.content === 'string' ? entry.content.trim() : '';
      if (mustKnow) state.mustKnowByThread.set(threadId, mustKnow);
    }
    if (entry.messageType === 'suggested_action') {
      const suggestedActions = normalizeSuggestedActionsPayload(entry?.payload, entry?.content);
      const primary = suggestedActions[0]?.actionType || entry?.payload?.actionType;
      const primaryType = normalizeActionType(primary);
      if (primaryType === 'open_link') {
        const externalAction = normalizeExternalActionPayload(
          suggestedActions[0]?.externalAction
          || entry?.payload?.externalAction
          || entry?.payload?.external_action
        );
        setSuggestedLinks(threadId, externalAction);
        setPendingSuggestedAction(threadId, 'open_link');
      } else if (primaryType === 'external_action') {
        setSuggestedLinks(threadId, null);
        setPendingSuggestedAction(threadId, 'external_action');
      } else {
        setSuggestedLinks(threadId, null);
        const suggested = normalizeSuggestedAction(primaryType);
        if (suggested) setPendingSuggestedAction(threadId, suggested);
      }
    }
    if (entry.messageType === 'draft_details') {
      setPendingCreate(threadId);
    }
  }

  function enqueueTranscriptEntry(threadId, entry, options = {}) {
    if (!threadId) return Promise.resolve();
    const manageTyping = options.manageTyping !== false;
    const delay = nextTypingDelay(entry.content || '');
    const queue = assistantQueues.get(threadId) || Promise.resolve();
    const next = queue.then(async () => {
      const isActive = threadId === state.activeId;
      if (delay > 0 && isActive && manageTyping) startAssistantTyping(threadId);
      if (delay > 0) await new Promise(resolve => window.setTimeout(resolve, delay));
      applyTranscriptEntryEffects(threadId, entry);
      state.timeline.push(entry);
      renderChat(threadId);
      if (delay > 0 && isActive && manageTyping) stopAssistantTyping(threadId);
    }).catch((err) => {
      console.error('Failed to enqueue transcript entry', err);
      if (manageTyping) stopAssistantTyping(threadId);
    });
    assistantQueues.set(threadId, next);
    return next;
  }

  function storePendingTranscript(threadId, entry) {
    const pending = pendingTranscripts.get(threadId) || [];
    pending.push(entry);
    pendingTranscripts.set(threadId, pending);
  }

  function revealPendingTranscripts(threadId) {
    const pending = pendingTranscripts.get(threadId);
    if (!pending || !pending.length) return;
    pendingTranscripts.delete(threadId);
    const isActive = threadId === state.activeId;
    if (isActive) startThreadLoadTyping(threadId);
    const chain = pending.reduce(
      (promise, entry) => promise.then(() => enqueueTranscriptEntry(threadId, entry, { manageTyping: false })),
      Promise.resolve()
    );
    if (isActive) {
      chain.finally(() => stopThreadLoadTyping(threadId));
    }
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
    startAssistantTyping(threadId);
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
      await enqueueAssistantMessage(threadId, reply || 'Here’s what I could pull together.');
      updateHint(threadId);
    } catch (err) {
      console.error('Review request failed', err);
      const message = err instanceof Error ? err.message : 'Unable to review this email.';
      setChatError(message);
    } finally {
      restore();
      stopAssistantTyping(threadId);
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
      await waitForAssistantSettled(threadId);
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

  async function unsubscribeCurrent(source) {
    if (!state.activeId || !refs.unsubscribeBtn) return;
    const threadId = state.activeId;
    const thread = state.lookup.get(threadId);
    if (!thread?.unsubscribe?.supported) {
      enqueueAssistantMessage(threadId, pickUnsubscribeUnavailableMessage());
      return;
    }
    const restoreBtn = withButtonBusy(refs.unsubscribeBtn, 'Unsubscribing…');
    toggleComposer(false);
    setChatError('');
    try {
      await executeActionForThread(threadId, 'unsubscribe');
    } finally {
      restoreBtn();
      toggleComposer(Boolean(state.activeId));
    }
  }

  async function openSuggestedLinks(source) {
    if (!state.activeId || !refs.openLinkBtn) return;
    const threadId = state.activeId;
    const suggestion = getSuggestedLinks(threadId);
    const links = Array.isArray(suggestion?.links) ? suggestion.links : [];
    const key = buildLinksKey(links);
    const progress = state.openLinkProgress.get(threadId);
    let index = progress && progress.key === key ? progress.index : 0;
    const target = links[index];
    const url = typeof target?.url === 'string' ? target.url.trim() : '';
    if (!url) {
      enqueueAssistantMessage(threadId, 'I don’t see any suggested links to open for this email.');
      return;
    }
    const restoreBtn = withButtonBusy(refs.openLinkBtn, 'Opening…');
    toggleComposer(false);
    setChatError('');
    try {
      const win = window.open(url, '_blank', 'noopener,noreferrer');
      if (!win) {
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.target = '_blank';
        anchor.rel = 'noopener noreferrer';
        document.body.appendChild(anchor);
        anchor.click();
        anchor.remove();
      }
      await executeActionForThread(threadId, 'open_link', { links: [url] });
      index += 1;
      state.openLinkProgress.set(threadId, { key, index });
      updateOpenLinkButton();
      renderChat(threadId);
      if (index < links.length) {
        const openedLabel = ordinalLabel(index);
        const nextLabel = ordinalLabel(index + 1);
        enqueueAssistantMessage(threadId, `Opened the ${openedLabel} link. When you’re back, open the ${nextLabel} link.`);
        setPendingSuggestedAction(threadId, 'open_link');
      } else {
        clearPendingSuggestedAction(threadId);
        promptArchiveAfterLinks(threadId);
      }
    } finally {
      restoreBtn();
      toggleComposer(Boolean(state.activeId));
    }
  }

  function removeCurrentFromQueue() {
    if (!state.activeId || !state.needs.length) return;
    const threadId = state.activeId;
    const order = getReviewOrder();
    const orderIndex = order.indexOf(threadId);
    const index = state.needs.indexOf(threadId);
    if (index === -1) return;
    state.needs.splice(index, 1);
    markReviewed(threadId);
    rebuildPriorityQueue();
    updateProgress();
    updateDrawerLists();
    updateHeaderCount();
    updatePriorityPill();
    updateQueuePill();
    closeTaskPanel(true);
    closeReplyPanel(true);
    clearPendingSuggestedAction(threadId);

    if (maybeSwitchToPriorityAfterResolve()) {
      return;
    }

    if (!state.needs.length) {
      if (state.hasMore) {
        autoLoadNextBatch();
      } else {
        setEmptyState('All emails reviewed. Nice work.');
        toggleComposer(false);
      }
      return;
    }
    const nextId = getNextUnreviewedAfter(threadId, {
      order,
      startIndex: orderIndex + 1,
      allowWrap: !state.hasMore
    });
    if (nextId) {
      scheduleAutoAdvance(nextId);
      return;
    }
    if (state.hasMore) {
      autoLoadNextBatch();
    } else {
      setEmptyState('All emails reviewed. Nice work.');
      toggleComposer(false);
    }
  }

  async function autoLoadNextBatch() {
    if (state.loadingMore || !state.hasMore) return;
    setEmptyState('Loading more emails…');
    toggleComposer(false);
    const added = await fetchNextPage('auto');
    if (added.length) {
      const nextId = getNextReviewCandidate();
      if (nextId) {
        scheduleAutoAdvance(nextId);
        return;
      }
    }
    if (!added.length) {
      const message = state.hasMore
        ? 'Unable to load more emails. Tap Review latest emails to retry.'
        : 'All emails reviewed. Nice work.';
      setEmptyState(message);
      toggleComposer(false);
    }
  }

  function advanceToNextThread(threadId = state.activeId) {
    const order = getReviewOrder();
    if (!threadId || !order.length) return;
    const index = order.indexOf(threadId);
    if (index === -1) return;
    if (index === state.needs.length - 1 && state.hasMore) {
      autoLoadNextBatch();
      return;
    }
    const nextIndex = order.length > 1 ? (index + 1) % order.length : -1;
    const nextId = nextIndex >= 0 ? order[nextIndex] : '';
    if (nextId && nextId !== threadId) {
      scheduleAutoAdvance(nextId);
    }
  }

  function skipCurrent(source) {
    if (!state.activeId) return;
    const threadId = state.activeId;
    const order = getReviewOrder();
    const index = order.indexOf(threadId);
    const hasRoomToAdvance = order.length > 1;
    markReviewed(threadId);
    rebuildPriorityQueue();
    updateProgress();
    updateDrawerLists();
    updateHeaderCount();
    updatePriorityPill();
    updateQueuePill();
    closeTaskPanel(true);
    closeReplyPanel(true);
    clearPendingSuggestedAction(threadId);

    if (maybeSwitchToPriorityAfterResolve()) {
      return;
    }
    const nextId = getNextUnreviewedAfter(threadId, {
      order,
      startIndex: index + 1,
      allowWrap: !state.hasMore
    });
    if (nextId) {
      scheduleAutoAdvance(nextId);
      return;
    }
    if (state.hasMore) {
      autoLoadNextBatch();
      return;
    }
    if (!hasRoomToAdvance) {
      setEmptyState('All emails reviewed. Nice work.');
      toggleComposer(false);
      return;
    }
    const nextIndex = index === -1 ? 0 : (index + 1) % order.length;
    const fallbackId = order[nextIndex] || order[0];
    if (fallbackId) {
      scheduleAutoAdvance(fallbackId);
    }
  }

  function getReviewOrder() {
    if (!state.needs.length) return [];
    const needsSet = new Set(state.needs);
    const priorityIds = state.priority.filter(id => needsSet.has(id) && !reviewedIds.has(id));
    const prioritySet = new Set(priorityIds);
    const remaining = state.needs.filter(id => !prioritySet.has(id));
    return [...priorityIds, ...sortThreadIdsByReceivedAt(remaining)];
  }

  function sortThreadIdsByReceivedAt(ids) {
    return ids.slice().sort((a, b) => {
      const aThread = state.lookup.get(a);
      const bThread = state.lookup.get(b);
      const aTime = aThread?.receivedAt ? new Date(aThread.receivedAt).getTime() : 0;
      const bTime = bThread?.receivedAt ? new Date(bThread.receivedAt).getTime() : 0;
      return bTime - aTime;
    });
  }

  function getNextReviewCandidate() {
    const order = getReviewOrder();
    for (let i = 0; i < order.length; i += 1) {
      const candidate = order[i];
      if (!reviewedIds.has(candidate)) return candidate;
    }
    return '';
  }

  function getNextUnreviewedAfter(threadId, options = {}) {
    const order = Array.isArray(options.order) ? options.order : state.needs;
    if (!order.length || !state.needs.length) return '';
    const allowWrap = options.allowWrap !== undefined ? Boolean(options.allowWrap) : true;
    const providedStart = Number.isInteger(options.startIndex) ? Number(options.startIndex) : null;
    const startIndex = providedStart !== null
      ? Math.max(0, providedStart)
      : Math.max(0, order.indexOf(threadId) + 1);
    const needsSet = new Set(state.needs);
    const maxOffset = allowWrap ? order.length : Math.max(0, order.length - startIndex);
    for (let offset = 0; offset < maxOffset; offset += 1) {
      const index = allowWrap ? (startIndex + offset) % order.length : startIndex + offset;
      const candidate = order[index];
      if (!needsSet.has(candidate)) continue;
      if (!reviewedIds.has(candidate)) return candidate;
    }
    return '';
  }

  function handleAutoIntent(intent, userText, options = {}) {
    if (!state.activeId || intent !== 'skip') return;
    const alreadyLogged = Boolean(options.alreadyLogged);
    const history = ensureHistory(state.activeId);
    if (!alreadyLogged) {
      appendTurn(state.activeId, { role: 'user', content: userText });
      renderChat();
    }

    const skipPromise = enqueueAssistantMessage(state.activeId, 'Skipping for now. It stays in Needs Review.');
    updateHint(state.activeId);

    clearAutoAdvance();
    const targetId = state.activeId;
    skipPromise.then(async () => {
      if (state.activeId !== targetId) return;
      await waitForAssistantSettled(targetId);
      if (state.activeId !== targetId) return;
      skipCurrent('auto');
      clearAutoAdvance();
    });
  }

  function clearAutoAdvance() {
    if (state.autoAdvanceTimer) {
      window.clearTimeout(state.autoAdvanceTimer);
      state.autoAdvanceTimer = 0;
    }
  }

  function scheduleAutoAdvance(threadId, options = {}) {
    if (!threadId) return;
    const delayMs = Number.isFinite(options.delayMs) ? options.delayMs : AUTO_THREAD_PAUSE_MS;
    clearAutoAdvance();
    state.autoAdvanceTimer = window.setTimeout(() => {
      state.autoAdvanceTimer = 0;
      setActiveThread(threadId, { source: 'auto' });
    }, Math.max(0, delayMs));
  }

  async function handleArchiveIntent(userText, options = {}) {
    if (!state.activeId) return;
    const alreadyLogged = Boolean(options.alreadyLogged);
    const history = ensureHistory(state.activeId);
    if (!alreadyLogged) {
      appendTurn(state.activeId, { role: 'user', content: userText });
      renderChat();
    }

    enqueueAssistantMessage(state.activeId, 'Archiving this email in Gmail…');
    updateHint(state.activeId);

    startAssistantTyping(state.activeId);
    try {
      await archiveCurrent('auto-intent');
    } finally {
      stopAssistantTyping(state.activeId);
    }
  }

  function handleCreateTaskIntent() {
    if (!state.activeId) return;
    requestDraft(state.activeId, 'generate');
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
      if (action === 'archive') {
        clearPendingSuggestedAction(state.activeId);
        await executeActionForThread(state.activeId, 'archive');
        return true;
      }
      if (action === 'reply') {
        clearPendingSuggestedAction(state.activeId);
        openReplyPanelFromPrompt(userText);
        return true;
      }
      if (action === 'skip') {
        clearPendingSuggestedAction(state.activeId);
        await executeActionForThread(state.activeId, 'skip');
        return true;
      }
      if (action === 'unsubscribe') {
        clearPendingSuggestedAction(state.activeId);
        await executeActionForThread(state.activeId, 'unsubscribe');
        return true;
      }
      if (action === 'create_task') {
        clearPendingSuggestedAction(state.activeId);
        handleCreateTaskIntent();
        return true;
      }
      if (action === 'more_info') {
        clearPendingSuggestedAction(state.activeId);
        await executeActionForThread(state.activeId, 'more_info');
        return true;
      }
      if (action === 'open_link') {
        clearPendingSuggestedAction(state.activeId);
        await openSuggestedLinks('suggested');
        return true;
      }
      if (action === 'external_action') {
        clearPendingSuggestedAction(state.activeId);
        await executeActionForThread(state.activeId, 'external_action');
        return true;
      }
    }

    if (isNegativeResponse(normalized)) {
      clearPendingSuggestedAction(state.activeId);
      enqueueAssistantMessage(state.activeId, 'Okay, I’ll hold off. Tell me what you’d like me to do instead.');
      return true;
    }

    const intent = await detectIntent(userText);
    if (intent === 'archive') {
      clearPendingSuggestedAction(state.activeId);
      await executeActionForThread(state.activeId, 'archive');
      return true;
    }
    if (intent === 'unsubscribe') {
      clearPendingSuggestedAction(state.activeId);
      await executeActionForThread(state.activeId, 'unsubscribe');
      return true;
    }
    if (intent === 'skip') {
      clearPendingSuggestedAction(state.activeId);
      await executeActionForThread(state.activeId, 'skip');
      return true;
    }
    if (intent === 'create_task') {
      clearPendingSuggestedAction(state.activeId);
      handleCreateTaskIntent();
      return true;
    }
    if (intent === 'open_link') {
      clearPendingSuggestedAction(state.activeId);
      await openSuggestedLinks('intent');
      return true;
    }
    if (intent === 'external_action') {
      clearPendingSuggestedAction(state.activeId);
      await executeActionForThread(state.activeId, 'external_action');
      return true;
    }
    if (intent === 'reply') {
      clearPendingSuggestedAction(state.activeId);
      openReplyPanelFromPrompt(userText);
      return true;
    }

    // User asked something else — clear pending and let normal flow handle it.
    clearPendingSuggestedAction(state.activeId);
    return false;
  }


  async function handleCreateConfirmationResponse(userText) {
    if (!state.activeId || !isCreateConfirmationPending(state.activeId)) return;
    const normalized = (userText || '').trim().toLowerCase();
    const flow = state.actionFlows.get(state.activeId);

    if (isAffirmativeResponse(normalized)) {
      if (flow?.actionType === 'create_task') {
        await executeActionForThread(state.activeId, 'create_task');
        return;
      }
      const result = await submitTask();
      if (result?.ok) {
        enqueueAssistantMessage(state.activeId, buildTaskCreatedMessage(result));
        clearPendingCreate();
        promptArchiveAfterTask(result, { includeSuccess: false });
      } else if (result?.error) {
        enqueueAssistantMessage(state.activeId, `Couldn't create the task: ${result.error}`);
      }
      return;
    }

    if (isNegativeResponse(normalized)) {
      clearPendingCreate();
      renderTaskPanel();
      enqueueAssistantMessage(state.activeId, 'Okay, I won’t create it. Adjust the fields or ask another action.');
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
    if (intent === 'reply') {
      clearPendingCreate();
      closeTaskPanel(true);
      openReplyPanelFromPrompt(userText);
      return;
    }
    if (intent === 'unsubscribe') {
      clearPendingCreate();
      closeTaskPanel(true);
      await executeActionForThread(state.activeId, 'unsubscribe');
      return;
    }
    if (intent === 'open_link') {
      clearPendingCreate();
      closeTaskPanel(true);
      await openSuggestedLinks('intent');
      return;
    }

    enqueueAssistantMessage(state.activeId, 'Please confirm: create the task as shown? (yes/no)');
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

  function pickUnsubscribeUnavailableMessage() {
    const options = [
      'Looks like Gmail does not offer an unsubscribe option for this sender.',
      'I can’t unsubscribe from this one—Gmail didn’t provide a supported unsubscribe link.',
      'No easy unsubscribe is available for this email. Want me to archive it instead?'
    ];
    return options[Math.floor(Math.random() * options.length)];
  }

  function promptArchiveAfterTask(result, options = {}) {
    const threadId = options.threadId || state.activeId;
    if (!threadId) return;
    const includeSuccess = Boolean(options.includeSuccess);
    setPendingArchive(threadId);
    if (includeSuccess) {
      enqueueAssistantMessage(threadId, buildTaskCreatedMessage(result));
    }
    const prompt = 'Archive this email and move on? I can keep it here if you want.';
    enqueueAssistantMessage(threadId, prompt);
    updateHint(threadId);
  }

  function promptArchiveAfterReply(options = {}) {
    const threadId = options.threadId || state.activeId;
    if (!threadId) return;
    setPendingArchive(threadId);
    enqueueAssistantMessage(threadId, 'Archive this email and move on? I can keep it here if you want.');
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
    if (intent === 'reply') {
      clearPendingArchive();
      openReplyPanelFromPrompt(userText);
      return;
    }

    if (isAffirmativeResponse(normalized)) {
      clearPendingArchive();
      await handleArchiveIntent(userText, { alreadyLogged: true });
      return;
    }

    if (isNegativeResponse(normalized)) {
      clearPendingArchive();
      enqueueAssistantMessage(state.activeId, 'Okay, leaving it in Needs Review. Ask to archive anytime.');
      return;
    }

    enqueueAssistantMessage(state.activeId, 'Want me to archive this email or keep it here?');
  }

  function setEmptyState(message) {
    const fallback = state.hasMore
      ? 'You reviewed everything loaded. Tap Review latest emails to keep going.'
      : 'All emails reviewed. Nice work.';
    const copy = message || fallback;
    closeTaskPanel(true);
    closeReplyPanel(true);
    state.activeId = '';
    refs.emailEmpty.classList.remove('hidden');
    if (refs.position) {
      refs.position.textContent = '';
      refs.position.classList.add('hidden');
    }
    if (refs.emailEmptyText) refs.emailEmptyText.textContent = copy;
    if (refs.loadMoreEmpty) {
      refs.loadMoreEmpty.classList.toggle('hidden', !state.needs.length);
    }
    renderChat();
    updateQueuePill();
    updateLoadMoreButtons();
  }

  function showLoadingOverlay(message) {
    if (!loadingOverlay) return;
    loadingOverlay.dataset.owner = OVERLAY_OWNER;
    loadingOverlay.classList.remove('hidden');
    if (loadingText && message) loadingText.textContent = message;
  }

  function hideLoadingOverlay() {
    if (!loadingOverlay) return;
    if (loadingOverlay.dataset.owner !== OVERLAY_OWNER) return;
    delete loadingOverlay.dataset.owner;
    loadingOverlay.classList.add('hidden');
  }

  function toggleQueueSection(sectionEl, toggleBtn) {
    if (!sectionEl || !toggleBtn) return;
    const isCollapsed = sectionEl.classList.toggle('is-collapsed');
    const expanded = !isCollapsed;
    toggleBtn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
    toggleBtn.textContent = expanded ? 'Collapse' : 'Expand';
  }

  function setQueueSectionExpanded(sectionEl, toggleBtn, expanded) {
    if (!sectionEl || !toggleBtn) return;
    sectionEl.classList.toggle('is-collapsed', !expanded);
    toggleBtn.setAttribute('aria-expanded', expanded ? 'true' : 'false');
    toggleBtn.textContent = expanded ? 'Collapse' : 'Expand';
  }

  function handleReviewLatestClick() {
    if (!state.needs.length) return;
    setQueueSectionExpanded(refs.priorityQueue, refs.priorityToggle, false);
    setQueueSectionExpanded(refs.reviewQueue, refs.queueToggle, true);
    const order = sortThreadIdsByReceivedAt(state.needs);
    const firstId = order[0];
    if (firstId) {
      markUserEngagedIfInbox(firstId);
      setActiveThread(firstId, { source: 'user' });
    }
  }

  async function detectIntent(text) {
    const cleaned = text.replace(/[.!?]/g, '').trim().toLowerCase();
    if (!cleaned) return '';
    const archivePhrases = ['archive', 'archive this', 'archive it', 'archive email', 'archive message', 'archive thread'];
    if (archivePhrases.includes(cleaned)) return 'archive';
    const openLinkPhrases = [
      'open link',
      'open the link',
      'open links',
      'open the links',
      'visit link',
      'visit the link',
      'visit links',
      'visit the links',
      'go to link',
      'go to the link',
      'click link',
      'click the link',
      'open url',
      'open website',
      'open the url',
      'open the site'
    ];
    if (openLinkPhrases.includes(cleaned)) return 'open_link';
    const unsubscribePhrases = [
      'unsubscribe',
      'unsubscribe me',
      'unsubscribe from this',
      'unsubscribe from this email',
      'opt out',
      'opt me out',
      'remove me',
      'remove me from this list',
      'stop these emails',
      'stop sending these',
      'stop sending me these',
      'stop emails like this'
    ];
    if (unsubscribePhrases.includes(cleaned)) return 'unsubscribe';
    const skipPhrases = ['skip', 'skip it', 'skip this', 'skip this one', 'skip this email', 'skip this thread'];
    if (skipPhrases.includes(cleaned)) return 'skip';
    const replyPhrases = [
      'reply',
      'reply back',
      'reply to',
      'respond',
      'respond to',
      'send a reply',
      'send a response',
      'draft a reply',
      'write back',
      'email back',
      'answer them'
    ];
    if (replyPhrases.includes(cleaned)) return 'reply';
    const replyStarters = [
      'reply ',
      'reply back ',
      'respond ',
      'send a reply ',
      'send a response ',
      'draft a reply ',
      'write back ',
      'email back ',
      'tell them ',
      'tell him ',
      'tell her ',
      'let them know ',
      'let him know ',
      'let her know ',
      'say ',
      'ask them ',
      'ask him ',
      'ask her '
    ];
    if (replyStarters.some(prefix => cleaned.startsWith(prefix))) return 'reply';
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
    return intent === 'archive' || intent === 'skip' || intent === 'create_task' || intent === 'reply' || intent === 'unsubscribe' || intent === 'open_link'
      ? intent
      : '';
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

  function formatQueueTimestamp(iso) {
    if (!iso) return '';
    const date = new Date(iso);
    if (!Number.isFinite(date.getTime())) return '';
    return date.toLocaleString(undefined, { dateStyle: 'short', timeStyle: 'short' });
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
    const normalized = replaceListsWithBullets(html);
    return sanitizeHtml(normalized);
  }

  function replaceListsWithBullets(html) {
    if (typeof document === 'undefined') return html;
    const container = document.createElement('div');
    container.innerHTML = html;
    const lists = container.querySelectorAll('ul, ol');
    lists.forEach((list) => {
      const prev = list.previousElementSibling;
      if (prev && prev.tagName === 'P') {
        prev.classList.add('tight-next');
      }
      const items = Array.from(list.children).filter((node) => node.tagName === 'LI');
      if (!items.length) {
        list.remove();
        return;
      }
      const block = document.createElement('div');
      block.className = 'bullet-block';
      items.forEach((item) => {
        const line = document.createElement('span');
        line.className = 'bullet-line';
        const content = item.innerHTML.trim();
        line.innerHTML = `• ${content}`;
        block.appendChild(line);
      });
      list.replaceWith(block);
    });
    return container.innerHTML;
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
