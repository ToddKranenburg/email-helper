(function () {
  const titleEl = document.getElementById('onboarding-title');
  const bodyEl = document.getElementById('onboarding-body');
  const pillsEl = document.getElementById('onboarding-pills');
  const progressEl = document.getElementById('onboarding-progress');
  const backBtn = document.getElementById('onboarding-back');
  const nextBtn = document.getElementById('onboarding-next');
  const chatLog = document.getElementById('sim-chat-log');
  const actionsEl = document.getElementById('sim-actions');
  const form = document.getElementById('onboarding-complete');
  const statusEl = document.getElementById('onboarding-status');

  if (!titleEl || !bodyEl || !pillsEl || !progressEl || !backBtn || !nextBtn || !chatLog || !actionsEl || !form) {
    return;
  }

  const steps = [
    {
      title: 'Meet your email secretary',
      body: 'Your inbox becomes a guided workspace. I surface what needs you and keep everything else easy to review.',
      pills: ['Priority queue', 'Guided actions', 'Full inbox view'],
      chat: [
        { role: 'assistant', text: "Hey, I'm your email secretary. I will triage your inbox and surface what needs you." },
        { role: 'assistant', text: 'I will keep each thread short, clear, and action-ready.' }
      ],
      actions: ['Summarize', 'Reply', 'Archive']
    },
    {
      title: 'Priority means focus',
      body: 'Anything urgent or decision-heavy gets promoted into the Priority Queue so you can clear it fast.',
      pills: ['Urgent first', 'Reason tags', 'Never lose context'],
      chat: [
        { role: 'assistant', text: 'I found two emails that need your attention.' },
        { role: 'assistant', text: 'Let\'s start with the renewal that needs approval.' },
        { role: 'user', text: 'Show me the key points.' }
      ],
      actions: ['More info', 'Approve', 'Schedule']
    },
    {
      title: 'Fast decisions, no guesswork',
      body: 'Every thread comes with a suggested next step so you can respond, archive, or create a task in one move.',
      pills: ['Suggested actions', 'Draft help', 'One-click follow-through'],
      chat: [
        { role: 'assistant', text: 'Here\'s the summary and the recommended next step.' },
        { role: 'assistant', text: 'I drafted a short reply. Want to send it?' }
      ],
      actions: ['Send reply', 'Edit', 'Skip']
    },
    {
      title: 'You stay in control',
      body: 'Browse the full inbox at any time. When you finish here, your priority queue should already be waiting.',
      pills: ['Full inbox access', 'Always editable', 'Human in the loop'],
      chat: [
        { role: 'assistant', text: 'All set. Your inbox is synced and ready.' },
        { role: 'assistant', text: 'Let\'s get through the first priority thread together.' }
      ],
      actions: ['Open priority', 'View inbox', 'Keep scanning']
    }
  ];

  let current = 0;
  let playing = false;

  function setStatusCopy(text) {
    if (!statusEl || !text) return;
    statusEl.textContent = text;
  }

  async function pollStatus() {
    try {
      const resp = await fetch('/ingest/status', {
        method: 'GET',
        headers: { 'Accept': 'application/json' },
        credentials: 'same-origin'
      });
      const data = await resp.json().catch(() => ({}));
      if (!resp.ok) return;
      if (data?.status === 'running') setStatusCopy('Syncing Gmail in the background…');
      if (data?.status === 'done') setStatusCopy('Inbox synced. Prioritizing now…');
      if (data?.status === 'error') setStatusCopy('Sync paused. We\'ll retry in the background.');
    } catch {
      // No-op
    }
  }

  function renderProgress() {
    const dots = Array.from(progressEl.querySelectorAll('.progress-dot'));
    dots.forEach((dot, index) => {
      dot.classList.toggle('active', index === current);
    });
  }

  function renderPills(list) {
    pillsEl.innerHTML = '';
    list.forEach(item => {
      const span = document.createElement('span');
      span.className = 'onboarding-pill';
      span.textContent = item;
      pillsEl.appendChild(span);
    });
  }

  function renderActions(list) {
    actionsEl.innerHTML = '';
    list.forEach(label => {
      const div = document.createElement('div');
      div.className = 'sim-action';
      div.textContent = label;
      actionsEl.appendChild(div);
    });
  }

  function renderChatMessage(message) {
    const bubble = document.createElement('div');
    bubble.className = `sim-bubble ${message.role}`;
    bubble.textContent = message.text;
    chatLog.appendChild(bubble);
    chatLog.scrollTop = chatLog.scrollHeight;
  }

  function renderTyping() {
    const typing = document.createElement('div');
    typing.className = 'sim-typing';
    typing.innerHTML = '<span class="sim-dot"></span><span class="sim-dot"></span><span class="sim-dot"></span>';
    chatLog.appendChild(typing);
    chatLog.scrollTop = chatLog.scrollHeight;
    return typing;
  }

  async function playChat(messages) {
    chatLog.innerHTML = '';
    for (const message of messages) {
      if (message.role === 'assistant') {
        const typing = renderTyping();
        await wait(500 + Math.random() * 600);
        typing.remove();
      }
      renderChatMessage(message);
      await wait(260 + Math.random() * 260);
    }
  }

  function wait(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  async function renderStep(index) {
    const step = steps[index];
    if (!step) return;
    current = index;
    playing = true;
    backBtn.disabled = current === 0;
    nextBtn.disabled = true;
    nextBtn.textContent = current === steps.length - 1 ? 'Open inbox' : 'Next';

    titleEl.textContent = step.title;
    bodyEl.textContent = step.body;
    renderPills(step.pills);
    renderProgress();
    renderActions(step.actions);
    await playChat(step.chat);

    playing = false;
    nextBtn.disabled = false;
  }

  backBtn.addEventListener('click', () => {
    if (playing || current === 0) return;
    renderStep(current - 1);
  });

  nextBtn.addEventListener('click', () => {
    if (playing) return;
    if (current === steps.length - 1) {
      form.submit();
      return;
    }
    renderStep(current + 1);
  });

  renderStep(0);
  pollStatus();
  setInterval(pollStatus, 4000);
})();
