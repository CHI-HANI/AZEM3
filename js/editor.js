/* ══════ PHASE 3: KEYBOARD SHORTCUTS ══════ */
const SHORTCUTS = [
  {key:'1-7', desc:'تغيير الثيم'}, {key:'8', desc:'تغيير الوضع'},
  {key:'M', desc:'تغيير الوضع'}, {key:'L', desc:'تغيير اللغة'},
  {key:'S', desc:'الصوت'}, {key:'← →', desc:'التنقل بين الأيام'},
  {key:'T', desc:'اليوم الحالي'}, {key:'Space', desc:'تشغيل/إيقاف'},
  {key:'N', desc:'التمرين التالي'}, {key:'P', desc:'التمرين السابق'},
  {key:'Enter', desc:'إنهاء المجموعة'}, {key:'Esc', desc:'إغلاق الجلسة'},
  {key:'R', desc:'إعادة المؤقت'}, {key:'C', desc:'تحديد مكتمل'},
  {key:'D', desc:'إنهاء اليوم'}, {key:'A', desc:'المدرب الذكي'},
  {key:'F', desc:'ملء الشاشة'}, {key:'?', desc:'هذه القائمة'},
];

function showShortcutsHelp() {
  const existing = document.getElementById('shortcuts-modal');
  if (existing) { existing.remove(); return; }
  const modal = document.createElement('div');
  modal.id = 'shortcuts-modal';
  modal.style.cssText = 'position:fixed;inset:0;z-index:9800;background:rgba(0,0,0,.75);display:flex;align-items:center;justify-content:center;backdrop-filter:blur(8px);';
  modal.innerHTML = `
    <div style="background:var(--night);border:1px solid var(--border);border-radius:20px;padding:24px;max-width:420px;width:90%;max-height:80vh;overflow-y:auto;">
      <div style="display:flex;align-items:center;margin-bottom:16px;">
        <span style="font-size:16px;font-weight:900;flex:1;">⌨️ اختصارات لوحة المفاتيح</span>
        <button onclick="document.getElementById('shortcuts-modal').remove()" style="font-size:20px;padding:4px 8px;border-radius:8px;background:var(--card);border:1px solid var(--border);">✕</button>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
        ${SHORTCUTS.map(s=>`
          <div style="display:flex;align-items:center;gap:8px;padding:6px 10px;border-radius:10px;background:var(--card);border:1px solid var(--border);">
            <kbd style="background:var(--card2);border:1px solid var(--border);border-radius:6px;padding:2px 8px;font-size:11px;font-weight:900;font-family:monospace;color:var(--gold);white-space:nowrap;">${s.key}</kbd>
            <span style="font-size:12px;color:var(--text2);">${s.desc}</span>
          </div>`).join('')}
      </div>
      <div style="margin-top:16px;padding-top:12px;border-top:1px solid var(--border);">
        <div style="font-size:13px;font-weight:900;margin-bottom:8px;">📺 ريموت التلفزيون</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;font-size:12px;color:var(--text2);">
          <div>🔴 ابدأ/إيقاف الجلسة</div>
          <div>🟢 تحديد مكتمل ✓</div>
          <div>🟡 المدرب الذكي AI</div>
          <div>🔵 هذه القائمة</div>
          <div>1-7 الثيمات</div>
          <div>8 تغيير الوضع</div>
          <div>OK تأكيد/تفاصيل</div>
          <div>Back رجوع</div>
        </div>
      </div>
    </div>`;
  modal.onclick = e => { if(e.target===modal) modal.remove(); };
  document.body.appendChild(modal);
}

document.addEventListener('keydown', (e) => {
  // Don't trigger shortcuts when typing in inputs
  if (['INPUT','TEXTAREA','SELECT'].includes(e.target.tagName)) return;
  
  const sessOpen = document.getElementById('session-overlay')?.classList.contains('open');
  const key = e.key;
  const code = e.code;

  // Theme shortcuts: 1-7
  if (!e.ctrlKey && !e.altKey && key >= '1' && key <= '7') {
    const themes = ['default','fire','ocean','nature','neon','purple','light'];
    setTheme(themes[parseInt(key)-1]);
    showMiniToast(THEME_ICONS[themes[parseInt(key)-1]] + ' ' + themes[parseInt(key)-1]);
    return;
  }

  // Mode: 8 or M
  if (!e.ctrlKey && key === '8') {
    cycleMode(); return;
  }
  if (key.toLowerCase() === 'm' && !sessOpen) {
    cycleMode(); return;
  }

  // Language: L
  if (key.toLowerCase() === 'l' && !sessOpen) {
    cycleLang(); return;
  }

  // Sound: S
  if (key.toLowerCase() === 's' && !sessOpen) {
    openSoundSheet(); return;
  }

  // Help: ?
  if (key === '?' || key === '/') {
    showShortcutsHelp(); return;
  }

  // Fullscreen: F
  if (key.toLowerCase() === 'f') {
    toggleFullscreen(); return;
  }

  // AI coach: A
  if (key.toLowerCase() === 'a' && !sessOpen) {
    if (typeof openAICoach === 'function') openAICoach();
    return;
  }

  // Day navigation: Arrow keys (when not in session)
  if (!sessOpen) {
    // RTL: ArrowRight = previous day (يمين = السابق), ArrowLeft = next day (يسار = التالي)
    if (key === 'ArrowRight') {
      if (S.currentDay > 1) { S.currentDay--; saveState(); render(); }
      return;
    }
    if (key === 'ArrowLeft') {
      const maxDay = S.user?.programDays || 30;
      if (S.currentDay < maxDay) { S.currentDay++; saveState(); render(); }
      return;
    }
    // ArrowDown/Up: scroll through days sequentially
    if (key === 'ArrowDown') {
      if (S.currentDay > 1) { S.currentDay--; saveState(); render(); }
      return;
    }
    if (key === 'ArrowUp') {
      const maxDay = S.user?.programDays || 30;
      if (S.currentDay < maxDay) { S.currentDay++; saveState(); render(); }
      return;
    }
    // T = today
    if (key.toLowerCase() === 't') {
      const maxDay = S.user?.programDays || 30; // FIX-G: use programDays not hardcoded 30
      const today = Math.min(maxDay, (S.completedDays.length||0) + 1);
      S.currentDay = today; saveState(); render();
      showMiniToast('📅 اليوم ' + today);
      return;
    }
    // D = mark day done
    if (key.toLowerCase() === 'd') { toggleDayDone(); return; }
    // C = mark exercise done (first unchecked)
    if (key.toLowerCase() === 'c') {
      const firstCheck = document.querySelector('.ex-check:not(.checked)');
      if (firstCheck) firstCheck.click();
      return;
    }
    // Enter = start guided session (TV remote OK button)
    if (key === 'Enter') { startGuidedSession(); return; }
    // Space = timer play/pause
    if (key === ' ' || code === 'Space') {
      e.preventDefault();
      timerToggle(); return;
    }
    // R = reset timer
    if (key.toLowerCase() === 'r') { timerReset(); return; }
  }

  // Session shortcuts
  if (sessOpen) {
    if (key === 'Escape') { confirmCloseSession(); return; }
    if (key === ' ' || code === 'Space') { e.preventDefault(); sessPause(); return; }
    if (key === 'ArrowRight' || key.toLowerCase() === 'p') { sessPrev(); return; }
    if (key === 'ArrowLeft' || key.toLowerCase() === 'n') { sessSkip(); return; }
    if (key === 'Enter') { sessMainAction(); return; }
  }

  // TV Remote color buttons (Hbb TV standard)
  // TV Remote color buttons (HbbTV standard VK codes)
  if (key === 'ColorF0Red'   || key === 'VK_RED')    { sessOpen ? sessPause() : startGuidedSession(); }
  if (key === 'ColorF1Green' || key === 'VK_GREEN')  { toggleDayDone(); }
  if (key === 'ColorF2Yellow'|| key === 'VK_YELLOW') { openActiveRest(); }
  if (key === 'ColorF3Blue'  || key === 'VK_BLUE')   { if(typeof openAICoach==='function') openAICoach(); }
  // Real TV remote keys available on most remotes:
  if (key === 'GoBack' || key === 'BrowserBack') { confirmCloseSession && sessOpen ? confirmCloseSession() : showMiniToast('اضغط ↑↓ لتغيير اليوم'); }
  if (key === 'HomePage' || key === 'BrowserHome') { cycleMode(); }
  // Also handle Play/Pause media keys common on smart TV remotes
  if (key === 'MediaPlayPause' || key === 'MediaPlay') { sessOpen ? sessPause() : startGuidedSession(); }
  if (key === 'MediaStop') { if(sessOpen && typeof confirmCloseSession==='function') confirmCloseSession(); }
});

// Show shortcuts hint on desktop/TV
function showShortcutHint() {
  if (window.innerWidth > 600) {
    showMiniToast('⌨️ اضغط ? لعرض الاختصارات');
  }
}

/* ══════ PHASE 2: EXERCISE EDITOR ══════ */
/* ══════ DAY EDITOR ══════ */

function getEffectiveSchedule(day) {
  // FIX#2: unified — always delegate to getDaySchedule which already
  // handles S.customSchedule and progressive overload correctly
  return getDaySchedule(day);
}

function openExEditor(exId, day) {
  const allEx = [...EXERCISES, ...(S.customExercises||[])];
  const ex = allEx.find(e=>e.id===exId);
  if (!ex) return;
  document.getElementById('ex-editor-id').value = exId;
  document.getElementById('ex-editor-day').value = day;
  document.getElementById('ex-editor-title').textContent = '✏️ ' + ex.nameAr;
  document.getElementById('ex-ed-nameAr').value = ex.nameAr;
  document.getElementById('ex-ed-nameEn').value = ex.nameEn || '';
  document.getElementById('ex-ed-icon').value = ex.icon || '💪';
  document.getElementById('ex-ed-muscles').value = ex.muscles || '';
  document.getElementById('ex-ed-sets').value = ex.sets;
  document.getElementById('ex-ed-reps').value = ex.reps;
  document.getElementById('ex-ed-type').value = ex.type || 'reps';
  document.getElementById('ex-ed-steps').value = (ex.steps||[]).join('\n');
  const restEl = document.getElementById('ex-ed-rest');
  if (restEl) restEl.value = (typeof ex.rest === 'number' && ex.rest > 0) ? ex.rest : 30;
  // Build "add from list" section
  // FIX-I: read from S.customSchedule first, fall back to getDaySchedule
  const sched = (S.customSchedule && S.customSchedule[day]) || getDaySchedule(day).exercises.map(e=>e.id);
  const available = allEx.filter(e => !sched.includes(e.id));
  document.getElementById('ex-add-from-list').innerHTML = available.map(e=>
    `<button onclick="addExToDay('${e.id}',${day})" style="padding:6px 12px;border-radius:20px;background:var(--card2);border:1px solid var(--border);font-size:12px;font-weight:700;cursor:pointer;">${e.icon} ${e.nameAr}</button>`
  ).join('');
  document.getElementById('new-ex-form').style.display = 'none';
  document.getElementById('ex-editor-modal').classList.add('open');
}

function exEdTypeChange() {
  const type = document.getElementById('ex-ed-type')?.value || 'reps';
  const lbl = document.getElementById('ex-ed-reps-lbl');
  const inp = document.getElementById('ex-ed-reps');
  if (!lbl || !inp) return;
  const labels = {reps:'التكرار', timer:'المدة (ث)', distance:'المسافة (م)'};
  const placeholders = {reps:'10', timer:'30', distance:'200'};
  lbl.textContent = labels[type] || 'التكرار';
  inp.placeholder = placeholders[type] || '10';
}

function closeExEditor() {
  document.getElementById('ex-editor-modal').classList.remove('open');
  document.getElementById('new-ex-form').style.display = 'none';
}

/* ══════════════════════════════════════════
   ACTIVE REST DAY
══════════════════════════════════════════ */
let arIdx = 0;
let arTimerInterval = null;
let arRemain = 0;
let arTotal = 0;

function openActiveRest() {
  arIdx = 0;
  document.getElementById('active-rest-overlay').style.display = 'flex';
  arLoad();
}

function closeActiveRest() {
  clearInterval(arTimerInterval);
  document.getElementById('active-rest-overlay').style.display = 'none';
}

function arLoad() {
  clearInterval(arTimerInterval);
  const ex = STRETCH_EXERCISES[arIdx];
  const total = STRETCH_EXERCISES.length;
  document.getElementById('ar-icon').textContent = ex.icon;
  document.getElementById('ar-name').textContent = ex.nameAr;
  document.getElementById('ar-steps').innerHTML = ex.steps.map(s => '• ' + s).join('<br>');
  document.getElementById('ar-progress').textContent = (arIdx+1) + ' / ' + total;
  document.getElementById('ar-main-btn').textContent = arIdx === total-1 ? '✅ انتهى' : 'التالي ←';
  // Dots
  document.getElementById('ar-dots').innerHTML = STRETCH_EXERCISES.map((_,i) =>
    `<div style="width:${i===arIdx?20:8}px;height:8px;border-radius:4px;background:${i<=arIdx?'var(--gold)':'rgba(255,255,255,.15)'};transition:all .3s;"></div>`
  ).join('');
  arTotal = ex.dur;
  arRemain = ex.dur;
  arUpdateRing();
  // Beep
  playTone(660, 0.15, 'sine', 0.3);
  arTimerInterval = setInterval(() => {
    arRemain--;
    arUpdateRing();
    if (arRemain <= 0) {
      clearInterval(arTimerInterval);
      playTone(880, 0.2, 'sine', 0.4);
      if (arIdx < STRETCH_EXERCISES.length - 1) {
        setTimeout(() => { arIdx++; arLoad(); }, 800);
      } else {
        setTimeout(() => {
          closeActiveRest();
          showMiniToast('🧘 أحسنت! انتهت جلسة الراحة النشطة');
        }, 800);
      }
    }
  }, 1000);
}

function arUpdateRing() {
  document.getElementById('ar-timer-num').textContent = arRemain;
  const pct = arRemain / arTotal;
  const circ = 377;
  document.getElementById('ar-ring').style.strokeDashoffset = circ * (1 - pct);
  // Color transition: gold → red as time runs out
  const hue = Math.round(pct * 40); // 40=gold, 0=red
  document.getElementById('ar-ring').style.stroke = pct > 0.3 ? 'var(--gold)' : '#ef4444';
}

function arSkip() {
  if (arIdx >= STRETCH_EXERCISES.length - 1) {
    closeActiveRest();
    showMiniToast('🧘 أحسنت! انتهت جلسة الراحة النشطة');
    return;
  }
  arIdx++;
  arLoad();
}

function arPrev() {
  if (arIdx > 0) { arIdx--; arLoad(); }
}

/* ══════════════════════════════════════════
   SMART STATS
══════════════════════════════════════════ */
function renderTrainingLogSection() {
  const el = document.getElementById('training-history-list');
  if (!el) return;
  const log = Object.values(S.trainingLog || {}).sort((a,b) => b.day - a.day).slice(0, 14);
  if (!log.length) {
    el.innerHTML = '<div style="text-align:center;padding:20px;color:var(--dim);font-size:13px;">أكمل تمريناً لترى سجلك هنا</div>';
    return;
  }
  el.innerHTML = log.map(entry => `
      <div style="display:flex;align-items:center;gap:12px;padding:12px 14px;background:var(--card);border-radius:14px;border:1px solid var(--border);margin-bottom:8px;">
        <div style="width:40px;height:40px;border-radius:12px;background:linear-gradient(135deg,var(--gold),var(--gd));display:flex;flex-direction:column;align-items:center;justify-content:center;flex-shrink:0;">
          <div style="font-size:14px;font-weight:900;color:var(--night);line-height:1;">${entry.day}</div>
          <div style="font-size:8px;color:var(--night);opacity:.8;">يوم</div>
        </div>
        <div style="flex:1;min-width:0;">
          <div style="font-size:13px;font-weight:700;color:var(--txt);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">
            ${(entry.exercises||[]).slice(0,3).join(' · ')}${entry.exCount > 3 ? ' ...' : ''}
          </div>
          <div style="font-size:11px;color:var(--dim);margin-top:3px;">${entry.date||''}</div>
        </div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;gap:3px;flex-shrink:0;">
          <div style="font-size:12px;font-weight:700;color:var(--gold);">🔥 ${entry.calories||0} كالوري</div>
          <div style="font-size:11px;color:var(--dim);">⏱ ${entry.duration||0} دقيقة</div>
        </div>
      </div>`).join('');
}

function renderSmartStats() {
  const el = document.getElementById('smart-stats');
  if (!el) return;
  const log = Object.values(S.trainingLog || {}).sort((a,b) => a.day - b.day);
  if (log.length < 2) {
    el.innerHTML = '<div style="text-align:center;padding:20px;color:var(--dim);font-size:13px;">أكمل 2 أيام على الأقل لرؤية الإحصاء</div>';
    return;
  }
  // Best streak
  const streak = S.streak || 0;
  // Completion rate
  const progDays = S.user?.programDays || 30;
  const rate = Math.round((S.completedDays.length / progDays) * 100);
  // Most frequent exercise
  const exCount = {};
  log.forEach(entry => (entry.exercises||[]).forEach(n => { exCount[n] = (exCount[n]||0)+1; }));
  const topEx = Object.entries(exCount).sort((a,b)=>b[1]-a[1])[0];
  // Weekly comparison
  const thisWeek = S.completedDays.filter(d => d > progDays - 7).length;
  const lastWeek = S.completedDays.filter(d => d > progDays - 14 && d <= progDays - 7).length;
  const weekTrend = thisWeek >= lastWeek ? '📈 تحسن' : '📉 تراجع';
  // Best day of week (from dates)
  const dayCount = {0:0,1:0,2:0,3:0,4:0,5:0,6:0};
  log.forEach(e => { if(e.ts) dayCount[new Date(e.ts).getDay()]++; });
  const DAY_NAMES = ['الأحد','الاثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'];
  const bestDayIdx = Object.entries(dayCount).sort((a,b)=>b[1]-a[1])[0]?.[0];
  const bestDay = bestDayIdx !== undefined ? DAY_NAMES[bestDayIdx] : '—';
  // Average calories per session
  const avgCal = log.length ? Math.round(log.reduce((s,e)=>s+(e.calories||150),0)/log.length) : 0;

  el.innerHTML = `
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">
      <div style="background:var(--card);border-radius:14px;padding:14px;text-align:center;">
        <div style="font-size:28px;font-weight:900;color:var(--gold);">${rate}%</div>
        <div style="font-size:11px;color:var(--dim);margin-top:4px;">${window.T('smartStatsCompletion','معدل الإتمام')}</div>
      </div>
      <div style="background:var(--card);border-radius:14px;padding:14px;text-align:center;">
        <div style="font-size:28px;font-weight:900;color:var(--gold);">${avgCal}</div>
        <div style="font-size:11px;color:var(--dim);margin-top:4px;">${window.T('smartStatsAvgCal','متوسط سعرة/جلسة')}</div>
      </div>
      <div style="background:var(--card);border-radius:14px;padding:14px;text-align:center;">
        <div style="font-size:14px;font-weight:900;color:var(--txt);">${topEx ? topEx[0] : '—'}</div>
        <div style="font-size:11px;color:var(--dim);margin-top:4px;">${window.T('smartStatsFav','تمرينك المفضل')}</div>
      </div>
      <div style="background:var(--card);border-radius:14px;padding:14px;text-align:center;">
        <div style="font-size:14px;font-weight:900;color:var(--txt);">${bestDay}</div>
        <div style="font-size:11px;color:var(--dim);margin-top:4px;">${window.T('smartStatsBestDay','أفضل يوم في الأسبوع')}</div>
      </div>
      <div style="background:var(--card);border-radius:14px;padding:14px;text-align:center;grid-column:1/-1;">
        <div style="font-size:14px;font-weight:700;color:var(--txt);">${weekTrend} ${window.T('smartStatsWeekTrend','مقارنة بالأسبوع الماضي')}</div>
        <div style="font-size:12px;color:var(--dim);margin-top:4px;">${window.T('smartStatsThisWeek','هذا الأسبوع')}: ${thisWeek} | ${window.T('smartStatsLastWeek','الأسبوع الماضي')}: ${lastWeek}</div>
      </div>
    </div>`;
}

/* ══════════════════════════════════════════
   HISTORY + SHARE + NOTIFICATIONS
══════════════════════════════════════════ */
function renderHistory() {
  // Unified with renderTrainingLogSection to avoid duplication
  renderTrainingLogSection();
}

function shareProgress() {
  const done = (S.completedDays||[]).length;
  const total = S.user?.programDays || 30;
  const pct = Math.round(done/total*100);
  const streak = S.streak || 0;
  const name = S.user?.name || 'بطل';
  const text = `💪 ${name} أكمل ${done} يوم من ${total} في AZEM (عزم)!
🔥 سلسلة ${streak} يوم متواصل
⚡ ${S.calories||0} سعرة محترقة
📊 التقدم: ${pct}%
${pct >= 50 ? '🏆 أكثر من النصف — استمر!' : '🚀 في البداية — لا تتوقف!'}`;
  if (navigator.share) {
    navigator.share({ title: 'AZEM (عزم)', text }).catch(()=>{});
  } else {
    navigator.clipboard?.writeText(text).then(()=> showMiniToast('✅ تم نسخ التقدم!'));
  }
}

// Notifications
let notifInterval = null;
function toggleNotifSwitch() {
  const toggle = document.getElementById('notif-toggle');
  const knob = document.getElementById('notif-toggle-knob');
  const sub = document.getElementById('notif-toggle-sub');
  const isOn = toggle.dataset.on === '1';
  if (isOn) {
    // Turn off
    toggle.dataset.on = '0';
    toggle.style.background = 'rgba(100,116,139,.3)';
    knob.style.transform = 'translateX(0)';
    if (sub) sub.textContent = 'اضغط لتفعيل التنبيهات';
    return;
  }
  // Turn on → request permission
  requestNotifPerm();
}
function _setNotifToggleOn() {
  const toggle = document.getElementById('notif-toggle');
  const knob = document.getElementById('notif-toggle-knob');
  const sub = document.getElementById('notif-toggle-sub');
  if (toggle) { toggle.dataset.on = '1'; toggle.style.background = 'var(--green,#22c55e)'; }
  if (knob) knob.style.transform = 'translateX(-24px)';
  if (sub) sub.textContent = '✅ التذكير مفعّل';
}
function requestNotifPerm() {
  if (!('Notification' in window)) {
    showMiniToast('المتصفح لا يدعم الإشعارات'); return;
  }
  Notification.requestPermission().then(perm => {
    if (perm === 'granted') {
      showMiniToast('✅ سيصلك تذكير في وقت تدريبك');
      scheduleTrainingNotif();
      _setNotifToggleOn();
    } else {
      showMiniToast('⚠️ الإشعارات غير مسموحة');
    }
  });
}

function scheduleTrainingNotif() {
  // FIX#9: notify via SW postMessage so it works when app is closed
  const trainTime = S.user?.trainTime;
  if (trainTime && 'serviceWorker' in navigator && navigator.serviceWorker.controller) {
    navigator.serviceWorker.controller.postMessage({
      type: 'SCHEDULE_REMINDER',
      trainTime,
      name: S.user?.name || 'بطل',
      day: S.currentDay
    });
  }
  // Fallback: setInterval for when app is open in browser
  if (notifInterval) clearInterval(notifInterval);
  notifInterval = setInterval(() => {
    const t = S.user?.trainTime;
    if (!t || Notification.permission !== 'granted') return;
    const now = new Date();
    const [h, m] = t.split(':').map(Number);
    if (now.getHours() === h && now.getMinutes() === m) {
      new Notification('AZEM (عزم) 🔥', {
        body: `${S.user?.name||'بطل'}، وقت تدريبك! اليوم ${S.currentDay} في انتظارك 💪`,
        icon: './icon-192-1.png',
        tag: 'azem-reminder',   // prevents duplicate toasts
        renotify: false
      });
    }
  }, 60000);
}

// Auto-schedule if permission already granted
if (typeof Notification !== 'undefined' && Notification.permission === 'granted') {
  scheduleTrainingNotif();
  setTimeout(() => {
    const btn = document.getElementById('notif-btn');
    if (btn) { btn.textContent = '✅ التذكير مفعّل'; }
  }, 500);
}


/* ══════════════════════════════════════════
   ONBOARDING
══════════════════════════════════════════ */
const OB_STEPS = [
  { id:'welcome', type:'welcome' },
  { id:'lang',    type:'lang' },
  { id:'name',    type:'input',  label:'ما اسمك؟',           field:'name',       placeholder:'مثال: أحمد',  kbd:'text'   },
  { id:'weight',  type:'input',  label:'وزنك الحالي (كغ)',   field:'weight',     placeholder:'مثال: 74',    kbd:'numeric'},
  { id:'height',  type:'input',  label:'طولك (سم)',           field:'height',     placeholder:'مثال: 177',   kbd:'numeric'},
  { id:'goal',    type:'choice', label:'ما هدفك الأساسي؟',   field:'goal',
    options:[
      {val:'burn',    icon:'🔥', label:'حرق الدهون'},
      {val:'muscle',  icon:'💪', label:'بناء العضلات'},
      {val:'fitness', icon:'🏃', label:'تحسين اللياقة'},
      {val:'health',  icon:'❤️', label:'الصحة العامة'},
    ]},
  { id:'days',    type:'choice', label:'كم يوماً برنامجك؟',  field:'programDays',
    options:[
      {val:15,  icon:'⚡', label:'15 يوم'},
      {val:30,  icon:'🗓️', label:'30 يوم'},
      {val:60,  icon:'🏆', label:'60 يوم'},
      {val:'custom', icon:'✏️', label:'تخصيص'},
    ]},
  { id:'time',    type:'input',  label:'وقت تدريبك المفضل',  field:'trainTime',  placeholder:'18:00', kbd:'time'},
  { id:'startDate', type:'date',  label:'متى تبدأ برنامجك؟',  field:'startDate' },
  { id:'install', type:'install' },
  { id:'aikey',   type:'aikey' },
  { id:'done',    type:'done' },
];
let obStep = 0;

// Shared helper — avoids duplicating the label map in 3 functions
function getObLabelMap() {
  return {
    name:      window.T ? window.T('obName','ما اسمك؟')                    : 'ما اسمك؟',
    weight:    window.T ? window.T('obWeight','وزنك (كغ)')                  : 'وزنك (كغ)',
    height:    window.T ? window.T('obHeight','طولك (سم)')                  : 'طولك (سم)',
    goal:      window.T ? window.T('obGoal','ما هدفك الأساسي؟')             : 'ما هدفك الأساسي؟',
    days:      window.T ? window.T('obDays','كم يوماً برنامجك؟')            : 'كم يوماً برنامجك؟',
    time:      window.T ? window.T('obTime','وقت تدريبك المفضل')            : 'وقت تدريبك المفضل',
    startDate: window.T ? window.T('obStart','متى تبدأ برنامجك؟')           : 'متى تبدأ برنامجك؟',
  };
}

function obInstallNow() {
  if (deferredPrompt) {
    deferredPrompt.prompt();
    deferredPrompt.userChoice.then(() => {
      deferredPrompt = null;
      const btn = document.getElementById('ob-inst-btn');
      if (btn) { btn.textContent = '\u2705 '+(currentLang==='en'?'Installed!':currentLang==='fr'?'Install\u00e9!':'\u062a\u0645 \u0627\u0644\u062a\u062b\u0628\u064a\u062a!'); btn.onclick=null; btn.style.opacity='0.7'; }
      setTimeout(obNext, 1200);
    });
  } else {
    const _isEn = currentLang==='en', _isFr = currentLang==='fr';
    const T = (ar,en,fr) => _isEn?en:_isFr?fr:ar;
    const btn = document.getElementById('ob-inst-btn');
    if (btn) btn.style.display = 'none';
    const existing = document.getElementById('ob-install-fallback');
    if (existing) return;
    const div = document.createElement('div');
    div.id = 'ob-install-fallback';
    div.style.cssText = 'background:var(--card);border-radius:16px;padding:16px;border:1px solid rgba(212,168,67,.3);text-align:right;margin-top:8px;';
    div.innerHTML = [['1','\u22ee',T('\u0627\u0636\u063a\u0637 \u0627\u0644\u0646\u0642\u0627\u0637 \u0627\u0644\u062b\u0644\u0627\u062b \u0641\u064a Chrome','Tap the 3 dots in Chrome','Appuie sur les 3 points')],['2','\ud83d\udcf2',T('\u0627\u062e\u062a\u0631 \u0625\u0636\u0627\u0641\u0629 \u0625\u0644\u0649 \u0627\u0644\u0634\u0627\u0634\u0629 \u0627\u0644\u0631\u0626\u064a\u0633\u064a\u0629','Choose Add to Home Screen','Choisir Ajouter')],['3','\u2705',T('\u0627\u0636\u063a\u0637 \u0625\u0636\u0627\u0641\u0629','Tap Add','Appuie')]].map(([n,ic,t])=>'<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;"><div style="width:26px;height:26px;border-radius:50%;background:var(--gold);color:var(--night);font-size:12px;font-weight:900;display:flex;align-items:center;justify-content:center;flex-shrink:0;">'+n+'</div><div style="font-size:13px;color:var(--txt);">'+ic+' '+t+'</div></div>').join('');
    if (btn) btn.insertAdjacentElement('afterend', div);
  }
}
function obSaveApiKey() {
  const val = (document.getElementById('ob-apikey-inp')?.value||'').trim();
  if (!val.startsWith('gsk_')) { showMiniToast(currentLang==='en'?'Key must start with gsk_':currentLang==='fr'?'La clé doit commencer par gsk_':'المفتاح يجب أن يبدأ بـ gsk_'); return; }
  S.apiKey = val; saveState();
  showMiniToast(currentLang==='en'?'✅ Key saved!':currentLang==='fr'?'✅ Clé sauvegardée!':'✅ تم حفظ المفتاح!');
  setTimeout(obNext, 700);
}
function obRequestNotif() {
  if (!('Notification' in window)) return;
  const timeVal = document.getElementById('ob-inp')?.value;
  if (timeVal) { S.user.trainTime = timeVal; saveState(); }
  Notification.requestPermission().then(perm => {
    const btn = document.getElementById('ob-notif-btn');
    if (perm === 'granted') {
      scheduleTrainingNotif();
      showMiniToast(currentLang==='en'?'✅ Reminder on!':currentLang==='fr'?'✅ Rappel activé!':'✅ سيصلك تذكير يومياً');
      if (btn) btn.outerHTML = `<div style="display:flex;align-items:center;gap:8px;padding:10px 14px;border-radius:12px;background:rgba(74,222,128,.1);border:1px solid rgba(74,222,128,.3);"><span>✅</span><span style="font-size:13px;font-weight:700;color:#4ade80;">${currentLang==='en'?'Reminder activated!':currentLang==='fr'?'Rappel activé!':'التذكير مفعّل!'}</span></div>`;
    } else {
      if (btn) { btn.textContent='⚠️ '+(currentLang==='en'?'Blocked':'محظورة'); btn.style.borderColor='#ef4444'; btn.style.color='#f87171'; btn.onclick=null; }
    }
  });
}
function obRefreshInstallStep() {
  if (OB_STEPS[obStep]?.type === 'install') {
    // Try to upgrade waiting div to a real install button
    const wait = document.getElementById('ob-inst-wait');
    if (wait && deferredPrompt) {
      wait.outerHTML = `<button id="ob-inst-btn" onclick="obInstallNow()" style="width:100%;padding:22px;border-radius:20px;background:linear-gradient(135deg,var(--gl),var(--gd));color:var(--night);font-family:'Cairo',sans-serif;font-size:20px;font-weight:900;cursor:pointer;border:none;box-shadow:0 10px 36px var(--glow-strong);letter-spacing:.3px;transition:transform .12s;" onmousedown="this.style.transform='scale(0.96)'" onmouseup="this.style.transform='scale(1)'" ontouchstart="this.style.transform='scale(0.96)'" ontouchend="this.style.transform='scale(1)'">
        📲 ${window.T ? window.T('installNowBtn','تثبيت الآن') : 'تثبيت الآن'}
      </button>`;
    } else {
      renderObStep();
    }
  }
}
function showOnboarding() {
  obStep = 0;
  document.getElementById('onboarding').style.display = 'flex';
  renderObStep();
}

function obSelectLang(code) {
  currentLang = code;
  S.lang = code;
  // Apply lang immediately so user sees translated steps
  applyLang(code);
  // Re-render current step to show checkmark
  renderObStep();
  // Auto-advance after short delay
  setTimeout(obNext, 400);
}

function obNext() {
  const _rawStep = OB_STEPS[obStep];
  const _obLabelMap = getObLabelMap();
  const step = { ..._rawStep, label: _obLabelMap[_rawStep.id] || _rawStep.label };
  if (step.type === 'input' || step.type === 'date') {
    const val = document.getElementById('ob-inp')?.value?.trim();
    if (step.type === 'input' && !val) { document.getElementById('ob-inp')?.focus(); return; }
    if (val) S.user[step.field] = (step.kbd === 'numeric') ? parseFloat(val)||0 : val;
  }
  if (step.type === 'done') {
    S.onboardingDone = true;
    saveState();
    document.getElementById('onboarding').style.display = 'none';
    render();
    setTimeout(startTutorial, 800);
    return;
  }
  obStep = Math.min(obStep + 1, OB_STEPS.length - 1);
  renderObStep();
}

function obSkip() {
  // Jump to the login/done step rather than immediately closing
  const doneIdx = OB_STEPS.findIndex(s => s.type === 'done');
  if (doneIdx >= 0 && obStep < doneIdx) {
    obStep = doneIdx;
    renderObStep();
    return;
  }
  // Already on done step or no done step — close
  S.onboardingDone = true;
  saveState();
  document.getElementById('onboarding').style.display = 'none';
  render();
}

function obChoose(val) {
  const _rawStep = OB_STEPS[obStep];
  const _obLabelMap = getObLabelMap();
  const step = { ..._rawStep, label: _obLabelMap[_rawStep.id] || _rawStep.label };
  if (val === 'custom') {
    // Switch this step to input mode for custom days
    const container = document.getElementById('ob-steps');
    container.innerHTML = `<div style="min-height:80vh;display:flex;flex-direction:column;justify-content:center;padding:40px 24px;">
      <div style="font-size:22px;font-weight:900;color:var(--txt);margin-bottom:28px;text-align:center;">أدخل عدد الأيام</div>
      <input id="ob-inp" type="number" min="7" max="365" placeholder="مثال: 45"
        style="width:100%;padding:18px 20px;border-radius:16px;background:var(--card);border:2px solid var(--gold);color:var(--txt);font-family:'Cairo',sans-serif;font-size:20px;font-weight:700;text-align:center;box-sizing:border-box;outline:none;"
        onkeydown="if(event.key==='Enter'){S.user.programDays=parseInt(this.value)||30;obStep=Math.min(obStep+1,OB_STEPS.length-1);renderObStep();}">
    </div>`;
    document.getElementById('ob-footer').style.display = 'flex';
    document.getElementById('ob-next-btn').onclick = () => {
      const v = parseInt(document.getElementById('ob-inp')?.value)||30;
      S.user.programDays = v;
      obStep = Math.min(obStep + 1, OB_STEPS.length - 1);
      document.getElementById('ob-next-btn').onclick = obNext;
      renderObStep();
    };
    setTimeout(()=>document.getElementById('ob-inp')?.focus(), 100);
    return;
  }
  S.user[step.field] = val;
  obStep = Math.min(obStep + 1, OB_STEPS.length - 1);
  renderObStep();
}

function renderObStep() {
  const _rawStep = OB_STEPS[obStep];
  const _obLabelMap = getObLabelMap();
  const step = { ..._rawStep, label: _obLabelMap[_rawStep.id] || _rawStep.label };
  const isLast = obStep === OB_STEPS.length - 1;
  document.getElementById('ob-next-btn').textContent = isLast ? '🚀 ابدأ الرحلة' : 'التالي ←';
  document.getElementById('ob-footer').style.display = (step.type === 'choice' || step.type === 'lang') ? 'none' : 'flex';

  const dataSteps = OB_STEPS.filter(s => s.type !== 'welcome' && s.type !== 'done');
  document.getElementById('ob-dots').innerHTML = dataSteps.map((s, i) => {
    // dataSteps[0] corresponds to obStep=1 (after welcome), etc.
    const stepObIdx = i + 1;
    const active = stepObIdx === obStep;
    const done = stepObIdx < obStep;
    return `<div style="width:${active?20:8}px;height:8px;border-radius:4px;background:${done||active?'var(--gold)':'rgba(255,255,255,.15)'};transition:all .3s;"></div>`;
  }).join('');

  let html = '';
  if (step.type === 'welcome') {
    html = `<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:40px 24px;text-align:center;">
      <div style="font-size:80px;margin-bottom:24px;animation:iconPulse 2s ease-in-out infinite;">⚡</div>
      <div style="font-size:28px;font-weight:900;color:var(--gold);margin-bottom:8px;">مرحباً بك في</div>
      <div style="font-size:36px;font-weight:900;color:var(--txt);margin-bottom:20px;">AZEM (عزم)</div>
      <div style="font-size:13px;color:var(--dim);line-height:2;max-width:290px;margin-bottom:28px;">تدرّب أينما كنت — منزل، جيم، أو هواء طلق.<br>برنامجك الشخصي، بأسلوبك.</div>
      <div style="display:flex;flex-direction:column;gap:10px;width:100%;max-width:300px;">
        ${[
          ['🤖','مدرب ذكاء اصطناعي يعرف حالتك','يعدّل برنامجك ويجيب على أسئلتك'],
          ['🏋️','منزل، جيم، أو بدون معدات','كل أنواع التمارين في مكان واحد'],
          ['📅','برنامج مرن — أسبوع لثلاثة أشهر','أنت تحدد المدة والهدف'],
          ['📊','تتبع تقدمك يومياً','إحصاء، شارات، وسجل كامل']
        ].map(([ic,t,s])=>
          `<div style="display:flex;align-items:center;gap:14px;background:var(--card);border-radius:16px;padding:14px 16px;border:1px solid rgba(212,168,67,.12);text-align:right;">
            <span style="font-size:26px;flex-shrink:0;">${ic}</span>
            <div><div style="font-size:13px;font-weight:800;color:var(--txt);">${t}</div><div style="font-size:11px;color:var(--dim);margin-top:2px;">${s}</div></div>
          </div>`
        ).join('')}
      </div>
    </div>`;
  } else if (step.type === 'lang') {
    const langs = [
      { code:'ar', flag:'🇩🇿', name:'العربية',    native:'العربية' },
      { code:'en', flag:'🇺🇸', name:'English',    native:'English' },
      { code:'fr', flag:'🇫🇷', name:'Français',   native:'Français' },
    ];
    html = `<div style="min-height:80vh;display:flex;flex-direction:column;justify-content:center;padding:40px 24px;">
      <div style="font-size:40px;text-align:center;margin-bottom:16px;">🌐</div>
      <div style="font-size:22px;font-weight:900;color:var(--txt);margin-bottom:8px;text-align:center;">اختر لغتك</div>
      <div style="font-size:22px;font-weight:900;color:var(--txt);margin-bottom:8px;text-align:center;">Choose your language</div>
      <div style="font-size:22px;font-weight:900;color:var(--txt);margin-bottom:28px;text-align:center;">Choisissez votre langue</div>
      <div style="display:flex;flex-direction:column;gap:12px;">
        ${langs.map(l => {
          const active = (S.lang || currentLang || 'ar') === l.code;
          return `<button onclick="obSelectLang('${l.code}')" style="display:flex;align-items:center;gap:16px;padding:18px 20px;border-radius:18px;background:${active ? 'rgba(212,168,67,.15)' : 'var(--card)'};border:2px solid ${active ? 'var(--gold)' : 'var(--border)'};cursor:pointer;width:100%;transition:all .2s;">
            <span style="font-size:32px;">${l.flag}</span>
            <div style="text-align:right;flex:1;">
              <div style="font-size:17px;font-weight:900;color:${active ? 'var(--gold)' : 'var(--txt)'};">${l.native}</div>
              <div style="font-size:12px;color:var(--dim);">${l.name}</div>
            </div>
            ${active ? '<span style="font-size:20px;color:var(--gold);">✓</span>' : ''}
          </button>`;
        }).join('')}
      </div>
    </div>`;
  } else if (step.type === 'date') {
    const today = new Date().toISOString().split('T')[0];
    html = `<div style="min-height:80vh;display:flex;flex-direction:column;justify-content:center;padding:40px 24px;">
      <div style="font-size:22px;font-weight:900;color:var(--txt);margin-bottom:28px;text-align:center;">${step.label}</div>
      <input id="ob-inp" type="date" value="${S.user[step.field]||today}"
        style="width:100%;padding:18px 20px;border-radius:16px;background:var(--card);border:1.5px solid var(--border);
        color:var(--txt);font-family:'Cairo',sans-serif;font-size:18px;box-sizing:border-box;outline:none;text-align:center;
        -webkit-appearance:none;"
        onkeydown="if(event.key==='Enter')obNext()">
      <div style="margin-top:16px;text-align:center;font-size:13px;color:var(--dim);">اختر تاريخ بدء برنامجك</div>
    </div>`;
  } else if (step.type === 'input') {
    const isTimeStep = step.kbd === 'time';
    const _isAr = currentLang==='ar', _isEn = currentLang==='en';
    const _T3 = (ar,en,fr) => _isAr?ar:_isEn?en:fr;
    const notifSupported = ('Notification' in window);
    const notifGranted = notifSupported && Notification.permission === 'granted';
    const notifDenied  = notifSupported && Notification.permission === 'denied';

    const notifBlock = isTimeStep ? `
      <div style="margin-top:20px;background:var(--card);border-radius:16px;padding:16px;border:1.5px solid rgba(212,168,67,.2);">
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">
          <span style="font-size:22px;">🔔</span>
          <div>
            <div style="font-size:13px;font-weight:800;color:var(--txt);">${_T3('تذكير التدريب اليومي','Daily Training Reminder','Rappel quotidien')}</div>
            <div style="font-size:11px;color:var(--dim);margin-top:2px;">${_T3('سنذكّرك في وقتك المحدد كل يوم','We\'ll notify you at your chosen time daily','On te rappelle chaque jour à l\'heure choisie')}</div>
          </div>
        </div>
        ${notifGranted ? `
          <div style="display:flex;align-items:center;gap:8px;padding:10px 14px;border-radius:12px;background:rgba(74,222,128,.1);border:1px solid rgba(74,222,128,.3);">
            <span>✅</span><span style="font-size:13px;font-weight:700;color:#4ade80;">${_T3('التذكير مفعّل!','Reminder activated!','Rappel activé!')}</span>
          </div>
        ` : notifDenied ? `
          <div style="padding:10px 14px;border-radius:12px;background:rgba(239,68,68,.08);border:1px solid rgba(239,68,68,.2);font-size:11px;color:#f87171;">
            ${_T3('⚠️ الإشعارات محظورة — فعّلها من إعدادات المتصفح','⚠️ Notifications blocked — enable in browser settings','⚠️ Notifications bloquées — activez dans les paramètres')}
          </div>
        ` : `
          <button id="ob-notif-btn" onclick="obRequestNotif()" style="width:100%;padding:13px;border-radius:12px;background:linear-gradient(135deg,rgba(212,168,67,.18),rgba(212,168,67,.06));border:1.5px solid var(--gold);color:var(--gold);font-family:'Cairo',sans-serif;font-size:14px;font-weight:800;cursor:pointer;">
            🔔 ${_T3('تفعيل التذكيرات','Enable Reminders','Activer les rappels')}
          </button>
        `}
      </div>` : '';

    html = `<div style="min-height:80vh;display:flex;flex-direction:column;justify-content:center;padding:40px 24px;">
      <div style="font-size:22px;font-weight:900;color:var(--txt);margin-bottom:28px;text-align:center;">${step.label}</div>
      <input id="ob-inp" type="${step.kbd==='numeric'?'number':step.kbd==='time'?'time':'text'}"
        value="${S.user[step.field]||''}" placeholder="${step.placeholder}"
        style="width:100%;padding:18px 20px;border-radius:16px;background:var(--card);border:2px solid var(--gold);color:var(--txt);font-family:'Cairo',sans-serif;font-size:20px;font-weight:700;text-align:center;box-sizing:border-box;outline:none;"
        onkeydown="if(event.key==='Enter')obNext()">
      ${notifBlock}
    </div>`;

  } else if (step.type === 'install') {
    const _ios = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
    const _standalone = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone;
    const _android = /android/i.test(navigator.userAgent);
    const _isAr = currentLang==='ar', _isEn = currentLang==='en';
    const _T3 = (ar,en,fr) => _isAr?ar:_isEn?en:fr;

    if (_standalone) { obStep = Math.min(obStep+1, OB_STEPS.length-1); renderObStep(); return; }

    let _actionHTML;
    if (deferredPrompt) {
      _actionHTML = `
      <button id="ob-inst-btn" onclick="obInstallNow()" style="width:100%;padding:22px;border-radius:20px;background:linear-gradient(135deg,var(--gl),var(--gd));color:var(--night);font-family:'Cairo',sans-serif;font-size:20px;font-weight:900;cursor:pointer;border:none;box-shadow:0 10px 36px var(--glow-strong),0 2px 8px rgba(0,0,0,.15);letter-spacing:.3px;transition:transform .12s,box-shadow .12s;" onmousedown="this.style.transform='scale(0.96)';this.style.boxShadow='0 4px 16px var(--glow)'" onmouseup="this.style.transform='scale(1)';this.style.boxShadow='0 10px 36px var(--glow-strong)'" ontouchstart="this.style.transform='scale(0.96)'" ontouchend="this.style.transform='scale(1)'">
        📲 ${_T3('تثبيت AZEM — نقرة واحدة','Install AZEM — One Tap','Installer AZEM — Un clic')}
      </button>
      <p style="font-size:11px;color:var(--dim);text-align:center;margin:10px 0 0;">${_T3('يضيف أيقونة على شاشتك ويعمل بدون إنترنت','Adds icon to home screen & works offline','Icône sur écran, fonctionne hors ligne')}</p>`;
    } else if (_ios) {
      _actionHTML = `<div style="background:var(--card);border-radius:18px;padding:18px;border:1.5px solid rgba(212,168,67,.3);">
        <div style="font-size:13px;font-weight:900;color:var(--gold);margin-bottom:14px;text-align:center;">📱 ${_T3('تثبيت على iPhone / iPad','Install on iPhone / iPad','Installer sur iPhone / iPad')}</div>
        ${[['1','📤',_T3('اضغط زر المشاركة في الأسفل','Tap the Share button at the bottom','Appuie sur Partager en bas')],['2','🏠',_T3('اختر "إضافة إلى الشاشة الرئيسية"','Choose "Add to Home Screen"','Choisir "Sur l’écran d’accueil"')],['3','✅',_T3('اضغط "إضافة"','Tap "Add"','Appuie sur "Ajouter"')]].map(([n,ic,t])=>`<div style="display:flex;align-items:center;gap:12px;margin-bottom:10px;"><div style="width:26px;height:26px;border-radius:50%;background:var(--gold);color:var(--night);font-size:12px;font-weight:900;display:flex;align-items:center;justify-content:center;flex-shrink:0;">${n}</div><div style="font-size:12px;color:var(--txt);">${ic} ${t}</div></div>`).join('')}
      </div>`;
    } else if (_android) {
      // Android but no deferredPrompt yet — show both: a waiting button + manual instructions
      _actionHTML = `
      <button id="ob-inst-btn" onclick="obInstallNow()" style="width:100%;padding:22px;border-radius:20px;background:linear-gradient(135deg,var(--gl),var(--gd));color:var(--night);font-family:'Cairo',sans-serif;font-size:18px;font-weight:900;cursor:pointer;border:none;box-shadow:0 10px 36px var(--glow-strong);margin-bottom:14px;">
        📲 ${_T3('تثبيت AZEM','Install AZEM','Installer AZEM')}
      </button>
      <div style="background:var(--card);border-radius:16px;padding:14px;border:1px solid rgba(212,168,67,.2);text-align:right;">
        <div style="font-size:11px;font-weight:900;color:var(--gold);margin-bottom:10px;">${_T3('أو ثبّت يدوياً:','Or install manually:','Ou installez manuellement:')}</div>
        ${[['1','⋮',_T3('اضغط النقاط الثلاث في Chrome','Tap the 3 dots in Chrome','Appuie sur les 3 points dans Chrome')],['2','📲',_T3('اختر \"إضافة إلى الشاشة الرئيسية\"','Choose \"Add to Home Screen\"','Choisir \"Ajouter à l\'écran\"')],['3','✅',_T3('اضغط إضافة','Tap Add','Appuie sur Ajouter')]].map(([n,ic,t])=>`<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;"><div style="width:22px;height:22px;border-radius:50%;background:var(--gold);color:var(--night);font-size:11px;font-weight:900;display:flex;align-items:center;justify-content:center;flex-shrink:0;">${n}</div><div style="font-size:11px;color:var(--txt);">${ic} ${t}</div></div>`).join('')}
      </div>`;
    } else {
      // Desktop or unknown
      _actionHTML = `<div style="background:var(--card);border-radius:16px;padding:18px;border:1px solid rgba(212,168,67,.2);text-align:center;">
        <div style="font-size:13px;color:var(--dim);line-height:1.8;">${_T3('افتح الموقع في Chrome أو Edge على جهازك وستظهر نافذة التثبيت تلقائياً','Open in Chrome or Edge on your device for automatic install','Ouvre dans Chrome ou Edge sur ton appareil')}</div>
      </div>`;
    }

    html = `<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:36px 20px;text-align:center;">
      <div style="font-size:70px;margin-bottom:10px;animation:iconPulse 2s ease-in-out infinite;">📲</div>
      <div style="font-size:24px;font-weight:900;color:var(--gold);margin-bottom:8px;">${_T3('ثبّت التطبيق على جهازك','Install the App','Installer l’application')}</div>
      <div style="font-size:13px;color:var(--dim);line-height:1.8;margin-bottom:22px;max-width:300px;">${_T3('أسرع وصول — يعمل بدون إنترنت — تذكيرات يومية','Faster access — works offline — daily reminders','Accès rapide — hors ligne — rappels quotidiens')}</div>
      <div style="display:flex;justify-content:center;gap:20px;margin-bottom:22px;">
        ${[['⚡',_T3('وصول فوري','Instant','Instantané')],['📴',_T3('بدون نت','Offline','Hors ligne')],['🔔',_T3('تذكيرات','Reminders','Rappels')]].map(([ic,lb])=>`<div style="display:flex;flex-direction:column;align-items:center;gap:5px;"><div style="width:50px;height:50px;background:var(--card);border-radius:14px;display:flex;align-items:center;justify-content:center;font-size:22px;border:1px solid rgba(212,168,67,.2);">${ic}</div><div style="font-size:10px;color:var(--dim);font-weight:700;">${lb}</div></div>`).join('')}
      </div>
      <div style="width:100%;max-width:340px;">
        ${_actionHTML}
        <button onclick="obNext()" style="width:100%;margin-top:12px;padding:12px;border-radius:14px;background:transparent;border:1px solid var(--border);color:var(--dim);font-family:'Cairo',sans-serif;font-size:13px;cursor:pointer;">
          ${_T3('تخطي — سأثبّت لاحقاً','Skip — I’ll install later','Passer — j’installerai plus tard')}
        </button>
      </div>
    </div>`;
    document.getElementById('ob-footer').style.display = 'none';
    document.getElementById('ob-steps').innerHTML = html;
    return;

  } else if (step.type === 'aikey') {
    const _isAr = currentLang==='ar', _isEn = currentLang==='en';
    const _T3 = (ar,en,fr) => _isAr?ar:_isEn?en:fr;
    const hasKey = !!(S.apiKey && S.apiKey.startsWith('gsk_'));
    html = `<div style="min-height:100vh;display:flex;flex-direction:column;padding:28px 20px;overflow-y:auto;">
      <div style="text-align:center;margin-bottom:18px;">
        <div style="font-size:50px;margin-bottom:8px;">🤖</div>
        <div style="font-size:22px;font-weight:900;color:var(--gold);">${_T3('المدرب الذكي','AI Coach','Coach IA')}</div>
        <div style="font-size:12px;color:var(--dim);margin-top:4px;">${_T3('اعرف الفرق بين الوضعين','Understand both modes','Comprendre les deux modes')}</div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px;">
        <div style="background:var(--card);border-radius:16px;padding:14px;border:1.5px solid var(--border);">
          <div style="font-size:11px;font-weight:900;color:var(--dim);margin-bottom:8px;text-align:center;">🤖 ${_T3('بدون مفتاح','No Key','Sans clé')}</div>
          ${[['✅',_T3('مجاني 100%','100% Free','Gratuit')],['✅',_T3('يعمل بدون نت','Works offline','Hors ligne')],['⚠️',_T3('إجابات محددة مسبقاً','Pre-set answers','Réponses prédéfinies')],['⚠️',_T3('لا يفهم أسئلة حرة','No free questions','Pas de questions libres')]].map(([ic,t])=>`<div style="display:flex;gap:5px;font-size:10px;color:var(--dim);margin-bottom:5px;">${ic} ${t}</div>`).join('')}
        </div>
        <div style="background:linear-gradient(135deg,rgba(212,168,67,.12),rgba(212,168,67,.04));border-radius:16px;padding:14px;border:1.5px solid rgba(212,168,67,.4);">
          <div style="font-size:11px;font-weight:900;color:var(--gold);margin-bottom:8px;text-align:center;">⚡ ${_T3('بمفتاح Groq','With Groq Key','Avec clé Groq')}</div>
          ${[['✅',_T3('ذكاء اصطناعي حقيقي','Real AI (Llama 70B)','IA réelle')],['✅',_T3('يفهم أي سؤال','Understands anything','Comprend tout')],['✅',_T3('يعدّل برنامجك','Adjusts your plan','Modifie le plan')],['✅',_T3('مفتاح مجاني!','Free key!','Clé gratuite!')]].map(([ic,t])=>`<div style="display:flex;gap:5px;font-size:10px;color:var(--dim);margin-bottom:5px;">${ic} ${t}</div>`).join('')}
        </div>
      </div>
      <div style="background:var(--card);border-radius:12px;padding:12px;margin-bottom:12px;border:1px solid var(--border);">
        <div style="font-size:10px;color:var(--dim);margin-bottom:7px;">💬 ${_T3('مثال: "أنا تعبان اليوم"','Example: "I’m exhausted today"','Exemple: "Je suis épuisé"')}</div>
        <div style="background:rgba(255,255,255,.04);border-radius:8px;padding:7px 10px;font-size:10px;color:var(--dim);margin-bottom:5px;line-height:1.5;">🤖 "${_T3('الراحة جزء من التدريب! استمر!','Rest is part of training! Keep going!','Le repos fait partie de l’entraînement!')}"</div>
        <div style="background:rgba(212,168,67,.07);border-radius:8px;padding:7px 10px;font-size:10px;color:var(--dim);line-height:1.5;border:1px solid rgba(212,168,67,.2);">⚡ "${_T3('بناءً على وزنك ويومك الحالي، التعب طبيعي. أقترح تخفيف الشدة اليوم. تريد أعدّل غداً؟','Based on your weight & today’s workout, fatigue is normal. I suggest reducing intensity. Reschedule tomorrow?','Vu ton profil, la fatigue est normale. Je suggère de réduire l’intensité. Réplanifier demain?')}"</div>
      </div>
      <div style="background:var(--card);border-radius:14px;padding:14px;margin-bottom:10px;border:1.5px solid rgba(212,168,67,.25);">
        <div style="font-size:12px;font-weight:800;color:var(--gold);margin-bottom:3px;">🔑 ${_T3('مفتاح Groq المجاني','Free Groq Key','Clé Groq Gratuite')}</div>
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;gap:8px;">
          <div style="font-size:10px;color:var(--dim);">${_T3('30 ثانية من','30 sec at','30 sec sur')} <span style="color:var(--gold);font-weight:700;">console.groq.com</span></div>
          <button onclick="showGroqHowTo()" style="padding:5px 10px;border-radius:8px;background:rgba(212,168,67,.12);border:1px solid rgba(212,168,67,.3);color:var(--gold);font-family:'Cairo',sans-serif;font-size:10px;font-weight:700;cursor:pointer;flex-shrink:0;">❓ ${_T3('كيف أحصل عليه؟','How to get it?','Comment l’obtenir?')}</button>
        </div>
        <div style="display:flex;gap:8px;">
          <input id="ob-apikey-inp" type="text" placeholder="gsk_..." value="${S.apiKey||''}"
            style="flex:1;padding:10px 12px;border-radius:10px;background:rgba(255,255,255,.05);border:1.5px solid ${hasKey?'#4ade80':'var(--border)'};color:var(--txt);font-family:monospace;font-size:12px;outline:none;min-width:0;"
            oninput="const v=this.value.trim();this.style.borderColor=v.startsWith('gsk_')?'#4ade80':'var(--border)';document.getElementById('ob-key-save').style.opacity=v.startsWith('gsk_')?'1':'0.4';">
          <button id="ob-key-save" onclick="obSaveApiKey()" style="padding:10px 14px;border-radius:10px;background:linear-gradient(135deg,var(--gl),var(--gd));color:var(--night);font-family:'Cairo',sans-serif;font-size:12px;font-weight:900;cursor:pointer;opacity:${hasKey?'1':'0.4'};flex-shrink:0;">
            ${_T3('حفظ','Save','Sauver')}
          </button>
        </div>
        ${hasKey?`<div style="margin-top:6px;font-size:11px;color:#4ade80;">✅ ${_T3('المفتاح محفوظ!','Key saved!','Clé sauvegardée!')}</div>`:''}
      </div>
      <button onclick="obNext()" style="width:100%;padding:12px;border-radius:14px;background:transparent;border:1px solid var(--border);color:var(--dim);font-family:'Cairo',sans-serif;font-size:13px;cursor:pointer;">
        ${_T3('متابعة بدون مفتاح','Continue without key','Continuer sans clé')}
      </button>
    </div>`;
    document.getElementById('ob-footer').style.display = 'none';
    document.getElementById('ob-steps').innerHTML = html;
    return;

  } else if (step.type === 'choice') {
    html = `<div style="min-height:80vh;display:flex;flex-direction:column;justify-content:center;padding:40px 24px;">
      <div style="font-size:22px;font-weight:900;color:var(--txt);margin-bottom:28px;text-align:center;">${step.label}</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;">
        ${step.options.map(o=>`
          <button onclick="obChoose(${typeof o.val==='number'?o.val:`'${o.val}'`})"
            style="display:flex;flex-direction:column;align-items:center;gap:10px;padding:24px 16px;border-radius:18px;background:${S.user[step.field]===o.val?'rgba(212,168,67,.2)':'var(--card)'};border:2px solid ${S.user[step.field]===o.val?'var(--gold)':'rgba(255,255,255,.06)'};cursor:pointer;transition:all .2s;">
            <span style="font-size:36px;">${o.icon}</span>
            <span style="font-size:14px;font-weight:700;color:var(--txt);">${(o.labelKey && window.T) ? window.T(o.labelKey, o.label) : o.label}</span>
          </button>`).join('')}
      </div>
    </div>`;
  } else if (step.type === 'done') {
    const name = S.user.name || 'بطل';
    html = `<div style="min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;padding:40px 24px;text-align:center;">
      <div style="font-size:80px;margin-bottom:20px;">🎉</div>
      <div style="font-size:26px;font-weight:900;color:var(--gold);margin-bottom:10px;">أهلاً ${name}!</div>
      <div style="font-size:13px;color:var(--dim);line-height:1.8;margin-bottom:20px;">برنامجك جاهز — المدرب الذكي في انتظارك!</div>
      <div style="background:var(--card);border-radius:18px;padding:16px;width:100%;max-width:300px;margin-bottom:20px;">
        ${[
          ['الاسم', S.user.name||'—'],
          ['الوزن', (S.user.weight||'—')+' كغ'],
          ['الطول', (S.user.height||'—')+' سم'],
          ['مدة البرنامج', (S.user.programDays||30)+' يوم'],
        ].map(([l,v])=>`<div style="display:flex;justify-content:space-between;padding:7px 0;border-bottom:1px solid var(--border);"><span style="color:var(--dim);font-size:12px;">${l}</span><span style="font-weight:700;font-size:12px;">${v}</span></div>`).join('')}
      </div>
      <div style="background:linear-gradient(135deg,rgba(66,133,244,.12),rgba(66,133,244,.06));border:1.5px solid rgba(66,133,244,.35);border-radius:16px;padding:18px;width:100%;max-width:300px;margin-bottom:16px;">
        <div style="font-size:14px;font-weight:900;color:#4285f4;margin-bottom:6px;">☁️ احفظ تقدمك على السحاب</div>
        <div style="font-size:12px;color:var(--dim);margin-bottom:14px;line-height:1.6;">سجّل بـ Google لمزامنة بياناتك على جميع أجهزتك</div>
        <button onclick="if(window.firebaseSignIn){window.firebaseSignIn().then(()=>{S.onboardingDone=true;saveState();document.getElementById('onboarding').style.display='none';render();setTimeout(startTutorial,800);}).catch(()=>{});}" style="width:100%;padding:12px;border-radius:12px;background:rgba(66,133,244,.2);border:1.5px solid rgba(66,133,244,.5);color:#4285f4;font-family:'Cairo',sans-serif;font-size:14px;font-weight:700;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;">
          <svg width="18" height="18" viewBox="0 0 48 48"><path fill="#4285F4" d="M44.5 20H24v8.5h11.8C34.7 33.9 30.1 37 24 37c-7.2 0-13-5.8-13-13s5.8-13 13-13c3.1 0 5.9 1.1 8.1 2.9l6.4-6.4C34.6 4.1 29.6 2 24 2 11.8 2 2 11.8 2 24s9.8 22 22 22c11 0 21-8 21-22 0-1.3-.2-2.7-.5-4z"/></svg>
          تسجيل الدخول بـ Google
        </button>
      </div>
      <button onclick="S.onboardingDone=true;saveState();document.getElementById('onboarding').style.display='none';render();setTimeout(startTutorial,800);" style="font-size:12px;color:var(--dim);background:none;border:none;cursor:pointer;padding:8px 16px;text-decoration:underline;">تخطي لاحقاً — ابدأ بدون تسجيل</button>
    </div>`;
    document.getElementById('ob-footer').style.display = 'none';
    document.getElementById('ob-steps').innerHTML = html;
    return;
  }
  document.getElementById('ob-steps').innerHTML = html;
  if (step.type === 'input') setTimeout(()=>document.getElementById('ob-inp')?.focus(), 100);
}

