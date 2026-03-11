/* ══════ TIMER ══════ */
let timerInterval = null, timerTotal = 60, timerRemain = 60, timerRunning = false;
function setTimer(secs) {
  timerReset();
  timerTotal = secs;
  timerRemain = secs;
  updateTimerUI();
}
function timerToggle() {
  if (timerTotal === 0) { setTimer(60); }
  if (timerRunning) {
    clearInterval(timerInterval);
    timerRunning = false;
    document.getElementById('t-play-btn').textContent = '▶ تشغيل';
    document.getElementById('t-lbl').textContent = 'موقف مؤقتاً';
  } else {
    if (timerRemain <= 0) { timerRemain = timerTotal; }
    timerRunning = true;
    document.getElementById('t-play-btn').textContent = '⏸ إيقاف';
    document.getElementById('t-lbl').textContent = 'يعمل';
    timerInterval = setInterval(() => {
      timerRemain--;
      if (S.tickOn) playTick();
      updateTimerUI();
      if (timerRemain <= 0) {
        clearInterval(timerInterval);
        timerRunning = false;
        document.getElementById('t-play-btn').textContent = '▶ تشغيل';
        document.getElementById('t-lbl').textContent = 'انتهى! ✅';
        playBeep();
      }
    }, 1000);
  }
}
function timerReset() {
  clearInterval(timerInterval);
  timerRunning = false;
  timerRemain = timerTotal;
  // Also stop tabata if running
  if (tabState.running) {
    tabState.running = false;
    tabState.curRound = 1;
    tabState.phase = 'work';
  }
  document.getElementById('t-play-btn').textContent = '▶ تشغيل';
  document.getElementById('t-lbl').textContent = 'مستعد';
  updateTimerUI();
}
function updateTimerUI() {
  const m = Math.floor(timerRemain/60), s = timerRemain%60;
  const timeStr = String(m).padStart(2,'0')+':'+String(s).padStart(2,'0');
  document.getElementById('t-display').textContent = timeStr;
  const pct = timerTotal > 0 ? timerRemain/timerTotal : 0;
  const dash = 2*Math.PI*88;
  document.getElementById('t-fill').style.strokeDashoffset = dash*(1-pct);
  // Sync desktop timer display
  const dtDisp = document.getElementById('dt-t-disp');
  if (dtDisp) dtDisp.textContent = timeStr;
}

/* Tabata */
const tabState = {work:20,rest:10,rounds:8,curRound:1,phase:'work',running:false};
function tabAdj(k,v) {
  S.tabata[k] = Math.max(k==='rounds'?1:5, (S.tabata[k]||tabState[k]) + v);
  saveState();
  document.getElementById('tab-'+k).textContent = S.tabata[k];
}
function startTabata() {
  tabState.work = S.tabata.work; tabState.rest = S.tabata.rest; tabState.rounds = S.tabata.rounds;
  tabState.curRound = 1; tabState.phase = 'work'; tabState.running = true;
  tabataNextPhase();
}

function tabataNextPhase() {
  if (!tabState.running) return;
  if (tabState.phase === 'work') {
    setTimer(tabState.work);
    document.getElementById('t-lbl').textContent = `🔥 جولة ${tabState.curRound}/${tabState.rounds} — عمل`;
    showMiniToast(`🔥 جولة ${tabState.curRound} — عمل ${tabState.work}ث`);
  } else {
    setTimer(tabState.rest);
    document.getElementById('t-lbl').textContent = `😤 راحة — جولة ${tabState.curRound}/${tabState.rounds}`;
    showMiniToast(`😤 راحة ${tabState.rest}ث`);
  }
  // Override the timer's onend to drive tabata progression
  clearInterval(timerInterval);
  timerRunning = true;
  document.getElementById('t-play-btn').textContent = '⏸ إيقاف';
  timerInterval = setInterval(() => {
    timerRemain--;
    if (S.tickOn) playTick();
    updateTimerUI();
    if (timerRemain <= 0) {
      clearInterval(timerInterval);
      timerRunning = false;
      playBeep();
      // Advance tabata state
      if (tabState.phase === 'work') {
        tabState.phase = 'rest';
        if (tabState.rest > 0) {
          tabataNextPhase();
        } else {
          // No rest — go directly to next round
          tabState.phase = 'work';
          tabState.curRound++;
          if (tabState.curRound > tabState.rounds) {
            tabataFinish();
          } else {
            tabataNextPhase();
          }
        }
      } else {
        tabState.phase = 'work';
        tabState.curRound++;
        if (tabState.curRound > tabState.rounds) {
          tabataFinish();
        } else {
          tabataNextPhase();
        }
      }
    }
  }, 1000);
}

function tabataFinish() {
  tabState.running = false;
  document.getElementById('t-lbl').textContent = '🏆 تاباتا منتهية!';
  document.getElementById('t-play-btn').textContent = '▶ تشغيل';
  // Calorie calculation: total work time in minutes × MET × kg / 60
  const tabWorkMins = (tabState.rounds * tabState.work) / 60;
  const tabKg = parseFloat(S.user?.weight) || 70;
  const tabCal = Math.max(20, Math.round(9.0 * tabKg * tabWorkMins / 60));
  S.calories += tabCal; saveState();
  showMiniToast(`🏆 أنهيت ${tabState.rounds} جولات تاباتا! 🔥 ${tabCal} سعرة`);
  playFanfare?.();
}

/* ══════ ROPE TRACKER ══════ */
let ropeInterval = null, ropeRunning = false, ropeStart = 0;
function ropeToggle() {
  if (ropeRunning) {
    clearInterval(ropeInterval);
    ropeRunning = false;
    // Calorie calculation for rope session: MET 10 × kg × elapsed minutes / 60
    const ropeElapsedMins = (Date.now() - ropeStart) / 60000;
    const ropeKg = parseFloat(S.user?.weight) || 70;
    const ropeCal = Math.max(5, Math.round(10.0 * ropeKg * ropeElapsedMins / 60));
    S.calories += ropeCal;
    S.ropeSessions++;
    // FIX-H: accumulate session jumps once on stop
    const sessionJumps = Math.floor(ropeElapsedMins * 100);
    S.ropeJumps = (S.ropeJumps || 0) + sessionJumps;
    saveState();
    document.getElementById('rope-btn').textContent = '▶ ابدأ';
    showMiniToast(`🪢 أحسنت! 🔥 ${ropeCal} سعرة`);
    checkBadges();
  } else {
    ropeRunning = true;
    ropeStart = Date.now();
    document.getElementById('rope-btn').textContent = '⏹ إيقاف';
    ropeInterval = setInterval(() => {
      const elapsed = (Date.now() - ropeStart) / 60000;
      const jumps = Math.floor(elapsed * 100);   // total jumps this session
      const meters = Math.floor(jumps * 1.8);
      const mins = Math.floor(elapsed);
      // Live calorie counter
      const liveKg = parseFloat(S.user?.weight) || 70;
      const liveCal = Math.round(10.0 * liveKg * elapsed / 60);
      // FIX-H: don't accumulate in S.ropeJumps during interval (causes exponential growth)
      // Just display the current session value live; S.ropeJumps updated once on stop
      document.getElementById('rope-jumps').textContent = jumps;
      document.getElementById('rope-meters').textContent = meters;
      document.getElementById('rope-mins').textContent = mins;
      document.getElementById('rope-sessions').textContent = S.ropeSessions;
      // Update live cal display if element exists
      const ropeCalEl = document.getElementById('rope-live-cal');
      if (ropeCalEl) ropeCalEl.textContent = liveCal;
    }, 1000);
  }
}
function ropeReset() {
  clearInterval(ropeInterval);
  ropeRunning = false;
  S.ropeJumps = 0; S.ropeSessions = 0;
  saveState();
  document.getElementById('rope-btn').textContent = '▶ ابدأ';
  document.getElementById('rope-jumps').textContent = '0';
  document.getElementById('rope-meters').textContent = '0';
  document.getElementById('rope-mins').textContent = '0';
  document.getElementById('rope-sessions').textContent = '0';
}

