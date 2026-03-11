/* ══════ RENDER ══════ */
function switchTab(tab, btn) {
  // Redirect tips to coach (tab-tips has no content)
  if (tab === 'tips') { switchTab('coach', document.querySelector(".tab-btn[onclick*='coach']") || btn); return; }
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  const tabEl = document.getElementById('tab-' + tab);
  if (tabEl) tabEl.classList.add('active');
  if (btn) btn.classList.add('active');
  if (tab === 'progress') renderProgress();
  if (tab === 'coach') renderCoach();
}

function render() {
  renderDayStrip();
  renderWorkoutTab();
  renderHeaderStats();
  if (document.body.classList.contains('desktop-mode')) dtRender();
  if (document.body.classList.contains('tv-mode')) tvRender();
}

function renderHeaderStats() {
  document.getElementById('hdr-streak').textContent = S.streak;
  document.getElementById('hdr-cal').textContent = S.calories > 999 ? (S.calories/1000).toFixed(1)+'k' : S.calories;
}

function renderDayStrip() {
  const strip = document.getElementById('day-strip');
  const progDays = S.user?.programDays || 30;
  strip.innerHTML = '';
  for (let d = 1; d <= progDays; d++) {
    const sched = getDaySchedule(d);
    const pill = document.createElement('div');
    pill.className = 'day-pill';
    if (d === S.currentDay) pill.classList.add('active');
    if (S.completedDays.includes(d)) pill.classList.add('completed');
    if (sched.type === 'rest') pill.classList.add('rest-day');
    if (d > S.currentDay) pill.classList.add('future');
    pill.innerHTML = `<span class="d-num">${d}</span><span class="d-type">${getScheduleLabel(sched)}</span>`;
    pill.onclick = () => selectDay(d);
    // Long press to edit (future feature placeholder)
    let pressTimer;
    pill.addEventListener('mousedown', () => { pressTimer = setTimeout(() => showMiniToast('محرر الجدول قريباً ✨'), 600); });
    pill.addEventListener('mouseup', () => clearTimeout(pressTimer));
    pill.addEventListener('touchstart', e => { pressTimer = setTimeout(() => showMiniToast('محرر الجدول قريباً ✨'), 600); }, {passive:true});
    pill.addEventListener('touchend', () => clearTimeout(pressTimer));
    strip.appendChild(pill);
  }
  // Scroll to current day
  setTimeout(() => {
    const active = strip.querySelector('.active');
    if (active) active.scrollIntoView({behavior:'smooth',block:'nearest',inline:'center'});
  }, 100);
  // Enable drag-to-scroll on mouse
  if (!strip._dragInited) {
    strip._dragInited = true;
    let isDown = false, startX, scrollLeft;
    strip.addEventListener('mousedown', e => { isDown=true; startX=e.pageX-strip.offsetLeft; scrollLeft=strip.scrollLeft; strip.style.cursor='grabbing'; });
    strip.addEventListener('mouseleave', () => { isDown=false; strip.style.cursor='grab'; });
    strip.addEventListener('mouseup', () => { isDown=false; strip.style.cursor='grab'; });
    strip.addEventListener('mousemove', e => { if(!isDown) return; e.preventDefault(); const x=e.pageX-strip.offsetLeft; strip.scrollLeft=scrollLeft-(x-startX); });
    // Mouse wheel scrolls horizontally when hovering over the strip
    strip.addEventListener('wheel', e => {
      if (Math.abs(e.deltaX) > Math.abs(e.deltaY)) return; // let natural horizontal scroll work
      e.preventDefault();
      strip.scrollLeft += e.deltaY * 2.5;
    }, { passive: false });
  }
}

function selectDay(d) {
  S.currentDay = d;
  saveState();
  render();
}

function renderWorkoutTab() {
  const tab = document.getElementById('tab-workout');
  const sched = getEffectiveSchedule(S.currentDay);
  const phase = getPhase(S.currentDay);
  const phaseProgress = Math.min(100, ((S.currentDay - phase.start) / (phase.end - phase.start + 1)) * 100);
  const totalProgress = Math.min(100, (S.completedDays.length / (S.user?.programDays || 30)) * 100);
  const isDone = S.completedDays.includes(S.currentDay);
  const _quotes = getQuotes(); const quote = _quotes[(S.currentDay - 1) % _quotes.length];

  if (sched.type === 'rest') {
    tab.innerHTML = `
      <div class="phase-card">
        <div class="phase-icon">${phase.icon}</div>
        <div class="phase-info"><div class="phase-name">${getPhaseName(phase)}</div><div class="phase-days">${window.T('days2')} ${phase.days}</div></div>
      </div>
      <div class="rest-card" style="margin:16px;border-radius:20px;background:var(--card);padding:28px;text-align:center;">
        <div style="font-size:56px;margin-bottom:12px;">😴</div>
        <div style="font-size:20px;font-weight:900;color:var(--txt);margin-bottom:8px;">${window.T('restDayTitle')}</div>
        <div style="font-size:13px;color:var(--dim);margin-bottom:20px;line-height:1.7;">${window.T('restDaySub').replace('\n','<br>')}</div>
        <button onclick="openActiveRest()" style="width:100%;padding:14px;border-radius:14px;background:linear-gradient(135deg,#6366f1,#8b5cf6);border:none;color:#fff;font-family:'Cairo',sans-serif;font-size:15px;font-weight:900;cursor:pointer;margin-bottom:10px;">${window.T('activeRest')}</button>
        <div style="font-size:12px;color:var(--dim);">${STRETCH_EXERCISES.length} ${window.T('stretchExLabel','تمارين تمدد')} • ${STRETCH_EXERCISES.reduce((s,e)=>s+e.dur,0)} ${window.T('lblSecs','ثانية')}</div>
      </div>
      <div style="padding:0 16px;">
        <button onclick="toggleDayDone()" style="width:100%;padding:13px;border-radius:14px;background:${isDone?'rgba(74,222,128,.15)':'rgba(212,168,67,.1)'};border:1.5px solid ${isDone?'#4ade80':'rgba(212,168,67,.3)'};color:${isDone?'#4ade80':'var(--gold)'};font-family:'Cairo',sans-serif;font-size:14px;font-weight:800;cursor:pointer;">
          ${isDone ? window.T('recordedRest') : window.T('recordRest')}
        </button>
      </div>`;
    return;
  }

  const exHTML = sched.exercises.map((ex, i) => {
    const doneKey = S.currentDay + '_' + ex.id;
    const isDoneEx = (S.completedExercises[doneKey] === true);
    const customImg = S.customImages?.[ex.id];
    const gifSrc = getExGif(ex.id);
    const imgHTML = customImg
      ? `<img class="ex-gif" src="${customImg}" alt="${ex.nameAr}" onerror="this.remove()">`
      : gifSrc
        ? `<img class="ex-gif" src="${gifSrc}" alt="${ex.nameAr}" onerror="this.remove()">`
        : `<div class="ex-placeholder" style="font-size:32px;">${ex.icon}</div>`;
    return `<div class="ex-card ${isDoneEx?'done-ex':''}" id="ex-card-${ex.id}">
      <div class="ex-main">
        <div class="ex-img-wrap" data-gif-id="${ex.id}" onclick="triggerImgUpload('${ex.id}')" style="background:var(--card);">
          ${imgHTML}
          <div class="ex-upload-overlay">📷</div>
          <input type="file" accept="image/*" style="display:none" id="img-inp-${ex.id}" onchange="handleImgUpload('${ex.id}',this)">
        </div>
        <div class="ex-info">
          <div class="ex-name">${currentLang==="ar"?ex.nameAr:(ex.nameEn||ex.nameAr)}</div>
          <div class="ex-name-en">${ex.nameEn}</div>
          <div class="chips">
            <span class="chip sets">${ex.sets} × ${ex.reps} ${getRepsLabel(ex)}</span>
            <span class="chip muscles">${ex.muscles}</span>
            <span class="chip" style="background:rgba(249,115,22,.12);color:#f97316;border-color:rgba(249,115,22,.25);">🔥 ~${calcExCal(ex)} ${window.T('exCalChip','كالوري')}</span>
          </div>
        </div>
      </div>
      <div class="ex-bottom">
        <div class="ex-check ${isDoneEx?'checked':''}" onclick="toggleExDone('${ex.id}')">
          ${isDoneEx?'✓':''}
        </div>
        <button class="ex-expand" onclick="toggleSteps('${ex.id}')">${window.T('steps')} ▼</button>
        <button class="ex-edit-btn" onclick="openExEditor('${ex.id}',${S.currentDay})">✏️</button>
      </div>
      <div class="ex-steps" id="steps-${ex.id}">
        ${ex.steps.map((s,i)=>`<div class="ex-step"><div class="step-num">${i+1}</div><div>${s}</div></div>`).join('')}
      </div>
    </div>`;
  }).join('');

  tab.innerHTML = `
    <div class="phase-card">
      <div class="phase-icon">${phase.icon}</div>
      <div class="phase-info">
        <div class="phase-name">${getPhaseName(phase)} — ${window.T('dayLabel','يوم')} ${S.currentDay}</div>
        <div class="phase-days">${getScheduleLabel(sched)} • ${getExercisesLabel(sched.exercises.length)}</div>
      </div>
      <svg width="44" height="44" viewBox="0 0 44 44">
        <circle cx="22" cy="22" r="18" fill="none" stroke="var(--bg-card)" stroke-width="5"/>
        <circle cx="22" cy="22" r="18" fill="none" stroke="var(--primary)" stroke-width="5" stroke-linecap="round"
          stroke-dasharray="${2*Math.PI*18}" stroke-dashoffset="${2*Math.PI*18*(1-phaseProgress/100)}"
          transform="rotate(-90 22 22)" style="filter:drop-shadow(0 0 4px var(--glow))"/>
        <text x="22" y="26" text-anchor="middle" class="prog-ring-val" font-size="10" font-weight="900" fill="var(--primary)">${Math.round(phaseProgress)}%</text>
      </svg>
    </div>
    <div class="progress-bar-wrap"><div class="progress-bar-fill" style="width:${totalProgress}%"></div></div>
    ${(()=>{
      const _standalone = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone;
      if (_standalone) return '';
      const _ios = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
      if (deferredPrompt) {
        return `<div id="home-install-banner" style="margin:0 0 2px;padding:10px 14px;background:linear-gradient(135deg,rgba(var(--gl-rgb,14,165,233),.12),rgba(var(--gd-rgb,2,132,199),.08));border-radius:14px;border:1.5px solid var(--border-strong);display:flex;align-items:center;gap:10px;">
          <span style="font-size:22px;flex-shrink:0;">📲</span>
          <div style="flex:1;min-width:0;">
            <div style="font-size:12px;font-weight:800;color:var(--txt);">${window.T('installAppTitle','ثبّت التطبيق على جهازك')}</div>
            <div style="font-size:10px;color:var(--dim);">${window.T('installAppSub','وصول فوري — يعمل بدون نت — تذكيرات')}</div>
          </div>
          <button onclick="homeInstallNow(this)" style="padding:8px 14px;border-radius:10px;background:linear-gradient(135deg,var(--gl),var(--gd));color:var(--night);border:none;font-family:'Cairo',sans-serif;font-size:12px;font-weight:900;cursor:pointer;white-space:nowrap;flex-shrink:0;">${window.T('installNowBtn','تثبيت الآن')}</button>
        </div>`;
      }
      if (_ios) {
        return `<div id="home-install-banner" style="margin:0 0 2px;padding:10px 14px;background:linear-gradient(135deg,rgba(14,165,233,.1),rgba(2,132,199,.06));border-radius:14px;border:1.5px solid var(--border-strong);display:flex;align-items:center;gap:10px;">
          <span style="font-size:22px;flex-shrink:0;">📲</span>
          <div style="flex:1;min-width:0;">
            <div style="font-size:12px;font-weight:800;color:var(--txt);">${window.T('installAppTitle','ثبّت التطبيق')}</div>
            <div style="font-size:10px;color:var(--dim);">اضغط <strong>⬆️ مشاركة</strong> ثم "إضافة إلى الشاشة الرئيسية"</div>
          </div>
          <button onclick="this.closest('#home-install-banner').remove()" style="background:none;border:none;color:var(--dim);font-size:18px;cursor:pointer;flex-shrink:0;">✕</button>
        </div>`;
      }
      return '';
    })()}
    <button class="start-btn" onclick="startGuidedSession()">${window.T('startGuided','▶ ابدأ الجلسة الموجهة')}</button>
    <button class="done-btn ${isDone?'done':''}" onclick="toggleDayDone()">
      ${isDone ? "✅ "+window.T("dayDone") : "☑️ "+window.T("markDone")}
    </button>
    <div style="display:flex;justify-content:center;gap:16px;margin:8px 16px;padding:10px 16px;background:rgba(249,115,22,.08);border-radius:14px;border:1px solid rgba(249,115,22,.2);">
      <span style="font-size:12px;color:var(--dim);">${window.T('estimatedCalToday','🔥 السعرات المقدرة لليوم:')}</span>
      <span style="font-size:13px;font-weight:900;color:#f97316;">~${calcScheduleCal(sched, parseFloat(S.user?.weight)||70)} ${window.T('calLabel','كالوري')}</span>
      <span style="font-size:11px;color:var(--dim);">( ${window.T('weightLbl','وزن')} ${parseFloat(S.user?.weight)||70} ${window.T('kgLbl','كغ')} )</span>
    </div>
    <div class="daily-quote"><div class="quote-text">"${quote.text}"</div><div class="quote-attr">— ${quote.attr}</div></div>
    <div class="section-title" style="display:flex;align-items:center;justify-content:space-between;">
      <span>${window.T('exercisesTodayLbl','تمارين اليوم')}</span>
      <button onclick="openDayEditor(${S.currentDay})" style="font-size:11px;padding:4px 10px;border-radius:8px;background:rgba(212,168,67,.15);border:1px solid rgba(212,168,67,.3);color:var(--gold);cursor:pointer;font-weight:700;">${window.T('editExercises','✏️ تعديل التمارين')}</button>
    </div>
    ${exHTML}`;
}

function toggleSteps(id) {
  const el = document.getElementById('steps-' + id);
  if (el) el.classList.toggle('show');
}
function toggleExDone(id) {
  const key = S.currentDay + '_' + id;
  S.completedExercises[key] = !S.completedExercises[key];
  saveState();
  const allExArr = [...EXERCISES, ...(S.customExercises||[])];
  const ex = allExArr.find(e => e.id === id);
  const cal = calcExCal(ex);
  if (S.completedExercises[key]) { S.calories += cal; } else { S.calories = Math.max(0, S.calories - cal); }
  saveState();
  renderWorkoutTab();
  renderHeaderStats();
  checkBadges();
}
function toggleDayDone() {
  const maxDay = S.user?.programDays || 30;
  if (S.completedDays.includes(S.currentDay)) {
    S.completedDays = S.completedDays.filter(d => d !== S.currentDay);
    // Delete using both possible key formats for backward compat
    delete (S.trainingLog||{})[S.currentDay];
    delete (S.trainingLog||{})['day_'+S.currentDay];
  } else {
    S.completedDays.push(S.currentDay);
    const _sched = getDaySchedule(S.currentDay);
    const _kg = parseFloat(S.user?.weight) || 70;
    const _dayCal = calcScheduleCal(_sched, _kg) || 150;
    S.calories += _dayCal;
    // Track morning workouts for 'early_bird' badge (before 10am)
    const hour = new Date().getHours();
    if (hour < 10) S.morningWorkouts = (S.morningWorkouts || 0) + 1;
    // Log training details
    if (!S.trainingLog) S.trainingLog = {};
    const sched = _sched;
    const logKey = 'day_' + S.currentDay;
    S.trainingLog[logKey] = {
      date: new Date().toLocaleDateString('ar-SA'),
      day: S.currentDay,
      type: sched.label || sched.type,
      exercises: sched.exercises.map(e => e.nameAr),
      exCount: sched.exercises.length,
      calories: _dayCal,
      duration: Math.max(20, sched.exercises.length * 5),
      ts: Date.now()
    };
    updateStreak();
    checkBadges();
    speakMotivation(S.currentDay, _dayCal);
    if (S.currentDay < maxDay) S.currentDay++;
  }
  saveState();
  render();
}
function updateStreak() {
  if (!S.completedDays.length) { S.streak = 0; return; }
  // Count from the highest completed day downward
  const maxDone = Math.max(...S.completedDays);
  let streak = 0;
  let d = maxDone;
  while (S.completedDays.includes(d)) { streak++; d--; }
  S.streak = streak;
}
function triggerImgUpload(id) { document.getElementById('img-inp-' + id)?.click(); }
function handleImgUpload(id, inp) {
  const file = inp.files[0];
  if (!file) return;
  // حد أقصى 500KB للصور لحماية localStorage
  if (file.size > 500 * 1024) {
    showMiniToast('⚠️ الصورة كبيرة جداً (الحد 500KB)');
    inp.value = '';
    return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    if (!S.customImages) S.customImages = {};
    S.customImages[id] = e.target.result;
    saveState();
    renderWorkoutTab();
    showMiniToast('✅ تم حفظ الصورة');
  };
  reader.readAsDataURL(file);
}

/* ══════ PROGRESS TAB ══════ */
function renderProgress() {
  document.getElementById('p-done').textContent = S.completedDays.length;
  document.getElementById('p-streak').textContent = S.streak;
  // Calories: sum from training log for accuracy
  const logCal = Object.values(S.trainingLog||{}).reduce((s,e)=>s+(e.calories||0),0);
  const displayCal = logCal || S.calories || 0;
  document.getElementById('p-cal').textContent = displayCal;
  const progDays = S.user?.programDays || 30;
  document.getElementById('p-remain').textContent = progDays - S.completedDays.length;
  // Show start date info below stats if available
  const sdInfo = document.getElementById('start-date-info');
  if (sdInfo && S.user?.startDate) {
    const sd = new Date(S.user.startDate);
    const endD = new Date(sd);
    endD.setDate(endD.getDate() + progDays);
    sdInfo.innerHTML = `<span>📅 البدء: ${sd.toLocaleDateString('ar-SA')}</span><span style="margin:0 8px;color:var(--border);">|</span><span>🏁 الانتهاء: ${endD.toLocaleDateString('ar-SA')}</span>`;
    sdInfo.style.display = 'flex';
  }
  renderPhases();
  renderWeeklyBars();
  renderCalendar();
  renderBadges();
  renderSmartStats();
  renderTrainingLogSection();
  renderNutrition();
}

/* ══════════════════════════════════════════
   NUTRITION PLAN
══════════════════════════════════════════ */
function renderNutrition() {
  const el = document.getElementById('nutrition-plan');
  if (!el) return;
  const kg = parseFloat(S.user?.weight) || 70;
  const goal = S.user?.goal || 'burn';
  const goalAr = {burn:'حرق الدهون', muscle:'بناء العضلات', fitness:'تحسين اللياقة', health:'الصحة العامة'}[goal] || 'عام';

  // Calorie targets per goal
  const bmr = Math.round(kg * 22 * 1.4); // rough TDEE
  const calorieTarget = goal === 'burn' ? Math.round(bmr * 0.8) : goal === 'muscle' ? Math.round(bmr * 1.1) : bmr;

  // Macros
  const protein = Math.round(kg * (goal === 'muscle' ? 2.0 : 1.6));
  const fat     = Math.round(kg * 0.8);
  const carb    = Math.round((calorieTarget - protein * 4 - fat * 9) / 4);
  const carbSafe = Math.max(50, carb);

  // Progress: today's burned vs target
  const todaySched = (() => { try { return getDaySchedule(S.currentDay); } catch(e){ return null; } })();
  const todayBurn  = todaySched ? calcScheduleCal(todaySched, kg) : 0;
  const netCal     = calorieTarget + todayBurn; // total you can eat on training days

  // Macro bars width %
  const maxMacro = Math.max(protein, carbSafe, fat);
  const pBar = (v) => Math.round(v / maxMacro * 100);

  const meals = goal === 'burn' ? [
    {time:'☀️ الفطور', foods:['بيضتان مسلوقتان','كوب شوفان بالماء','تفاحة صغيرة']},
    {time:'🥗 الغداء',  foods:['صدر دجاج 150غ مشوي','أرز بني ½ كوب','سلطة خضراء بزيت زيتون']},
    {time:'🌙 العشاء',  foods:['تونا علبة','خبز أسمر قطعتان','خيار وطماطم']},
    {time:'💪 بعد التمرين', foods:['موزة','كوب حليب أو لبن خالي دسم']},
  ] : goal === 'muscle' ? [
    {time:'☀️ الفطور', foods:['3 بيضات + بياض 2','شوفان كوب كامل','موزة']},
    {time:'🥗 الغداء',  foods:['صدر دجاج 200غ','أرز كوب كامل','خضار مشوية']},
    {time:'🌙 العشاء',  foods:['لحم أحمر أو سمك 200غ','بطاطا حلوة','سلطة']},
    {time:'💪 بعد التمرين', foods:['شيك بروتين أو لبن يوناني','موزة + ملعقة عسل']},
  ] : [
    {time:'☀️ الفطور', foods:['بيضتان','خبز أسمر','فاكهة']},
    {time:'🥗 الغداء',  foods:['بروتين 150غ','نشويات متوسطة','خضار']},
    {time:'🌙 العشاء',  foods:['بروتين خفيف','سلطة','شوربة']},
    {time:'💪 وجبة صغيرة', foods:['مكسرات حفنة','تمر 3 حبات']},
  ];

  el.innerHTML = `
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:6px;">
      <div style="font-size:12px;color:var(--dim);">الهدف: <strong style="color:var(--green);">${goalAr}</strong> · ${kg} كغ</div>
      <div style="font-size:13px;font-weight:900;color:var(--green);">${calorieTarget} كالوري/يوم</div>
    </div>

    <!-- Macro bars -->
    <div style="display:flex;flex-direction:column;gap:10px;margin-bottom:16px;background:var(--card2);border-radius:14px;padding:14px;">
      ${[
        {label:'🥩 بروتين', val:protein, unit:'غ', cal:protein*4, color:'#f97316'},
        {label:'🍚 كارب',   val:carbSafe,unit:'غ', cal:carbSafe*4, color:'#38bdf8'},
        {label:'🥑 دهون',   val:fat,     unit:'غ', cal:fat*9,     color:'#a78bfa'},
      ].map(m => `
        <div>
          <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
            <span style="font-size:12px;font-weight:700;color:var(--txt);">${m.label}</span>
            <span style="font-size:12px;color:var(--dim);">${m.val}${m.unit} · ${m.cal} كال</span>
          </div>
          <div style="background:rgba(255,255,255,.06);border-radius:6px;height:8px;overflow:hidden;">
            <div style="width:${pBar(m.val)}%;height:100%;background:${m.color};border-radius:6px;transition:width .6s;"></div>
          </div>
        </div>`).join('')}
    </div>

    ${todayBurn > 0 ? `<div style="background:rgba(249,115,22,.08);border:1px solid rgba(249,115,22,.2);border-radius:10px;padding:10px 14px;margin-bottom:14px;font-size:12px;color:var(--dim);">
      🔥 يوم تدريب: يمكنك أكل <strong style="color:#f97316;">${netCal} كالوري</strong> (+${todayBurn} تعويض التمرين)
    </div>` : ''}

    <!-- Meal plan -->
    <div style="font-size:12px;font-weight:800;color:var(--dim);margin-bottom:8px;text-transform:uppercase;letter-spacing:1px;">وجباتك اليوم</div>
    <div style="display:flex;flex-direction:column;gap:8px;">
      ${meals.map(meal => `
        <div style="background:var(--card);border-radius:12px;padding:12px 14px;border:1px solid var(--border);">
          <div style="font-size:12px;font-weight:800;color:var(--green);margin-bottom:6px;">${meal.time}</div>
          <div style="display:flex;flex-wrap:wrap;gap:6px;">
            ${meal.foods.map(f => `<span style="background:rgba(34,197,94,.08);border:1px solid rgba(34,197,94,.2);border-radius:20px;padding:3px 10px;font-size:11px;color:var(--txt);">${f}</span>`).join('')}
          </div>
        </div>`).join('')}
    </div>

    <div style="margin-top:12px;padding:10px 14px;background:rgba(56,189,248,.07);border-radius:10px;font-size:11px;color:var(--dim);line-height:1.8;">
      💧 اشرب ${(kg * 0.035).toFixed(1)} لتر ماء يومياً + 500 مل إضافية يوم التمرين
    </div>`;
}
function renderPhases() {
  const list = document.getElementById('phases-list');
  if (!list) return;
  list.innerHTML = PHASES.map(p => {
    const cur = S.currentDay;
    const done = S.completedDays.filter(d => d >= p.start && d <= p.end).length;
    const total = p.end - p.start + 1;
    const pct = Math.min(100, (done / total) * 100);
    const cls = cur >= p.start && cur <= p.end ? 'ph-active' : done >= total ? 'ph-done' : '';
    return `<div class="phase-item ${cls}">
      <div class="ph-icon">${p.icon}</div>
      <div class="ph-info">
        <div class="ph-name">${p.name} (${p.days})</div>
        <div class="ph-days">${done}/${total} أيام</div>
        <div class="ph-bar"><div class="ph-fill" style="width:${pct}%"></div></div>
      </div>
    </div>`;
  }).join('');
}
function renderWeeklyBars() {
  const el = document.getElementById('weekly-bars');
  if (!el) return;
  const progDays = S.user?.programDays || 30;
  const numWeeks = Math.ceil(progDays / 7);

  // Build weekly stats: days completed + calories
  const weeks = Array.from({length: numWeeks}, (_, i) => {
    const w = i + 1;
    const start = (w-1)*7+1, end = Math.min(w*7, progDays);
    const done = S.completedDays.filter(d => d>=start && d<=end).length;
    const cal = Object.values(S.trainingLog||{})
      .filter(e => e.day >= start && e.day <= end)
      .reduce((s,e)=>s+(e.calories||0),0);
    return { w, start, end, done, max: end-start+1, cal };
  });

  const maxDone = Math.max(...weeks.map(w=>w.done), 1);
  const maxCal  = Math.max(...weeks.map(w=>w.cal), 1);

  // Toggle between "days" and "calories" views
  const view = el.dataset.view || 'days';
  const vals = view === 'cal' ? weeks.map(w=>w.cal) : weeks.map(w=>w.done);
  const maxVal = view === 'cal' ? maxCal : maxDone;

  el.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;">
      <div style="font-size:12px;color:var(--dim);">${view==='days'?'أيام التدريب الأسبوعية':'السعرات الأسبوعية'}</div>
      <div style="display:flex;gap:6px;">
        <button onclick="document.getElementById('weekly-bars').dataset.view='days';renderWeeklyBars();"
          style="padding:4px 10px;border-radius:8px;background:${view==='days'?'rgba(212,168,67,.2)':'var(--card)'};border:1px solid ${view==='days'?'var(--gold)':'var(--border)'};color:${view==='days'?'var(--gold)':'var(--dim)'};font-size:11px;cursor:pointer;font-family:'Cairo',sans-serif;">📅 أيام</button>
        <button onclick="document.getElementById('weekly-bars').dataset.view='cal';renderWeeklyBars();"
          style="padding:4px 10px;border-radius:8px;background:${view==='cal'?'rgba(249,115,22,.2)':'var(--card)'};border:1px solid ${view==='cal'?'#f97316':'var(--border)'};color:${view==='cal'?'#f97316':'var(--dim)'};font-size:11px;cursor:pointer;font-family:'Cairo',sans-serif;">🔥 كالوري</button>
      </div>
    </div>
    <div style="display:flex;align-items:flex-end;gap:6px;height:90px;padding:0 4px;">
      ${weeks.map((w, i) => {
        const val = vals[i];
        const h = maxVal > 0 ? Math.max(4, Math.round((val / maxVal) * 78)) : 4;
        const cur = S.currentDay >= w.start && S.currentDay <= w.end;
        const color = cur
          ? (view==='cal'?'#f97316':'var(--gold)')
          : val > 0
            ? (view==='cal'?'rgba(249,115,22,.5)':'rgba(212,168,67,.45)')
            : 'rgba(255,255,255,.06)';
        const lbl = view==='cal' ? (val>=1000?(val/1000).toFixed(1)+'k':val||'') : (val||'');
        return `<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;gap:4px;">
          <div style="font-size:9px;color:${cur?'var(--gold)':'var(--dim)'};font-weight:${cur?'900':'400'};">${lbl}</div>
          <div style="width:100%;border-radius:5px 5px 3px 3px;background:${color};height:${h}px;transition:height .5s ease;cursor:default;min-height:4px;"
            title="أسبوع ${w.w}: ${view==='cal'?val+' كالوري':val+'/'+w.max+' أيام'}"></div>
          <div style="font-size:10px;color:${cur?'var(--gold)':'var(--dim)'};font-weight:${cur?'800':'400'};">أ${w.w}</div>
        </div>`;
      }).join('')}
    </div>
    ${weeks.some(w=>w.done>0) ? `
    <div style="margin-top:12px;display:flex;gap:10px;flex-wrap:wrap;">
      <div style="background:var(--card);border-radius:10px;padding:8px 12px;flex:1;text-align:center;min-width:80px;">
        <div style="font-size:16px;font-weight:900;color:var(--gold);">${weeks.reduce((s,w)=>s+w.done,0)}</div>
        <div style="font-size:10px;color:var(--dim);">يوم مكتمل</div>
      </div>
      <div style="background:var(--card);border-radius:10px;padding:8px 12px;flex:1;text-align:center;min-width:80px;">
        <div style="font-size:16px;font-weight:900;color:#f97316;">${weeks.reduce((s,w)=>s+w.cal,0)}</div>
        <div style="font-size:10px;color:var(--dim);">إجمالي كالوري</div>
      </div>
      <div style="background:var(--card);border-radius:10px;padding:8px 12px;flex:1;text-align:center;min-width:80px;">
        <div style="font-size:16px;font-weight:900;color:#4ade80;">${Math.round(weeks.filter(w=>w.done>0).reduce((s,w)=>s+w.done,0)/Math.max(1,weeks.filter(w=>w.done>0).length))}</div>
        <div style="font-size:10px;color:var(--dim);">متوسط/أسبوع</div>
      </div>
    </div>` : ''}
  `;
}
function renderCalendar() {
  const el = document.getElementById('cal-grid');
  if (!el) return;
  const progDays = S.user?.programDays || 30;
  el.innerHTML = Array.from({length:progDays},(_,i)=>{
    const d = i+1;
    const sched = getDaySchedule(d);
    let cls = 'cal-cell';
    if (S.completedDays.includes(d)) cls += ' cc-done';
    else if (d === S.currentDay) cls += ' cc-today';
    else if (sched.type === 'rest') cls += ' cc-rest';
    else if (d > S.currentDay) cls += ' cc-future';
    return `<div class="${cls}" title="يوم ${d}">${d}</div>`;
  }).join('');
}
function renderBadges() {
  const el = document.getElementById('badges-grid');
  if (!el) return;
  el.innerHTML = BADGES.map(b => {
    const unlocked = S.unlockedBadges.includes(b.id);
    return `<div class="badge-item${unlocked?' unlocked':''}" title="${b.desc}">
      <div class="b-icon">${b.icon}</div>
      <div class="b-name">${b.name}</div>
    </div>`;
  }).join('');
}
function renderTips() {
  const el = document.getElementById('tips-list');
  if (!el) return;
  el.innerHTML = TIPS.map(t => `<div class="tip-card"><div class="tip-cat">${t.cat}</div><div class="tip-text">${t.text}</div></div>`).join('');
}
function saveMeasurements() {
  const date = new Date().toISOString().split('T')[0];
  S.bodyMeasurements[date] = {
    weight: document.getElementById('m-weight').value,
    waist: document.getElementById('m-waist').value,
    chest: document.getElementById('m-chest').value,
  };
  saveState();
  showMiniToast('✅ تم حفظ القياسات');
}

