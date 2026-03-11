/* ══════ AUDIO (Web Audio API) ══════ */
let audioCtx = null;
function getCtx() {
  if (!audioCtx) audioCtx = new (window.AudioContext||window.webkitAudioContext)();
  return audioCtx;
}
function playTone(freq, dur, type='sine', vol=0.3) {
  if (!S.soundOn) return;
  try {
    const ctx = getCtx();
    const o = ctx.createOscillator();
    const g = ctx.createGain();
    o.connect(g); g.connect(ctx.destination);
    o.type = type; o.frequency.value = freq;
    g.gain.setValueAtTime((S.volume/100)*vol, ctx.currentTime);
    g.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime+dur);
    o.start(ctx.currentTime); o.stop(ctx.currentTime+dur);
  } catch(e){}
}
function playBeep() { playTone(880, 0.2, 'sine', 0.4); }
function playTick() { if(S.tickOn) playTone(440, 0.05, 'sine', 0.15); }
function playStart() { playTone(660, 0.1); setTimeout(()=>playTone(880,0.15),120); }
function playRest()  { playTone(440, 0.15); setTimeout(()=>playTone(330,0.2),150); }
function playFanfare() {
  if (!S.soundOn) return;
  const notes = [523,659,784,1047];
  notes.forEach((f,i) => setTimeout(()=>playTone(f,0.3,'triangle',0.5),i*150));
}

/* ══════════════════════════════════════════
   MOTIVATIONAL ARABIC TTS
   يستخدم Web Speech API المدمج في المتصفح
══════════════════════════════════════════ */
const MOTIVATIONAL_SPEECHES = [
  (name, day, cal) => `أحسنت يا ${name}! أنهيت اليوم ${day} من البرنامج. أنت أقوى مما تتخيل. استمر!`,
  (name, day, cal) => `رائع يا ${name}! ${cal} سعرة حرقتها اليوم. جسمك يشكرك. غداً أفضل!`,
  (name, day, cal) => `يا ${name}، كل يوم تتدرب فيه هو استثمار في نفسك. يوم ${day} مكتمل. فخور بك!`,
  (name, day, cal) => `أنت بطل يا ${name}! اليوم ${day} انتهى. الانضباط يصنع الفرق. لا تتوقف!`,
  (name, day, cal) => `ممتاز يا ${name}! أنهيت تمريناً آخر. ${cal} سعرة أقل. هدفك أقرب!`,
  (name, day, cal) => `يا ${name}، الجهد لا يكذب. اليوم ${day} مكتمل. استرح جيداً واستعد للغد!`,
  (name, day, cal) => `عظيم يا ${name}! كل تمرين يبني نسخة أفضل منك. يوم ${day} منجز!`,
];

function speakMotivation(day, cal) {
  // Check if TTS is enabled and supported
  if (!S.ttsOn) return;
  if (!window.speechSynthesis) return;
  const name = S.user?.name || 'بطل';
  const idx = (day - 1) % MOTIVATIONAL_SPEECHES.length;
  const text = MOTIVATIONAL_SPEECHES[idx](name, day, cal);
  // Cancel any ongoing speech
  window.speechSynthesis.cancel();
  const utt = new SpeechSynthesisUtterance(text);
  utt.lang = 'ar-SA';
  utt.rate = 0.9;
  utt.pitch = 1.0;
  utt.volume = (S.volume || 80) / 100;
  // Try to find an Arabic voice
  const voices = window.speechSynthesis.getVoices();
  const arVoice = voices.find(v => v.lang.startsWith('ar')) || null;
  if (arVoice) utt.voice = arVoice;
  // Small delay to let fanfare finish
  setTimeout(() => window.speechSynthesis.speak(utt), 900);
}

function testMotivationSpeech() {
  const orig = S.ttsOn;
  S.ttsOn = true;
  speakMotivation(S.currentDay || 1, 250);
  S.ttsOn = orig;
}

/* ══════ SOUND SHEET ══════ */
function openSoundSheet()  { document.getElementById('sound-sheet').classList.add('open'); }
function closeSoundSheet() { document.getElementById('sound-sheet').classList.remove('open'); }

function openSettingsSheet() {
  const sheet = document.getElementById('settings-sheet');
  sheet.style.display = 'flex';
  // Populate fields
  document.getElementById('settings-apikey').value = S.apiKey || '';
  document.getElementById('set-name').value = S.user?.name || '';
  document.getElementById('set-weight').value = S.user?.weight || '';
  document.getElementById('set-height').value = S.user?.height || '';
  document.getElementById('set-traintime').value = S.user?.trainTime || '';
  const sdEl = document.getElementById('set-startdate');
  if (sdEl) sdEl.value = S.user?.startDate || new Date().toISOString().split('T')[0];
  // Update TTS toggle state
  const ttsBtn = document.getElementById('tts-toggle');
  if (ttsBtn) ttsBtn.classList.toggle('on', S.ttsOn !== false);
  // Update program switcher highlight
  const prog = S.user?.program || 'standard';
  const btnStd = document.getElementById('pgm-btn-standard');
  const btnBeg = document.getElementById('pgm-btn-beginner');
  if (btnStd) { btnStd.style.background = prog==='standard'?'rgba(212,168,67,.2)':'var(--bg)'; btnStd.style.border = '2px solid '+(prog==='standard'?'var(--gold)':'var(--border)'); }
  if (btnBeg) { btnBeg.style.background = prog==='beginner'?'rgba(34,197,94,.2)':'var(--bg)'; btnBeg.style.border = '2px solid '+(prog==='beginner'?'#22c55e':'var(--border)'); }
  // Update install section state
  updateInstallSection();
}
function closeSettingsSheet() { document.getElementById('settings-sheet').style.display = 'none'; }
function saveApiKey() {
  const key = document.getElementById('settings-apikey').value.trim();
  S.apiKey = key;
  saveState();
  showMiniToast(key ? '✅ تم حفظ المفتاح' : '🗑️ تم حذف المفتاح');
}
function saveProfile() {
  if (!S.user) S.user = {};
  S.user.name   = document.getElementById('set-name').value.trim();
  S.user.weight = parseFloat(document.getElementById('set-weight').value) || 0;
  S.user.height = parseFloat(document.getElementById('set-height').value) || 0;
  S.user.trainTime = document.getElementById('set-traintime').value;
  const sd = document.getElementById('set-startdate')?.value;
  if (sd) S.user.startDate = sd;
  saveState();
  showMiniToast('✅ تم حفظ الملف الشخصي');
  closeSettingsSheet();
}
function switchProgram(prog) {
  // FIX#6: double confirmation + backup
  const progName = prog === 'beginner' ? 'المبتدئين (21 يوم)' : 'المتقدم (30 يوم)';
  if (!confirm('هل تريد التبديل لبرنامج ' + progName + '?\nسيعود اليوم الحالي إلى 1.')) return;
  const hasProgress = (S.completedDays||[]).length > 3;
  if (hasProgress && !confirm('\u26a0\ufe0f تنبيه: لديك ' + S.completedDays.length + ' يوماً مكتملاً.\nهل أنت متأكد؟ لا يمكن التراجع.')) return;
  // Backup current progress
  S._programBackup = {
    program: S.user?.program,
    programDays: S.user?.programDays,
    currentDay: S.currentDay,
    completedDays: [...(S.completedDays||[])],
    calories: S.calories,
    streak: S.streak,
    backedUpAt: Date.now()
  };
  if (!S.user) S.user = {};
  S.user.program = prog;
  S.user.programDays = prog === 'beginner' ? 21 : 30;
  S.currentDay = 1;
  S.completedDays = [];
  S.completedExercises = {};
  S.calories = 0;
  S.streak = 0;
  saveState();
  const d = document.getElementById('pgm-desc');
  if (d) d.textContent = prog==='beginner' ? '🌱 برنامج مريح للبدء: 3 أيام تدريب × 4 أسابيع، تمارين أخف وراحة كافية.' : '🔥 برنامج متكامل: HIIT + قوة + كور. مناسب لمن لديهم لياقة أساسية.';
  closeSettingsSheet();
  render();
  showMiniToast(prog === 'beginner' ? '🌱 برنامج المبتدئين 21 يوم — يلا نبدأ!' : '🔥 البرنامج المتقدم — جاهز!');
}
function toggleTTS() {
  S.ttsOn = !S.ttsOn; saveState();
  document.getElementById('tts-toggle').classList.toggle('on', S.ttsOn);
  showMiniToast(S.ttsOn ? '🎙️ الصوت التحفيزي مفعّل' : '🔇 الصوت التحفيزي متوقف');
}
function toggleSound() {
  S.soundOn = !S.soundOn; saveState();
  document.getElementById('snd-toggle').classList.toggle('on', S.soundOn);
}
function toggleTick() {
  S.tickOn = !S.tickOn; saveState();
  document.getElementById('tick-toggle').classList.toggle('on', S.tickOn);
}
function setVolume(v) { S.volume = parseInt(v); saveState(); }

