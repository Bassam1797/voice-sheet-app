export function setupSpeech({ onNumber, onStatus, getSilenceMs, getButton }){
  const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
  const btn = getButton();
  let rec = null;
  let silenceTimer = null;

  function start(){
    if (!SpeechRec){ onStatus('Speech Recognition not supported in this browser.'); return; }
    if (rec){ stop(); return; }
    rec = new SpeechRec();
    rec.lang = 'en-US';
    rec.continuous = true;
    rec.interimResults = true;
    rec.onstart = ()=>{ btn.textContent='Mic ■'; onStatus('Listening…'); };
    rec.onend = ()=>{ btn.textContent='Mic ▶'; onStatus('Ready'); rec=null; };
    rec.onerror = (e)=>{ onStatus('Mic error: '+e.error); };
    rec.onresult = (e)=>{
      // find the last final result
      for (let i = e.resultIndex; i < e.results.length; i++){
        const res = e.results[i];
        const txt = res[0].transcript.trim();
        if (res.isFinal){
          // accept only numeric
          const cleaned = txt.replace(/[,\s]/g,''); // "1 000" -> "1000"
          if (/^[-+]?\d*(?:\.\d+)?$/.test(cleaned) && cleaned !== ''){
            onNumber(cleaned);
          }
          resetSilence();
        } else {
          resetSilence(); // keep timer alive while interim coming in
        }
      }
    };
    rec.start();
    resetSilence();
  }
  function resetSilence(){
    clearTimeout(silenceTimer);
    silenceTimer = setTimeout(()=>{
      // simulate auto-advance trigger by sending empty noop; grid handles advance
      onStatus('Auto-advance after silence.');
      // do nothing else — advance is done by grid after valid number already entered
    }, getSilenceMs());
  }
  function stop(){
    if (rec){ rec.stop(); rec=null; }
    btn.textContent='Mic ▶';
  }

  btn.addEventListener('click', ()=>{ if (rec) stop(); else start(); });
  return { start, stop };
}
