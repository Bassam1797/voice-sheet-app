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
      for (let i = e.resultIndex; i < e.results.length; i++){
        const res = e.results[i];
        const txt = res[0].transcript.trim();
        if (res.isFinal){
          const cleaned = txt.replace(/[,\s]/g,''); 
          if (/^[-+]?\d*(?:\.\d+)?$/.test(cleaned) && cleaned !== ''){
            onNumber(cleaned);
          }
          resetSilence();
        } else {
          resetSilence();
        }
      }
    };
    rec.start();
    resetSilence();
  }
  function resetSilence(){
    clearTimeout(silenceTimer);
    silenceTimer = setTimeout(()=>{
      onStatus('Auto-advance after silence.');
    }, getSilenceMs());
  }
  function stop(){
    if (rec){ rec.stop(); rec=null; }
    btn.textContent='Mic ▶';
  }

  btn.addEventListener('click', ()=>{ if (rec) stop(); else start(); });
  return { start, stop };
}
