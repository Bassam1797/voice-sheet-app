interface SpeechSetupOptions {
  onNumber: (val: string) => void;
  onStatus: (status: string) => void;
  getSilenceMs: () => number;
  getButton: () => HTMLButtonElement;
}

export function setupSpeech({
  onNumber,
  onStatus,
  getSilenceMs,
  getButton,
}: SpeechSetupOptions): { start: () => void; stop: () => void } {
  const SpeechRec = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
  const btn = getButton();
  let rec: SpeechRecognition | null = null;
  let silenceTimer: ReturnType<typeof setTimeout> | null = null;

  function start(): void {
    if (!SpeechRec) {
      onStatus('Speech Recognition not supported in this browser.');
      return;
    }

    if (rec) {
      stop();
      return;
    }

    rec = new SpeechRec();
    rec.lang = 'en-US';
    rec.continuous = true;
    rec.interimResults = true;

    rec.onstart = () => {
      btn.textContent = 'Mic ■';
      btn.setAttribute('aria-pressed', 'true');
      onStatus('Listening…');
    };

    rec.onend = () => {
      btn.textContent = 'Mic ▶';
      btn.setAttribute('aria-pressed', 'false');
      onStatus('Ready');
      rec = null;
    };

    rec.onerror = (e) => {
      onStatus('Mic error: ' + e.error);
    };

    rec.onresult = (e: SpeechRecognitionEvent) => {
      for (let i = e.resultIndex; i < e.results.length; i++) {
        const res = e.results[i];
        const txt = res[0].transcript.trim();
        if (res.isFinal) {
          const cleaned = txt.replace(/[,\s]/g, '');
          if (/^[-+]?\d*(?:\.\d+)?$/.test(cleaned) && cleaned !== '') {
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

  function resetSilence(): void {
    clearTimeout(silenceTimer!);
    silenceTimer = setTimeout(() => {
      onStatus('Auto-advance after silence.');
    }, getSilenceMs());
  }

  function stop(): void {
    if (rec) {
      rec.stop();
      rec = null;
    }
    btn.textContent = 'Mic ▶';
    btn.setAttribute('aria-pressed', 'false');
  }

  btn.addEventListener('click', () => {
    if (rec) stop();
    else start();
  });

  return { start, stop };
}
