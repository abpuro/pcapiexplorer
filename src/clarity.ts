const clarityProjectId = import.meta.env.VITE_CLARITY_PROJECT_ID?.trim();

declare global {
  interface Window {
    clarity?: ((command: string, value?: string) => void) & {
      q?: unknown[];
    };
  }
}

export function initClarity() {
  if (!clarityProjectId || window.clarity) {
    return;
  }

  window.clarity = function clarity(command: string, value?: string) {
    window.clarity!.q = window.clarity!.q ?? [];
    window.clarity!.q.push([command, value]);
  };

  const script = document.createElement("script");
  script.async = true;
  script.src = `https://www.clarity.ms/tag/${encodeURIComponent(clarityProjectId)}`;
  document.head.appendChild(script);
}

export function trackClarityEvent(name: string) {
  if (!clarityProjectId || !window.clarity) {
    return;
  }

  window.clarity("event", name);
}
