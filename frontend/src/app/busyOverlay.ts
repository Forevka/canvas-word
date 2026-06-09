// A full-screen "busy" overlay shown while a blocking async action runs: opening
// a .docx from disk, joining a shared document via a ?collab link, or publishing
// a share link. Mirrors the modal pattern of identityPopup.ts and is styled by
// ui/styles.ts (.cw-loading-*). Covering the screen also debounces double-clicks
// for the duration of the operation, so callers need no extra button-disable.

export interface BusyHandle {
  /** Update the label; pass pct (0–100) to reveal a determinate progress bar. */
  update(label: string, pct?: number): void;
  /** Remove the overlay. Idempotent — safe to call from a finally + an early path. */
  done(): void;
}

export function showBusy(label: string): BusyHandle {
  const back = document.createElement("div");
  back.className = "cw-loading-overlay";

  const card = document.createElement("div");
  card.className = "cw-loading-card";

  const spinner = document.createElement("div");
  spinner.className = "cw-spinner";

  const text = document.createElement("div");
  text.className = "cw-loading-label";
  text.textContent = label;

  const progress = document.createElement("div");
  progress.className = "cw-progress";
  progress.style.display = "none";
  const bar = document.createElement("div");
  bar.className = "cw-progress-bar";
  progress.appendChild(bar);

  card.append(spinner, text, progress);
  back.appendChild(card);
  document.body.appendChild(back);

  let removed = false;
  return {
    update(nextLabel: string, pct?: number): void {
      text.textContent = nextLabel;
      if (pct === undefined || Number.isNaN(pct)) {
        progress.style.display = "none";
      } else {
        progress.style.display = "";
        bar.style.width = `${Math.max(0, Math.min(100, pct))}%`;
      }
    },
    done(): void {
      if (removed) return;
      removed = true;
      back.remove();
    },
  };
}
