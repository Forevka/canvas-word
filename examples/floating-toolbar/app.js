// Floating-toolbar customization demo. The `floatingToolbar` constructor option
// controls Word's selection mini-toolbar: pass `true`/`false` to toggle the default
// set, or an object to pick which controls appear (`buttons`), reorder them, add
// your own custom buttons, and decide whether it also shows at a bare caret
// (`onCaret`). This demo re-mounts the editor with each preset so you can compare.
import { WordCanvas } from "@forevka/wordcanvas";

const SAMPLE =
  "Select any of this text to see the floating format toolbar appear just above it. " +
  "Switch presets above to reshape the bar — trim it to a couple of buttons, reorder " +
  "them, show it at a bare caret, or add your own custom button.";

// Each preset is a value for the `floatingToolbar` option.
const PRESETS = {
  default: {
    label: "Default",
    // `true` (or omitting the option) shows the full built-in set.
    config: true,
  },
  minimal: {
    label: "Minimal (B / I / U)",
    // A trimmed set — only the toggles you list, in the order you list them.
    config: { buttons: ["bold", "italic", "underline"] },
  },
  reordered: {
    label: "Reordered, no font picker",
    // Built-in ids + "|" separators, arranged however you like.
    config: {
      buttons: ["fontSize", "|", "bold", "italic", "|", "color", "highlight", "|", "clearFormat"],
    },
  },
  onCaret: {
    label: "Show at caret too",
    // Also pop the bar at a collapsed caret, not only over a selection.
    config: { onCaret: true },
  },
  custom: {
    label: "With a custom button",
    // Mix built-ins with your own button. Its onClick gets the same
    // RibbonActionContext a custom ribbon button gets (editor handle + helpers).
    config: {
      buttons: [
        "bold",
        "italic",
        "|",
        "font",
        "fontSize",
        "|",
        {
          id: "demo.date",
          icon: "📅",
          tooltip: "Insert today's date",
          onClick: (ctx) => ctx.insertText(new Date().toLocaleDateString()),
        },
      ],
    },
  },
  disabled: {
    label: "Disabled",
    // `false` hides it entirely (the ribbon still has every command).
    config: false,
  },
};

const container = document.getElementById("editor");
const barEl = document.getElementById("bar");
const loader = document.getElementById("loader");
const progress = loader.querySelector(".bar");
const label = loader.querySelector(".label");
const PHASE_LABEL = { bundle: "Loading editor…", fonts: "Loading fonts…", ready: "Ready" };

let editor = null;
let current = "default";

async function mount(key) {
  current = key;
  for (const b of barEl.querySelectorAll("button")) b.classList.toggle("active", b.dataset.key === key);

  editor?.destroy();
  loader.classList.remove("done");
  editor = new WordCanvas({
    container,
    floatingToolbar: PRESETS[key].config,
    onLoadProgress: ({ phase, percent }) => {
      progress.style.width = `${Math.round(percent * 100)}%`;
      label.textContent = PHASE_LABEL[phase] ?? "";
      if (phase === "ready") loader.classList.add("done");
    },
  });
  window.__wc = editor;

  // Seed some text so the toolbar is immediately demonstrable.
  await editor.whenReady();
  await editor.insertText(SAMPLE);
}

// Build the preset switcher.
for (const [key, { label: text }] of Object.entries(PRESETS)) {
  const b = document.createElement("button");
  b.textContent = text;
  b.dataset.key = key;
  b.addEventListener("click", () => mount(key));
  barEl.appendChild(b);
}

void mount("default");
