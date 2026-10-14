import { createRoot } from "react-dom/client";
import { App } from "./ui/App";
import { createHost } from "./services/host";
import "./ui/styles.css";

function mount(host: Office.HostType | null) {
  const demoKind = new URLSearchParams(location.search).get("demo");
  const info = createHost(host, demoKind);
  createRoot(document.getElementById("root")!).render(<App host={info} />);
}

if (typeof Office !== "undefined" && Office.onReady) {
  // Outside Office, onReady resolves with host = null → demo mode.
  Office.onReady((i) => mount(i.host ?? null));
} else {
  mount(null);
}
