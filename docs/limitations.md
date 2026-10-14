# Known limitations

Each item below is a requirement that cannot be met safely within the Office.js APIs, or has not been verified yet. They are recorded here rather than simulated. The product reports each one to the user where it applies, as a task limitation, an *unverified* or *unsupported* status, or a manual step in an acceptance record.

## Not possible, or not safe, with Office.js

| Requirement | Limitation | What the product does instead |
|---|---|---|
| Rendered-slide QA on every PowerPoint | `Slide.getImageAsBase64` exists only in PowerPointApi 1.8. Older builds can only be inspected as objects | `presentation.check_rendering` returns `status: "unsupported"`. The specialist must say the slides were not visually checked, and acceptance records list a visual step for a person. It is never shown as passed |
| Exact text overflow without rendering | Office.js does not report rendered line breaks or text extents | Overflow is estimated from font size, box size and text (`basis: "text-estimate"`) and labelled as an estimate. The rendered check confirms it where supported |
| Slide size | Not exposed by Office.js | 960×540 (16:9) is assumed and stated in layout findings |
| Speaker notes in live PowerPoint | No Office.js API | The notes capability is not offered; the Presentation Author reports notes it could not add, and the action is *partial* |
| Native PowerPoint charts | Office.js cannot create chart objects in PowerPoint | Bar charts are drawn with editable shapes, and labels and bar heights are verified |
| Verifying an Outlook reply window | `displayReplyForm` opens a form the add-in cannot read back | The draft is reported *unverified* and the task *partial*. In compose mode, inserted text is read back and verified |
| Sending email | Deliberately not offered (invariant) | Drafts only; the user sends from Outlook |
| Reading another open document from an add-in | Each add-in instance sees only its own host document | Cross-app sources require the user to press *Share* in the source app and attach the source in the target app. Sources pass through the companion's memory, expire after 30 minutes and are read-only |
| Batch-approving destructive changes | Policy floors require confirmation for destructive, send and external operations | Those operations always prompt, even inside a batch-approved plan |
| Undo for every change | Office.js has no transactional rollback | Compensations exist where they are safe (restore text or cells, delete added slides or paragraphs); otherwise the UI points to Office's own Undo |

## Not yet verified

| Area | Status | How to close it |
|---|---|---|
| Live Office | No scenario has been run in a real Office host. All 7 acceptance scenarios pass in **mocked** mode only | Run the acceptance suite in each Office build ([acceptance.md](acceptance.md)) and commit the records |
| Real models | No model has been evaluated. The harness is tested against reference scripts and an Ollama stand-in only | Run `loai-companion eval` on real hardware and commit the run files ([model-compatibility.md](model-compatibility.md)) |
| Rendered checks on real slides | The pixel analysis is tested on synthetic renders. Real `getImageAsBase64` output (anti-aliasing, themes, pictures) has not been seen | Run the PowerPoint acceptance scenario on PowerPointApi 1.8 and compare with the visual steps |
| Office on the web | Not a launch target: a web-hosted pane generally cannot reach `localhost` | — |
| Installers on clean machines | Release packaging, checksums and signatures are verified in CI; the installers have not been run on clean macOS/Windows machines | Install from a signed release on each OS and record the result |
