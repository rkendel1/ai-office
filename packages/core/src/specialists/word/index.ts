import { documentEditor } from "./editor.js";
import { documentReviewer } from "./reviewer.js";
import { documentSummarizer } from "./summarizer.js";
import { documentWriter } from "./writer.js";

export { documentEditor, documentReviewer, documentSummarizer, documentWriter };
export const WORD_SPECIALISTS = [documentWriter, documentEditor, documentReviewer, documentSummarizer];
