import { actionExtractor } from "./action-extractor.js";
import { replyWriter } from "./reply-writer.js";
import { threadAnalyst } from "./thread-analyst.js";

export { actionExtractor, replyWriter, threadAnalyst };
export const OUTLOOK_SPECIALISTS = [threadAnalyst, replyWriter, actionExtractor];
