import { deckReviewer } from "./deck-reviewer.js";
import { presentationAuthor } from "./presentation-author.js";
import { slideDesigner } from "./slide-designer.js";

export { deckReviewer, presentationAuthor, slideDesigner };
export const POWERPOINT_SPECIALISTS = [presentationAuthor, slideDesigner, deckReviewer];
