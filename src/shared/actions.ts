/**
 * This file is shared by the task pane and the helper server.
 *
 * Keeping actions in one place prevents bugs where the UI sends an action
 * value that the helper does not understand.
 */

export type ActionId =
  | "rewrite"
  | "summarise"
  | "simplify"
  | "expand"
  | "shorten"
  | "formal"
  | "friendly"
  | "clarity"
  | "grammar"
  | "to_bullets"
  | "bullets_to_prose"
  | "academic"
  | "plain_english"
  | "email_polish"
  | "student_feedback"
  | "action_list"
  | "key_points";

export interface TextAction {
  id: ActionId;
  label: string;
  helperInstruction: string;
}

export const TEXT_ACTIONS: TextAction[] = [
  {
    id: "rewrite",
    label: "Rewrite",
    helperInstruction: "Rewrite the text to improve clarity, grammar, and flow while preserving the original meaning."
  },
  {
    id: "summarise",
    label: "Summarise",
    helperInstruction: "Produce a concise summary containing only the most important ideas."
  },
  {
    id: "simplify",
    label: "Simplify",
    helperInstruction: "Rewrite the text for a non-expert reader using plain, accessible language."
  },
  {
    id: "expand",
    label: "Expand",
    helperInstruction: "Elaborate on the text with more explanation and smoother transitions without inventing facts."
  },
  {
    id: "shorten",
    label: "Shorten",
    helperInstruction: "Make the text shorter while keeping the essential meaning and important details."
  },
  {
    id: "formal",
    label: "Make more formal",
    helperInstruction: "Rewrite the text in a more formal tone while keeping it natural and readable."
  },
  {
    id: "friendly",
    label: "Make more friendly",
    helperInstruction: "Rewrite the text in a warmer, friendlier tone without becoming overly casual."
  },
  {
    id: "clarity",
    label: "Improve clarity",
    helperInstruction: "Improve clarity, remove ambiguity, and make the text easier to follow."
  },
  {
    id: "grammar",
    label: "Fix grammar",
    helperInstruction: "Correct grammar, spelling, punctuation, and sentence structure while preserving meaning."
  },
  {
    id: "to_bullets",
    label: "Convert to bullet points",
    helperInstruction: "Convert the text into clear bullet points. Keep only useful content and preserve meaning."
  },
  {
    id: "bullets_to_prose",
    label: "Convert bullets to prose",
    helperInstruction: "Convert the bullet points into coherent prose with smooth transitions."
  },
  {
    id: "academic",
    label: "Academic tone",
    helperInstruction: "Rewrite in a clear academic style without sounding pompous or adding unsupported claims."
  },
  {
    id: "plain_english",
    label: "Plain English",
    helperInstruction: "Remove jargon and make the text easier to understand while preserving technical accuracy."
  },
  {
    id: "email_polish",
    label: "Email polish",
    helperInstruction: "Polish the text as a clear, professional email while preserving the intended message."
  },
  {
    id: "student_feedback",
    label: "Student feedback tone",
    helperInstruction: "Rewrite as constructive student feedback: clear, kind, specific, and actionable."
  },
  {
    id: "action_list",
    label: "Action list extraction",
    helperInstruction: "Return a short bullet list of concrete actions only. Do not include background explanation."
  },
  {
    id: "key_points",
    label: "Key points only",
    helperInstruction: "Extract only the key points as a concise bullet list."
  }
];

export function findAction(actionId: string): TextAction | undefined {
  return TEXT_ACTIONS.find((action) => action.id === actionId);
}
