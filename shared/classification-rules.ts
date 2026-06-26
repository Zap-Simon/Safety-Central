// Single source of truth for how a staff member's free-text submission in the
// Teams Improve+ Submit tab is categorised and routed to a SharePoint list.
//
// Both the server (which builds the AI classification prompt in
// `server/openai-service.ts`) and the client (which renders the confirm/
// follow-up cards in `client/src/pages/teams/SubmitTab.tsx`) read from this
// file, so the category list, list routing, and follow-up questions can never
// drift apart. To add or change a category, edit ONLY this file (plus the
// React icon/colour entry in SubmitTab's CATEGORY_META).

export type ListTarget = "near-miss" | "safety-ideas" | "business-ideas";

// Ordered list of every category. Used to derive the Category union type and
// the validation lists. Order is the order the AI sees them in the prompt.
export const CATEGORY_NAMES = [
  "Near Miss",
  "Safety Observation",
  "Improvement Idea",
  "Business Improvement",
  "Supply Request",
  "Near Miss Meeting Agenda Item",
  "Safety Meeting Agenda Item",
  "Business Meeting Agenda Item",
  "Other",
] as const;

export type Category = (typeof CATEGORY_NAMES)[number];

export const LIST_TARGETS: ListTarget[] = ["near-miss", "safety-ideas", "business-ideas"];

// Maps a list target to the SharePoint listType string used by the
// /api/sharepoint/create-item endpoint.
export const LIST_TARGET_LIST_TYPE: Record<ListTarget, string> = {
  "near-miss": "Near Miss",
  "safety-ideas": "Safety Ideas",
  "business-ideas": "Business Ideas",
};

export interface CategoryRule {
  // The category name the AI returns and the app keys off.
  name: Category;
  // Which SharePoint list this category is filed into. This is authoritative:
  // the server derives the final listTarget from the category's rule, so the
  // AI can't route a category to the wrong list.
  listTarget: ListTarget;
  // Human-readable destination shown on the confirm card ("Goes to: ...").
  listLabel: string;
  // Plain-English description injected into the AI system prompt.
  aiDescription: string;
  // Questions offered when the AI's confidence is low (< 0.8).
  followUpQuestions: string[];
}

export const CATEGORY_RULES: CategoryRule[] = [
  {
    name: "Near Miss",
    listTarget: "near-miss",
    listLabel: "Near Miss Register",
    aiDescription: "something that almost caused injury or damage but didn't",
    followUpQuestions: [
      "Where did this happen?",
      "What task were you doing?",
      "What was the potential injury or damage?",
      "What caused it?",
      "Was any PPE being worn?",
    ],
  },
  {
    name: "Safety Observation",
    listTarget: "safety-ideas",
    listLabel: "Safety Ideas List",
    aiDescription: "a hazard, unsafe condition, or safety concern observed",
    followUpQuestions: [
      "Where is the hazard located?",
      "How serious is the risk?",
      "Who could be affected?",
    ],
  },
  {
    name: "Improvement Idea",
    listTarget: "safety-ideas",
    listLabel: "Safety Ideas List",
    aiDescription: "an idea to improve safety processes, equipment, or procedures",
    followUpQuestions: [
      "Which area or process does this improve?",
      "What is the expected benefit?",
    ],
  },
  {
    name: "Business Improvement",
    listTarget: "business-ideas",
    listLabel: "Business Ideas List",
    aiDescription: "an idea to improve business processes, efficiency, or operations",
    followUpQuestions: [
      "Which area or process does this improve?",
      "What problem does it solve?",
    ],
  },
  {
    name: "Supply Request",
    listTarget: "business-ideas",
    listLabel: "Business Ideas List",
    aiDescription: "a need for supplies, equipment, or materials",
    followUpQuestions: [
      "What quantity is needed?",
      "Is this urgent?",
      "Which job or area is it for?",
    ],
  },
  {
    name: "Near Miss Meeting Agenda Item",
    listTarget: "near-miss",
    listLabel: "Near Miss Register",
    aiDescription:
      "an agenda item to discuss a specific near-miss incident, accident, or its investigation at the next H&S meeting",
    followUpQuestions: [
      "Which incident does this relate to?",
      "Why is this important to discuss?",
    ],
  },
  {
    name: "Safety Meeting Agenda Item",
    listTarget: "safety-ideas",
    listLabel: "Safety Ideas List",
    aiDescription:
      "an agenda item about a safety topic — hazards, the safety or hazard register, PPE, safety processes, or safety training — to discuss at the next H&S meeting",
    followUpQuestions: [
      "Why is this important to discuss?",
      "Is there a deadline?",
    ],
  },
  {
    name: "Business Meeting Agenda Item",
    listTarget: "business-ideas",
    listLabel: "Business Ideas List",
    aiDescription:
      "an agenda item about a business or operational topic — rosters, pricing, customers, scheduling, or efficiency — to discuss at the next H&S meeting",
    followUpQuestions: [
      "Why is this important to discuss?",
      "Is there a deadline?",
    ],
  },
  {
    name: "Other",
    listTarget: "business-ideas",
    listLabel: "Business Ideas List",
    aiDescription: "doesn't fit any of the above",
    followUpQuestions: [],
  },
];

export function getCategoryRule(name: string): CategoryRule | undefined {
  return CATEGORY_RULES.find((rule) => rule.name === name);
}
