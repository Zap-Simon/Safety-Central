export const LIKELIHOODS_LIST = ["Rare", "Unlikely", "Possible", "Very Likely", "Almost Certain"] as const;
export const CONSEQUENCES_LIST = ["Superficial", "Minor", "Moderate", "Major", "Catastrophic"] as const;

const RISK_MATRIX: Record<string, Record<string, string>> = {
  "Rare":          { "Superficial": "Low",      "Minor": "Low",      "Moderate": "Low",      "Major": "Low",      "Catastrophic": "Moderate" },
  "Unlikely":      { "Superficial": "Low",      "Minor": "Low",      "Moderate": "Low",      "Major": "Moderate", "Catastrophic": "High"     },
  "Possible":      { "Superficial": "Low",      "Minor": "Low",      "Moderate": "Moderate", "Major": "Moderate", "Catastrophic": "High"     },
  "Very Likely":   { "Superficial": "Low",      "Minor": "Moderate", "Moderate": "Moderate", "Major": "High",     "Catastrophic": "High"     },
  "Almost Certain":{ "Superficial": "Moderate", "Minor": "Moderate", "Moderate": "High",     "Major": "High",     "Catastrophic": "Extreme"  },
};

export function getRiskLevelForCell(likelihood: string, consequence: string): string {
  return RISK_MATRIX[likelihood]?.[consequence] ?? "";
}
