import { Badge } from "@/components/ui/badge";

export type Likelihood = "Rare" | "Unlikely" | "Possible" | "Very Likely" | "Almost Certain";
export type Consequence = "Superficial" | "Minor" | "Moderate" | "Major" | "Catastrophic";
export type RiskLevel = "Low" | "Moderate" | "High" | "Extreme";

const LIKELIHOODS: Likelihood[] = ["Rare", "Unlikely", "Possible", "Very Likely", "Almost Certain"];
const CONSEQUENCES: Consequence[] = ["Superficial", "Minor", "Moderate", "Major", "Catastrophic"];
const LIKELIHOOD_LABELS: Record<Likelihood, string> = {
  "Rare": "Rare (<5%)",
  "Unlikely": "Unlikely (5-10%)",
  "Possible": "Possible (10-50%)",
  "Very Likely": "Very Likely (50-75%)",
  "Almost Certain": "Almost Certain (>75%)",
};

const RISK_MATRIX: Record<Likelihood, Record<Consequence, RiskLevel>> = {
  "Rare":          { "Superficial": "Low",      "Minor": "Low",      "Moderate": "Low",      "Major": "Low",      "Catastrophic": "Moderate" },
  "Unlikely":      { "Superficial": "Low",      "Minor": "Low",      "Moderate": "Low",      "Major": "Moderate", "Catastrophic": "High"     },
  "Possible":      { "Superficial": "Low",      "Minor": "Low",      "Moderate": "Moderate", "Major": "Moderate", "Catastrophic": "High"     },
  "Very Likely":   { "Superficial": "Low",      "Minor": "Moderate", "Moderate": "Moderate", "Major": "High",     "Catastrophic": "High"     },
  "Almost Certain":{ "Superficial": "Moderate", "Minor": "Moderate", "Moderate": "High",     "Major": "High",     "Catastrophic": "Extreme"  },
};

const RISK_COLORS: Record<RiskLevel, string> = {
  "Low":      "bg-green-500 text-white",
  "Moderate": "bg-yellow-400 text-gray-900",
  "High":     "bg-red-500 text-white",
  "Extreme":  "bg-black text-white",
};

const CELL_COLORS: Record<RiskLevel, string> = {
  "Low":      "bg-green-400 hover:bg-green-500",
  "Moderate": "bg-yellow-300 hover:bg-yellow-400",
  "High":     "bg-red-400 hover:bg-red-500",
  "Extreme":  "bg-gray-900 hover:bg-black text-white",
};

export function getRiskLevel(likelihood: Likelihood, consequence: Consequence): RiskLevel {
  return RISK_MATRIX[likelihood]?.[consequence] ?? "Low";
}

interface RiskMatrixProps {
  selectedLikelihood: Likelihood | "";
  selectedConsequence: Consequence | "";
  onSelect: (likelihood: Likelihood, consequence: Consequence) => void;
}

export default function RiskMatrix({ selectedLikelihood, selectedConsequence, onSelect }: RiskMatrixProps) {
  const selectedRisk = selectedLikelihood && selectedConsequence
    ? getRiskLevel(selectedLikelihood as Likelihood, selectedConsequence as Consequence)
    : null;

  return (
    <div className="space-y-3">
      <div className="overflow-x-auto">
        <table className="w-full border-collapse text-xs">
          <thead>
            <tr>
              <th className="w-28 p-1" aria-hidden="true"></th>
              <th colSpan={CONSEQUENCES.length} className="p-1 text-center font-semibold text-gray-700 text-[10px] uppercase tracking-wide">
                Consequence →
              </th>
            </tr>
            <tr>
              <th className="w-28 text-left p-1 font-semibold text-gray-700 text-[10px] uppercase tracking-wide">
                Likelihood ↓
              </th>
              {CONSEQUENCES.map(c => (
                <th key={c} className="p-1 text-center font-medium text-gray-700 text-[10px] min-w-[72px]">{c}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {LIKELIHOODS.map(likelihood => (
              <tr key={likelihood}>
                <td className="p-1 text-[10px] text-gray-600 font-medium whitespace-nowrap pr-2">
                  {LIKELIHOOD_LABELS[likelihood]}
                </td>
                {CONSEQUENCES.map(consequence => {
                  const risk = RISK_MATRIX[likelihood][consequence];
                  const isSelected = likelihood === selectedLikelihood && consequence === selectedConsequence;
                  return (
                    <td
                      key={consequence}
                      onClick={() => onSelect(likelihood, consequence)}
                      className={`p-1 text-center font-bold text-[11px] cursor-pointer transition-all border-2 rounded ${CELL_COLORS[risk]} ${
                        isSelected ? "ring-2 ring-offset-1 ring-blue-600 border-blue-600 scale-105 shadow-md z-10 relative" : "border-transparent"
                      }`}
                    >
                      {risk}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {selectedRisk && (
        <div className="flex items-center gap-3 p-3 bg-gray-50 rounded-lg border border-gray-200">
          <div className="text-sm text-gray-600">Risk Score:</div>
          <Badge className={`text-sm font-bold px-3 py-1 ${RISK_COLORS[selectedRisk]}`}>
            {selectedRisk}
          </Badge>
          <div className="text-xs text-gray-500">
            {selectedLikelihood} × {selectedConsequence}
          </div>
        </div>
      )}
    </div>
  );
}
