import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Plus, Trash2 } from "lucide-react";
import { LIKELIHOODS_LIST, CONSEQUENCES_LIST, getRiskLevelForCell } from "./riskUtils";

export interface HazardRow {
  id: string;
  hazard: string;
  likelihood: string;
  consequence: string;
  risk: string;
  control: string;
}

const RISK_COLORS: Record<string, string> = {
  "Low":      "bg-green-100 text-green-800 border-green-300",
  "Moderate": "bg-yellow-100 text-yellow-800 border-yellow-300",
  "High":     "bg-red-100 text-red-800 border-red-300",
  "Extreme":  "bg-gray-900 text-white border-gray-700",
};

interface HazardTableProps {
  rows: HazardRow[];
  onChange: (rows: HazardRow[]) => void;
}

export default function HazardTable({ rows, onChange }: HazardTableProps) {
  const addRow = () => {
    const newRow: HazardRow = {
      id: `h-${Date.now()}`,
      hazard: "",
      likelihood: "",
      consequence: "",
      risk: "",
      control: "",
    };
    onChange([...rows, newRow]);
  };

  const updateRow = (id: string, field: keyof HazardRow, value: string) => {
    const updated = rows.map(r => {
      if (r.id !== id) return r;
      const updatedRow = { ...r, [field]: value };
      if (field === "likelihood" || field === "consequence") {
        const likelihood = field === "likelihood" ? value : r.likelihood;
        const consequence = field === "consequence" ? value : r.consequence;
        if (likelihood && consequence) {
          updatedRow.risk = getRiskLevelForCell(likelihood, consequence);
        }
      }
      return updatedRow;
    });
    onChange(updated);
  };

  const removeRow = (id: string) => {
    onChange(rows.filter(r => r.id !== id));
  };

  return (
    <div className="space-y-3">
      <div className="overflow-x-auto rounded-lg border border-gray-200">
        <table className="w-full text-xs">
          <thead className="bg-gray-50">
            <tr>
              <th className="text-left p-2 font-semibold text-gray-700 min-w-[140px]">Hazard</th>
              <th className="text-left p-2 font-semibold text-gray-700 min-w-[110px]">Likelihood</th>
              <th className="text-left p-2 font-semibold text-gray-700 min-w-[110px]">Consequence</th>
              <th className="text-left p-2 font-semibold text-gray-700 min-w-[80px]">Risk</th>
              <th className="text-left p-2 font-semibold text-gray-700 min-w-[180px]">Control / Response</th>
              <th className="p-2 w-8"></th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 && (
              <tr>
                <td colSpan={6} className="p-4 text-center text-gray-400 italic text-sm">
                  No hazards added yet. Click "Add Hazard" below.
                </td>
              </tr>
            )}
            {rows.map(row => (
              <tr key={row.id} className="border-t border-gray-100 hover:bg-gray-50">
                <td className="p-1.5">
                  <input
                    type="text"
                    value={row.hazard}
                    onChange={e => updateRow(row.id, "hazard", e.target.value)}
                    placeholder="Describe hazard..."
                    className="w-full border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:border-orange-400"
                  />
                </td>
                <td className="p-1.5">
                  <select
                    value={row.likelihood}
                    onChange={e => updateRow(row.id, "likelihood", e.target.value)}
                    className="w-full border border-gray-200 rounded px-1 py-1 text-xs focus:outline-none focus:border-orange-400"
                  >
                    <option value="">Select...</option>
                    {LIKELIHOODS_LIST.map(l => (
                      <option key={l} value={l}>{l}</option>
                    ))}
                  </select>
                </td>
                <td className="p-1.5">
                  <select
                    value={row.consequence}
                    onChange={e => updateRow(row.id, "consequence", e.target.value)}
                    className="w-full border border-gray-200 rounded px-1 py-1 text-xs focus:outline-none focus:border-orange-400"
                  >
                    <option value="">Select...</option>
                    {CONSEQUENCES_LIST.map(c => (
                      <option key={c} value={c}>{c}</option>
                    ))}
                  </select>
                </td>
                <td className="p-1.5">
                  {row.risk ? (
                    <Badge className={`text-[10px] font-bold px-1.5 py-0.5 border ${RISK_COLORS[row.risk] || "bg-gray-100 text-gray-700"}`}>
                      {row.risk.toUpperCase()}
                    </Badge>
                  ) : (
                    <span className="text-gray-300 text-[10px]">—</span>
                  )}
                </td>
                <td className="p-1.5">
                  <input
                    type="text"
                    value={row.control}
                    onChange={e => updateRow(row.id, "control", e.target.value)}
                    placeholder="Control measure..."
                    className="w-full border border-gray-200 rounded px-2 py-1 text-xs focus:outline-none focus:border-orange-400"
                  />
                </td>
                <td className="p-1.5">
                  <button
                    onClick={() => removeRow(row.id)}
                    className="text-gray-300 hover:text-red-500 transition-colors"
                    title="Remove hazard"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <Button
        type="button"
        variant="outline"
        onClick={addRow}
        className="w-full border-dashed border-orange-300 text-orange-600 hover:bg-orange-50 text-xs h-8"
      >
        <Plus className="h-3.5 w-3.5 mr-1" />
        Add Hazard
      </Button>
    </div>
  );
}
