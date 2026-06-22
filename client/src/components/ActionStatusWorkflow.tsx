import { Circle, Loader2, ClipboardCheck, CheckCircle2, PauseCircle, CalendarClock } from "lucide-react";

export interface ActionStatusUpdate {
  actionStatus?: string;
  reconsiderDate?: string;
}

export interface ActionStatusWorkflowProps {
  status: string;
  reconsiderDate?: string;
  onChange: (updates: ActionStatusUpdate) => void;
  compact?: boolean;
  // When false, the "Completed" step can't be selected here (e.g. Near Miss items
  // are only signed off in the meeting minutes once the investigation is done).
  allowComplete?: boolean;
}

const STAGES = [
  {
    value: "",
    label: "Not Started",
    short: "Not Started",
    desc: "Logged and waiting — nobody has picked it up yet.",
    icon: Circle,
  },
  {
    value: "In Progress",
    label: "In Progress",
    short: "In Progress",
    desc: "Someone is actively working on this now.",
    icon: Loader2,
  },
  {
    value: "Ready to Close",
    label: "Ready to Close",
    short: "Ready to Close",
    desc: "Work looks done — bring it to the next meeting for a final check before closing. You can sit here as long as you need while getting input.",
    icon: ClipboardCheck,
  },
  {
    value: "Completed",
    label: "Completed",
    short: "Completed",
    desc: "Signed off and closed. No further action needed.",
    icon: CheckCircle2,
  },
];

const stageIndex = (status: string) => {
  const idx = STAGES.findIndex((s) => s.value === (status || ""));
  return idx === -1 ? 0 : idx;
};

export default function ActionStatusWorkflow({
  status,
  reconsiderDate,
  onChange,
  compact = false,
  allowComplete = true,
}: ActionStatusWorkflowProps) {
  const normalized = status || "";
  const onHold = normalized === "On Hold";
  const currentIdx = onHold ? -1 : stageIndex(normalized);
  const activeStage = onHold ? null : STAGES[currentIdx];

  // Moving to any normal stage clears a leftover reconsider date so it can't
  // silently resurface the item later. Completing is blocked when allowComplete
  // is false (the item is finished elsewhere, e.g. in the meeting minutes).
  const goToStage = (value: string) => {
    if (value === "Completed" && !allowComplete) return;
    onChange({ actionStatus: value, reconsiderDate: "" });
  };

  return (
    <div className="space-y-2.5">
      {/* Stepper row */}
      <div className={`flex items-stretch ${onHold ? "opacity-50" : ""}`}>
        {STAGES.map((stage, i) => {
          const Icon = stage.icon;
          const isDone = !onHold && i < currentIdx;
          const isCurrent = !onHold && i === currentIdx;
          const isLocked = stage.value === "Completed" && !allowComplete;

          const circleClasses = isCurrent
            ? "bg-amber-500 text-white ring-4 ring-amber-200 border-amber-500"
            : isDone
            ? "bg-green-500 text-white border-green-500"
            : "bg-white text-gray-400 border-gray-300";

          return (
            <div key={stage.value || "not-started"} className="flex-1 flex flex-col items-center relative">
              {/* connector line to the left */}
              {i > 0 && (
                <div
                  className={`absolute top-4 right-1/2 left-0 h-0.5 -translate-y-1/2 ${
                    !onHold && i <= currentIdx ? "bg-green-500" : "bg-gray-300"
                  }`}
                />
              )}
              <button
                type="button"
                onClick={() => goToStage(stage.value)}
                disabled={isLocked}
                className={`relative z-10 flex items-center justify-center rounded-full border-2 transition-all ${circleClasses} ${
                  compact ? "h-7 w-7" : "h-8 w-8"
                } ${isLocked ? "cursor-not-allowed opacity-50" : "hover:scale-105"}`}
                title={isLocked ? "Near Miss items are signed off in the meeting minutes once the investigation is complete." : stage.desc}
                data-testid={`status-step-${stage.value || "not-started"}`}
              >
                <Icon className={`${compact ? "h-3.5 w-3.5" : "h-4 w-4"} ${isCurrent && stage.value === "In Progress" ? "animate-spin" : ""}`} />
              </button>
              <span
                className={`mt-1 text-center leading-tight ${compact ? "text-[9px]" : "text-[10px]"} ${
                  isCurrent ? "font-bold text-amber-700" : isDone ? "font-medium text-green-700" : "text-gray-500"
                }`}
              >
                {stage.short}
              </span>
            </div>
          );
        })}
      </div>

      {/* Active-stage description */}
      {!onHold && activeStage && (
        <div className="flex items-start gap-2 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">
          <activeStage.icon className="h-3.5 w-3.5 text-amber-600 mt-0.5 flex-shrink-0" />
          <p className="text-xs text-amber-900 leading-snug">{activeStage.desc}</p>
        </div>
      )}

      {/* Completion is handled in the meeting minutes for these items */}
      {!allowComplete && (
        <p className="text-[10px] text-gray-500 leading-snug px-0.5">
          This item is signed off and completed in the meeting minutes once the investigation form is finished.
        </p>
      )}

      {/* On Hold branch */}
      <div className={`rounded-lg border-2 transition-all ${onHold ? "border-orange-400 bg-orange-50" : "border-dashed border-gray-300 bg-white"}`}>
        <button
          type="button"
          onClick={() => (onHold ? onChange({ actionStatus: "In Progress", reconsiderDate: "" }) : onChange({ actionStatus: "On Hold" }))}
          className="w-full flex items-center gap-2 px-3 py-2 text-left"
          data-testid="status-step-on-hold"
        >
          <PauseCircle className={`${compact ? "h-4 w-4" : "h-5 w-5"} flex-shrink-0 ${onHold ? "text-orange-600" : "text-gray-400"}`} />
          <div className="flex-1 min-w-0">
            <div className={`text-xs font-semibold ${onHold ? "text-orange-800" : "text-gray-600"}`}>
              {onHold ? "On Hold — paused" : "Put On Hold"}
            </div>
            <div className="text-[10px] text-gray-500 leading-tight">
              Not right now / maybe later. Set a date to bring it back.
            </div>
          </div>
          <span className={`text-[10px] font-medium px-2 py-0.5 rounded-full ${onHold ? "bg-orange-200 text-orange-800" : "bg-gray-100 text-gray-500"}`}>
            {onHold ? "Tap to resume" : "Tap to pause"}
          </span>
        </button>

        {onHold && (
          <div className="px-3 pb-3 pt-1 space-y-1.5 border-t border-orange-200">
            <label className="text-[10px] font-semibold text-orange-800 uppercase tracking-wide flex items-center gap-1">
              <CalendarClock className="h-3 w-3" /> Reconsider on
            </label>
            <input
              type="date"
              value={reconsiderDate ? reconsiderDate.split("T")[0] : ""}
              onChange={(e) => onChange({ reconsiderDate: e.target.value })}
              className="w-full text-xs border-2 border-orange-300 rounded-md px-2 py-1.5 bg-white focus:border-orange-500 focus:outline-none"
              data-testid="input-reconsider-date"
            />
            <p className="text-[10px] text-orange-700 leading-tight">
              On that date this item pops back into the meeting so the team can review it again.
            </p>
          </div>
        )}
      </div>
    </div>
  );
}
