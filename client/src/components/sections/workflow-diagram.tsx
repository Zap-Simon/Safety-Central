import { useEffect, useState, useRef } from "react";
import { Dialog, DialogContent, DialogTrigger, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Expand, Move, ZoomIn, ZoomOut, RotateCcw } from "lucide-react";

export default function WorkflowDiagram() {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [scale, setScale] = useState(1);
  const [position, setPosition] = useState({ x: 0, y: 0 });
  const [isDragging, setIsDragging] = useState(false);
  const [dragStart, setDragStart] = useState({ x: 0, y: 0 });
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    // Re-render mermaid diagram when component mounts
    if (window.mermaid) {
      window.mermaid.init();
    }
  }, []);

  useEffect(() => {
    // Re-render mermaid when modal opens
    if (isModalOpen && window.mermaid) {
      setTimeout(() => {
        window.mermaid.init();
      }, 100);
    }
  }, [isModalOpen]);

  const handlePointerDown = (e: React.PointerEvent) => {
    e.preventDefault();
    setIsDragging(true);
    setDragStart({
      x: e.clientX - position.x,
      y: e.clientY - position.y
    });
  };

  const handlePointerMove = (e: React.PointerEvent) => {
    if (!isDragging) return;
    e.preventDefault();
    
    setPosition({
      x: e.clientX - dragStart.x,
      y: e.clientY - dragStart.y
    });
  };

  const handlePointerUp = () => {
    setIsDragging(false);
  };

  const zoomIn = () => setScale(prev => Math.min(prev + 0.2, 3));
  const zoomOut = () => setScale(prev => Math.max(prev - 0.2, 0.5));
  const resetView = () => {
    setScale(1);
    setPosition({ x: 0, y: 0 });
  };

  const mermaidDiagram = `graph TD
    A0[Staff Access Teams<br/>Viva Card] --> A[User Submits Form<br/>Safety/Business/Near Miss]
    A --> B[Power Automate Triggered]
    B --> C[Add to Relevant List]
    B --> D[Teams Notification]
    C --> E[Status: Submitted]
    E --> F[Fortnightly Meeting<br/>Review All Lists]
    F --> G[Review Submitted Items<br/>Safety, Business & Near Miss]
    F --> H[Review In Discussion Items<br/>Progress & Actions Achieved]
    G --> I{New Item Decision}
    H --> J{Action Progress}
    I -->|Needs Discussion| K[Status: In Discussion]
    I -->|Approved for Action| L[Status: Actioned]
    J -->|Still Active| M[Continue In Discussion]
    J -->|Completed| N[Mark as Actioned]
    K --> O[Next Meeting Review]
    L --> P[Power Automate<br/>Copy to Action List]
    N --> P
    P --> Q[Action List<br/>Status: Not Started]
    Q --> R[Admin Assignment & Tracking]
    R --> S[In Progress → Completed]
    S --> T[Auto-Update Original Item]
    T --> U[Status: Closed<br/>Removed from Board]
    O --> F
    
    style A0 fill:#6B46C1,stroke:#553C9A,color:#fff
    style A fill:#00BCF2,stroke:#0078D4,color:#fff
    style F fill:#FFB900,stroke:#CC9400,color:#000
    style G fill:#E3F2FD,stroke:#1976D2,color:#000
    style H fill:#E8F5E8,stroke:#388E3C,color:#000
    style U fill:#107C10,stroke:#0B5A0B,color:#fff`;

  return (
    <section id="workflow-diagram" className="bg-gradient-to-br from-white to-gray-50 rounded-xl shadow-lg p-4 sm:p-6 lg:p-8 mb-6 sm:mb-8 border border-gray-200">
      <div className="flex items-center justify-between mb-6">
        <div className="flex items-center space-x-3">
          <div className="w-10 h-10 bg-gradient-to-r from-blue-600 to-blue-700 rounded-xl flex items-center justify-center shadow-lg flex-shrink-0">
            <i className="fas fa-project-diagram text-white text-lg"></i>
          </div>
          <div className="min-w-0 flex-1">
            <h2 className="text-xl sm:text-2xl font-bold ms-gray-900">Process Workflow</h2>
            <p className="text-xs sm:text-sm text-gray-600">Complete process flow from submission to completion</p>
          </div>
        </div>
        
        {/* Expand Button */}
        <Dialog open={isModalOpen} onOpenChange={setIsModalOpen}>
          <DialogTrigger asChild>
            <button className="flex items-center space-x-2 px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition-colors touch-manipulation text-sm font-medium shadow-sm">
              <Expand className="w-4 h-4" />
              <span className="hidden sm:inline">Expand View</span>
            </button>
          </DialogTrigger>
          <DialogContent className="max-w-7xl w-[95vw] h-[90vh] p-0">
            <DialogHeader className="p-6 pb-2 border-b">
              <DialogTitle className="flex items-center justify-between">
                <div className="flex items-center space-x-3">
                  <div className="w-8 h-8 bg-gradient-to-r from-blue-600 to-blue-700 rounded-lg flex items-center justify-center">
                    <i className="fas fa-project-diagram text-white text-sm"></i>
                  </div>
                  <span className="text-xl">Process Workflow - Expanded View</span>
                </div>
                
                {/* Controls */}
                <div className="flex items-center space-x-2">
                  <div className="hidden md:flex items-center space-x-1 text-sm text-gray-500 mr-4">
                    <Move className="w-4 h-4" />
                    <span>Drag to move • Scroll to zoom</span>
                  </div>
                  
                  <div className="flex items-center space-x-1 bg-gray-100 rounded-lg p-1">
                    <button
                      onClick={zoomOut}
                      className="p-2 hover:bg-white rounded-md transition-colors touch-manipulation"
                      title="Zoom Out"
                    >
                      <ZoomOut className="w-4 h-4" />
                    </button>
                    
                    <span className="px-2 text-sm font-medium min-w-[3rem] text-center">
                      {Math.round(scale * 100)}%
                    </span>
                    
                    <button
                      onClick={zoomIn}
                      className="p-2 hover:bg-white rounded-md transition-colors touch-manipulation"
                      title="Zoom In"
                    >
                      <ZoomIn className="w-4 h-4" />
                    </button>
                    
                    <button
                      onClick={resetView}
                      className="p-2 hover:bg-white rounded-md transition-colors touch-manipulation"
                      title="Reset View"
                    >
                      <RotateCcw className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              </DialogTitle>
            </DialogHeader>
            
            <div 
              className="flex-1 overflow-hidden bg-gray-50"
              ref={containerRef}
              onPointerDown={handlePointerDown}
              onPointerMove={handlePointerMove}
              onPointerUp={handlePointerUp}
              onPointerLeave={handlePointerUp}
              onWheel={(e) => {
                e.preventDefault();
                const delta = e.deltaY > 0 ? -0.1 : 0.1;
                setScale(prev => Math.max(0.3, Math.min(3, prev + delta)));
              }}
              style={{ touchAction: 'none' }}
            >
              <div className="w-full h-full flex items-center justify-center p-6">
                <div 
                  className="bg-white rounded-lg border border-gray-200 shadow-lg overflow-visible"
                  style={{
                    transform: `translate(${position.x}px, ${position.y}px) scale(${scale})`,
                    transformOrigin: 'center center',
                    cursor: isDragging ? 'grabbing' : 'grab',
                    transition: isDragging ? 'none' : 'transform 0.2s ease'
                  }}
                >
                  <div className="mermaid p-8 min-w-[800px] min-h-[600px] select-none">
                    {mermaidDiagram}
                  </div>
                </div>
              </div>
            </div>
          </DialogContent>
        </Dialog>
      </div>
      
      <div className="bg-white rounded-lg border border-gray-200 p-3 sm:p-6 shadow-sm">
        <div className="overflow-x-auto scrollbar-thin scrollbar-thumb-gray-300 scrollbar-track-gray-100">
          <div className="mermaid min-w-max touch-pan-x">
            {mermaidDiagram}
          </div>
        </div>
        <div className="flex items-center justify-between mt-3">
          <p className="text-xs text-gray-500 sm:hidden">Swipe horizontally to view full diagram</p>
          <button 
            onClick={() => setIsModalOpen(true)}
            className="text-xs text-blue-600 hover:text-blue-700 font-medium sm:hidden"
          >
            Tap to expand →
          </button>
        </div>
      </div>
    </section>
  );
}
