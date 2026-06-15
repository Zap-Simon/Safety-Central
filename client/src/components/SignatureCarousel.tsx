import { useState, useRef, useEffect, useCallback } from "react";
import { Button } from "@/components/ui/button";
import { Progress } from "@/components/ui/progress";
import { CheckCircle, ChevronLeft, PenLine, Monitor, UserX, RotateCcw, ArrowRight } from "lucide-react";

type SignatureStatus = 'signed' | 'remote' | 'absent';

interface SignatureRecord {
  status: SignatureStatus;
  signatureData: string | null;
  signedAt: string;
}

interface Attendee {
  name: string;
  role: string;
}

interface SignatureCarouselProps {
  meetingDate: string;
  attendees: Attendee[];
  existingSignatures: Record<string, SignatureRecord>;
  onSaveSignature: (attendeeName: string, status: SignatureStatus, signatureData: string | null, signedAt: string) => Promise<void>;
  onComplete: () => void;
  onClose: () => void;
}

function getInitials(name: string): string {
  return name.split(' ').map(n => n[0]).join('').slice(0, 2).toUpperCase();
}

function generateRemoteSignatureImage(name: string): string {
  const canvas = document.createElement('canvas');
  canvas.width = 320;
  canvas.height = 80;
  const ctx = canvas.getContext('2d')!;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.font = 'italic 36px Georgia, "Times New Roman", serif';
  ctx.fillStyle = '#1f2937';
  ctx.textBaseline = 'middle';
  ctx.fillText(name, 10, 44);
  return canvas.toDataURL('image/png');
}

export default function SignatureCarousel({
  meetingDate,
  attendees,
  existingSignatures,
  onSaveSignature,
  onComplete,
  onClose
}: SignatureCarouselProps) {
  const [currentIndex, setCurrentIndex] = useState(0);
  const [localSignatures, setLocalSignatures] = useState<Record<string, SignatureRecord>>(() => {
    const copy: Record<string, SignatureRecord> = {};
    for (const a of attendees) {
      if (existingSignatures[a.name]) copy[a.name] = existingSignatures[a.name];
    }
    return copy;
  });
  const [mode, setMode] = useState<'choose' | 'draw' | 'done'>('choose');
  const [hasDrawn, setHasDrawn] = useState(false);
  const [isSaving, setIsSaving] = useState(false);
  const [isComplete, setIsComplete] = useState(false);

  const canvasRef = useRef<HTMLCanvasElement>(null);
  const isDrawing = useRef(false);
  const lastPoint = useRef<{ x: number; y: number } | null>(null);

  const currentAttendee = attendees[currentIndex];
  const alreadySigned = currentAttendee && !!localSignatures[currentAttendee.name];

  const signedCount = Object.values(localSignatures).filter(s => s.status !== undefined).length;
  const progress = attendees.length > 0 ? (signedCount / attendees.length) * 100 : 0;

  useEffect(() => {
    setMode(alreadySigned ? 'choose' : 'choose');
    setHasDrawn(false);
    clearCanvas();
  }, [currentIndex]);

  const clearCanvas = useCallback(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    setHasDrawn(false);
  }, []);

  const getCanvasPos = (e: MouseEvent | Touch, canvas: HTMLCanvasElement) => {
    const rect = canvas.getBoundingClientRect();
    const scaleX = canvas.width / rect.width;
    const scaleY = canvas.height / rect.height;
    return {
      x: (e.clientX - rect.left) * scaleX,
      y: (e.clientY - rect.top) * scaleY
    };
  };

  const handleMouseDown = (e: React.MouseEvent<HTMLCanvasElement>) => {
    const canvas = canvasRef.current;
    if (!canvas) return;
    isDrawing.current = true;
    lastPoint.current = getCanvasPos(e.nativeEvent, canvas);
  };

  const handleMouseMove = (e: React.MouseEvent<HTMLCanvasElement>) => {
    if (!isDrawing.current) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const point = getCanvasPos(e.nativeEvent, canvas);
    if (lastPoint.current) {
      ctx.beginPath();
      ctx.moveTo(lastPoint.current.x, lastPoint.current.y);
      ctx.lineTo(point.x, point.y);
      ctx.strokeStyle = '#1f2937';
      ctx.lineWidth = 2.5;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.stroke();
      setHasDrawn(true);
    }
    lastPoint.current = point;
  };

  const handleMouseUp = () => {
    isDrawing.current = false;
    lastPoint.current = null;
  };

  const handleTouchStart = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    const canvas = canvasRef.current;
    if (!canvas) return;
    isDrawing.current = true;
    lastPoint.current = getCanvasPos(e.touches[0], canvas);
  };

  const handleTouchMove = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    if (!isDrawing.current) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext('2d');
    if (!ctx) return;
    const point = getCanvasPos(e.touches[0], canvas);
    if (lastPoint.current) {
      ctx.beginPath();
      ctx.moveTo(lastPoint.current.x, lastPoint.current.y);
      ctx.lineTo(point.x, point.y);
      ctx.strokeStyle = '#1f2937';
      ctx.lineWidth = 2.5;
      ctx.lineCap = 'round';
      ctx.lineJoin = 'round';
      ctx.stroke();
      setHasDrawn(true);
    }
    lastPoint.current = point;
  };

  const handleTouchEnd = (e: React.TouchEvent<HTMLCanvasElement>) => {
    e.preventDefault();
    isDrawing.current = false;
    lastPoint.current = null;
  };

  const saveSignature = async (status: SignatureStatus, signatureData: string | null) => {
    if (!currentAttendee) return;
    setIsSaving(true);
    const signedAt = new Date().toISOString();
    try {
      await onSaveSignature(currentAttendee.name, status, signatureData, signedAt);
      const record: SignatureRecord = { status, signatureData, signedAt };
      setLocalSignatures(prev => ({ ...prev, [currentAttendee.name]: record }));
      advance();
    } finally {
      setIsSaving(false);
    }
  };

  const advance = () => {
    if (currentIndex < attendees.length - 1) {
      setCurrentIndex(i => i + 1);
    } else {
      setIsComplete(true);
    }
  };

  const handleSign = () => setMode('draw');

  const handleConfirmSign = async () => {
    const canvas = canvasRef.current;
    if (!canvas || !hasDrawn) return;
    const dataUrl = canvas.toDataURL('image/png');
    await saveSignature('signed', dataUrl);
    setMode('choose');
  };

  const handleRemote = async () => {
    const img = generateRemoteSignatureImage(currentAttendee.name);
    await saveSignature('remote', img);
  };

  const handleAbsent = async () => {
    await saveSignature('absent', null);
  };

  const handleBack = () => {
    if (currentIndex > 0) {
      setCurrentIndex(i => i - 1);
      setMode('choose');
    }
  };

  const getStatusIcon = (status: SignatureStatus) => {
    if (status === 'signed') return '✍️';
    if (status === 'remote') return '🖥️';
    return '—';
  };

  const getStatusLabel = (status: SignatureStatus) => {
    if (status === 'signed') return 'Signed';
    if (status === 'remote') return 'Remote';
    return 'Absent';
  };

  const getStatusColour = (status: SignatureStatus) => {
    if (status === 'signed') return 'bg-green-100 text-green-700 border-green-200';
    if (status === 'remote') return 'bg-purple-100 text-purple-700 border-purple-200';
    return 'bg-gray-100 text-gray-500 border-gray-200';
  };

  if (isComplete) {
    return (
      <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4">
        <div className="bg-white rounded-2xl shadow-2xl w-full max-w-md p-8 text-center">
          <div className="w-20 h-20 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-5">
            <CheckCircle className="h-10 w-10 text-green-600" />
          </div>
          <h2 className="text-2xl font-bold text-gray-900 mb-2">All Done!</h2>
          <p className="text-gray-500 mb-2">All signatures collected for this meeting.</p>
          <p className="text-sm text-gray-400 mb-8">You can export the HTML minutes now to include the signatures, or collect more later.</p>

          <div className="space-y-2 mb-6 text-left">
            {attendees.map(a => {
              const sig = localSignatures[a.name];
              if (!sig) return null;
              return (
                <div key={a.name} className={`flex items-center gap-3 px-3 py-2 rounded-lg border text-sm ${getStatusColour(sig.status)}`}>
                  <span className="text-base">{getStatusIcon(sig.status)}</span>
                  <div className="flex-1 min-w-0">
                    <div className="font-medium truncate">{a.name}</div>
                    <div className="text-xs opacity-70">{a.role}</div>
                  </div>
                  <span className="text-xs font-medium">{getStatusLabel(sig.status)}</span>
                </div>
              );
            })}
          </div>

          <Button onClick={onComplete} className="w-full bg-green-600 hover:bg-green-700 text-white">
            Close
          </Button>
        </div>
      </div>
    );
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-2 sm:p-4">
      <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg flex flex-col overflow-hidden" style={{ maxHeight: '95dvh' }}>

        {/* Header */}
        <div className="bg-gradient-to-r from-blue-600 to-blue-700 text-white px-5 py-4 flex-shrink-0">
          <div className="flex items-center justify-between mb-3">
            <div>
              <h2 className="font-bold text-lg leading-tight">Meeting Sign-Off</h2>
              <p className="text-blue-200 text-sm">{meetingDate}</p>
            </div>
            <div className="text-right">
              <div className="text-2xl font-bold">{signedCount}<span className="text-blue-300 text-base font-normal"> / {attendees.length}</span></div>
              <div className="text-blue-200 text-xs">signed</div>
            </div>
          </div>
          <Progress value={progress} className="h-2 bg-blue-500 [&>div]:bg-white" />
        </div>

        {/* Attendee mini-track */}
        <div className="flex gap-1.5 px-5 py-3 bg-gray-50 border-b border-gray-100 overflow-x-auto flex-shrink-0">
          {attendees.map((a, i) => {
            const sig = localSignatures[a.name];
            const isCurrent = i === currentIndex;
            return (
              <button
                key={a.name}
                onClick={() => { setCurrentIndex(i); setMode('choose'); }}
                title={a.name}
                className={`flex-shrink-0 w-9 h-9 rounded-full text-xs font-bold transition-all border-2 ${
                  isCurrent ? 'border-blue-500 bg-blue-600 text-white scale-110 shadow' :
                  sig ? (sig.status === 'signed' ? 'border-green-400 bg-green-100 text-green-700' :
                         sig.status === 'remote' ? 'border-purple-400 bg-purple-100 text-purple-700' :
                         'border-gray-300 bg-gray-100 text-gray-400') :
                  'border-gray-200 bg-white text-gray-500 hover:border-blue-300'
                }`}
              >
                {sig ? getStatusIcon(sig.status) : getInitials(a.name)}
              </button>
            );
          })}
        </div>

        {/* Main content */}
        <div className="flex-1 overflow-y-auto px-5 py-5">

          {/* Attendee card */}
          <div className="flex items-center gap-4 mb-5">
            <div className="w-14 h-14 rounded-full bg-gradient-to-br from-blue-500 to-blue-600 flex items-center justify-center text-white font-bold text-lg flex-shrink-0 shadow-md">
              {getInitials(currentAttendee?.name ?? '')}
            </div>
            <div className="flex-1 min-w-0">
              <div className="text-xl font-bold text-gray-900 truncate">{currentAttendee?.name}</div>
              <div className="text-sm text-gray-500">{currentAttendee?.role}</div>
              {alreadySigned && (
                <span className={`inline-flex items-center gap-1 text-xs px-2 py-0.5 rounded-full border mt-1 ${getStatusColour(localSignatures[currentAttendee.name].status)}`}>
                  {getStatusIcon(localSignatures[currentAttendee.name].status)} {getStatusLabel(localSignatures[currentAttendee.name].status)}
                </span>
              )}
            </div>
          </div>

          {mode === 'choose' && (
            <div className="space-y-3">
              {alreadySigned && localSignatures[currentAttendee.name].signatureData && (
                <div className="rounded-xl border border-gray-200 bg-gray-50 p-3 mb-2">
                  <div className="text-xs text-gray-500 mb-1">Current signature:</div>
                  <img src={localSignatures[currentAttendee.name].signatureData!} alt="signature" className="max-h-14 max-w-full" />
                </div>
              )}
              <p className="text-sm text-gray-500 mb-4">{alreadySigned ? 'Re-sign or change status:' : 'How is this person attending?'}</p>
              <Button
                onClick={handleSign}
                className="w-full h-14 text-base bg-blue-600 hover:bg-blue-700 text-white flex items-center gap-3"
                disabled={isSaving}
              >
                <PenLine className="h-5 w-5" />
                Sign ✍️
              </Button>
              <Button
                onClick={handleRemote}
                variant="outline"
                className="w-full h-14 text-base border-purple-300 text-purple-700 hover:bg-purple-50 flex items-center gap-3"
                disabled={isSaving}
              >
                <Monitor className="h-5 w-5" />
                Attended Remotely 🖥️
              </Button>
              <Button
                onClick={handleAbsent}
                variant="outline"
                className="w-full h-14 text-base border-gray-300 text-gray-600 hover:bg-gray-50 flex items-center gap-3"
                disabled={isSaving}
              >
                <UserX className="h-5 w-5" />
                Absent
              </Button>
            </div>
          )}

          {mode === 'draw' && (
            <div>
              <p className="text-sm text-gray-500 mb-3">Draw your signature below:</p>
              <div className="border-2 border-dashed border-gray-300 rounded-xl overflow-hidden bg-gray-50 touch-none">
                <canvas
                  ref={canvasRef}
                  width={480}
                  height={160}
                  className="w-full cursor-crosshair block"
                  style={{ touchAction: 'none' }}
                  onMouseDown={handleMouseDown}
                  onMouseMove={handleMouseMove}
                  onMouseUp={handleMouseUp}
                  onMouseLeave={handleMouseUp}
                  onTouchStart={handleTouchStart}
                  onTouchMove={handleTouchMove}
                  onTouchEnd={handleTouchEnd}
                />
              </div>
              <div className="flex items-center justify-between mt-3 gap-3">
                <Button variant="outline" size="sm" onClick={() => { clearCanvas(); setMode('choose'); }} className="flex items-center gap-1.5 text-gray-600">
                  <RotateCcw className="h-4 w-4" />
                  Back
                </Button>
                <Button variant="outline" size="sm" onClick={clearCanvas} className="flex items-center gap-1.5 text-gray-500">
                  <RotateCcw className="h-3.5 w-3.5" />
                  Clear
                </Button>
                <Button
                  onClick={handleConfirmSign}
                  disabled={!hasDrawn || isSaving}
                  className="flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white"
                  size="sm"
                >
                  {isSaving ? 'Saving…' : <>Confirm <ArrowRight className="h-4 w-4" /></>}
                </Button>
              </div>
            </div>
          )}
        </div>

        {/* Footer navigation */}
        <div className="px-5 py-4 border-t border-gray-100 flex items-center justify-between flex-shrink-0 bg-white">
          <Button
            variant="ghost"
            onClick={handleBack}
            disabled={currentIndex === 0}
            className="flex items-center gap-1.5 text-gray-500"
          >
            <ChevronLeft className="h-4 w-4" />
            Back
          </Button>
          <span className="text-sm text-gray-400">{currentIndex + 1} of {attendees.length}</span>
          <Button variant="ghost" onClick={onClose} className="text-gray-400 text-sm">
            Save &amp; Close
          </Button>
        </div>
      </div>
    </div>
  );
}
