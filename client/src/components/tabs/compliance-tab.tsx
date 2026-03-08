import { Link } from "wouter";
import { useState, useEffect } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { FileText, Shield, Users, ExternalLink, DollarSign, type LucideIcon } from "lucide-react";
import { apiRequest } from "@/lib/queryClient";
import { authService } from "@/auth/authService";

interface DragState {
  draggedIndex: number | null;
  draggedOverIndex: number | null;
}

interface ComplianceCard {
  id: string;
  title: string;
  description: string;
  icon: LucideIcon;
  iconColor: string;
  backgroundColor: string;
  href?: string;
  externalUrl?: string;
  comingSoon?: boolean;
  requiresAuth?: boolean;
  protectedPath?: string;
}

const defaultComplianceCards: ComplianceCard[] = [
  {
    id: "contract-management",
    title: "Contract Management",
    description: "Contract compliance & tracking",
    icon: FileText,
    iconColor: "text-blue-600",
    backgroundColor: "bg-blue-100",
    comingSoon: true
  },
  {
    id: "insurances",
    title: "Insurances",
    description: "Insurance policies & coverage",
    icon: Shield,
    iconColor: "text-green-600",
    backgroundColor: "bg-green-100",
    comingSoon: true
  },
  {
    id: "hr",
    title: "HR",
    description: "Human resources compliance",
    icon: Users,
    iconColor: "text-orange-600",
    backgroundColor: "bg-orange-100",
    externalUrl: "https://app.brighthr.com.au/dashboard"
  },
  {
    id: "sales-pricing-standards",
    title: "Sales & Pricing Standards",
    description: "Sales compliance & pricing standards",
    icon: DollarSign,
    iconColor: "text-emerald-600",
    backgroundColor: "bg-emerald-100",
    comingSoon: true
  },
  {
    id: "aroflo-templates",
    title: "Aroflo Templates",
    description: "Document templates & resources",
    icon: FileText,
    iconColor: "text-purple-600",
    backgroundColor: "bg-purple-100",
    externalUrl: "https://templates.glaziermate.tech/"
  },
  {
    id: "xero-price-exporter",
    title: "Xero Price Exporter",
    description: "Xero price exporter to Aroflo",
    icon: DollarSign,
    iconColor: "text-teal-600",
    backgroundColor: "bg-teal-100",
    externalUrl: "https://sales.glaziermate.tech"
  }
];

export default function ComplianceTab() {
  const [cards, setCards] = useState<ComplianceCard[]>(defaultComplianceCards);
  const [dragMode, setDragMode] = useState(false);
  const [dragState, setDragState] = useState<DragState>({
    draggedIndex: null,
    draggedOverIndex: null
  });
  const [hasLoadedFromDB, setHasLoadedFromDB] = useState(false);
  
  const queryClient = useQueryClient();

  // Load card ordering from the database for compliance cards
  const { data: cardOrdering } = useQuery({
    queryKey: ['/api/compliance-card-ordering']
  });

  // Save card ordering to the database for compliance cards
  const saveCardOrderingMutation = useMutation({
    mutationFn: async (cardOrders: { cardId: string; position: number }[]) => {
      const response = await apiRequest('POST', '/api/compliance-card-ordering', { cardOrders });
      return response.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['/api/compliance-card-ordering'] });
    }
  });

  // Apply saved ordering when data is loaded (only once)
  useEffect(() => {
    if (cardOrdering && (cardOrdering as any)?.success && (cardOrdering as any).cardOrdering && !hasLoadedFromDB) {
      const savedOrder = (cardOrdering as any).cardOrdering;
      
      if (savedOrder.length > 0) {
        // Create ordered cards based on saved ordering
        const orderedCards = [...defaultComplianceCards].sort((a, b) => {
          const aOrder = savedOrder.find((item: any) => item.cardId === a.id);
          const bOrder = savedOrder.find((item: any) => item.cardId === b.id);
          
          if (!aOrder && !bOrder) return 0;
          if (!aOrder) return 1;
          if (!bOrder) return -1;
          
          return aOrder.position - bOrder.position;
        });
        
        setCards(orderedCards);
        setHasLoadedFromDB(true);
      } else {
        // No saved order, use default
        setCards([...defaultComplianceCards]);
        setHasLoadedFromDB(true);
      }
    }
  }, [cardOrdering, hasLoadedFromDB]);

  const handleDragStart = (e: React.DragEvent, index: number) => {
    setDragState(prev => ({ ...prev, draggedIndex: index }));
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/html', '');
  };

  const handleDragOver = (e: React.DragEvent, index: number) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    setDragState(prev => ({ ...prev, draggedOverIndex: index }));
  };

  const handleDragLeave = () => {
    setDragState(prev => ({ ...prev, draggedOverIndex: null }));
  };

  const handleDrop = (e: React.DragEvent, dropIndex: number) => {
    e.preventDefault();
    const { draggedIndex } = dragState;
    
    if (draggedIndex === null || draggedIndex === dropIndex) {
      setDragState({ draggedIndex: null, draggedOverIndex: null });
      return;
    }

    const newCards = [...cards];
    const draggedCard = newCards[draggedIndex];
    
    // Remove the dragged card and insert it at the new position
    newCards.splice(draggedIndex, 1);
    newCards.splice(dropIndex, 0, draggedCard);
    
    setCards(newCards);
    setDragState({ draggedIndex: null, draggedOverIndex: null });
    
    // Save the new ordering to the database
    const cardOrders = newCards.map((card, index) => ({
      cardId: card.id,
      position: index
    }));
    
    saveCardOrderingMutation.mutate(cardOrders);
  };

  const handleDragEnd = () => {
    setDragState({ draggedIndex: null, draggedOverIndex: null });
  };

  const renderCard = (card: ComplianceCard, index: number) => {
    const IconComponent = card.icon;
    const isDragging = dragState.draggedIndex === index;
    const isDraggedOver = dragState.draggedOverIndex === index;
    
    const cardClasses = `
      relative rounded-xl p-4 transition-all text-left w-full h-32 flex flex-col group
      ${card.externalUrl 
        ? 'bg-gradient-to-br from-blue-50 to-blue-100 border-2 border-blue-200 hover:from-blue-100 hover:to-blue-150 hover:border-blue-300 hover:shadow-lg' 
        : 'bg-white border border-gray-200 hover:shadow-md'
      }
      ${dragMode ? 'cursor-move' : (card.href || card.externalUrl) ? 'cursor-pointer' : 'cursor-default'}
      ${isDragging ? 'opacity-50 scale-95' : ''}
      ${isDraggedOver ? 'border-blue-400 bg-blue-50' : ''}
    `;

    const cardContent = (
      <>
        <div className={`w-10 h-10 ${card.backgroundColor} rounded-lg flex items-center justify-center mb-3`}>
          <IconComponent className={card.iconColor} size={20} />
        </div>
        <h4 className="font-semibold text-gray-900 mb-1 flex items-center gap-2">
          {card.title}
          {card.externalUrl && (
            <ExternalLink size={14} className="text-blue-600" />
          )}
        </h4>
        <p className={`text-sm flex-1 ${card.externalUrl ? 'text-blue-700' : 'text-gray-600'}`}>
          {card.description}
        </p>
        {card.externalUrl && (
          <span className="absolute top-2 right-2 bg-blue-600 text-white text-xs px-2 py-1 rounded-full font-medium">
            External Link
          </span>
        )}
        {card.comingSoon && (
          <span className="absolute top-2 right-2 bg-yellow-100 text-yellow-800 text-xs px-2 py-1 rounded-full font-medium">
            Coming Soon
          </span>
        )}
      </>
    );

    if (card.externalUrl) {
      // External URL - open in new window/tab
      const handleExternalClick = () => {
        if (dragMode) return; // Don't handle clicks in drag mode
        window.open(card.externalUrl!, '_blank', 'noopener,noreferrer');
      };

      return (
        <div
          key={card.id}
          className={cardClasses}
          onClick={handleExternalClick}
          draggable={dragMode}
          onDragStart={dragMode ? (e) => handleDragStart(e, index) : undefined}
          onDragOver={dragMode ? (e) => handleDragOver(e, index) : undefined}
          onDragLeave={dragMode ? handleDragLeave : undefined}
          onDrop={dragMode ? (e) => handleDrop(e, index) : undefined}
          onDragEnd={dragMode ? handleDragEnd : undefined}
        >
          {cardContent}
        </div>
      );
    }

    if (card.href) {
      // Internal routing for all other cards
      return (
        <Link key={card.id} href={card.href} className="block">
          <div
            className={cardClasses}
            draggable={dragMode}
            onDragStart={dragMode ? (e) => handleDragStart(e, index) : undefined}
            onDragOver={dragMode ? (e) => handleDragOver(e, index) : undefined}
            onDragLeave={dragMode ? handleDragLeave : undefined}
            onDrop={dragMode ? (e) => handleDrop(e, index) : undefined}
            onDragEnd={dragMode ? handleDragEnd : undefined}
          >
            {cardContent}
          </div>
        </Link>
      );
    }

    if (card.requiresAuth && card.protectedPath) {
      const handleProtectedClick = async () => {
        if (dragMode) return; // Don't handle clicks in drag mode
        
        try {
          if (card.id === 'contract-management') {
            await authService.openHealthSafetyPolicy();
          } else if (card.id === 'insurances') {
            await authService.openHealthSafetyPolicy();
          } else if (card.id === 'hr') {
            await authService.openHealthSafetyPolicy();
          }
        } catch (error) {
          console.error('Failed to open protected resource:', error);
        }
      };

      return (
        <div
          key={card.id}
          className={cardClasses}
          draggable={dragMode}
          onClick={handleProtectedClick}
          onDragStart={dragMode ? (e) => handleDragStart(e, index) : undefined}
          onDragOver={dragMode ? (e) => handleDragOver(e, index) : undefined}
          onDragLeave={dragMode ? handleDragLeave : undefined}
          onDrop={dragMode ? (e) => handleDrop(e, index) : undefined}
          onDragEnd={dragMode ? handleDragEnd : undefined}
        >
          {cardContent}
        </div>
      );
    }

    return (
      <div
        key={card.id}
        className={cardClasses}
        draggable={dragMode}
        onDragStart={dragMode ? (e) => handleDragStart(e, index) : undefined}
        onDragOver={dragMode ? (e) => handleDragOver(e, index) : undefined}
        onDragLeave={dragMode ? handleDragLeave : undefined}
        onDrop={dragMode ? (e) => handleDrop(e, index) : undefined}
        onDragEnd={dragMode ? handleDragEnd : undefined}
      >
        {cardContent}
      </div>
    );
  };

  const renderExternalCard = (card: ComplianceCard) => {
    const IconComponent = card.icon;
    
    const handleExternalClick = () => {
      window.open(card.externalUrl!, '_blank', 'noopener,noreferrer');
    };

    return (
      <div
        key={card.id}
        onClick={handleExternalClick}
        className="relative rounded-lg p-3 transition-all text-left cursor-pointer bg-gradient-to-br from-blue-50 to-blue-100 border border-blue-200 hover:from-blue-100 hover:to-blue-150 hover:border-blue-300 hover:shadow-md group"
      >
        <div className="flex items-center gap-3">
          <div className={`w-8 h-8 ${card.backgroundColor} rounded-md flex items-center justify-center flex-shrink-0`}>
            <IconComponent className={card.iconColor} size={16} />
          </div>
          <div className="flex-1 min-w-0">
            <h5 className="font-semibold text-gray-900 text-sm flex items-center gap-1.5">
              {card.title}
              <ExternalLink size={12} className="text-blue-600" />
            </h5>
            <p className="text-xs text-blue-700 truncate">
              {card.description}
            </p>
          </div>
        </div>
      </div>
    );
  };

  const internalCards: ComplianceCard[] = [];
  const externalCards: ComplianceCard[] = [];

  cards.forEach(card => {
    if (card.externalUrl) {
      externalCards.push(card);
    } else {
      internalCards.push(card);
    }
  });

  return (
    <div className="p-6 space-y-8">
      {/* Main Compliance Cards */}
      <div className="space-y-6">
        <div className="flex justify-end mb-6">
          <div className="flex items-center space-x-3">
            <div className="flex items-center">
              <button
                type="button"
                className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors duration-200 ease-in-out focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 ${
                  dragMode ? 'bg-blue-600' : 'bg-gray-200'
                }`}
                onClick={() => setDragMode(!dragMode)}
              >
                <span
                  className={`inline-block h-4 w-4 rounded-full bg-white transition-transform duration-200 ease-in-out ${
                    dragMode ? 'translate-x-6' : 'translate-x-1'
                  }`}
                />
              </button>
            </div>
            <label className="text-sm text-gray-600 cursor-pointer" onClick={() => setDragMode(!dragMode)}>
              Reorder mode
            </label>
          </div>
        </div>
        
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          {cards.map((card, index) => {
            if (card.externalUrl) return null;
            return renderCard(card, index);
          })}
        </div>
      </div>

      {/* External Links Section */}
      {externalCards.length > 0 && (
        <div className="space-y-4 pt-6 border-t border-gray-200">
          <h3 className="text-lg font-semibold text-gray-900 flex items-center gap-2">
            <ExternalLink size={20} className="text-blue-600" />
            External Links
          </h3>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-3">
            {externalCards.map((card) => renderExternalCard(card))}
          </div>
        </div>
      )}
    </div>
  );
}