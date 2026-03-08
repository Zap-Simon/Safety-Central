import { Link } from "wouter";
import { useState, useEffect } from "react";
import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { defaultCardOrder, type CardData } from "@/lib/card-data";
import { apiRequest } from "@/lib/queryClient";
import { authService } from "@/auth/authService";

interface DragState {
  draggedIndex: number | null;
  draggedOverIndex: number | null;
}

export default function OverviewTab() {
  const [cards, setCards] = useState<CardData[]>(defaultCardOrder);
  const [dragMode, setDragMode] = useState(false);
  const [dragState, setDragState] = useState<DragState>({
    draggedIndex: null,
    draggedOverIndex: null
  });
  const [hasLoadedFromDB, setHasLoadedFromDB] = useState(false);
  
  const queryClient = useQueryClient();

  // Load card ordering from the database
  const { data: cardOrdering } = useQuery({
    queryKey: ['/api/card-ordering']
  });

  // Save card ordering to the database
  const saveCardOrderingMutation = useMutation({
    mutationFn: async (cardOrders: { cardId: string; position: number }[]) => {
      const response = await apiRequest('POST', '/api/card-ordering', { cardOrders });
      return response.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ['/api/card-ordering'] });
    }
  });

  // Apply saved ordering when data is loaded (only once)
  useEffect(() => {
    if (cardOrdering && (cardOrdering as any)?.success && (cardOrdering as any).cardOrdering && !hasLoadedFromDB) {
      const savedOrder = (cardOrdering as any).cardOrdering;
      
      if (savedOrder.length > 0) {
        // Create ordered cards based on saved ordering
        const orderedCards = [...defaultCardOrder].sort((a, b) => {
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
        setCards([...defaultCardOrder]);
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

  const renderCard = (card: CardData, index: number) => {
    const IconComponent = card.icon;
    const isDragging = dragState.draggedIndex === index;
    const isDraggedOver = dragState.draggedOverIndex === index;
    
    const cardClasses = `
      relative bg-white border border-gray-200 rounded-xl p-4 transition-all text-left w-full h-32 flex flex-col group
      ${dragMode ? 'cursor-move' : card.href ? 'cursor-pointer' : 'cursor-default'}
      ${isDragging ? 'opacity-50 scale-95' : 'hover:shadow-md'}
      ${isDraggedOver ? 'border-blue-400 bg-blue-50' : ''}
    `;

    const cardContent = (
      <>
        
        <div className={`w-10 h-10 ${card.backgroundColor} rounded-lg flex items-center justify-center mb-3`}>
          <IconComponent className={card.iconColor} size={20} />
        </div>
        <h4 className="font-semibold text-gray-900 mb-1">{card.title}</h4>
        <p className="text-gray-600 text-sm flex-1">{card.description}</p>
        {card.comingSoon && (
          <span className="absolute top-2 right-2 bg-yellow-100 text-yellow-800 text-xs px-2 py-1 rounded-full font-medium">
            Coming Soon
          </span>
        )}
      </>
    );

    if (card.href) {
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
          if (card.id === 'health-safety-policy') {
            await authService.openHealthSafetyPolicy();
          }
        } catch (error) {
          console.error('Failed to open protected resource:', error);
          // Could show a toast notification here
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

  return (
    <div className="p-6 space-y-6">
      {/* Health & Safety Navigation */}
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
          {cards.map((card, index) => renderCard(card, index))}
        </div>
      </div>
    </div>
  );
}