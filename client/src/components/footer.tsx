export default function Footer() {
  const currentDate = new Date().toLocaleDateString('en-US', { 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  });

  return (
    <footer className="bg-white border-t border-gray-200 mt-8 sm:mt-16 print-hidden">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 sm:py-8">
        <div className="text-center text-xs sm:text-sm text-gray-600 space-y-1 sm:space-y-2">
          <p>Cranfield Glass Christchurch - HSEQ - Standards & Compliance</p>
          <p>System Documentation | Generated {currentDate} | Internal Use Only</p>
        </div>
      </div>
    </footer>
  );
}
