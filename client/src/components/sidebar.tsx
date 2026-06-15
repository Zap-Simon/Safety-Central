import { useEffect, useState } from "react";
import { MicrosoftLogo } from "./ui/microsoft-logo";

const navigationItems = [
  { id: "overview", label: "Staff Engagement System", icon: "fas fa-users", type: "icon" },
  { id: "visual-overview", label: "Teams Viva & Data Flow", icon: "microsoft", type: "custom" },
  { id: "workflow-diagram", label: "Process Workflow", icon: "fas fa-project-diagram", type: "icon" },
  { id: "workflow-steps", label: "Process Walkthrough", icon: "fas fa-list-ol", type: "icon" },
  { id: "list-columns", label: "Data Structures", icon: "fas fa-table", type: "icon" },
  { id: "benefits", label: "Key Benefits", icon: "fas fa-star", type: "icon" },
];

interface SidebarProps {
  onBeforeNavigate?: () => void;
}

export default function Sidebar({ onBeforeNavigate }: SidebarProps = {}) {
  const [activeSection, setActiveSection] = useState("");

  useEffect(() => {
    const updateActiveNav = () => {
      const sections = document.querySelectorAll('section[id]');
      let current = '';
      const headerOffset = 150;
      
      const isNearBottom = (window.innerHeight + window.scrollY) >= document.body.offsetHeight - 200;
      
      if (isNearBottom) {
        const lastSection = sections[sections.length - 1];
        current = lastSection?.getAttribute('id') || '';
      } else {
        sections.forEach(section => {
          const sectionTop = (section as HTMLElement).offsetTop;
          if (window.scrollY >= sectionTop - headerOffset) {
            current = section.getAttribute('id') || '';
          }
        });
      }
      
      setActiveSection(current);
    };

    window.addEventListener('scroll', updateActiveNav);
    updateActiveNav();
    
    return () => {
      window.removeEventListener('scroll', updateActiveNav);
    };
  }, []);

  const handleNavClick = (id: string) => {
    if (onBeforeNavigate) {
      onBeforeNavigate();
      setTimeout(() => {
        const element = document.getElementById(id);
        if (element) {
          const headerHeight = 80;
          const elementPosition = element.offsetTop - headerHeight;
          window.scrollTo({ top: elementPosition, behavior: 'smooth' });
        }
      }, 320);
    } else {
      const element = document.getElementById(id);
      if (element) {
        const headerHeight = 80;
        const elementPosition = element.offsetTop - headerHeight;
        window.scrollTo({ top: elementPosition, behavior: 'smooth' });
      }
    }
  };

  return (
    <nav className="lg:col-span-1 print-hidden">
      <div className="bg-white/95 backdrop-blur-md rounded-lg shadow-lg border border-gray-200/50 p-4 lg:p-6 sticky top-4 lg:top-8 lg:fixed lg:top-[13rem] lg:w-64 lg:z-20 lg:max-h-[calc(100vh-15rem)] lg:overflow-y-auto lg:left-[max(1.5rem,calc((100vw-80rem)/2+1.5rem))]">
        <h2 className="font-semibold ms-gray-900 mb-3 lg:mb-4 text-sm lg:text-base">Navigation</h2>
        <ul className="space-y-1 lg:space-y-2">
          {navigationItems.map((item) => (
            <li key={item.id}>
              <button
                onClick={() => handleNavClick(item.id)}
                className={`nav-link flex items-center px-3 py-2 lg:py-2 rounded-md text-xs lg:text-sm w-full text-left transition-all duration-200 touch-manipulation ${
                  activeSection === item.id 
                    ? 'bg-blue-600 text-white shadow-md' 
                    : 'text-gray-700 hover:bg-gray-100 active:bg-gray-200 hover:shadow-sm'
                }`}
              >
                {item.type === "custom" && item.icon === "microsoft" ? (
                  <div className="mr-3">
                    <MicrosoftLogo 
                      size={12} 
                      variant={activeSection === item.id ? "white" : "grey"} 
                    />
                  </div>
                ) : (
                  <i className={`${item.icon} mr-3 text-xs ${activeSection === item.id ? 'text-white' : 'text-gray-400'}`}></i>
                )}
                {item.label}
              </button>
            </li>
          ))}
        </ul>
      </div>
    </nav>
  );
}
