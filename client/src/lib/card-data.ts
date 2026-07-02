import { Calendar, Shield, Users, ClipboardList, HardHat, Settings, Target, Leaf, Heart } from "lucide-react";
import { LucideIcon } from "lucide-react";

export interface CardData {
  id: string;
  title: string;
  description: string;
  icon: LucideIcon;
  iconColor: string;
  backgroundColor: string;
  comingSoon?: boolean;
  href?: string;
  requiresAuth?: boolean;
  protectedPath?: string;
}

export const defaultCardOrder: CardData[] = [
  // Functional cards first (no comingSoon property)
  {
    id: "meeting-minutes",
    title: "Meeting Minutes",
    description: "Toolbox & H&S meeting records",
    icon: Calendar,
    iconColor: "text-blue-600",
    backgroundColor: "bg-blue-100",
    href: "/meeting-history",
  },
  {
    id: "staff-training",
    title: "Staff Training",
    description: "Job workflows & training materials",
    icon: HardHat,
    iconColor: "text-orange-600",
    backgroundColor: "bg-orange-100",
    href: "/staff-training",
  },
  {
    id: "equipment-maintenance",
    title: "Equipment, Maintenance & Inspections",
    description: "Equipment maintenance & inspection records",
    icon: Settings,
    iconColor: "text-green-600",
    backgroundColor: "bg-green-100",
    href: "/equipment-maintenance",
  },
  {
    id: "accreditations",
    title: "Accreditations & Pre-quals",
    description: "Certification & qualification records",
    icon: Shield,
    iconColor: "text-purple-600",
    backgroundColor: "bg-purple-100",
    comingSoon: true,
  },
  {
    id: "chemicals",
    title: "Chemicals & Hazardous Substances",
    description: "Chemical register & safety data sheets",
    icon: Shield,
    iconColor: "text-orange-600",
    backgroundColor: "bg-orange-100",
    comingSoon: true,
  },
  {
    id: "contractors",
    title: "Contractors",
    description: "Contractor management & induction",
    icon: Users,
    iconColor: "text-blue-600",
    backgroundColor: "bg-blue-100",
    comingSoon: true,
  },
  {
    id: "emergency-procedures",
    title: "Emergency Procedures",
    description: "Emergency response plans",
    icon: Shield,
    iconColor: "text-green-600",
    backgroundColor: "bg-green-100",
    comingSoon: true,
  },
  {
    id: "health-monitoring",
    title: "Health Monitoring",
    description: "Staff health tracking & wellness programs",
    icon: Heart,
    iconColor: "text-pink-600",
    backgroundColor: "bg-pink-100",
    comingSoon: true,
  },
  // Actions is now fully functional
  {
    id: "actions",
    title: "Actions",
    description: "Meeting actions tracking & administration",
    icon: ClipboardList,
    iconColor: "text-amber-600",
    backgroundColor: "bg-amber-100",
    href: "/actions",
  },
  {
    id: "hazard-register",
    title: "Hazard Register",
    description: "Operational hazards & controls",
    icon: Shield,
    iconColor: "text-orange-600",
    backgroundColor: "bg-orange-100",
    href: "/hazard-register",
  },
  {
    id: "health-safety-policy",
    title: "Health & Safety Policy",
    description: "Official signed policy document",
    icon: Shield,
    iconColor: "text-red-600",
    backgroundColor: "bg-red-100",
    href: "/policy/health-safety",
  },
  {
    id: "incidents-reports",
    title: "Incidents & Reports",
    description: "Accident register & investigation",
    icon: Shield,
    iconColor: "text-red-600",
    backgroundColor: "bg-red-100",
    comingSoon: true,
  },
  {
    id: "task-analysis",
    title: "Task Analysis",
    description: "Job safety analysis & procedures",
    icon: Shield,
    iconColor: "text-indigo-600",
    backgroundColor: "bg-indigo-100",
    comingSoon: true,
  },
];