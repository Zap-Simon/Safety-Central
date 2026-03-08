interface MicrosoftLogoProps {
  size?: number;
  className?: string;
  variant?: "colored" | "white" | "grey";
}

export function MicrosoftLogo({ size = 16, className = "", variant = "colored" }: MicrosoftLogoProps) {
  const colors = variant === "white" 
    ? {
        red: "#FFFFFF",
        green: "#FFFFFF", 
        blue: "#FFFFFF",
        yellow: "#FFFFFF"
      }
    : variant === "grey"
    ? {
        red: "#9CA3AF",
        green: "#9CA3AF",
        blue: "#9CA3AF",
        yellow: "#9CA3AF"
      }
    : {
        red: "#F25022",
        green: "#7FBA00",
        blue: "#00A4EF", 
        yellow: "#FFB900"
      };

  return (
    <svg 
      width={size} 
      height={size} 
      viewBox="0 0 16 16" 
      className={className}
      xmlns="http://www.w3.org/2000/svg"
    >
      {/* Top left */}
      <rect x="0" y="0" width="7" height="7" fill={colors.red} />
      {/* Top right */}
      <rect x="8.5" y="0" width="7" height="7" fill={colors.green} />
      {/* Bottom left */}
      <rect x="0" y="8.5" width="7" height="7" fill={colors.blue} />
      {/* Bottom right */}
      <rect x="8.5" y="8.5" width="7" height="7" fill={colors.yellow} />
    </svg>
  );
}