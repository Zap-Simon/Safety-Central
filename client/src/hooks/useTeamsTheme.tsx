import { createContext, useContext, useState, useEffect } from "react";
import type { ReactNode } from "react";
import * as microsoftTeams from "@microsoft/teams-js";

export type TeamsTheme = "default" | "dark" | "contrast";

interface TeamsThemeContextValue {
  theme: TeamsTheme;
  inTeams: boolean;
  isDark: boolean;
}

const TeamsThemeContext = createContext<TeamsThemeContextValue>({
  theme: "default",
  inTeams: false,
  isDark: false,
});

export function TeamsThemeProvider({ children }: { children: ReactNode }) {
  const [theme, setTheme] = useState<TeamsTheme>("default");
  const [inTeams, setInTeams] = useState(false);

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        await microsoftTeams.app.initialize();
        if (cancelled) return;
        const ctx = await microsoftTeams.app.getContext();
        if (cancelled) return;
        setInTeams(true);
        setTheme((ctx.app.theme as TeamsTheme) ?? "default");
        microsoftTeams.app.registerOnThemeChangeHandler((t) => {
          if (!cancelled) setTheme((t as TeamsTheme) ?? "default");
        });
      } catch {
        // running outside Teams — context stays at defaults
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <TeamsThemeContext.Provider value={{ theme, inTeams, isDark: theme !== "default" }}>
      {children}
    </TeamsThemeContext.Provider>
  );
}

export function useTeamsTheme(): TeamsThemeContextValue {
  return useContext(TeamsThemeContext);
}
