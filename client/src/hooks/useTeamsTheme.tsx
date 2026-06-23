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

  // Drive Tailwind's `.dark` class off the Teams theme so every non-Fluent
  // element (shadcn inputs, the body background, the safe-area filler) follows
  // dark/light in lockstep with the Fluent components. Scoped to the Teams app
  // only: this provider lives under the Teams routes, so the main website never
  // gets `.dark` and its always-light styling is unaffected. Cleanup on unmount
  // (e.g. navigating back to a main-site route) removes the class again.
  useEffect(() => {
    const root = document.documentElement;
    root.classList.toggle("dark", theme !== "default");
    return () => {
      root.classList.remove("dark");
    };
  }, [theme]);

  return (
    <TeamsThemeContext.Provider value={{ theme, inTeams, isDark: theme !== "default" }}>
      {children}
    </TeamsThemeContext.Provider>
  );
}

export function useTeamsTheme(): TeamsThemeContextValue {
  return useContext(TeamsThemeContext);
}
