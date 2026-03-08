declare global {
  interface Window {
    mermaid: any;
  }
}

export function initializeMermaid() {
  if (window.mermaid) {
    window.mermaid.initialize({
      startOnLoad: true,
      theme: 'base',
      themeVariables: {
        primaryColor: '#0078D4',
        primaryTextColor: '#323130',
        primaryBorderColor: '#106EBE',
        lineColor: '#605E5C',
        secondaryColor: '#F3F2F1',
        tertiaryColor: '#00BCF2'
      }
    });
  }
}
