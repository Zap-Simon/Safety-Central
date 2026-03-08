/**
 * Predictive Text Engine for Cranfield Glass
 * Learns from user input patterns and provides intelligent suggestions
 */

import Fuse from 'fuse.js';
import { rankItem } from '@tanstack/match-sorter';

interface TextPattern {
  id: string;
  text: string;
  frequency: number;
  context: 'safety' | 'business' | 'meeting_notes' | 'general';
  lastUsed: Date;
}

interface PredictiveTextEngine {
  addPattern: (text: string, context: string) => void;
  getSuggestions: (input: string, context: string, limit?: number) => string[];
  getInlineCompletion: (input: string, context: string) => string | null;
  learnFromInput: (input: string, context: string) => void;
  getStorageKey: () => string;
}

class CranfieldPredictiveText implements PredictiveTextEngine {
  private patterns: TextPattern[] = [];
  private fuse: Fuse<TextPattern>;
  private storageKey = 'cranfield-predictive-text-patterns';

  // Cranfield Glass specific terms and phrases for continuous completion
  private glazingTerms = [
    // Complete phrases that help with continuous typing
    "safety equipment needs to be replaced", "protective equipment must be worn", "glass handling requires proper training",
    "installation procedures should be followed", "tempered glass specifications need review", "laminated glass quality control checks",
    "workplace safety is our top priority", "PPE requirements include safety glasses", "risk assessment shows potential hazards",
    "incident reporting must be completed", "near miss investigation reveals issues", "safety protocols require immediate action",
    "training requirements for all staff", "compliance with health and safety", "glass specifications meet customer requirements",
    "cost reduction through process improvement", "efficiency improvement in workflow", "workflow optimization can reduce costs",
    "team coordination needs better communication", "delivery scheduling requires planning", "material procurement should be optimized",
    "meeting discussion about safety issues", "action items need follow-up", "follow-up required for implementation",
    "implementation of new safety procedures", "review process shows good results", "approval needed from management",
    "budget considerations for new equipment", "timeline needs to be established", "staff feedback indicates improvements needed",
    "process improvement through better training", "equipment maintenance is scheduled", "van inspections must be completed weekly",
    "tool safety checks are mandatory", "workspace organization improves efficiency", "quality control measures are essential",
    // Common single words and shorter phrases for word completion
    "safety", "equipment", "protective", "glass", "handling", "installation", "procedures",
    "tempered", "laminated", "quality", "control", "workplace", "requirements", "assessment",
    "reporting", "investigation", "protocols", "training", "compliance", "specifications",
    "customer", "improvement", "workflow", "optimization", "coordination", "scheduling",
    "procurement", "discussion", "implementation", "review", "approval", "budget",
    "timeline", "feedback", "maintenance", "inspections", "organization", "measures"
  ];

  constructor() {
    this.loadPatterns();
    this.initializeFuse();
    this.seedInitialPatterns();
  }

  private initializeFuse() {
    this.fuse = new Fuse(this.patterns, {
      keys: ['text'],
      threshold: 0.4,
      includeScore: true,
      minMatchCharLength: 3
    });
  }

  private loadPatterns() {
    try {
      const stored = localStorage.getItem(this.storageKey);
      if (stored) {
        const parsed = JSON.parse(stored);
        this.patterns = parsed.map((p: any) => ({
          ...p,
          lastUsed: new Date(p.lastUsed)
        }));
      }
    } catch (error) {
      console.warn('Failed to load predictive text patterns:', error);
      this.patterns = [];
    }
  }

  private savePatterns() {
    try {
      localStorage.setItem(this.storageKey, JSON.stringify(this.patterns));
    } catch (error) {
      console.warn('Failed to save predictive text patterns:', error);
    }
  }

  private seedInitialPatterns() {
    // Only seed if we have no patterns stored
    if (this.patterns.length === 0) {
      this.glazingTerms.forEach((term, index) => {
        this.patterns.push({
          id: `seed-${index}`,
          text: term,
          frequency: 1,
          context: 'general',
          lastUsed: new Date()
        });
      });
      this.savePatterns();
      this.initializeFuse();
    }
  }

  addPattern(text: string, context: string): void {
    const cleanText = text.trim().toLowerCase();
    if (cleanText.length < 3) return;

    const existingPattern = this.patterns.find(p => 
      p.text.toLowerCase() === cleanText && p.context === context
    );

    if (existingPattern) {
      existingPattern.frequency++;
      existingPattern.lastUsed = new Date();
    } else {
      this.patterns.push({
        id: `pattern-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
        text: cleanText,
        frequency: 1,
        context: context as any,
        lastUsed: new Date()
      });
    }

    this.savePatterns();
    this.initializeFuse();
  }

  /**
   * Get inline completion for current word being typed
   * Returns the completion text that should appear inline
   */
  getInlineCompletion(input: string, context: string): string | null {
    if (input.length < 2) return null;

    const trimmedInput = input.trim().toLowerCase();
    const words = trimmedInput.split(/\s+/);
    const currentWord = words[words.length - 1];
    
    if (currentWord.length < 2) return null;

    // Strategy 1: Find patterns where the current word is the beginning of a known phrase
    const phraseMatches = this.patterns
      .filter(pattern => {
        const patternLower = pattern.text.toLowerCase();
        const patternWords = patternLower.split(/\s+/);
        
        // Check if any word in the pattern starts with current word
        return patternWords.some(word => 
          word.startsWith(currentWord) && 
          word.length > currentWord.length
        ) && (pattern.context === context || pattern.context === 'general');
      })
      .map(pattern => {
        const patternLower = pattern.text.toLowerCase();
        const patternWords = patternLower.split(/\s+/);
        
        // Find the word that matches and return completion
        for (let i = 0; i < patternWords.length; i++) {
          if (patternWords[i].startsWith(currentWord) && patternWords[i].length > currentWord.length) {
            // Return completion for current word + any following words from the phrase
            const completion = patternWords[i].substring(currentWord.length);
            const remainingWords = patternWords.slice(i + 1);
            return {
              completion: remainingWords.length > 0 ? completion + ' ' + remainingWords.join(' ') : completion,
              pattern,
              score: this.calculateScore(pattern, 0, currentWord)
            };
          }
        }
        return null;
      })
      .filter(Boolean)
      .sort((a, b) => b!.score - a!.score);

    if (phraseMatches.length > 0) {
      return phraseMatches[0]!.completion;
    }

    // Strategy 2: Find patterns that contain the current context and suggest next words
    const contextMatches = this.patterns
      .filter(pattern => {
        const patternLower = pattern.text.toLowerCase();
        // Check if the input so far appears in this pattern
        return patternLower.includes(trimmedInput) && 
               patternLower.length > trimmedInput.length &&
               (pattern.context === context || pattern.context === 'general');
      })
      .map(pattern => {
        const patternLower = pattern.text.toLowerCase();
        const startIndex = patternLower.indexOf(trimmedInput);
        if (startIndex !== -1) {
          const completion = pattern.text.substring(startIndex + trimmedInput.length);
          return {
            completion: completion.trim(),
            pattern,
            score: this.calculateScore(pattern, 0, trimmedInput)
          };
        }
        return null;
      })
      .filter(Boolean)
      .sort((a, b) => b!.score - a!.score);

    if (contextMatches.length > 0) {
      const bestMatch = contextMatches[0]!;
      // Limit completion to next few words to avoid overwhelming
      const completionWords = bestMatch.completion.split(/\s+/).slice(0, 3);
      return completionWords.join(' ');
    }

    return null;
  }

  getSuggestions(input: string, context: string, limit: number = 5): string[] {
    if (input.length < 2) return [];

    const cleanInput = input.trim().toLowerCase();
    
    // Get fuzzy search results
    const fuseResults = this.fuse.search(cleanInput);
    
    // Filter by context preference and score results
    const contextMatches = fuseResults
      .filter(result => 
        result.item.context === context || 
        result.item.context === 'general' ||
        (context === 'meeting_notes' && ['safety', 'business'].includes(result.item.context))
      )
      .map(result => ({
        text: result.item.text,
        score: this.calculateScore(result.item, result.score || 0, cleanInput),
        item: result.item
      }));

    // Also check for partial matches in our patterns
    const partialMatches = this.patterns
      .filter(pattern => 
        pattern.text.includes(cleanInput) &&
        (pattern.context === context || pattern.context === 'general')
      )
      .map(pattern => ({
        text: pattern.text,
        score: this.calculateScore(pattern, 0.2, cleanInput),
        item: pattern
      }));

    // Combine and deduplicate
    const allMatches = [...contextMatches, ...partialMatches];
    const uniqueMatches = allMatches.reduce((acc, match) => {
      const existing = acc.find(m => m.text === match.text);
      if (!existing || existing.score < match.score) {
        acc = acc.filter(m => m.text !== match.text);
        acc.push(match);
      }
      return acc;
    }, [] as typeof allMatches);

    // Sort by score (higher is better) and return top results
    return uniqueMatches
      .sort((a, b) => b.score - a.score)
      .slice(0, limit)
      .map(match => this.formatSuggestion(match.text, input));
  }

  private calculateScore(pattern: TextPattern, fuseScore: number, input: string): number {
    let score = 1 - fuseScore; // Invert fuse score (lower fuse score = better match)
    
    // Boost for frequency
    score += Math.log(pattern.frequency + 1) * 0.1;
    
    // Boost for recent usage
    const daysSinceUsed = (Date.now() - pattern.lastUsed.getTime()) / (1000 * 60 * 60 * 24);
    score += Math.max(0, (30 - daysSinceUsed) / 30) * 0.2;
    
    // Boost for exact prefix matches
    if (pattern.text.toLowerCase().startsWith(input.toLowerCase())) {
      score += 0.3;
    }
    
    return score;
  }

  private formatSuggestion(suggestion: string, input: string): string {
    // Capitalize first letter and maintain original case for the rest
    return suggestion.charAt(0).toUpperCase() + suggestion.slice(1);
  }

  learnFromInput(input: string, context: string): void {
    const sentences = input.split(/[.!?]+/).filter(s => s.trim().length > 0);
    
    sentences.forEach(sentence => {
      const cleanSentence = sentence.trim();
      if (cleanSentence.length > 5) {
        this.addPattern(cleanSentence, context);
      }
      
      // Also learn common phrases (3-5 words)
      const words = cleanSentence.split(/\s+/);
      for (let i = 0; i <= words.length - 3; i++) {
        const phrase = words.slice(i, i + Math.min(5, words.length - i)).join(' ');
        if (phrase.length > 10 && phrase.length < 100) {
          this.addPattern(phrase, context);
        }
      }
    });
  }

  getStorageKey(): string {
    return this.storageKey;
  }

  // Debug method to inspect stored patterns
  getPatterns(): TextPattern[] {
    return this.patterns;
  }

  // Clear all learned patterns (useful for reset)
  clearPatterns(): void {
    this.patterns = [];
    localStorage.removeItem(this.storageKey);
    this.seedInitialPatterns();
  }
}

// Export singleton instance
export const predictiveText = new CranfieldPredictiveText();

// Export types for use in components
export type { TextPattern, PredictiveTextEngine };