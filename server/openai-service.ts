import OpenAI from 'openai';
import pino from 'pino';
import { CATEGORY_RULES, getCategoryRule, type ListTarget } from '../shared/classification-rules';

// Use same logger configuration as server
const logger = pino({
  level: process.env.LOG_LEVEL || (process.env.NODE_ENV === 'production' ? 'info' : 'debug'),
  transport: process.env.NODE_ENV === 'development' ? {
    target: 'pino-pretty'
  } : undefined
});

// the newest OpenAI model is "gpt-4o" which was released May 13, 2024. do not change this unless explicitly requested by the user
// Uses Replit-managed AI Integrations (billed through Replit, no personal OpenAI key needed)
const openai = new OpenAI({
  apiKey: process.env.AI_INTEGRATIONS_OPENAI_API_KEY,
  baseURL: process.env.AI_INTEGRATIONS_OPENAI_BASE_URL,
});

export class OpenAIService {
  // Verify OpenAI API is working properly
  static async verifyAPIConnection(): Promise<{ working: boolean; error?: string; testTitle?: string }> {
    try {
      const testContent = "Test content for API verification";
      const testType = "Safety Ideas";
      
      const response = await openai.chat.completions.create({
        model: "gpt-4o", // the newest OpenAI model is "gpt-4o" which was released May 13, 2024. do not change this unless explicitly requested by the user
        messages: [
          {
            role: "system",
            content: `Generate a short, professional title (max 80 characters) for this ${testType.toLowerCase()} submission. The title should be clear and descriptive based on the content.`
          },
          {
            role: "user",
            content: testContent
          }
        ],
        max_tokens: 50,
        temperature: 0.3
      });

      const testTitle = response.choices[0].message.content?.trim() || '';
      
      return {
        working: true,
        testTitle
      };
    } catch (error: any) {
      return {
        working: false,
        error: error.message || 'Unknown error'
      };
    }
  }

  static async generateSmartTitle(content: string, type: string): Promise<string> {
    if (!content || content.trim().length === 0) {
      return `${type} Item`;
    }

    try {
      const response = await openai.chat.completions.create({
        model: "gpt-4o-mini", // Much faster than gpt-4o
        messages: [
          {
            role: "system",
            content: `You write short, plain titles for a glass and glazing company's health & safety and business idea submissions. Write like a tradesperson would talk — direct, practical, no corporate jargon. Use everyday language.

The title MUST describe what actually happened — the hazard, incident, problem, or idea itself. Never make the title just a location, a job/site name, an address, or a person's name. If a location or job is mentioned, it may be added for context but only after the incident (e.g. prefer "Sash dropped on hand at Airedale Courts" over "3b Airedale Courts"). If someone's name appears, describe what happened to or around them, don't use the name as the title.

Examples of good titles: "Broken step near loading bay", "Sash dropped on hand removing screws", "Near miss with forklift on site", "Need better lighting in workshop", "Glass off-cuts piling up — trip hazard", "Idea to speed up order processing". Bad titles to avoid: a bare location like "3b Airedale Courts", just a person's name, "Structural Integrity Compromise", "Operational Efficiency Enhancement", "Workplace Safety Incident Report". Keep it under 10 words.`
          },
          {
            role: "user",
            content: `Write a title for this ${type} submission: ${content}`
          }
        ],
        max_tokens: 50,
        temperature: 0.3
      });

      const title = response.choices[0].message.content?.trim();
      
      if (title && title.length > 0) {
        // Ensure title stays within SharePoint's 255-character limit
        if (title.length > 250) {
          const truncated = title.substring(0, 250);
          const lastSpaceIndex = truncated.lastIndexOf(' ');
          return lastSpaceIndex > 40 ? truncated.substring(0, lastSpaceIndex) + '...' : truncated + '...';
        }
        return title;
      }
      
      // Fallback to first sentence if AI fails
      return this.extractFirstSentence(content, type);
      
    } catch (error) {
      logger.error({ err: error }, 'OpenAI title generation failed');
      // Fallback to first sentence extraction
      return this.extractFirstSentence(content, type);
    }
  }

  private static extractFirstSentence(content: string, type: string): string {
    if (!content) return `${type} Item`;
    
    // Remove HTML and clean text
    const cleanContent = content
      .replace(/<[^>]*>/g, '')
      .replace(/&#58;/g, ':')
      .replace(/&#39;/g, "'")
      .replace(/&amp;/g, '&')
      .replace(/\s+/g, ' ')
      .trim();
    
    const sentences = cleanContent.split(/[.!?]+/).filter(s => s.trim().length > 0);
    
    if (sentences.length === 0) return `${type} Item`;
    
    let title = sentences[0].trim();
    if (title.length > 60) {
      const truncated = title.substring(0, 60);
      const lastSpaceIndex = truncated.lastIndexOf(' ');
      title = lastSpaceIndex > 40 ? truncated.substring(0, lastSpaceIndex) + '...' : truncated + '...';
    }
    
    return title;
  }

  static async bulkGenerateTitles(items: Array<{content: string, type: string}>): Promise<string[]> {
    const titles: string[] = [];
    
    // More cost-efficient: Process in smaller batches of 3, with longer delays
    for (let i = 0; i < items.length; i += 3) {
      const batch = items.slice(i, i + 3);
      
      try {
        // Use single batch request to reduce API calls
        const batchPromises = batch.map(item => this.generateSmartTitle(item.content, item.type));
        const batchTitles = await Promise.all(batchPromises);
        titles.push(...batchTitles);
        
        if (process.env.NODE_ENV === 'development') {
          console.log(`[DEBUG] AI generated ${batchTitles.length} titles (batch ${Math.floor(i/3) + 1})`);
        }
      } catch (error) {
        logger.error({ err: error }, 'AI batch failed, using fallback titles');
        // Graceful fallback for failed batch
        titles.push(...batch.map(item => this.extractFirstSentence(item.content, item.type)));
      }
      
      // Longer delay between batches to respect quota limits
      if (i + 3 < items.length) {
        await new Promise(resolve => setTimeout(resolve, 500));
      }
    }
    
    return titles;
  }

  static async enhanceNotes(content: string, itemType: string): Promise<string> {
    if (!content || content.trim().length === 0) {
      return content;
    }

    try {
      const response = await openai.chat.completions.create({
        model: "gpt-4o-mini", // Much faster than gpt-4o
        messages: [
          {
            role: "system",
            content: `You are a light-touch proofreader for workplace meeting notes. Fix ONLY spelling, grammar, and punctuation mistakes. Do NOT rephrase, reword, expand, summarise, formalise, or embellish anything. Keep every sentence's original wording, tone, and length as close to the input as possible. If the text is already correct, return it unchanged. Return only the corrected text with no commentary.`
          },
          {
            role: "user",
            content: content
          }
        ],
        max_tokens: 800, // Enough headroom so longer notes aren't cut off mid-sentence
        temperature: 0.1  // Lower temperature for faster, more consistent responses
      });

      const enhancedContent = response.choices[0].message.content?.trim();
      
      if (enhancedContent && enhancedContent.length > 0) {
        return enhancedContent;
      }
      
      // Return original content if AI enhancement fails
      return content;
      
    } catch (error) {
      logger.error({ err: error }, 'OpenAI note enhancement failed');
      // Return original content on error
      return content;
    }
  }

  static async classifySubmission(text: string): Promise<{
    category: string;
    listTarget: 'near-miss' | 'safety-ideas' | 'business-ideas';
    confidence: number;
    reasoning: string;
    followUpQuestions: string[];
    orderItem: string;
  }> {
    // Build the prompt from the shared classification rules so the categories,
    // routing, and follow-up questions can never drift from the client.
    const categoryLines = CATEGORY_RULES
      .map((rule) => `- "${rule.name}" → ${rule.listTarget} list (${rule.aiDescription})`)
      .join('\n');

    const followUpLines = CATEGORY_RULES
      .filter((rule) => rule.followUpQuestions.length > 0)
      .map((rule) => `- ${rule.name}: ${rule.followUpQuestions.map((q) => `"${q}"`).join(', ')}`)
      .join('\n');

    const categoryUnion = CATEGORY_RULES.map((rule) => `"${rule.name}"`).join(' | ');

    const systemPrompt = `You are a health & safety and business improvement classifier for Cranfield Glass Christchurch, a glass and glazing company in New Zealand.

Given a staff member's free-text description, classify it into exactly one of these categories and provide follow-up questions if confidence is low.

Categories and their SharePoint list targets:
${categoryLines}

When the submission mentions bringing in, buying, receiving, trialling, or starting to use a chemical, hazardous substance, or product (e.g. acetone, solvents, thinners, adhesives, silicones, sealants, paints, cleaners, aerosols, gases, lubricants), classify it as "Chemical Register" so it can be reviewed and added to the chemical / hazardous substances register. Examples: "Purchased a drum of acetone", "We're trialling a new adhesive from a supplier", "Got a new solvent-based cleaner in the workshop" → "Chemical Register". Note: a pure ordering/restock request for supplies that are NOT chemicals is still "Supply Request".

When the submission is something to raise or discuss at the next H&S meeting (an agenda item), pick the agenda category that matches the TOPIC, not the format:
- "Near Miss Meeting Agenda Item" when it's about a specific incident, accident, or near-miss event/investigation.
- "Safety Meeting Agenda Item" when it's about a safety topic such as the safety/hazard register, PPE, hazards, or safety processes. Example: "Talk about the updates to our safety hazard register, how it works and continuous improvement" → "Safety Meeting Agenda Item".
- "Business Meeting Agenda Item" only for business/operational topics such as rosters, pricing, customers, or scheduling.

For each category, these are the key follow-up questions when confidence < 0.8:
${followUpLines}

When the category is "Supply Request", also extract a short, plain item name suitable for a team ordering list — just the thing(s) to buy, in everyday words, no full sentence. Examples: "Safety gloves", "Squeegee rubbers", "Masking tape". Leave it as an empty string for every other category.

Respond with valid JSON only:
{
  "category": ${categoryUnion},
  "listTarget": "near-miss" | "safety-ideas" | "business-ideas",
  "confidence": 0.0 to 1.0,
  "reasoning": "brief explanation",
  "followUpQuestions": ["question 1", "question 2"],
  "orderItem": "short item name when category is Supply Request, otherwise empty string"
}`;

    try {
      const response = await openai.chat.completions.create({
        model: 'gpt-4o-mini', // Much faster than gpt-4o; classification is a simple task
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: text.trim() }
        ],
        max_tokens: 300,
        temperature: 0.1,
        response_format: { type: 'json_object' }
      });

      const content = response.choices[0].message.content?.trim();
      if (!content) throw new Error('Empty response from AI');

      const parsed = JSON.parse(content);
      if (!parsed.category || !parsed.listTarget || typeof parsed.confidence !== 'number') {
        throw new Error('Invalid AI response structure');
      }

      const followUpQuestions = parsed.confidence < 0.8 ? (parsed.followUpQuestions || []) : [];

      // The category's rule is authoritative for routing — derive the list
      // target from it so the AI can't file a category into the wrong list.
      // Fall back to the AI's listTarget (then business-ideas) for an unknown
      // category.
      const rule = getCategoryRule(parsed.category);
      const validTargets: ListTarget[] = ['near-miss', 'safety-ideas', 'business-ideas'];
      const listTarget: ListTarget = rule
        ? rule.listTarget
        : validTargets.includes(parsed.listTarget) ? parsed.listTarget : 'business-ideas';

      return {
        category: parsed.category,
        listTarget,
        confidence: parsed.confidence,
        reasoning: parsed.reasoning || '',
        followUpQuestions,
        orderItem: parsed.category === 'Supply Request' && typeof parsed.orderItem === 'string' ? parsed.orderItem.trim() : ''
      };
    } catch (error) {
      logger.error({ err: error }, 'OpenAI classification failed');
      return {
        category: 'Other',
        listTarget: 'business-ideas',
        confidence: 0.5,
        reasoning: 'Classification failed, defaulting to Other',
        followUpQuestions: ['Can you describe what happened in more detail?', 'Which area of the business does this relate to?'],
        orderItem: ''
      };
    }
  }

  static async generateSuggestions(content: string, type: string): Promise<string[]> {
    if (!content || content.trim().length < 10) {
      return [];
    }

    try {
      const response = await openai.chat.completions.create({
        model: "gpt-4o",
        messages: [
          {
            role: "system",
            content: `You are an AI assistant helping staff at Cranfield Glass Christchurch write professional ${type} submissions. Based on the partial content the user has typed, provide 3 helpful suggestions to complete or improve their idea.

Guidelines for suggestions:
- Keep suggestions concise (max 50 characters each)
- Make suggestions relevant to glazing/construction industry
- Focus on safety, efficiency, and practical improvements
- Suggestions should complete thoughts or add important details
- For Safety Ideas: suggest safety equipment, procedures, or hazard prevention
- For Business Ideas: suggest efficiency improvements, cost savings, or process enhancements
- For Near Miss: suggest specific details, prevention measures, or investigation points

Return exactly 3 suggestions as a JSON array of strings. Example: ["suggestion 1", "suggestion 2", "suggestion 3"]

Return ONLY the JSON array, no other text.`
          },
          {
            role: "user",
            content: `Current ${type} content: "${content}"`
          }
        ],
        max_tokens: 200,
        temperature: 0.7
      });

      const suggestionsText = response.choices[0].message.content?.trim();
      
      if (suggestionsText) {
        try {
          const suggestions = JSON.parse(suggestionsText);
          if (Array.isArray(suggestions) && suggestions.length > 0) {
            return suggestions.slice(0, 3); // Ensure max 3 suggestions
          }
        } catch (parseError) {
          logger.error({ err: parseError }, 'Failed to parse AI suggestions JSON');
        }
      }
      
      // Fallback suggestions if AI fails
      return [
        "Add specific location details",
        "Include potential safety impact", 
        "Suggest implementation steps"
      ];
      
    } catch (error) {
      logger.error({ err: error }, 'OpenAI suggestions failed');
      // Return helpful fallback suggestions
      if (type === 'Safety Ideas') {
        return [
          "Include safety equipment needed",
          "Mention specific hazards addressed",
          "Add implementation timeline"
        ];
      } else {
        return [
          "Include cost/time savings",
          "Add implementation details",
          "Mention staff training needed"
        ];
      }
    }
  }

  static async analyzeJobListingSection(sectionText: string, sectionType: string): Promise<{
    overallRating: number;
    analysis: string;
    strengths: string[];
    improvements: string[];
    newZealandTone: number;
    mobileFriendliness: number;
    candidateAppeal: number;
  }> {
    try {
      const response = await openai.chat.completions.create({
        model: "gpt-4o",
        messages: [
          {
            role: "system",
            content: `You are an expert recruitment consultant and job listing analyzer specializing in TradeMe New Zealand job postings for building industry and service-based businesses.

Analyze the provided job listing section and provide a comprehensive rating and analysis. Consider:

1. **New Zealand Conversational Tone** (0-10): How well does it use authentic Kiwi language and avoid corporate jargon?
2. **Mobile Friendliness** (0-10): How scannable and readable is it on mobile devices?
3. **Candidate Appeal** (0-10): How compelling is it for building/service industry candidates?
4. **Overall Effectiveness** (0-10): Combined rating for TradeMe success

Focus on:
- Authentic NZ conversational style vs corporate speak
- Clarity and directness without being intimidating
- Appeal to practical, hands-on workers
- Mobile scanning and readability
- Immediate start urgency without being pushy
- Focus on personality over complex requirements

Return your analysis as a JSON object with this exact structure:
{
  "overallRating": number,
  "analysis": "detailed paragraph analysis",
  "strengths": ["strength 1", "strength 2", "strength 3"],
  "improvements": ["improvement 1", "improvement 2", "improvement 3"],
  "newZealandTone": number,
  "mobileFriendliness": number,
  "candidateAppeal": number
}

Be thorough but concise. Focus on actionable insights.`
          },
          {
            role: "user",
            content: `Section Type: ${sectionType}

Section Text: "${sectionText}"`
          }
        ],
        max_tokens: 800,
        temperature: 0.3
      });

      const analysisText = response.choices[0].message.content?.trim();
      
      if (analysisText) {
        try {
          // Clean up markdown code blocks if present
          const cleanedText = analysisText
            .replace(/```json\n?/g, '')
            .replace(/```\n?/g, '')
            .trim();
          
          const analysis = JSON.parse(cleanedText);
          return analysis;
        } catch (parseError) {
          logger.error({ err: parseError }, 'Failed to parse AI analysis JSON');
        }
      }
      
      // Fallback analysis if AI fails
      return {
        overallRating: 6,
        analysis: "Unable to provide detailed analysis due to technical issues. The section appears functional but may benefit from review.",
        strengths: ["Clear communication", "Practical approach", "Professional tone"],
        improvements: ["Consider more conversational language", "Add mobile optimization", "Enhance candidate appeal"],
        newZealandTone: 5,
        mobileFriendliness: 6,
        candidateAppeal: 6
      };
      
    } catch (error) {
      logger.error({ err: error }, 'OpenAI job listing analysis failed');
      return {
        overallRating: 5,
        analysis: "Analysis unavailable due to technical issues. Manual review recommended.",
        strengths: ["Section exists", "Conveys basic information"],
        improvements: ["Review for clarity", "Optimize for mobile", "Enhance tone"],
        newZealandTone: 5,
        mobileFriendliness: 5,
        candidateAppeal: 5
      };
    }
  }

  static async generateJobListing(previousListing: string, requirements: string): Promise<string> {
    try {
      const response = await openai.chat.completions.create({
        model: "gpt-4o",
        messages: [
          {
            role: "system",
            content: `You are an expert recruitment consultant for TradeMe New Zealand, specializing in creating standout job listings for building industry and service-based businesses. 

Create an optimized TradeMe job listing that uses proper Markdown formatting. The listing should:
- Use compelling, immediate-start language for Q4 urgency
- Target candidates with building industry or service business background
- Emphasize practical skills over complex tech knowledge
- Use TradeMe's Markdown syntax: **bold**, *italic*, # headings, - bullet points
- Be mobile-optimized and scannable (3-minute read test)
- Highlight what makes this role unique
- Include clear structure with headings and bullet points
- Appeal to quick learners who want hands-on work

TradeMe best practices:
- Job title should be clear and searchable
- Use bullet points for key requirements and benefits
- Bold important information like salary/benefits
- Keep sections concise but comprehensive
- Include urgency without being pushy
- Focus on immediate opportunities and growth potential

Return ONLY the formatted job listing content, ready to paste into TradeMe.`
          },
          {
            role: "user",
            content: `Previous listing: ${previousListing}

New requirements: ${requirements}`
          }
        ],
        max_tokens: 1500,
        temperature: 0.7
      });

      const jobListing = response.choices[0].message.content?.trim();
      
      if (jobListing && jobListing.length > 0) {
        return jobListing;
      }
      
      // Fallback if AI fails
      return "# Office Coordinator - Immediate Start\n\n**Cranfield Glass Christchurch** is looking for a dynamic team member to join our growing business.\n\nContact us for more details.";
      
    } catch (error) {
      logger.error({ err: error }, 'OpenAI job listing generation failed');
      return "# Office Coordinator - Immediate Start\n\n**Cranfield Glass Christchurch** is looking for a dynamic team member to join our growing business.\n\nContact us for more details.";
    }
  }
}