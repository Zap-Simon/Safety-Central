import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, BorderStyle, WidthType, HeadingLevel, AlignmentType } from "docx";
import { buildActionRequiredLines, isEmptyActionPlaceholder, getDisplayItemStatus, buildReadyToCloseActions, type ActionLine, type ReadyToCloseAction } from "./meeting-export-shared";

interface MeetingItem {
  id: string;
  title: string;
  description: string;
  type: string;
  status: string;
  meetingDate: string;
  meetingNotes: string;
  submittedBy: string;
  submittedDate: string;
  ideaType?: string;
  assignedTo?: string;
  secondaryDescription?: string;
  actionAssignedTo?: string;
  actionStatus?: string;
  actionPriority?: string;
  actionStartDate?: string;
  actionDueDate?: string;
  actionNotes?: string;
}

interface AttendanceData {
  [meetingDate: string]: string[];
}

interface TemplateData {
  companyName: string;
  meetingTitle: string;
  meetingDate: string;
  currentDate: string;
  currentTime: string;
  items: Array<{
    title: string;
    description: string;
    type: string;
    status: string;
    statusColor: string;
    submittedBy: string;
    submittedDate: string;
    ideaType?: string;
    meetingNotes: string;
    secondaryDescription?: string;
    actionLines: ActionLine[];
  }>;
  managementTeam: Array<{
    name: string;
    role: string;
    present: boolean;
  }>;
  glaziers: Array<{
    name: string;
    present: boolean;
  }>;
  totalItems: number;
  businessIdeasCount: number;
  safetyIdeasCount: number;
  nearMissCount: number;
  readyToCloseActions: ReadyToCloseAction[];
}

export class AdvancedWordTemplateEngine {
  
  /**
   * Generate a professional Word document using advanced templating
   */
  static async generateMeetingDocument(data: {
    meetingData: MeetingItem[];
    selectedMeeting: string;
    selectedType: string;
    meetingAttendance: AttendanceData;
  }): Promise<Buffer> {
    
    // Create template data
    const templateData = this.prepareTemplateData(data);
    
    // Generate document using docx library with advanced styling
    return this.createAdvancedDocument(templateData);
  }

  /**
   * Prepare structured data for templating
   */
  private static prepareTemplateData(data: {
    meetingData: MeetingItem[];
    selectedMeeting: string;
    selectedType: string;
    meetingAttendance: AttendanceData;
  }): TemplateData {
    
    // Filter and deduplicate data
    let filteredData = data.meetingData;
    if (data.selectedMeeting && data.selectedMeeting !== 'all') {
      const selectedDate = new Date(data.selectedMeeting);
      filteredData = data.meetingData.filter(item => {
        if (!item.meetingDate || item.meetingDate === 'unknown-meeting') return false;
        const itemDate = new Date(item.meetingDate);
        // Compare just the date parts, ignoring time
        return itemDate.toDateString() === selectedDate.toDateString();
      });
    }
    
    const uniqueItems = filteredData.filter((item, index, self) => 
      index === self.findIndex(t => t.id === item.id)
    );

    // Prepare meeting info with proper date correction
    const meetingDate = data.selectedMeeting === 'all' ? 'Multiple Meetings' : 
      AdvancedWordTemplateEngine.formatMeetingDate(data.selectedMeeting);

    // Prepare attendance data
    const managementTeam = [
      { name: 'Hoani Hunt', role: 'Company Director' },
      { name: 'Simon Hubbard', role: 'Health & Safety Coordinator' },
      { name: 'James Waites', role: 'Glazing Supervisor' },
      { name: 'Emma White', role: 'Administrator' }
    ];

    const glaziers = ['Kevin Young', 'Ryan Newman', 'Isaac Ensor', 'Struan O\'Donnell', 'Sam Chang'];
    
    // Find attendance data by matching date (ignoring time)
    let attendance: string[] = [];
    if (data.meetingAttendance && data.selectedMeeting !== 'all') {
      const selectedDate = new Date(data.selectedMeeting);
      
      // Look for attendance data with matching date
      for (const [meetingDateKey, attendees] of Object.entries(data.meetingAttendance)) {
        const meetingDate = new Date(meetingDateKey);
        if (meetingDate.toDateString() === selectedDate.toDateString()) {
          attendance = attendees;
          break;
        }
      }
      
      // If no attendance found, default to all present
      if (attendance.length === 0) {
        attendance = [
          ...managementTeam.map(m => m.name),
          ...glaziers
        ];
      }
    }

    // Process items with enhanced data
    const processedItems = uniqueItems.map(item => ({
      title: item.title || 'Untitled Item',
      description: item.description || '',
      type: item.type,
      status: getDisplayItemStatus(item),
      statusColor: this.getStatusColor(item.status),
      submittedBy: item.submittedBy || 'Unknown',
      submittedDate: new Date(item.submittedDate).toLocaleDateString('en-GB'),
      ideaType: item.ideaType,
      meetingNotes: item.meetingNotes || 'No notes recorded',
      secondaryDescription: item.secondaryDescription || '',
      actionLines: buildActionRequiredLines(item)
    }));

    // Count items by type
    const businessIdeasCount = processedItems.filter(item => item.type === 'Business Ideas').length;
    const safetyIdeasCount = processedItems.filter(item => item.type === 'Safety Ideas').length;
    const nearMissCount = processedItems.filter(item => item.type === 'Near Miss').length;

    return {
      companyName: 'CRANFIELD GLASS CHRISTCHURCH',
      meetingTitle: 'Health & Safety Meeting Minutes',
      meetingDate,
      currentDate: new Date().toLocaleDateString('en-GB'),
      currentTime: new Date().toLocaleTimeString('en-GB'),
      items: processedItems,
      managementTeam: managementTeam.map(person => ({
        name: person.name,
        role: person.role,
        present: attendance.includes(person.name)
      })),
      glaziers: glaziers.map(name => ({
        name,
        present: attendance.includes(name)
      })),
      totalItems: processedItems.length,
      businessIdeasCount,
      safetyIdeasCount,
      nearMissCount,
      // Ready-to-Close actions come from the FULL backlog, not just the selected
      // meeting, so every export lists all outstanding actions awaiting sign-off.
      readyToCloseActions: buildReadyToCloseActions(data.meetingData)
    };
  }

  /**
   * Create a clean, presentable Word document
   */
  private static async createAdvancedDocument(data: TemplateData): Promise<Buffer> {
    const sections: any[] = [];

    // Clean Company Header
    sections.push(
      new Paragraph({
        children: [
          new TextRun({
            text: data.companyName,
            bold: true,
            size: 48, // 24pt
            color: "1F2937"
          })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 }
      }),
      
      new Paragraph({
        children: [
          new TextRun({
            text: "Health & Safety Meeting",
            size: 28, // 14pt
            color: "6B7280"
          })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 }
      }),
      
      new Paragraph({
        children: [
          new TextRun({
            text: data.meetingTitle,
            bold: true,
            size: 32, // 16pt
            color: "1F2937"
          })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 }
      }),
      
      new Paragraph({
        children: [
          new TextRun({
            text: `Generated: ${data.currentDate} at ${data.currentTime}`,
            size: 22, // 11pt
            color: "6B7280",
            italics: true
          })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 300 }
      }),
      
      // Meeting Information Table
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 2, color: "E5E7EB" },
          bottom: { style: BorderStyle.SINGLE, size: 2, color: "E5E7EB" },
          left: { style: BorderStyle.SINGLE, size: 2, color: "E5E7EB" },
          right: { style: BorderStyle.SINGLE, size: 2, color: "E5E7EB" },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "F3F4F6" },
          insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "F3F4F6" }
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ text: "Meeting Date:", bold: true, size: 24, // 12pt
                    color: "374151" })]
                })],
                width: { size: 30, type: WidthType.PERCENTAGE },
                shading: { fill: "F9FAFB" },
                margins: { top: 100, bottom: 100, left: 150, right: 100 }
              }),
              new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ text: data.meetingDate, size: 24, // 12pt
                    color: "4B5563" })]
                })],
                margins: { top: 100, bottom: 100, left: 150, right: 100 }
              })
            ]
          }),
          new TableRow({
            children: [
              new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ text: "Total Items:", bold: true, size: 24, // 12pt
                    color: "374151" })]
                })],
                shading: { fill: "F9FAFB" },
                margins: { top: 100, bottom: 100, left: 150, right: 100 }
              }),
              new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ text: `${data.totalItems} (Business: ${data.businessIdeasCount}, Safety: ${data.safetyIdeasCount}, Near Miss: ${data.nearMissCount})`, size: 24, // 12pt
                    color: "4B5563" })]
                })],
                margins: { top: 100, bottom: 100, left: 150, right: 100 }
              })
            ]
          })
        ]
      })
    );

    // Meeting Items Section
    if (data.items.length > 0) {
      sections.push(
        new Paragraph({
          children: [
            new TextRun({
              text: "Meeting Items",
              bold: true,
              size: 32, // 16pt
              color: "1F2937"
            })
          ],
          spacing: { before: 300, after: 200 }
        }),
        
        new Paragraph({
          children: [
            new TextRun({
              text: `${data.totalItems} items (Business: ${data.businessIdeasCount}, Safety: ${data.safetyIdeasCount}, Near Miss: ${data.nearMissCount})`,
              size: 22, // 11pt
              color: "6B7280"
            })
          ],
          spacing: { after: 200 }
        })
      );

      // Clean items table
      const itemsTable = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 2, color: "D1D5DB" },
          bottom: { style: BorderStyle.SINGLE, size: 2, color: "D1D5DB" },
          left: { style: BorderStyle.SINGLE, size: 2, color: "D1D5DB" },
          right: { style: BorderStyle.SINGLE, size: 2, color: "D1D5DB" },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "E5E7EB" },
          insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "E5E7EB" }
        },
        rows: [
          // Items header row
          new TableRow({
            children: [
              new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ 
                    text: "#", 
                    bold: true, 
                    size: 24 // 12pt
                  })],
                  alignment: AlignmentType.CENTER
                })],
                width: { size: 5, type: WidthType.PERCENTAGE },
                shading: { fill: "F9FAFB" },
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
              }),
              new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ 
                    text: "Title & Description", 
                    bold: true, 
                    size: 24 // 12pt
                  })],
                  alignment: AlignmentType.LEFT
                })],
                width: { size: 45, type: WidthType.PERCENTAGE },
                shading: { fill: "F9FAFB" },
                margins: { top: 80, bottom: 80, left: 100, right: 100 }
              }),
              new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ 
                    text: "Type", 
                    bold: true, 
                    size: 24 // 12pt
                  })],
                  alignment: AlignmentType.CENTER
                })],
                width: { size: 15, type: WidthType.PERCENTAGE },
                shading: { fill: "F9FAFB" },
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
              }),
              new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ 
                    text: "Status", 
                    bold: true, 
                    size: 24 // 12pt
                  })],
                  alignment: AlignmentType.CENTER
                })],
                width: { size: 15, type: WidthType.PERCENTAGE },
                shading: { fill: "F9FAFB" },
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
              }),
              new TableCell({
                children: [new Paragraph({
                  children: [new TextRun({ 
                    text: "Submitted By", 
                    bold: true, 
                    size: 24 // 12pt
                  })],
                  alignment: AlignmentType.CENTER
                })],
                width: { size: 20, type: WidthType.PERCENTAGE },
                shading: { fill: "F9FAFB" },
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
              })
            ]
          }),
          // Items data rows
          ...data.items.map((item, index) => {
            return new TableRow({
              children: [
                // Item number
                new TableCell({
                  children: [new Paragraph({
                    children: [new TextRun({
                      text: `${index + 1}`,
                      size: 22 // 11pt
                    })],
                    alignment: AlignmentType.CENTER
                  })],
                  shading: { fill: index % 2 === 0 ? "F8FAFC" : "FFFFFF" },
                  margins: { top: 80, bottom: 80, left: 50, right: 50 }
                }),
                // Title and description with meeting notes
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [new TextRun({
                        text: item.title,
                        bold: true,
                        size: 22, // 11pt
                        color: "1F2937"
                      })],
                      spacing: { after: 80 }
                    }),
                    new Paragraph({
                      children: [new TextRun({
                        text: item.description,
                        size: 20, // 10pt
                        color: "4B5563"
                      })],
                      spacing: { after: 0 }
                    }),
                    // "How it happened" follow-up belongs to the agenda submission
                    ...(item.secondaryDescription ? [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "How it happened: ",
                            bold: true,
                            size: 20, // 10pt
                            color: "374151"
                          }),
                          new TextRun({
                            text: item.secondaryDescription,
                            size: 20, // 10pt
                            color: "4B5563"
                          })
                        ],
                        spacing: { before: 60 }
                      })
                    ] : []),
                    // Add meeting notes if present
                    ...(item.meetingNotes && item.meetingNotes !== 'No notes recorded' ? [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "Meeting Discussion: ",
                            bold: true,
                            size: 20, // 10pt
                            color: "374151"
                          }),
                          new TextRun({
                            text: item.meetingNotes,
                            size: 20, // 10pt
                            color: "4B5563",
                            italics: true
                          })
                        ],
                        spacing: { before: 80 }
                      })
                    ] : []),
                    // Action Required — shared, compliant action lines (omit empty placeholder)
                    ...(!isEmptyActionPlaceholder(item.actionLines) ? [
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "Action Required:",
                            bold: true,
                            size: 20, // 10pt
                            color: "374151"
                          })
                        ],
                        spacing: { before: 80 }
                      }),
                      ...item.actionLines.map(line => new Paragraph({
                        children: [
                          ...(line.label ? [new TextRun({
                            text: `${line.label}: `,
                            bold: true,
                            size: 18, // 9pt
                            color: "374151"
                          })] : []),
                          new TextRun({
                            text: line.value,
                            size: 18, // 9pt
                            color: "4B5563"
                          })
                        ]
                      }))
                    ] : [])
                  ],
                  shading: { fill: index % 2 === 0 ? "F8FAFC" : "FFFFFF" },
                  margins: { top: 80, bottom: 80, left: 100, right: 100 }
                }),
                // Type
                new TableCell({
                  children: [new Paragraph({
                    children: [new TextRun({
                      text: item.type,
                      size: 20, // 10pt
                      color: "1F2937"
                    })],
                    alignment: AlignmentType.CENTER
                  })],
                  shading: { fill: index % 2 === 0 ? "F8FAFC" : "FFFFFF" },
                  margins: { top: 80, bottom: 80, left: 50, right: 50 }
                }),
                // Status
                new TableCell({
                  children: [new Paragraph({
                    children: [new TextRun({
                      text: item.status,
                      size: 20, // 10pt
                      color: "1F2937"
                    })],
                    alignment: AlignmentType.CENTER
                  })],
                  shading: { fill: index % 2 === 0 ? "F8FAFC" : "FFFFFF" },
                  margins: { top: 80, bottom: 80, left: 50, right: 50 }
                }),
                // Submitted by
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [new TextRun({
                        text: item.submittedBy,
                        size: 20, // 10pt
                        color: "1F2937"
                      })],
                      spacing: { after: 40 }
                    }),
                    new Paragraph({
                      children: [new TextRun({
                        text: item.submittedDate,
                        size: 18, // 9pt
                        color: "6B7280"
                      })]
                    })
                  ],
                  shading: { fill: index % 2 === 0 ? "F8FAFC" : "FFFFFF" },
                  margins: { top: 80, bottom: 80, left: 50, right: 50 }
                })
              ]
            });
          })
        ]
      });

      sections.push(itemsTable);
    }

    // Actions Ready to Close — drawn from the whole backlog (all meetings)
    if (data.readyToCloseActions.length > 0) {
      sections.push(
        new Paragraph({
          children: [
            new TextRun({
              text: "Actions Ready to Close",
              bold: true,
              size: 32, // 16pt
              color: "1F2937"
            })
          ],
          spacing: { before: 300, after: 150 }
        })
      );
      sections.push(this.createReadyToCloseTable(data.readyToCloseActions));
    }

    // Attendance Section Header
    sections.push(
      new Paragraph({
        children: [
          new TextRun({
            text: "Meeting Attendance",
            bold: true,
            size: 32, // 16pt
            color: "1F2937"
          })
        ],
        spacing: { before: 300, after: 200 }
      })
    );

    // Attendance Table
    sections.push(this.createProfessionalAttendanceTable(data));

    // Create document
    const doc = new Document({
      sections: [{
        properties: {},
        children: sections
      }]
    });

    return await Packer.toBuffer(doc);
  }

  /**
   * Table of actions parked at "Ready to Close" — completed work awaiting a group
   * review + sign-off. Pulled from the whole backlog and stamped with each due
   * date, since the same action can recur across consecutive meetings.
   */
  private static createReadyToCloseTable(actions: ReadyToCloseAction[]): Table {
    const headerCell = (text: string, width: number) => new TableCell({
      children: [new Paragraph({
        children: [new TextRun({ text, bold: true, size: 20, color: "166534" })],
      })],
      width: { size: width, type: WidthType.PERCENTAGE },
      shading: { fill: "DCFCE7" },
      margins: { top: 80, bottom: 80, left: 100, right: 100 }
    });

    const bodyCell = (text: string, width: number, bold = false) => new TableCell({
      children: [new Paragraph({
        children: [new TextRun({ text: text || "—", size: 18, color: bold ? "1F2937" : "4B5563", bold })],
      })],
      width: { size: width, type: WidthType.PERCENTAGE },
      margins: { top: 80, bottom: 80, left: 100, right: 100 }
    });

    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 2, color: "86EFAC" },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: "86EFAC" },
        left: { style: BorderStyle.SINGLE, size: 2, color: "86EFAC" },
        right: { style: BorderStyle.SINGLE, size: 2, color: "86EFAC" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "BBF7D0" },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "BBF7D0" }
      },
      rows: [
        new TableRow({
          tableHeader: true,
          children: [
            headerCell("Item", 28),
            headerCell("Type", 14),
            headerCell("Actioned By", 16),
            headerCell("Due Date", 13),
            headerCell("What Was Done", 29),
          ]
        }),
        ...actions.map((action) => new TableRow({
          children: [
            bodyCell(action.title, 28, true),
            bodyCell(action.type, 14),
            bodyCell(action.assignedTo, 16),
            bodyCell(action.dueDate, 13),
            bodyCell(action.outcome, 29),
          ]
        }))
      ]
    });
  }

  /**
   * Create clean attendance table
   */
  private static createProfessionalAttendanceTable(data: TemplateData): Table {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 2, color: "D1D5DB" },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: "D1D5DB" },
        left: { style: BorderStyle.SINGLE, size: 2, color: "D1D5DB" },
        right: { style: BorderStyle.SINGLE, size: 2, color: "D1D5DB" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "E5E7EB" },
        insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "E5E7EB" }
      },
      rows: [
        // Header row
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({
                children: [new TextRun({ text: "Management Team", bold: true, size: 24 // 12pt
                  })],
                alignment: AlignmentType.CENTER
              })],
              width: { size: 40, type: WidthType.PERCENTAGE },
              shading: { fill: "F9FAFB" },
              margins: { top: 80, bottom: 80, left: 100, right: 100 }
            }),
            new TableCell({
              children: [new Paragraph({
                children: [new TextRun({ text: "Present", bold: true, size: 24 // 12pt
                  })],
                alignment: AlignmentType.CENTER
              })],
              width: { size: 10, type: WidthType.PERCENTAGE },
              shading: { fill: "F9FAFB" },
              margins: { top: 80, bottom: 80, left: 50, right: 50 }
            }),
            new TableCell({
              children: [new Paragraph({
                children: [new TextRun({ text: "Glaziers", bold: true, size: 24 // 12pt
                  })],
                alignment: AlignmentType.CENTER
              })],
              width: { size: 40, type: WidthType.PERCENTAGE },
              shading: { fill: "F9FAFB" },
              margins: { top: 80, bottom: 80, left: 100, right: 100 }
            }),
            new TableCell({
              children: [new Paragraph({
                children: [new TextRun({ text: "Present", bold: true, size: 24 // 12pt
                  })],
                alignment: AlignmentType.CENTER
              })],
              width: { size: 10, type: WidthType.PERCENTAGE },
              shading: { fill: "F9FAFB" },
              margins: { top: 80, bottom: 80, left: 50, right: 50 }
            })
          ]
        }),
        // Data rows
        ...Array.from({ length: Math.max(data.managementTeam.length, data.glaziers.length) }, (_, i) => {
          const mgmtPerson = data.managementTeam[i];
          const glazier = data.glaziers[i];

          return new TableRow({
            children: [
              new TableCell({
                children: [new Paragraph({
                  children: mgmtPerson ? [
                    new TextRun({
                      text: `${mgmtPerson.name} - ${mgmtPerson.role}`,
                      size: 20, // 10pt
                      color: "374151"
                    })
                  ] : [new TextRun({ text: "", size: 1 })]
                })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" },
                margins: { top: 80, bottom: 80, left: 100, right: 100 }
              }),
              new TableCell({
                children: [new Paragraph({
                  children: mgmtPerson ? [
                    new TextRun({
                      text: mgmtPerson.present ? "Yes" : "No",
                      size: 20, // 10pt
                      color: mgmtPerson.present ? "059669" : "DC2626"
                    })
                  ] : [new TextRun({ text: "", size: 1 })]
                })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" },
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
              }),
              new TableCell({
                children: [new Paragraph({
                  children: glazier ? [
                    new TextRun({
                      text: glazier.name,
                      size: 20, // 10pt
                      color: "374151"
                    })
                  ] : [new TextRun({ text: "", size: 1 })]
                })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" },
                margins: { top: 80, bottom: 80, left: 100, right: 100 }
              }),
              new TableCell({
                children: [new Paragraph({
                  children: glazier ? [
                    new TextRun({
                      text: glazier.present ? "Yes" : "No",
                      size: 20, // 10pt
                      color: glazier.present ? "059669" : "DC2626"
                    })
                  ] : [new TextRun({ text: "", size: 1 })]
                })],
                shading: { fill: i % 2 === 0 ? "F8FAFC" : "FFFFFF" },
                margins: { top: 80, bottom: 80, left: 50, right: 50 }
              })
            ]
          });
        })
      ]
    });
  }

  /**
   * Create professional signature section
   */
  private static createProfessionalSignatureSection(): any[] {
    return [
      new Paragraph({
        children: [new TextRun({
          text: "Meeting Approval & Sign-off",
          bold: true,
          size: 36, // 18pt
          color: "1F2937"
        })],
        spacing: { before: 600, after: 300 },
        alignment: AlignmentType.CENTER
      }),
      
      new Paragraph({
        children: [
          new TextRun({ text: "H&S Coordinator: ", bold: true, size: 24, color: "374151" }),
          new TextRun({ text: "_________________________________ ", size: 24, color: "6B7280" }),
          new TextRun({ text: "Date: ", bold: true, size: 24, color: "374151" }),
          new TextRun({ text: "_______________", size: 24, color: "6B7280" })
        ],
        spacing: { before: 200, after: 200 }
      }),
      
      new Paragraph({
        children: [
          new TextRun({ text: "H&S Manager: ", bold: true, size: 24, color: "374151" }),
          new TextRun({ text: "_________________________________ ", size: 24, color: "6B7280" }),
          new TextRun({ text: "Date: ", bold: true, size: 24, color: "374151" }),
          new TextRun({ text: "_______________", size: 24, color: "6B7280" })
        ],
        spacing: { after: 200 }
      }),
      
      new Paragraph({
        children: [
          new TextRun({ text: "General Manager: ", bold: true, size: 24, color: "374151" }),
          new TextRun({ text: "_________________________________ ", size: 24, color: "6B7280" }),
          new TextRun({ text: "Date: ", bold: true, size: 24, color: "374151" }),
          new TextRun({ text: "_______________", size: 24, color: "6B7280" })
        ],
        spacing: { after: 400 }
      }),
      
      new Paragraph({
        children: [new TextRun({
          text: "Next Meeting Details",
          bold: true,
          size: 36, // 18pt
          color: "1F2937"
        })],
        spacing: { before: 400, after: 200 },
        alignment: AlignmentType.CENTER
      }),
      
      new Paragraph({
        children: [
          new TextRun({ text: "Date: ", bold: true, size: 24, color: "374151" }),
          new TextRun({ text: "_________________________________ ", size: 24, color: "6B7280" }),
          new TextRun({ text: "Time: ", bold: true, size: 24, color: "374151" }),
          new TextRun({ text: "_________________", size: 24, color: "6B7280" })
        ],
        spacing: { before: 100, after: 150 }
      }),
      
      new Paragraph({
        children: [
          new TextRun({ text: "Location: ", bold: true, size: 24, color: "374151" }),
          new TextRun({ text: "_________________________________________________________________", size: 24, color: "6B7280" })
        ],
        spacing: { after: 300 }
      }),
      

    ];
  }

  /**
   * Format meeting date
   */
  private static formatMeetingDate(dateString: string): string {
    if (!dateString) return 'Unknown Date';
    
    const parsedDate = new Date(dateString);
    
    if (isNaN(parsedDate.getTime())) {
      return 'Invalid Date';
    }
    
    return parsedDate.toLocaleDateString('en-GB', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric'
    });
  }

  /**
   * Get status color for styling
   */
  private static getStatusColor(status: string): string {
    switch (status?.toLowerCase()) {
      case 'submitted': return '3B82F6';
      case 'in discussion': return 'F59E0B';
      case 'actions': return '10B981';
      case 'closed': return '6B7280';
      default: return '4B5563';
    }
  }
}