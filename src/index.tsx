import React, { Component } from "react";
import { createRoot } from "react-dom/client";
import "./style.css";
import { saveAs } from "file-saver";
// import { Packer } from "docx";

import {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  TabStopPosition,
  TabStopType,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  VerticalAlign,
  ImageRun,
  convertInchesToTwip,
} from "docx";
const PHONE_NUMBER = "07534563401";
const PROFILE_URL = "https://www.linkedin.com/in/dolan1";
const EMAIL = "docx@docx.com";
// import { experiences, education, skills, achievements } from "./cv-data";
// import { DocumentCreator } from "./cv-generator.ts";

interface AppProps {}

interface AppState {
  name: string;
}

class App extends Component<AppProps, AppState> {
  constructor(props: AppProps) {
    super(props);
    this.state = {
      name: "React",
    };
    this.generate = this.generate.bind(this);
  }

  async generate(): Promise<void> {
    try {
      const response = await fetch('http://localhost:3001/api/documents/generate-certificate');
      if (!response.ok) {
        throw new Error('Failed to generate certificate');
      }
      const blob = await response.blob();
      saveAs(blob, "certificate.docx");
      console.log("Document created successfully");
    } catch (error) {
      console.error("Error generating document:", error);
      alert("Failed to generate document. Please try again.");
    }
  }

  render() {
    return (
      <div>
        <h1>Certificate of Analysis Generator</h1>
        <p>
          Click the button below to generate the certificate
          <button onClick={this.generate}>
            Generate Certificate of Analysis
          </button>
        </p>
      </div>
    );
  }
}

const root = createRoot(document.getElementById("root")!);
root.render(<App />);

export const experiences = [
  {
    company: {
      name: "Example Company",
    },
    position: "Software Developer",
    startDate: {
      year: 2020,
    },
    endDate: {
      year: 2023,
    },
    summary:
      "Developed and maintained web applications using React and Node.js",
  },
];

export const education = [
  {
    schoolName: "University of Example",
    degree: "Bachelor of Science",
    fieldOfStudy: "Computer Science",
    startDate: {
      year: 2016,
    },
    endDate: {
      year: 2020,
    },
    notes: "",
  },
];

export const skills = [
  {
    name: "JavaScript",
  },
  {
    name: "TypeScript",
  },
  {
    name: "React",
  },
  {
    name: "Node.js",
  },
];

export const achievements = [
  {
    name: "Successfully implemented a document generation system",
  },
  {
    name: "Improved application performance by 40%",
  },
];

export class DocumentCreator {
  // tslint:disable-next-line: typedef
  public create([experiences, educations, skills, achivements]): Document {
    const document = new Document({
      sections: [
        {
          children: [
            this.createCertificateOfAnalysis(),
            new Paragraph({
              text: "Dolan Miu",
              heading: HeadingLevel.TITLE,
            }),
            this.createContactInfo(PHONE_NUMBER, PROFILE_URL, EMAIL),
            this.createHeading("Education"),
            ...educations
              .map((education) => {
                const arr: Paragraph[] = [];
                arr.push(
                  this.createInstitutionHeader(
                    education.schoolName,
                    `${education.startDate.year} - ${education.endDate.year}`
                  )
                );
                arr.push(
                  this.createRoleText(
                    `${education.fieldOfStudy} - ${education.degree}`
                  )
                );

                const bulletPoints = this.splitParagraphIntoBullets(
                  education.notes
                );
                bulletPoints.forEach((bulletPoint) => {
                  arr.push(this.createBullet(bulletPoint));
                });

                return arr;
              })
              .reduce((prev, curr) => prev.concat(curr), []),
            this.createHeading("Experience"),
            ...experiences
              .map((position) => {
                const arr: Paragraph[] = [];

                arr.push(
                  this.createInstitutionHeader(
                    position.company.name,
                    this.createPositionDateText(
                      position.startDate,
                      position.endDate,
                      position.isCurrent
                    )
                  )
                );
                arr.push(this.createRoleText(position.title));

                const bulletPoints = this.splitParagraphIntoBullets(
                  position.summary
                );

                bulletPoints.forEach((bulletPoint) => {
                  arr.push(this.createBullet(bulletPoint));
                });

                return arr;
              })
              .reduce((prev, curr) => prev.concat(curr), []),
            this.createHeading("Skills, Achievements and Interests"),
            this.createSubHeading("Skills"),
            this.createSkillList(skills),
            this.createSubHeading("Achievements"),
            ...this.createAchivementsList(achivements),
            this.createSubHeading("Interests"),
            this.createInterests(
              "Programming, Technology, Music Production, Web Design, 3D Modelling, Dancing."
            ),
            this.createHeading("References"),
            new Paragraph(
              "Dr. Dean Mohamedally Director of Postgraduate Studies Department of Computer Science, University College London Malet Place, Bloomsbury, London WC1E d.mohamedally@ucl.ac.uk"
            ),
            new Paragraph("More references upon request"),
            new Paragraph({
              text: "This CV was generated in real-time based on my Linked-In profile from my personal website www.dolan.bio.",
              alignment: AlignmentType.CENTER,
            }),
            this.createExperienceTable(experiences),
            this.createEducationTable(educations),
            this.createSkillsList(skills),
            this.createAchievementsList(achivements),
          ],
        },
      ],
    });

    return document;
  }

  public createContactInfo(
    phoneNumber: string,
    profileUrl: string,
    email: string
  ): Paragraph {
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun(
          `Mobile: ${phoneNumber} | LinkedIn: ${profileUrl} | Email: ${email}`
        ),
        new TextRun({
          text: "Address: 58 Elm Avenue, Kent ME4 6ER, UK",
          break: 1,
        }),
      ],
    });
  }

  public createHeading(text: string): Paragraph {
    return new Paragraph({
      text: text,
      heading: HeadingLevel.HEADING_1,
      thematicBreak: true,
    });
  }

  public createSubHeading(text: string): Paragraph {
    return new Paragraph({
      text: text,
      heading: HeadingLevel.HEADING_2,
    });
  }

  public createInstitutionHeader(
    institutionName: string,
    dateText: string
  ): Paragraph {
    return new Paragraph({
      tabStops: [
        {
          type: TabStopType.RIGHT,
          position: TabStopPosition.MAX,
        },
      ],
      children: [
        new TextRun({
          text: institutionName,
          bold: true,
        }),
        new TextRun({
          text: `\t${dateText}`,
          bold: true,
        }),
      ],
    });
  }

  public createRoleText(roleText: string): Paragraph {
    return new Paragraph({
      children: [
        new TextRun({
          text: roleText,
          italics: true,
        }),
      ],
    });
  }

  public createBullet(text: string): Paragraph {
    return new Paragraph({
      text: text,
      bullet: {
        level: 0,
      },
    });
  }

  // tslint:disable-next-line:no-any
  public createSkillList(skills: any[]): Paragraph {
    return new Paragraph({
      children: [
        new TextRun(skills.map((skill) => skill.name).join(", ") + "."),
      ],
    });
  }

  // tslint:disable-next-line:no-any
  public createAchivementsList(achivements: any[]): Paragraph[] {
    return achivements.map(
      (achievement) =>
        new Paragraph({
          text: achievement.name,
          bullet: {
            level: 0,
          },
        })
    );
  }

  public createInterests(interests: string): Paragraph {
    return new Paragraph({
      children: [new TextRun(interests)],
    });
  }

  public splitParagraphIntoBullets(text: string): string[] {
    return text ? text.split("\n\n") : [];
  }

  // tslint:disable-next-line:no-any
  public createPositionDateText(
    startDate: any,
    endDate: any,
    isCurrent: boolean
  ): string {
    const startDateText =
      this.getMonthFromInt(startDate.month) + ". " + startDate.year;
    const endDateText = isCurrent
      ? "Present"
      : `${this.getMonthFromInt(endDate.month)}. ${endDate.year}`;

    return `${startDateText} - ${endDateText}`;
  }

  public getMonthFromInt(value: number): string {
    switch (value) {
      case 1:
        return "Jan";
      case 2:
        return "Feb";
      case 3:
        return "Mar";
      case 4:
        return "Apr";
      case 5:
        return "May";
      case 6:
        return "Jun";
      case 7:
        return "Jul";
      case 8:
        return "Aug";
      case 9:
        return "Sept";
      case 10:
        return "Oct";
      case 11:
        return "Nov";
      case 12:
        return "Dec";
      default:
        return "N/A";
    }
  }

  public createExperienceTable(experiences) {
    const rows = experiences.map(
      (exp) =>
        new TableRow({
          children: [
            new TableCell({
              width: {
                size: 100,
                type: WidthType.PERCENTAGE,
              },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: exp.company.name, bold: true }),
                    new TextRun({
                      text: ` (${exp.startDate.year} - ${exp.endDate.year})`,
                      bold: true,
                    }),
                  ],
                }),
                new Paragraph({ text: exp.title }),
                new Paragraph({ text: exp.summary }),
              ],
            }),
          ],
        })
    );

    return new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ text: "Experience", bold: true })],
            }),
          ],
        }),
        ...rows,
      ],
    });
  }

  public createEducationTable(education) {
    const rows = education.map(
      (edu) =>
        new TableRow({
          children: [
            new TableCell({
              width: {
                size: 100,
                type: WidthType.PERCENTAGE,
              },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: edu.schoolName, bold: true }),
                    new TextRun({
                      text: ` (${edu.startDate.year} - ${edu.endDate.year})`,
                      bold: true,
                    }),
                  ],
                }),
                new Paragraph({ text: edu.degree }),
              ],
            }),
          ],
        })
    );

    return new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ text: "Education", bold: true })],
            }),
          ],
        }),
        ...rows,
      ],
    });
  }

  public createSkillsList(skills) {
    return new Paragraph({
      children: [
        new TextRun({ text: "Skills: ", bold: true }),
        new TextRun({ text: skills.map((skill) => skill.name).join(", ") }),
      ],
    });
  }

  public createAchievementsList(achievements) {
    return new Paragraph({
      children: [
        new TextRun({ text: "Achievements: ", bold: true }),
        new TextRun({
          text: achievements.map((achievement) => achievement.name).join(". "),
        }),
      ],
    });
  }

  public createCertificateOfAnalysis(): Table {
    return new Table({
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      borders: {
        top: { style: "single", size: 1, color: "000000" },
        bottom: { style: "single", size: 1, color: "000000" },
        left: { style: "single", size: 1, color: "000000" },
        right: { style: "single", size: 1, color: "000000" },
        insideHorizontal: { style: "single", size: 1, color: "000000" },
        insideVertical: { style: "single", size: 1, color: "000000" },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 15,
                type: WidthType.PERCENTAGE,
              },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: "C\nE\nR\nT\nI\nF\nI\nC\nA\nT\nE\n\nO\nF\n\nA\nN\nA\nL\nY\nS\nI\nS",
                      bold: true,
                      size: 24,
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              width: {
                size: 85,
                type: WidthType.PERCENTAGE,
              },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: "Product Information", bold: true }),
                  ],
                }),
                ...this.createProductInfoRows().map(
                  (info) =>
                    new Paragraph({
                      children: [
                        new TextRun({ text: info.label + ": ", bold: true }),
                        new TextRun({ text: info.value }),
                      ],
                    })
                ),
              ],
            }),
          ],
        }),
        // Test Headers with background color
        new TableRow({
          children: [
            new TableCell({
              shading: {
                fill: "D3E3C3",
                color: "auto",
                type: "clear",
              },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "TESTS", bold: true })],
                }),
              ],
            }),
            new TableCell({
              shading: {
                fill: "D3E3C3",
                color: "auto",
                type: "clear",
              },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: "SPECIFICATION", bold: true }),
                  ],
                }),
              ],
            }),
            new TableCell({
              shading: {
                fill: "D3E3C3",
                color: "auto",
                type: "clear",
              },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "RESULTS", bold: true })],
                }),
              ],
            }),
          ],
        }),
        // Section headers and content with background colors
        ...this.createAssayRows(),
        ...this.createPhysicalControlRows(),
        ...this.createChemicalControlRows(),
        ...this.createMicrobiologicalRows(),
        // Storage Section
        new TableRow({
          children: [
            new TableCell({
              columnSpan: 3,
              shading: {
                fill: "D3E3C3",
                color: "auto",
                type: "clear",
              },
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "STORAGE", bold: true })],
                }),
              ],
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              columnSpan: 3,
              children: [
                new Paragraph({
                  text: "Store in a well-sealed container in a cool and dry place away from direct sunlight.",
                }),
                new Paragraph({
                  text: "Conclusion: Conforms to USP33",
                  spacing: {
                    before: 200,
                  },
                }),
              ],
            }),
          ],
        }),
      ],
    });
  }

  private createProductInfoRows() {
    const productInfo = [
      { label: "Product Name", value: "L-Leucine" },
      { label: "Batch No", value: "55233389" },
      { label: "Manufacturing Date", value: "2023-08-14" },
      { label: "Expiry Date", value: "2026-08-13" },
      { label: "CAS No", value: "61-90-5" },
      { label: "Standard", value: "USP33" },
    ];

    return productInfo;
  }

  private createAssayRows(): TableRow[] {
    return [
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 3,
            shading: {
              fill: "D3E3C3",
              color: "auto",
              type: "clear",
            },
            children: [
              new Paragraph({
                children: [new TextRun({ text: "ASSAY", bold: true })],
              }),
            ],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            width: {
              size: 30,
              type: WidthType.PERCENTAGE,
            },
            children: [new Paragraph({ text: "Assay" })],
          }),
          new TableCell({
            width: {
              size: 35,
              type: WidthType.PERCENTAGE,
            },
            children: [new Paragraph({ text: "98.5% ~101.5%" })],
          }),
          new TableCell({
            width: {
              size: 35,
              type: WidthType.PERCENTAGE,
            },
            children: [new Paragraph({ text: "99.7%" })],
          }),
        ],
      }),
    ];
  }

  private createPhysicalControlRows(): TableRow[] {
    const physicalControls = [
      {
        test: "Appearance",
        spec: "White crystals or crystalline powder",
        result: "Conforms",
      },
      { test: "Identification", spec: "Positive", result: "Conforms" },
      { test: "Specific Rotation", spec: "+14.9° ~ +17.3°", result: "+15.5°" },
      { test: "pH", spec: "5.5 ~ 7.0", result: "5.8" },
      { test: "Loss on drying", spec: "<0.20%", result: "0.14%" },
      { test: "Residue on Ignition", spec: "<0.40%", result: "0.03%" },
    ];

    return [
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 3,
            shading: {
              fill: "D3E3C3",
              color: "auto",
              type: "clear",
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "PHYSICAL CONTROL", bold: true }),
                ],
              }),
            ],
          }),
        ],
      }),
      ...physicalControls.map(
        (control) =>
          new TableRow({
            children: [
              new TableCell({
                width: {
                  size: 30,
                  type: WidthType.PERCENTAGE,
                },
                children: [new Paragraph({ text: control.test })],
              }),
              new TableCell({
                width: {
                  size: 35,
                  type: WidthType.PERCENTAGE,
                },
                children: [new Paragraph({ text: control.spec })],
              }),
              new TableCell({
                width: {
                  size: 35,
                  type: WidthType.PERCENTAGE,
                },
                children: [new Paragraph({ text: control.result })],
              }),
            ],
          })
      ),
    ];
  }

  private createChemicalControlRows(): TableRow[] {
    const chemicalControls = [
      { test: "Heavy Metals", spec: "≤0.0015%", result: "<0.0015%" },
      { test: "Lead (Pb)", spec: "≤2ppm", result: "<2ppm" },
      { test: "Arsenic (As)", spec: "≤2ppm", result: "<2ppm" },
      { test: "Cadmium (Cd)", spec: "≤1ppm", result: "≤1ppm" },
      { test: "Mercury (Hg)", spec: "≤1ppm", result: "≤1ppm" },
      { test: "Iron (Fe)", spec: "≤0.003%", result: "<0.003%" },
      { test: "Chloride", spec: "<0.05%", result: "<0.05%" },
      { test: "Sulfate", spec: "<0.03%", result: "<0.03%" },
    ];

    return [
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 3,
            shading: {
              fill: "D3E3C3",
              color: "auto",
              type: "clear",
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "CHEMICAL CONTROL", bold: true }),
                ],
              }),
            ],
          }),
        ],
      }),
      ...chemicalControls.map(
        (control) =>
          new TableRow({
            children: [
              new TableCell({
                width: {
                  size: 30,
                  type: WidthType.PERCENTAGE,
                },
                children: [new Paragraph({ text: control.test })],
              }),
              new TableCell({
                width: {
                  size: 35,
                  type: WidthType.PERCENTAGE,
                },
                children: [new Paragraph({ text: control.spec })],
              }),
              new TableCell({
                width: {
                  size: 35,
                  type: WidthType.PERCENTAGE,
                },
                children: [new Paragraph({ text: control.result })],
              }),
            ],
          })
      ),
    ];
  }

  private createMicrobiologicalRows(): TableRow[] {
    const microControls = [
      { test: "Total plate Count", spec: "≤5,000cfu/g", result: "<1000cfu/g" },
      { test: "Yeast & Mold", spec: "≤100cfu/g", result: "<100cfu/g" },
      { test: "Salmonella", spec: "Negative", result: "Negative" },
      { test: "E.Coli", spec: "Negative", result: "Negative" },
      { test: "S.Aureus", spec: "Negative", result: "Negative" },
    ];

    return [
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 3,
            shading: {
              fill: "D3E3C3",
              color: "auto",
              type: "clear",
            },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "MICROBIOLOGICAL CONTROL", bold: true }),
                ],
              }),
            ],
          }),
        ],
      }),
      ...microControls.map(
        (control) =>
          new TableRow({
            children: [
              new TableCell({
                width: {
                  size: 30,
                  type: WidthType.PERCENTAGE,
                },
                children: [new Paragraph({ text: control.test })],
              }),
              new TableCell({
                width: {
                  size: 35,
                  type: WidthType.PERCENTAGE,
                },
                children: [new Paragraph({ text: control.spec })],
              }),
              new TableCell({
                width: {
                  size: 35,
                  type: WidthType.PERCENTAGE,
                },
                children: [new Paragraph({ text: control.result })],
              }),
            ],
          })
      ),
    ];
  }
}
