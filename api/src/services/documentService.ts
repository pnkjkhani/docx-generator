import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  ShadingType,
} from "docx";
import * as fs from 'fs';
import * as path from 'path';

interface ProductInfoRow {
  label: string;
  value: string;
}

interface Test {
  test: string;
  spec: string;
  result: string;
}

interface SubTable {
  name: string;
  tests: Test[];
}

interface TestTable {
  mainHeading: string;
  columns: string[];
  columnWidths: number[];
  subTables: SubTable[];
}

interface CertificateData {
  productInfo: {
    rows: ProductInfoRow[];
  };
  testTables: TestTable[];
}

export class DocumentService {
  private data: CertificateData;

  constructor() {
    const dataPath = path.join(__dirname, '..', 'data', 'certificate-data.json');
    this.data = JSON.parse(fs.readFileSync(dataPath, 'utf8'));
  }

  public getTestTables(): TestTable[] {
    return this.data.testTables;
  }

  public async generateDocument(): Promise<Buffer> {
    const document = new Document({
      sections: [
        {
          children: [
            this.createProductInfoTable(),
            new Paragraph({ text: "" }), // Add spacing between tables
            ...this.data.testTables.map(table => this.createTestTable(table)),
          ],
        },
      ],
    });

    return await Packer.toBuffer(document);
  }

  public createProductInfoTable(): Table {
    const greenShading = {
      fill: "E9EFE1",
      color: "auto",
      type: ShadingType.CLEAR,
    };

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
      rows: this.data.productInfo.rows.map(
        (info) =>
          new TableRow({
            children: [
              new TableCell({
                shading:
                  info.label === "Product Name" ? greenShading : undefined,
                width: {
                  size: 50,
                  type: WidthType.PERCENTAGE,
                },
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: info.label,
                        bold: true,
                      }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                shading:
                  info.label === "Product Name" ? greenShading : undefined,
                width: {
                  size: 50,
                  type: WidthType.PERCENTAGE,
                },
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: info.value,
                      }),
                    ],
                  }),
                ],
              }),
            ],
          })
      ),
    });
  }

  public createTestTable(tableData: TestTable): Table {
    const headerShading = {
      fill: "E9EFE1",
      color: "auto",
      type: ShadingType.CLEAR,
    };

    const categoryShading = {
      fill: "E9EFE1",
      color: "auto",
      type: ShadingType.CLEAR,
    };

    const rows: TableRow[] = [
      // Main heading row
      new TableRow({
        children: [
          new TableCell({
            shading: headerShading,
            columnSpan: 3,
            children: [new Paragraph({
              children: [new TextRun({ text: tableData.mainHeading, bold: true })]
            })],
          }),
        ],
      }),
      // Column headers row
      new TableRow({
        children: tableData.columns.map((header, index) => 
          new TableCell({
            shading: headerShading,
            width: { size: tableData.columnWidths[index], type: WidthType.PERCENTAGE },
            children: [new Paragraph({
              children: [new TextRun({ text: header, bold: true })]
            })],
          })
        ),
      }),
    ];

    // Add sub-tables
    tableData.subTables.forEach(subTable => {
      // Add category header
      rows.push(
        new TableRow({
          children: [
            new TableCell({
              shading: categoryShading,
              columnSpan: 3,
              children: [new Paragraph({
                children: [new TextRun({ text: subTable.name, bold: true })]
              })],
            }),
          ],
        })
      );

      // Add test rows for this category
      subTable.tests.forEach(test => {
        rows.push(
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ text: test.test })] }),
              new TableCell({ children: [new Paragraph({ text: test.spec })] }),
              new TableCell({ children: [new Paragraph({ text: test.result })] }),
            ],
          })
        );
      });
    });

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
      rows: rows,
    });
  }
}
