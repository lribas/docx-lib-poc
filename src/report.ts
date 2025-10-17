import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  ImageRun,
  SectionType,
  PageOrientation,
  Header,
} from "docx";
import fs from "fs";
import path from "path";

async function generateReport() {
  const colorPrincipal = "ff0000"; // cor da borda em hex (sem “#”)
  //   const logoPath = path.join(__dirname, "../logo_empresa.png"); // ajuste conforme local da logo

  // Exemplo de dados
  const data = {
    cliente: "GRUPO TESTE",
    representadoPor: "Espólio de Daniel Figlarz representado(a) por Sonia",
    cidade: "Curitiba",
    data: "31 de julho de 2025",
    quantidadeAcoes: { ativo: 1, passivo: 1, terceiro: 0 },
    valores: {
      ativo: "R$ 306.000,00",
      passivo: "R$ 390.000,00",
      terceiro: "R$ 0,00",
    },
    processos: [
      {
        numero: "0018169-08.2018.8.16.0188",
        assunto: "Inventário e Partilha",
        vara: "1ª Vara de Sucessões de Curitiba",
        valor: "R$ 306.000,00",
        magistrado: "Ronaldo Sansone Guerra",
        instancia: "Primeira",
        sistema: "Projudi Paraná",
        ultimaMov: "14/07/2025 16:03",
        resumo:
          "Inventário em tramitação regular. Última movimentação: confirmação de comunicação eletrônica.",
      },
      {
        numero: "0014836-72.2023.8.16.0188",
        assunto: "Inventário e Partilha",
        vara: "1ª Vara de Sucessões de Curitiba",
        valor: "R$ 390.000,00",
        magistrado: "Ronaldo Sansone Guerra",
        instancia: "Primeira",
        sistema: "Projudi Paraná",
        ultimaMov: "25/07/2025 17:13",
        resumo: "Avaliação judicial de imóvel em andamento, aguardando manifestação das partes.",
      },
    ],
  };

  // Função auxiliar para célula de tabela
  function makeCell(content: string | number, isHeader = false, backgroundColor?: string): TableCell {
    return new TableCell({
      shading: backgroundColor ? { fill: backgroundColor, color: "FFFFFF" } : undefined,
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: String(content),
              bold: isHeader,
              color: backgroundColor ? "FFFFFF" : "000000",
            }),
          ],
          alignment: AlignmentType.CENTER,
        }),
      ],
      width: { size: 100 / 3, type: WidthType.PERCENTAGE },
    });
  }

  function _createSummaryTable(q: { ativo: number; passivo: number; terceiro: number }): Table {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            makeCell("Polo Ativo", true, colorPrincipal),
            makeCell("Polo Passivo", true, colorPrincipal),
            makeCell("Terceiro", true, colorPrincipal),
          ],
        }),
        new TableRow({
          children: [makeCell(q.ativo), makeCell(q.passivo), makeCell(q.terceiro)],
        }),
      ],
    });
  }

  function _createValuesTable(v: { ativo: string; passivo: string; terceiro: string }): Table {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            makeCell("Valor de Causa (Ativo)", true, colorPrincipal),
            makeCell("Valor de Causa (Passivo)", true, colorPrincipal),
            makeCell("Valor de Causa (Terceiro)", true, colorPrincipal),
          ],
        }),
        new TableRow({
          children: [makeCell(v.ativo), makeCell(v.passivo), makeCell(v.terceiro)],
        }),
      ],
    });
  }

  // Cria o documento

  const doc = new Document({
    styles: {
      paragraphStyles: [
        {
          id: "Roboto",
          name: "Roboto",
          basedOn: "Normal",
          next: "Normal",
          run: {
            font: "Roboto",
          },
        },
      ],
    },
    sections: [
      // Seção da capa
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            margin: {
              top: 20, // Margem zero para a logo ocupar toda largura
              right: 20,
              bottom: 20,
              left: 20,
            },
            borders: {
              pageBorderTop: {
                color: colorPrincipal,
                size: 20,
                style: BorderStyle.DOUBLE,
              },
              pageBorderBottom: {
                color: colorPrincipal,
                size: 12,
                style: BorderStyle.SINGLE,
              },
              pageBorderLeft: {
                color: colorPrincipal,
                size: 12,
                style: BorderStyle.SINGLE,
              },
              pageBorderRight: {
                color: colorPrincipal,
                size: 12,
                style: BorderStyle.SINGLE,
              },
            },
          },
        },
        children: [
          // Logo cobrindo toda a largura no topo
          new Paragraph({
            children: [
              new ImageRun({
                data: fs.readFileSync("./lt_logo_blue_full_darkblue_1.png"),
                // data: fs.readFileSync("./logo1.png"),
                transformation: {
                  width: 800, // 3535 × 1024
                  height: 232,
                },
                type: "png",
              }),
            ],
            alignment: AlignmentType.END,
            // spacing: { before: 0, after: 2000 },
          }),

          // Conteúdo da capa
          new Paragraph({
            children: [
              new TextRun({
                text: "RELATÓRIO DE PRESTAÇÃO DE SERVIÇOS JURÍDICOS",
                bold: true,
                size: 48,
                font: "Roboto",
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 2000, after: 1000 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: data.cliente,
                size: 32,
                font: "Roboto",
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 500 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: `${data.cidade}, ${data.data}`,
                size: 28,
                font: "Roboto",
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 3000 },
          }),
        ],
      },

      // Seção do conteúdo
      {
        properties: {
          page: {
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440,
            },
          },
        },
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: `À ${data.representadoPor}`,
                font: "Roboto",
                size: 24,
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Prezados(as) Senhores(as),",
                font: "Roboto",
                size: 24,
              }),
            ],
            spacing: { after: 200 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "Tem a presente e especial finalidade de encaminhar relatório referente aos processos judiciais, sob o patrocínio de nosso escritório até esta data.",
                font: "Roboto",
                size: 24,
              }),
            ],
            spacing: { after: 400 },
          }),

          new Paragraph({
            children: [
              new TextRun({
                text: "QUANTIDADE DE AÇÕES",
                font: "Roboto",
                bold: true,
              }),
            ],
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 },
          }),
          createSummaryTable(data.quantidadeAcoes.ativo),

          new Paragraph({
            children: [
              new TextRun({
                text: "VALORES ENVOLVIDOS",
                font: "Roboto",
                bold: true,
              }),
            ],
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 },
          }),
          createValuesTable(data.valores.ativo),

          new Paragraph({
            children: [
              new TextRun({
                text: "AÇÕES",
                font: "Roboto",
                bold: true,
              }),
            ],
            heading: HeadingLevel.HEADING_2,
            spacing: { before: 400, after: 200 },
          }),

          ...data.processos.map(
            (p, i) =>
              new Paragraph({
                children: [
                  new TextRun({
                    text: `${i + 1}. ${p.assunto} (${p.numero})`,
                    bold: true,
                    size: 32,
                    font: "Roboto",
                  }),
                  new TextRun({
                    text: `\nVara: ${p.vara}\nValor da causa: ${p.valor}\nMagistrado: ${p.magistrado}\nInstância: ${p.instancia}\nSistema: ${p.sistema}\nÚltima movimentação: ${p.ultimaMov}\n\nResumo: ${p.resumo}`,
                    size: 24,
                    font: "Roboto",
                  }),
                ],
                spacing: { after: 400 },
              })
          ),
        ],
      },
    ],
  });

  // Funções auxiliares para tabelas (atualizadas com fonte Roboto)
  function createSummaryTable(quantidadeAcoes: number) {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "Tipo", bold: true, font: "Roboto" })],
                }),
              ],
              shading: { fill: "E6E6E6" },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "Quantidade", bold: true, font: "Roboto" })],
                }),
              ],
              shading: { fill: "E6E6E6" },
            }),
          ],
        }),
        ...Object.entries(quantidadeAcoes).map(
          ([tipo, quantidade]) =>
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: tipo, font: "Roboto" })],
                    }),
                  ],
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: quantidade.toString(), font: "Roboto" })],
                    }),
                  ],
                }),
              ],
            })
        ),
      ],
    });
  }

  function createValuesTable(valores: string) {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "Descrição", bold: true, font: "Roboto" })],
                }),
              ],
              shading: { fill: "E6E6E6" },
            }),
            new TableCell({
              children: [
                new Paragraph({
                  children: [new TextRun({ text: "Valor", bold: true, font: "Roboto" })],
                }),
              ],
              shading: { fill: "E6E6E6" },
            }),
          ],
        }),
        ...Object.entries(valores).map(
          ([descricao, valor]) =>
            new TableRow({
              children: [
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [new TextRun({ text: descricao, font: "Roboto" })],
                    }),
                  ],
                }),
                new TableCell({
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: valor.toString(),
                          // typeof valor === "number"
                          //   ? `R$ ${valor.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}`
                          //   : valor.toString(),
                          font: "Roboto",
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            })
        ),
      ],
    });
  }

  const buffer = await Packer.toBuffer(doc);
  const outPath = path.join(process.cwd(), "RelatorioPrestacaoServicos.docx");
  fs.writeFileSync(outPath, buffer);
  console.log("Relatório gerado:", outPath);
}

// Executa
generateReport().catch((err) => {
  console.error("Erro ao gerar relatório:", err);
});
