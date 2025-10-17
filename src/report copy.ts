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
    representadoPor: "Espólio de Sbrubles representado(a) por Sonia",
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
        magistrado: "Juca Chaves",
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
        magistrado: "Juca Chaves",
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

  function createSummaryTable(q: { ativo: number; passivo: number; terceiro: number }): Table {
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

  function createValuesTable(v: { ativo: string; passivo: string; terceiro: string }): Table {
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
    sections: [
      // Seção da capa
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            borders: {
              pageBorderTop: {
                color: colorPrincipal,
                size: 12,
                style: BorderStyle.SINGLE,
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

            // orientation: PageOrientation.PORTRAIT,
          },
        },

        headers: {
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new ImageRun({
                    data: fs.readFileSync("./lt_logo_blue_full_darkblue_1.png"),
                    transformation: {
                      width: 600,
                      height: 200,
                    },
                    type: "png",
                  }),
                ],
              }),
            ],
          }),
        },
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: "RELATÓRIO DE PRESTAÇÃO DE SERVIÇOS JURÍDICOS",
                bold: true,
                size: 48, // 48 half-points = 24pt
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 4000, after: 1000 },
          }),
          new Paragraph({
            text: data.cliente,
            alignment: AlignmentType.CENTER,
            spacing: { after: 500 },
          }),
          new Paragraph({
            text: `${data.cidade}, ${data.data}`,
            alignment: AlignmentType.CENTER,
          }),
        ],
      },

      // Seção do conteúdo
      {
        // headers: {
        //   default: new Header({
        //     children: [
        //       new Paragraph({
        //         children: [
        //           new ImageRun({
        //             type: "png",
        //             data: fs.readFileSync("./lt_logo_blue_full_darkblue_1.png"),
        //             transformation: {
        //               width: 0,
        //               height: 0,
        //             },
        //             floating: {
        //               horizontalPosition: {
        //                 offset: 0,
        //               },
        //               verticalPosition: {
        //                 offset: 0,
        //               },
        //             },
        //           }),
        //         ],
        //       }),
        //       new Paragraph({
        //         children: [
        //           new ImageRun({
        //             type: "png",
        //             data: fs.readFileSync("./logo0.png"),
        //             transformation: {
        //               width: 1000,
        //               height: 300,
        //             },
        //           }),
        //         ],
        //       }),
        //       new Paragraph({
        //         children: [
        //           new ImageRun({
        //             type: "png",
        //             data: fs.readFileSync("./logo1.png"),
        //             transformation: {
        //               width: 100,
        //               height: 100,
        //             },
        //           }),
        //         ],
        //       }),
        //     ],
        //   }),
        // },
        children: [
          new Paragraph({
            text: `À ${data.representadoPor}`,
            spacing: { after: 200 },
          }),
          new Paragraph({
            text: "Prezados(as) Senhores(as),\n\nTem a presente e especial finalidade de encaminhar relatório referente aos processos judiciais, sob o patrocínio de nosso escritório até esta data.",
            spacing: { after: 400 },
          }),

          new Paragraph({
            text: "QUANTIDADE DE AÇÕES",
            heading: HeadingLevel.HEADING_2,
            // bold: true,
            spacing: { before: 400, after: 200 },
          }),
          createSummaryTable(data.quantidadeAcoes),

          new Paragraph({
            text: "VALORES ENVOLVIDOS",
            heading: HeadingLevel.HEADING_2,
            // bold: true,
            spacing: { before: 400, after: 200 },
          }),
          createValuesTable(data.valores),

          new Paragraph({
            text: "AÇÕES",
            heading: HeadingLevel.HEADING_2,
            // bold: true,
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
                  }),
                  new TextRun({
                    text: `\nVara: ${p.vara}\nValor da causa: ${p.valor}\nMagistrado: ${p.magistrado}\nInstância: ${p.instancia}\nSistema: ${p.sistema}\nÚltima movimentação: ${p.ultimaMov}\n\nResumo: ${p.resumo}`,
                    size: 24,
                  }),
                ],
                spacing: { after: 400 },
              })
          ),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  const outPath = path.join(process.cwd(), "RelatorioPrestacaoServicos.docx");
  fs.writeFileSync(outPath, buffer);
  console.log("Relatório gerado:", outPath);
}

// Executa
generateReport().catch((err) => {
  console.error("Erro ao gerar relatório:", err);
});
