import {
    Paragraph,
    TextRun,
    AlignmentType,
    TableCell,
    TableRow,
    BorderStyle,
    WidthType,
} from "docx";
import docData from "./tools/DocData.js";
class cTools {
    static getToolsTableRow(tools) {
        let tablerow = new TableRow({
            children: [
                new TableCell({
                    children: [
                        docData.getSubTitle("Outils :"),
                    ],
                    borders: {
                        top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                        bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                        left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                        right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    },
                    width: { size: 20, type: WidthType.PERCENTAGE },
                }),
                new TableCell({
                    children: [new Paragraph({
                        children: [
                            new TextRun({
                                text: tools.trim(),
                                alignment: AlignmentType.LEFT,
                                font: "Century Gothic",
                                size: 20,
                            }),
                        ],
                        spacing: {
                            line: 250,
                        },
                    }), ],
                    borders: {
                        top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                        bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                        left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                        right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    },
                    width: { size: 80, type: WidthType.PERCENTAGE },
                })
            ],
        });
        return tablerow;
    }
}
export default cTools;