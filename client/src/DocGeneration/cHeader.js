import {
    Header,
    ImageRun,
    AlignmentType,
    Document,
    HeadingLevel,
    Packer,
    Paragraph,
    Tab,
    TabStopPosition,
    TabStopType,
    TextRun,
    HorizontalPositionAlign,
    VerticalPositionAlign,
    ExternalHyperlink,
    PageNumber,
    FrameAnchorType,
    ShadingType,
} from "docx";
import docData from "./tools/DocData";
import enumImg from "../_helpers/enum-Img.js";
class cHeader {
    static getFirstPageHeader(familyname, firstname, poste, nbexps, filename) {
        return new Header({
            // The header on first page when the 'Different First Page' option is activated
            children: [
                this.getHeader(familyname, firstname, poste, nbexps, filename),
                this.getHeaderLogo1stPage(),
                this.getHeaderGauche(),
            ],
        });
    }
    static getDefaultPageHeader() {
        return new Header({
            // The standard default header on every page or header on odd pages when the 'Different Odd & Even Pages' option is activated
            children: [
                this.getHeaderLogoDefault(),
            ],
        });
    }
    static getHeader(familyname, firstname, poste, nbexps, filename) {
        let name = firstname.trim() + " " + familyname.trim().toUpperCase();
        if (filename == 2) {
            name = firstname.trim().substr(0, 1).toUpperCase() + familyname.trim().substr(0, 2).toUpperCase();
        }
        return new Paragraph({
            children: [
                new TextRun({
                    text: name,
                    bold: true,
                    font: "Century Gothic",
                    size: 36,
                    color: "#e6f1eb",
                    alignment: AlignmentType.CENTER,
                }),
                new TextRun({
                    text: "",
                    break: 1,
                }),
                new TextRun({
                    text: this.getPosteAndExps(poste, nbexps),
                    bold: true,
                    font: "Century Gothic",
                    size: 24,
                    color: "#e6f1eb",
                    alignment: AlignmentType.CENTER,
                }),
            ],
            floating: {
                horizontalPosition: {
                    align: HorizontalPositionAlign.CENTER,
                },
                verticalPosition: {
                    align: VerticalPositionAlign.TOP,
                },
            },
        });
    }
    static getHeaderLogo1stPage() {
        return new Paragraph({
            children: [
                new ImageRun({
                    type: "png",
                    data: docData.urlToBlob(enumImg.Logo2),
                    transformation: {
                        width: 180,
                        height: 85,
                    },

                    floating: {
                        horizontalPosition: {
                            align: HorizontalPositionAlign.RIGHT,
                        },
                        verticalPosition: {
                            align: VerticalPositionAlign.TOP,
                        },
                    },
                }),
            ],
        });
    }
    static getHeaderGauche() {
        return new Paragraph({
            children: [
                new ImageRun({
                    type: "png",
                    data: docData.urlToBlob(enumImg.HeaderGauche),
                    transformation: {
                        width: 600,
                        height: 110,
                    },
                    floating: {
                        behindDocument: true,
                        horizontalPosition: {
                            align: HorizontalPositionAlign.LEFT,
                        },
                        verticalPosition: {
                            align: VerticalPositionAlign.TOP,
                        },
                    },
                }),
            ],
        });
    }
    static getHeaderLogoDefault() {
        return new Paragraph({
            children: [
                new ImageRun({
                    type: "png",
                    data: docData.urlToBlob(enumImg.Logo1),
                    transformation: {
                        width: 75,
                        height: 105,
                    },

                    floating: {
                        horizontalPosition: {
                            align: HorizontalPositionAlign.RIGHT,
                        },
                        verticalPosition: {
                            align: VerticalPositionAlign.TOP,
                        },
                    },
                }),
            ],
        });
    }
    static getPosteAndExps(poste, nbexps) {
        if (poste != "" && poste != null && poste.length > 0) {
            if (nbexps > 0)
                return poste + " (" + nbexps + " ans)";
            return poste;
        }
        return "";
    }

}

export default cHeader;