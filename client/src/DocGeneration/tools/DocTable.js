import exppro from "../../DocGeneration/cExpPro";
import expperso from "../../DocGeneration/cExpPerso";
import comp from "../../DocGeneration/cComps";
import certs from "../../DocGeneration/cCerts";
import bref from "../../DocGeneration/cBref";
import lang from "../../DocGeneration/cLang";
import system from "../../DocGeneration/cSystems.js";
import languages from "../../DocGeneration/cLanguages.js";
import tools from "../cTools.js";
import databases from "../../DocGeneration/cDatabases.js";
import environments from "../../DocGeneration/cEnvironments.js";
import tbrow from "./tableRow";
import {
    Table,
    TableRow,
    TableCell,
    BorderStyle,
} from "docx";
import docData from "./DocData";

class DocTable {
    static getTable(docjs) {
            const table = new Table({
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({
                                children: [
                                    comp.getTitle(),
                                    docData.LineBreak(),
                                ],
                                columnSpan: 2,
                                borders: {
                                    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                                    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                                    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                                    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                                },
                            }),
                        ],
                    }),
                ],
            });
            if (docjs.functionalAbilities != "" && docjs.functionalAbilities != null && docjs.functionalAbilities.length > 0) {
                table.addChildElement(comp.getCompfunctionals(docjs.functionalAbilities));
            }
            if (docjs.technicalAbilities != "" && docjs.technicalAbilities != null && docjs.technicalAbilities.length > 0) {
                table.addChildElement(comp.getComptechnicals(docjs.technicalAbilities));
            }

            if (docjs.skills.environments != "" && docjs.skills.environments != null && docjs.skills.environments.length > 0) {
                table.addChildElement(environments.getEnvironmentsTableRow(docjs.skills.environments));
                table.addChildElement(tbrow.getBlankTableRowFont5LineBreak());
            }
            if (docjs.skills.languages != "" && docjs.skills.languages != null && docjs.skills.languages.length > 0) {
                table.addChildElement(languages.getLanguagesTableRow(docjs.skills.languages));
                table.addChildElement(tbrow.getBlankTableRowFont5LineBreak());
            }
            if (docjs.skills.databases != "" && docjs.skills.databases != null && docjs.skills.databases.length > 0) {
                table.addChildElement(databases.getDatabasesTableRow(docjs.skills.databases));
                table.addChildElement(tbrow.getBlankTableRowFont5LineBreak());
            }
            if (docjs.skills.tools != "" && docjs.skills.tools != null && docjs.skills.tools.length > 0) {
                table.addChildElement(tools.getToolsTableRow(docjs.skills.tools));
                table.addChildElement(tbrow.getBlankTableRowFont5LineBreak());
            }
            if (docjs.skills.systems != "" && docjs.skills.systems != null && docjs.skills.systems.length > 0) {
                table.addChildElement(system.getSystemsTableRow(docjs.skills.systems));
                table.addChildElement(tbrow.getBlankTableRowFont5LineBreak());
            }
            if (docjs.languages != "" && docjs.languages != null && docjs.languages.length > 0) {
                table.addChildElement(lang.getLanguesTableRow(docjs.languages));
            }

            if (docjs.certifications != "" && docjs.certifications != null && docjs.certifications.length > 0) {
                table.addChildElement(tbrow.getBlankTableRowSingleLineBreak());
                table.addChildElement(certs.getCertsTitleTableRow());
                table.addChildElement(tbrow.getBlankTableRowSingleLineBreak());
                table.addChildElement(certs.getCertsTableRow(docjs.certifications));
            }

            if (docjs.bref != "" && docjs.bref != null && docjs.bref.length > 0) {
                table.addChildElement(tbrow.getBlankTableRowSingleLineBreak());
                table.addChildElement(bref.getBrefTitleTableRow());
                table.addChildElement(tbrow.getBlankTableRowSingleLineBreak());
                table.addChildElement(bref.getBrefTableRow(docjs.bref));
            }
            //EXP PROFESSIONNELLES *********    START   *****************
            let nblines = 0;
            let linesperpages = [];
            let page = 0;
            let nbexps = 0;
            let nblinesaffectedsurpage = 0;
            let exppros = [];
            if (docjs.experiencesPro != "" && docjs.experiencesPro != null && docjs.experiencesPro.length > 0) {
                table.addChildElement(tbrow.getBlankTableRowPageBreak());
                table.addChildElement(tbrow.getExpTitle("Principales Expériences Professionnelles"));
                table.addChildElement(tbrow.getBlankTableRowDoubleLineBreak());

                for (let index = 0; index < docjs.experiencesPro.length; index++) {
                    nbexps += 1;
                    //get number of lines per exp and add to a variable
                    nblines += exppro.getNblinesPro(docjs.experiencesPro[index], index);
                    let lens = docjs.experiencesPro.length - 1;
                    //compare the number of lines and should not exceed 40 lines or must be a last exp in the list
                    if (nblines > 40 || index == lens) {
                        //alert("Iam here: " + index + "  nblines: " + nblines + "lens: " + lens);
                        //if the below condition is true, update number of lines affected in the current page
                        if (index == lens && nblines < 41) {
                            nblinesaffectedsurpage = nblines;
                        }
                        nbexps -= 1;
                        if (nbexps == 0) {
                            nbexps = 1;
                            exppros.push(docjs.experiencesPro[index]);
                        } else {
                            //verify the below condition and add the exp to the current page
                            if (index == lens && nblines <= 40) {
                                exppros.push(docjs.experiencesPro[index]);
                            }
                        }
                        page += 1;
                        //For verification and manipulation of experience index in the list, one list of pages is created
                        linesperpages.push({ num: page, nblinespage: nblinesaffectedsurpage, profexp: exppros });
                        let temp = 0;
                        linesperpages.forEach(element => {
                            //alert("element.profexp.length: " + element.profexp.length);
                            temp += element.profexp.length;
                        });
                        let compare = index + 1;
                        //alert("temp: " + temp + "**** compare: " + compare)
                        if (temp != compare) {
                            index -= 1;
                        }
                        //Finally initiate the variables as zero or null
                        nblines = 0;
                        nbexps = 0;
                        exppros = [];
                    } else {
                        //update number of lines affected in the current page
                        nblinesaffectedsurpage = nblines;
                        //add the exp to the current page
                        exppros.push(docjs.experiencesPro[index]);
                    }
                }
                //Verification of the creation of the pages while debugging
                /* for (let k = 0; k < linesperpages.length; k++) {
                     alert("Page num: " + linesperpages[k].num);
                     alert("nblinespage: " + linesperpages[k].nblinespage);
                     alert("Nb exps pour ce page: " + linesperpages[k].profexp.length);
                 }*/
                //Created page list is added to the document table
                linesperpages.forEach(function(element, idx, array) {
                    element.profexp.forEach(function(elem, ind, arr) {
                        table.addChildElement(tbrow.getExpPosteTitle(elem.company));
                        table.addChildElement(tbrow.getBlankTableRowFont5LineBreak());
                        table.addChildElement(tbrow.getExpTableRow(elem));
                        if (ind != arr.length - 1) {
                            table.addChildElement(tbrow.getBlankTableRowDoubleLineBreak());
                        }
                    });
                    if (idx != array.length - 1) {
                        table.addChildElement(tbrow.getBlankTableRowPageBreak());
                    }
                });
            }
            //EXP PROFESSIONNELLES *********    END *****************


            //EXP PERSONNELLES *********    START   *****************
            /*   nblines = 0;
               linesperpages = [];
               page = 0;
               nbexps = 0;
               nblinesaffectedsurpage = 0;
               exppros = [];
               if (docjs.projectsPerso != "" && docjs.projectsPerso != null && docjs.projectsPerso.length > 0) {
                   table.addChildElement(tbrow.getBlankTableRowPageBreak());
                   table.addChildElement(tbrow.getExpTitle("Principales Expériences Personnelles"));
                   table.addChildElement(tbrow.getBlankTableRowDoubleLineBreak());
                   for (let index = 0; index < docjs.projectsPerso.length; index++) {
                       nbexps += 1;
                       //get number of lines per exp and add to a variable
                       nblines += expperso.getNblinesPerso(docjs.projectsPerso[index], index);
                       let lens = docjs.projectsPerso.length - 1;

                       //compare the number of lines and should not exceed 40 lines or must be a last exp in the list
                       if (nblines > 40 || index == lens) {
                           // alert("Iam here: " + index);
                           //if the below condition is true, update number of lines affected in the current page
                           if (index == lens && nblines < 41) {
                               nblinesaffectedsurpage = nblines;
                           }
                           nbexps -= 1;
                           //Find if the page has one and only one expapply below condition
                           if (nbexps == 0) {
                               nbexps = 1;
                               exppros.push(docjs.projectsPerso[index]);
                           } else {
                               //verify the below condition and add the exp to the current page
                               if (index == lens && nblines <= 40) {
                                   exppros.push(docjs.experiencesPro[index]);
                               }
                           }
                           page += 1;
                           //For verification and manipulation of experience index in the list, one list of pages is created
                           linesperpages.push({ num: page, nblinespage: nblinesaffectedsurpage, profexp: exppros });
                           let temp = 0;
                           linesperpages.forEach(element => {
                               temp += element.profexp.length;
                           });
                           let compare = index + 1;
                           if (temp != compare) {
                               index -= 1;
                           }
                           //Finally initiate the variables as zero or null
                           nblines = 0;
                           nbexps = 0;
                           exppros = [];
                       } else {
                           //update number of lines affected in the current page
                           nblinesaffectedsurpage = nblines;
                           //add the exp to the current page
                           exppros.push(docjs.projectsPerso[index]);
                       }
                   }
                   //Verification of the creation of the pages while debugging
                   /*for (let k = 0; k < linesperpages.length; k++) {
                       alert("Page num: " + linesperpages[k].num);
                       alert("nblinespage: " + linesperpages[k].nblinespage);
                       alert("Nb exps pour ce page: " + linesperpages[k].profexp.length);
                   }*/
            //Created page list is added to the document table
            /* linesperpages.forEach(function(element, idx, array) {
                    element.profexp.forEach(function(elem, ind, arr) {
                        table.addChildElement(tbrow.getProjectsTableRow(elem));
                        if (ind != arr.length - 1) {
                            table.addChildElement(tbrow.getBlankTableRowDoubleLineBreak());
                        }
                    });
                    if (idx != array.length - 1) {
                        table.addChildElement(tbrow.getBlankTableRowPageBreak());
                    }
                });
            }
            //EXP PERSONNELLES *********    END    *****************/
            return table;
        }
        /*static getNbRowBreak(nb) {
            if (nb > 15) return 0;
            if (nb < 10) return 4;
            switch (nb) {
                case 15:
                    return 1;
                case 14:
                    return 2;
                case 13:
                    return 3;
                case 12:
                    return 3;
                case 11:
                    return 3;
                case 10:
                    return 3;
            }
        }
        static getTitleRow() {
            return new TableRow({
                children: [
                    new TableCell({
                        children: [this.getTitle(), docData.LineBreak()],
                        columnSpan: 2,
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                        },
                    }),
                ],
            });
        }*/



}

export default DocTable;