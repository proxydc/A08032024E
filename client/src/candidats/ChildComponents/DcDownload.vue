<template>
  <div class="hello">
    <div class="example ex1">
            <h5>File Name</h5>
            <label class="radio red">
              <input type="radio" v-model="filename" value="1" name="group1" />
              <span>Complet</span>
            </label>
            <label class="radio blue">
              <input type="radio" v-model="filename" value="2" name="group1" checked />
              <span>Trigram</span>
            </label>
          </div>
    <button class="btn" v-on:click="createDoc(documentId)">
      DOWNLOAD DOCUMENT
    </button>
  </div>
</template>
<script>
import { Document, Packer } from "docx";
import docData from "../../DocGeneration/tools/DocData";
import docTable from "../../DocGeneration/tools/DocTable";
const FileSaver = require("file-saver");
import urldc from "../../_helpers/urllist.js";
import axios from "axios";
import header from "../../DocGeneration/cHeader";
import footer from "../../DocGeneration/cFooter";
export default {
  data() {
    return {
      form: [],
      documentId: "",
      dbDoc: {},
      errormsg: "",
      filename: 2,
    };
  },
  mounted() {
    try {
      this.documentId = this.$route.params.id;
      this.filename = 2;
      this.createDoc();
    } catch (err) {
      this.errormsg = err.message;
    }
  },
  methods: {
    createDoc() {
      try {
        const url = urldc.getDcDocUrl(this.documentId);
        axios.get(url).then((res) => {
          console.log(
            "docdata: " + res.data.document
          );
          this.dbDoc = res.data;
        });
      } catch (err) {
        this.errormsg = err;
      }
      console.log("docdata: " + this.dbDoc);
      let docjs = this.dbDoc.document;
      const firstname = this.dbDoc.firstname;
      const familyname = this.dbDoc.familyname;
      const poste = this.dbDoc.poste;
      const nbexps = this.dbDoc.nbexps; 
      const doc = new Document({
        sections: [
          {
            properties: {
              titlePage: true,
            },
            headers: {
              first: header.getFirstPageHeader(familyname, firstname, poste, nbexps, this.filename),
              default: header.getDefaultPageHeader(),
            },
            footers: {
              default: footer.getDefaultPageFooter(),
              first: footer.getFirstPageFooter(),
            },
            children: [docTable.getTable(docjs)],
          },
        ],
      });
      // To export into a .docx file
      var filen = docData.getFileName(familyname, firstname, this.filename);
      this.saveDocumentToFile(doc, filen);
    },
    saveDocumentToFile(doc, fileName) {
      const mimeType =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
      Packer.toBlob(doc).then((blob) => {
        const docblob = blob.slice(0, blob.size, mimeType);
        FileSaver.saveAs(docblob, fileName);
      });
    },
  },
};
</script>