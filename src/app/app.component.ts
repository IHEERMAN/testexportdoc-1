import { Component } from '@angular/core';
import {  
  AlignmentType,
  BorderStyle,
  convertInchesToTwip,
  Document,
  Header,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TabStopPosition,
  TabStopType,
  TextRun,
  UnderlineType,
  WidthType } from 'docx';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {

  title = 'testexportdoc';

  private DOC_CREATOR:string="มหาวิทยาลัยมหาสารคาม";
  private DOC_DESCRIPTION:string="หนังสือบันทึกข้อความภายใน มหาวิทยาลัยมหาสารคาม";
  private DOC_TITLE:string="หนังสือบันทึกข้อความ ภายใน";
  private FILE_NAME:string="บันทึกข้อความภายใน.docx";
  private DOC_DEFAULT_FONT:string ="TH SarabunPSK";

  public exprotWord() {
    // รูปแบบเอกสาร
    const doc = new Document({
      creator: this.DOC_CREATOR,
      description:this.DOC_DESCRIPTION,
      title:this.DOC_TITLE,
      sections: [{
        properties: {},

        children: [
          // บันทึกข้อความ
            new Paragraph({
              children: [
                new TextRun({
                  text: "บันทึกข้อความ",
                  bold: true,
                  font: this.DOC_DEFAULT_FONT,
                  size: 60, //30pt
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),

          // ส่วนราชการ
          new Paragraph({
            tabStops: [{type: TabStopType.LEFT, position: 1300},{type: TabStopType.RIGHT,position: TabStopPosition.MAX}],
            children: [
              new TextRun({
                text: "ส่วนราชการ",
                bold: true,
                font: this.DOC_DEFAULT_FONT,
                size: 32, //16pt
              }),
              new TextRun({
                text: `\tสำนักคอมพิวเตอร์ มหาวิทยาลัยมหาสารคาม\t`,
                font: this.DOC_DEFAULT_FONT,
                size: 32,
                underline: { type: UnderlineType.DOTTED, color: "gray" },
              }),
            ],
            alignment: AlignmentType.LEFT,
          }),

          // ที่  วันที่
          new Paragraph({
            tabStops: [
              {type: TabStopType.LEFT, position: 250},
              {type: TabStopType.LEFT, position: 4500},
              {type: TabStopType.LEFT, position: 5100},
              {type: TabStopType.RIGHT,position: TabStopPosition.MAX}],
            children: [
              new TextRun({
                text: "ที่",
                bold: true,
                font: this.DOC_DEFAULT_FONT,
                size: 32, //16pt
              }),
              new TextRun({
                text: `\t อว.123456 \t`,
                font: this.DOC_DEFAULT_FONT,
                size: 32,
                underline: { type: UnderlineType.DOTTED, color: "gray" },
              }),
              new TextRun({
                text: "วันที่",
                bold: true,
                font: this.DOC_DEFAULT_FONT,
                size: 32, //16pt
              }),
              new TextRun({
                text: `\t 17 กรกฎาคม 2564 \t`,
                font: this.DOC_DEFAULT_FONT,
                size: 32,
                underline: { type: UnderlineType.DOTTED, color: "gray" },
              }),
            ],
            alignment: AlignmentType.LEFT,
          }),

          // เรื่อง
          new Paragraph({
            tabStops: [{type: TabStopType.LEFT, position: 500},{type: TabStopType.RIGHT,position: TabStopPosition.MAX}],
            children: [
              new TextRun({
                text: "เรื่อง",
                bold: true,
                font: this.DOC_DEFAULT_FONT,
                size: 32, //16pt
              }),
              new TextRun({
                text: `\t ขอความอนุเคราะห์ \t`,
                font: this.DOC_DEFAULT_FONT,
                size: 32,
                underline: { type: UnderlineType.DOTTED, color: "gray" },
              }),
            ],
            alignment: AlignmentType.LEFT,
          }),

          // เรียน
          new Paragraph({
            tabStops: [{type: TabStopType.LEFT, position: 500},{type: TabStopType.RIGHT,position: TabStopPosition.MAX}],
            children: [
              new TextRun({
                text: "เรียน",
                bold: true,
                font: this.DOC_DEFAULT_FONT,
                size: 32, //16pt
              }),
              new TextRun({
                text: `\t อธิการบดีมหาวิทยาลัยมหาสารคาม \t`,
                font: this.DOC_DEFAULT_FONT,
                size: 32,
                underline: { type: UnderlineType.DOTTED, color: "gray" },
              }),
            ],
            alignment: AlignmentType.LEFT,
          }),

          // เรียน
          new Paragraph({
            indent:{
              // firstLine:-convertInchesToTwip(0.6),
              firstLine:convertInchesToTwip(0.8),
              start:convertInchesToTwip(0)
            },
            // tabStops: [{type: TabStopType.LEFT, position: 500},{type: TabStopType.RIGHT,position: TabStopPosition.MAX}],
            children: [
              new TextRun({
                text: `ตามมติที่ประชุมคณะกรรมการกองทุนสำรองเลี้ยงซีพ มหาวิทยาลัยมหาสารคาม คราวประชุม
                ครั้งที่ 1/2564 เมื่อวันที่ 17 กุมภาพันธ์ 2566 เวลา 10.00 น. ณ ห้องประชุมกองคลังและพัสดุ ขั้น 2 อาคาร
                บรมราชกุมารี สำนักงานอธิการบดี ได้พิจารณา Model เกี่ยวกับการเพิ่มเงินนำส่งเงินสะสมและเงินสมทบ
                เข้ากองทุนสำรองเลี้ยงชีพ มหาวิทยาลัยมหาสารคาม นั้น`,
                font: this.DOC_DEFAULT_FONT,
                size: 32,
              }),
            ],
            alignment: AlignmentType.LEFT,
          }),
          new Paragraph({
            indent:{
              // firstLine:-convertInchesToTwip(0.6),
              firstLine:convertInchesToTwip(0.8),
              start:convertInchesToTwip(0)
            },
            // tabStops: [{type: TabStopType.LEFT, position: 500},{type: TabStopType.RIGHT,position: TabStopPosition.MAX}],
            children: [
              new TextRun({
                text: `ในการนี้ ฝ่ายเลขานุการ จึงใคร่ขอสำรวจความคิดเห็นสมาชิกกองทุนสำรองเลี้ยงชีพเกี่ยวกับ
                การเพิ่มเงินนำส่งเงินสะสมและเงินสมทบเข้ากองทุนสำรองเลี้ยงชีพ มหาวิทยาลัยมหาสารคาม โดยให้สมาชิก
                กองทุนสำรองเลี้ยงชีพตอบแบบสำรวจความคิดเห็นผ่าน QR-Code ที่แนบมาพร้อมนี้`,
                font: this.DOC_DEFAULT_FONT,
                size: 32,
              }),
            ],
            alignment: AlignmentType.LEFT,
          }),


        ],
      }]
    });


    // สร้างเอกสาร
    Packer.toBuffer(doc).then((buffer: any) => {
      var FileSaver = require('file-saver');
      var blob = new Blob([buffer], { type: 'text/plain;charset=utf-8' });
      FileSaver.saveAs(blob, this.FILE_NAME);
    });

  }
}
