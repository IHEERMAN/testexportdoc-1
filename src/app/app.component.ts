import * as fs from 'fs';
import { Component } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import {
  AlignmentType,
  BorderStyle,
  convertInchesToTwip,
  Document,
  FrameAnchorType,
  Header,
  HeadingLevel,
  HorizontalPositionAlign,
  HorizontalPositionRelativeFrom,
  ImageRun,
  Media,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableBorders,
  TableCell,
  TableRow,
  TabStopPosition,
  TabStopType,
  TextDirection,
  TextRun,
  UnderlineType,
  VerticalAlign,
  VerticalPositionAlign,
  VerticalPositionRelativeFrom,
  WidthType,
} from 'docx';
import { Observable } from 'rxjs';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
})
export class AppComponent {
  title = 'testexportdoc';

  private DOC_CREATOR: string = 'มหาวิทยาลัยมหาสารคาม';
  private DOC_DESCRIPTION: string =
    'หนังสือบันทึกข้อความภายใน มหาวิทยาลัยมหาสารคาม';
  private DOC_TITLE: string = 'หนังสือบันทึกข้อความ ภายใน';
  private FILE_NAME: string = 'บันทึกข้อความภายใน.docx';
  private DOC_DEFAULT_FONT: string = 'TH SarabunPSK';

  constructor(private http: HttpClient) {}

  getImage(imageUrl: string): Observable<any> {
    // return this.http.get(imageUrl, { responseType: 'blob' });
    return this.http.get(imageUrl, { responseType: 'arraybuffer' });
  }

  ngOnInit() {
    //console.log(getBase64("../assets/krut.jpg"));

    /* const imageBase64Data = "data:./assets/krut.jpg;base64";//Base64 Image String
    var ot = this.getBase64ImageFromUrl("../assets/krut.jpg");
    console.log(ot);
    */

    this.http
      .get('../assets/krut.jpg', { responseType: 'arraybuffer' })
      .subscribe((data) => {
        //  console.log(data);
      });
  }

  async getBase64ImageFromUrl(imageUrl: any) {
    var res = await fetch(imageUrl);
    var blob = await res.blob();

    return new Promise((resolve, reject) => {
      var reader = new FileReader();
      reader.addEventListener(
        'load',
        function () {
          resolve(reader.result);
        },
        false
      );

      reader.onerror = () => {
        // return reject(this);
      };
      reader.readAsDataURL(blob);
    });
  }

  public exprotWord() {
    //  var blob = new Blob(["ss"], { type: 'text/plain;charset=utf-8' });
    // const image = Media.addImage(document, "data:./assets/images/team/add.png;base64,", 400, 400);
    // const imageb64 = Media.addImage(document, Uint8Array.from(atob(base64data), c => c.charCodeAt(0)), 80, 80);

    /* const image = new ImageRun({
      data: fs.readFileSync("assets/krut.jpg"),
      transformation: {
          width: 200,
          height: 200,
      }
     
  })*/
    // รูปแบบเอกสาร

    /*const imageb64 = Media.addImage(doc, Uint8Array.from(atob(base64data), c => c.charCodeAt(0)), 80, 80);
     */
    global.Buffer = global.Buffer || require('buffer').Buffer;
    const imageBase64Data = `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`;

    const doc = new Document({
      creator: this.DOC_CREATOR,
      description: this.DOC_DESCRIPTION,
      title: this.DOC_TITLE,
      sections: [
        {
          properties: {},
          
          children: [
            new Table({
              borders: TableBorders.NONE,
              columnWidths: [4500, 7505],
               rows: [
                  
                   new TableRow({
                       children: [
                         new TableCell({
                           width: {
                               size: 4500,
                               type: WidthType.DXA,
                           },
                           children: [
                            new Paragraph({
              
                              children: [
                                new ImageRun({
                                  data: Buffer.from(imageBase64Data, "base64"),
                                  transformation: {
                                      width: 80,
                                      height: 80,
                                  },
                              })
                              ],
                            }),
                           ],
                       }),
                           new TableCell({
                             width: {
                               size: 7505,
                               type: WidthType.DXA,
                           },
                               children: [
                                new Paragraph({
                                  children: [
                                    new TextRun({
                                      text: 'บันทึกข้อความ',
                                      bold: true,
                                      font: this.DOC_DEFAULT_FONT,
                                      size: 44, //30pt
                                    }),
                                  ],
                                  alignment: AlignmentType.LEFT,
                                }),
                                   
                               ],
                               verticalAlign: VerticalAlign.CENTER,
                               
                           }),
                           
                           
                       ],
                       
                   }),
                    
                   
               ],
               
               
           }),
            
            // บันทึกข้อความ
            

            // ส่วนราชการ
            new Paragraph({
              tabStops: [
                { type: TabStopType.LEFT, position: 1300 },
                { type: TabStopType.RIGHT, position: TabStopPosition.MAX },
              ],
              children: [
                new TextRun({
                  text: 'ส่วนราชการ',
                  bold: true,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32, //16pt
                }),
                new TextRun({
                  text: `\tกองการเจ้าหน้าที่ สำนักงานอธิการบดี มหาวิทยาลัยมหาสารคาม โทร 1327, 1286\t`,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32,
                  underline: { type: UnderlineType.DOTTED, color: 'gray' },
                }),
              ],
              spacing: {
                before: 200,
              },
              alignment: AlignmentType.LEFT,
            }),

            // ที่  วันที่
            new Paragraph({
              tabStops: [
                { type: TabStopType.LEFT, position: 250 },
                { type: TabStopType.LEFT, position: 4500 },
                { type: TabStopType.LEFT, position: 3500 },
                { type: TabStopType.LEFT, position: 4500},
              ],
              children: [
                new TextRun({
                  text: 'ที่',
                  bold: true,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32, //16pt
                }),
                new TextRun({
                  text: `\t อว 0605.1(5.2)/412	 \t`,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32,
                  underline: { type: UnderlineType.DOTTED, color: 'gray' },
                }),
                new TextRun({
                  text: 'วันที่',
                  bold: true,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32, //16pt
                }),
                new TextRun({
                  text: `\t 17 กรกฎาคม 2564 \t`,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32,
                  underline: { type: UnderlineType.DOTTED, color: 'gray' },
                }),
              ],
              alignment: AlignmentType.LEFT,
            }),

            // เรื่อง
            new Paragraph({
              tabStops: [
                { type: TabStopType.LEFT, position: 500 },
                { type: TabStopType.RIGHT, position: TabStopPosition.MAX },
              ],
              children: [
                new TextRun({
                  text: 'เรื่อง',
                  bold: true,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32, //16pt
                }),
                new TextRun({
                  text: `\t ขอความอนุเคราะห์ระบบการลงเวลามาปฏิบัติงาน \t`,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32,
                  underline: { type: UnderlineType.DOTTED, color: 'gray' },
                }),
              ],
              alignment: AlignmentType.LEFT,
            }),

            // เรียน
            new Paragraph({
              tabStops: [
                { type: TabStopType.LEFT, position: 500 },
                { type: TabStopType.RIGHT, position: TabStopPosition.MAX },
              ],
              children: [
                new TextRun({
                  text: 'เรียน',
                  bold: true,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32, //16pt
                }),
                new TextRun({
                  text: `\t ผู้อำนวยการสำนักคอมพิวเตอร์ \t`,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32,
                  underline: { type: UnderlineType.DOTTED, color: 'gray' },
                }),
              ],
              spacing: {
                before: 200,
              },
              alignment: AlignmentType.LEFT,
            }),

            // เรียน
            new Paragraph({
              indent: {
                // firstLine:-convertInchesToTwip(0.6),
                firstLine: convertInchesToTwip(0.8),
                start: convertInchesToTwip(0),
              },
              // tabStops: [{type: TabStopType.LEFT, position: 500},{type: TabStopType.RIGHT,position: TabStopPosition.MAX}],
              children: [
                new TextRun({
                  text: `ตามที่ หน่วยงานของท่านมีระบบการลงเวลาปฏิบัติงานที่มีประสิทธิภาพ ตอบสนองต่อ 
                  การบริหารจัดการภายในหน่วยงานได้เป็นอย่างดี ทั้งนี้  กองการเจ้าหน้าที่ได้ตระหนักถึงการปฏิบัติงาน
                  ที่มีคุณภาพ และเพื่อให้การบริหารจัดการภายในหน่วยงานเป็นไปด้วยความเรียบร้อยและมีประสิทธิภาพ
                  จึงใคร่ขอความอนุเคราะห์ระบบดังกล่าวพร้อมเจ้าหน้าที่ เพื่อติดตั้งระบบการลงเวลาปฏิบัติงาน
                  ในการนี้ กองการเจ้าหน้าที่ได้จัดเตรียมอุปกรณ์ต่างๆ ที่เกี่ยวข้องไว้เรียบร้อยแล้ว
                  `,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32,
                }),
              ],
              spacing: {
                before: 200,
              },
              alignment: AlignmentType.LEFT,
            }),
            new Paragraph({
              indent: {
                // firstLine:-convertInchesToTwip(0.6),
                firstLine: convertInchesToTwip(0.8),
                start: convertInchesToTwip(0),
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
             new Paragraph({
              indent: {
                // firstLine:-convertInchesToTwip(0.6),
                firstLine: convertInchesToTwip(0.8),
                start: convertInchesToTwip(0),
              },
              // tabStops: [{type: TabStopType.LEFT, position: 500},{type: TabStopType.RIGHT,position: TabStopPosition.MAX}],
              children: [
                new TextRun({
                  text: `จึงเรียนมาเพื่อโปรดพิจารณาอนุเคราะห์
                  `,
                  font: this.DOC_DEFAULT_FONT,
                  size: 32,
                }),
              ],
              spacing: {
                before: 200,
                after: 1000,
              },
              alignment: AlignmentType.LEFT,
            }),
            
              new Table({
               borders: TableBorders.NONE,
               columnWidths: [2000, 7505],
                rows: [
                   
                    new TableRow({
                        children: [
                          new TableCell({
                            width: {
                                size: 2000,
                                type: WidthType.DXA,
                            },
                            children: [],
                        }),
                            new TableCell({
                              width: {
                                size: 7505,
                                type: WidthType.DXA,
                            },
                                children: [
                                    new Paragraph({
                                       
                                        children: [
                                          new TextRun({
                                            text: `(นายนิวัฒ พัฒนิบูลย์)
                                            `,
                                            font: this.DOC_DEFAULT_FONT,
                                            size: 32,
                                          }),
                                        ],
                                        alignment: AlignmentType.CENTER,
                                    }),
                                    
                                ],
                                verticalAlign: VerticalAlign.CENTER,
                                
                            }),
                            
                            
                        ],
                        
                    }),
                     new TableRow({
                        children: [
                          new TableCell({
                            width: {
                                size: 2000,
                                type: WidthType.DXA,
                            },
                            children: [],
                        }),
                            new TableCell({
                              width: {
                                size: 7505,
                                type: WidthType.DXA,
                            },
                                children: [
                                    new Paragraph({
                                       
                                        children: [
                                          new TextRun({
                                            text: `ผู้อำนวยการกองการเจ้าหน้าที่
                                            `,
                                            font: this.DOC_DEFAULT_FONT,
                                            size: 32,
                                          }),
                                        ],
                                        alignment: AlignmentType.CENTER,
                                    }),
                                    
                                ],
                                verticalAlign: VerticalAlign.CENTER,
                                
                            }),
                            
                            
                        ],
                        
                    }),
                    
                ],
                
                
            }),
            
          ],
        },
      ],
    });

    // สร้างเอกสาร
    Packer.toBuffer(doc).then((buffer: any) => {
      
      var FileSaver = require('file-saver');
      var blob = new Blob([buffer], { type: 'text/plain;charset=utf-8' });
      FileSaver.saveAs(blob, this.FILE_NAME);
    });
  }
  
}
