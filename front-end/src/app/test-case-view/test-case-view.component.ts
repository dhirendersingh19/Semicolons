import { Component, ElementRef, Input, ViewChild } from '@angular/core';
import WebViewer from '@pdftron/webviewer';

@Component({
  selector: 'app-test-case-view',
  templateUrl: './test-case-view.component.html',
  styleUrls: ['./test-case-view.component.scss']
})
export class TestCaseViewComponent {
  @Input() file_url : string = "../../assets/sheets/OwnCloud.xlsx";
  // doc: string = "https://docs.google.com/uc?export=download&id=10VGT0b8z0OuJnZ88SVLCwJBeVD4rVUd3";
  // test: URL = new URL(this.doc);
  // doc: string = "./yourfile.pptx";

  @ViewChild('viewer') viewerRef!: ElementRef;



  ngAfterViewInit(): void {
    WebViewer({
      path: '../../assets/lib',
      initialDoc: `${this.file_url}`,
      // extension: 'docx'
    }, this.viewerRef.nativeElement).then(instance => {

    })
  }
}
