import { Component, AfterViewInit, ViewChild, ElementRef, Input } from '@angular/core';
import WebViewer from '@pdftron/webviewer';


@Component({
  selector: 'app-slide-view',
  templateUrl: './slide-view.component.html',
  styleUrls: ['./slide-view.component.scss']
})
export class SlideViewComponent implements AfterViewInit {
  @Input() file_url : string = "../../assets/ppts/SPRINT1-02172023.pptx";
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

//Above original

// // import { Component } from '@angular/core';
// // import WebViewer from '@pdftron/webviewer'

// // @Component({
// //   selector: 'app-slide-view',
// //   templateUrl: './slide-view.component.html',
// //   styleUrls: ['./slide-view.component.scss']
// // })
// // export class SlideViewComponent {



// // }

// import { Component, ViewChild, OnInit, Output, EventEmitter, ElementRef, AfterViewInit } from '@angular/core';
// import { Subject } from 'rxjs';
// import WebViewer, { WebViewerInstance } from '@pdftron/webviewer';

// @Component({
//   selector: 'app-slide-view',
//   templateUrl: './slide-view.component.html',
//   styleUrls: ['./slide-view.component.scss']
// })
// export class SlideViewComponent implements OnInit, AfterViewInit {
//   @ViewChild('viewer') viewer!: ElementRef;
//   wvInstance!: WebViewerInstance;
//   @Output() coreControlsEvent:EventEmitter<string> = new EventEmitter(); 

//   private documentLoaded$: Subject<void>;

//   constructor() {
//     this.documentLoaded$ = new Subject<void>();
//   }

//   ngAfterViewInit(): void {

//     WebViewer({
//       path: '/lib',
//       initialDoc: '../yourfile.pptx'
//     }, this.viewer.nativeElement).then(instance => {
//       this.wvInstance = instance;

//       this.coreControlsEvent.emit(instance.UI.LayoutMode.Single);

//       const { documentViewer, Annotations, annotationManager } = instance.Core;

//       instance.UI.openElements(['notesPanel']);

//       documentViewer.addEventListener('annotationsLoaded', () => {
//         console.log('annotations loaded');
//       });

//       documentViewer.addEventListener('documentLoaded', () => {
//         this.documentLoaded$.next();
//         const rectangleAnnot = new Annotations.RectangleAnnotation({
//           PageNumber: 1,
//           // values are in page coordinates with (0, 0) in the top left
//           X: 100,
//           Y: 150,
//           Width: 200,
//           Height: 50,
//           Author: annotationManager.getCurrentUser()
//         });
//         annotationManager.addAnnotation(rectangleAnnot);
//         annotationManager.redrawAnnotation(rectangleAnnot);
//       });
//     })
//   }

//   ngOnInit() {
//   }

//   getDocumentLoadedObservable() {
//     return this.documentLoaded$.asObservable();
//   }
// }