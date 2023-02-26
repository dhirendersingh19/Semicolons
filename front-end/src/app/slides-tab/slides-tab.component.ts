import { Component, Input } from '@angular/core';

@Component({
  selector: 'app-slides-tab',
  templateUrl: './slides-tab.component.html',
  styleUrls: ['./slides-tab.component.scss']
})
export class SlidesTabComponent {
  @Input() item_list: any = [
    {name: "Slide 1", url: "../../assets/ppts/yourfile.pptx"},
    {name: "Slide 2", url: "../../assets/ppts/yourfile.pptx"},
    {name: "Slide 3", url: "../../assets/ppts/yourfile.pptx"},
    {name: "Slide 4", url: "../../assets/ppts/yourfile.pptx"},
    {name: "Slide 5", url: "../../assets/ppts/yourfile.pptx"},
  ]
  // test(num?: number)
  // {
  //   console.log("test", num);
  // }

  download(url?: string)
  {
    console.log("Download -> ", url);
    let a = document.createElement('a');
    a.href = url || "";
    a.click();
  }

  preview(url?: string)
  {
    console.log("Preview -> ", url);
  }
}
