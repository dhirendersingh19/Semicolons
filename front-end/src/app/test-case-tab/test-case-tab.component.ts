import { Component, Input } from '@angular/core';

@Component({
  selector: 'app-test-case-tab',
  templateUrl: './test-case-tab.component.html',
  styleUrls: ['./test-case-tab.component.scss']
})
export class TestCaseTabComponent {
  @Input() item_list: any = [
    {name: "Own Cloud", url: "../../assets/sheets/OwnCloud.xlsx"},
    // {name: "Test Case 2", url: "../../assets/sheets/sheet1.xlsx"},
    // {name: "Test Case 3", url: "../../assets/sheets/sheet1.xlsx"},
    // {name: "Test Case 4", url: "../../assets/sheets/sheet1.xlsx"},
    // {name: "Test Case 5", url: "../../assets/sheets/sheet1.xlsx"},
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
