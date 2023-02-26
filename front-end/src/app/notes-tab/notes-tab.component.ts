import { Component, Input } from '@angular/core';
import { TranferDataToViewService } from '../shared/tranfer-data-to-view.service';

@Component({
  selector: 'app-notes-tab',
  templateUrl: './notes-tab.component.html',
  styleUrls: ['./notes-tab.component.scss']
})
export class NotesTabComponent {
  @Input() item_list: any = [
    {name: "Sprint 1", text: "Hi this is alston, as a project manager I will be managing thise team. Siddesh started working on developing APIS for file-upload feature, JIRA-ticket:1401. Siyashree is working oc creating the cypress test cases for owncloud, JIRA-ticket:1402. She is has completed the test cases for the login page and Dashboard and other feature test cases are pending. Diraj is creating UI components for owncloud. He has completed the login component,, JIRA-ticket:1403 and other components are pending. Dhirender is working on mongoDB api to create, delete, update and fetch users from the mongoDb database, JIRA-ticket:1404."},
    // {name: "Sprint 2", text: "Hi this is alston, I am currently working on creating the cypress test cases for the dashbpard, JIRA-ticket:1405. Siddesh will continue with the API creation for the upload file feature, JIRA-ticket:1401. Siyashree has completed JIRA-ticket:1402 that is the test creation for the login page but there are 3 test case failures and 2 are passing and she is currently debugging the failing test cases, JIRA-ticket:1406. Diraj is working on the UI dashboard component, JIRA-ticket:1406. Dhirender is continuing to work on mongoDB api to create, delete, update and fetch users from the mongoDb database, JIRA-ticket:1404."},
    // {name: "Sprint 3", text: "another dummy sentence"},
    // {name: "Sprint 4", text: "demo assist rocks"},
    // {name: "Sprint 5", text: "thankss for reading"},
  ]

  constructor(private tranfer: TranferDataToViewService) {}

  preview(index: number = -1)
  {
    if(index == -1) return;

    // console.log("Inside preview, index = ",index);

    // call transfer-url service with index parameter item text
    // to show the text in notes-view component
    this.tranfer.setDataToShow(this.item_list[index]['text']);
    // this.tranfer.getDataToShow();
  }
}
