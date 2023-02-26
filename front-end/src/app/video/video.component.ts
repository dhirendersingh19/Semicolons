
import { NgModule, Component, Input } from '@angular/core';
import { CommonMaterialModule } from '../common.material.module';
import { StepperModule } from '../stepper/stepper.component';
import { VideoPlayerModule } from '../video-player/video-player.component'
import { TranferDataToViewService } from '../shared/tranfer-data-to-view.service';

@Component({
  selector: 'app-presentation',
  templateUrl: './video.component.html',
  styleUrls: ['./video.component.scss']
})
export class VideoComponent {
  @Input() item_list: any = [{name: "Demo 1", url: "../../assets/video/fridayoutput.mp4"},
    {name: "Demo 2", url: "../../assets/video/satoutput.mp4"}]

  constructor(private tranfer: TranferDataToViewService) {
    localStorage.setItem('video-link', '../../assets/video/fridayoutput.mp4');
  }
  
  preview(i: number = -1)
  {
    // if(i == -1) return;

    // this.tranfer.setDataToShow(this.item_list[i]['url']);
    // localStorage.setItem('video-link', this.item_list[i]['url']);
    // console.log("Preview ", url);
  }

  download(url?: string)
  {
    let a = document.createElement('a');
    a.href = url || "";
    a.click();
    a.remove();
  }
}

@NgModule({
  imports:[CommonMaterialModule, VideoPlayerModule, StepperModule],
  declarations:[VideoComponent]
})
export class VideoComponentModule {

}

