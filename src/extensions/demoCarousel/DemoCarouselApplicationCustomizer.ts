import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'DemoCarouselApplicationCustomizerStrings';
import { css } from 'office-ui-fabric-react';

const LOG_SOURCE: string = 'DemoCarouselApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDemoCarouselApplicationCustomizerProperties {
  // This is an example; replace with your own property
  // testMessage: string;
  cssUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class DemoCarouselApplicationCustomizer
  extends BaseApplicationCustomizer<IDemoCarouselApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let cssUrl = this.properties.cssUrl;
   
    if(cssUrl){
      const head: any = document.getElementsByTagName("head")[0] || document.documentElement;
      let customStyle: HTMLLinkElement = document.createElement("link");
      customStyle.href = cssUrl;
      customStyle.rel = "stylesheet";
      customStyle.type = "text/css";
      head.insertAdjacentElement("beforeEnd", customStyle);
    }

    return Promise.resolve();
  }
}
