import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import {SPHttpClient} from '@microsoft/sp-http';

import * as strings from 'HomeUrlRedirectApplicationCustomizerStrings';

const LOG_SOURCE: string = 'HomeUrlRedirectApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHomeUrlRedirectApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class HomeUrlRedirectApplicationCustomizer
  extends BaseApplicationCustomizer<IHomeUrlRedirectApplicationCustomizerProperties> {

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const url: string = `https://bannerbankcorp.sharepoint.com/sites/mybannernetdev/siteassets/config.json`;
    
    await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then(res => res.json()).then(data => {
      debugger;
      const htmlElem = document.querySelector(data.htmlElementTarget) as HTMLAnchorElement;
      htmlElem.href = data.redirectUrl;
    }).catch(e => {
      const htmlElem = document.querySelector(`a[class='sp-appBar-link']`) as HTMLAnchorElement;
      htmlElem.href = `/_layouts/sharepoint.aspx`;
    });

    return Promise.resolve();
  }
}
