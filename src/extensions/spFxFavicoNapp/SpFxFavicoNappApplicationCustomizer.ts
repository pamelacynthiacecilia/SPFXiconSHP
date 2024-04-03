import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";

import * as strings from "SpFxFavicoNappApplicationCustomizerStrings";
import { override } from "@microsoft/decorators";

const LOG_SOURCE: string = "SpFxFavicoNappApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpFxFavicoNappApplicationCustomizerProperties {
  // This is an example; replace with your own property

  favicon: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpFxFavicoNappApplicationCustomizer extends BaseApplicationCustomizer<ISpFxFavicoNappApplicationCustomizerProperties> {
  @override
  public onInit(): Promise<void> {
    let fileURL: string = this.properties.favicon;

    if (!fileURL) {
      Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    } else {
      let link =
        (document.querySelector("link[rel*='icon']") as HTMLElement) ||
        (document.createElement("link") as HTMLElement);
      link.setAttribute("type", "image/x icon");
      link.setAttribute("rel", "shortcut icon");
      link.setAttribute("href", fileURL);
      //document.getElementTagName("head")[0].appendChild(link);
      document.getElementsByTagName("head")[0].appendChild(link);
    }
    return Promise.resolve();
  }
}
