import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'FeedbackApplicationCustomizerStrings';

const LOG_SOURCE: string = 'FeedbackApplicationCustomizer';

export interface IFeedbackApplicationCustomizerProperties {
  testMessage: string;
}

export default class FeedbackApplicationCustomizer
  extends BaseApplicationCustomizer<IFeedbackApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`).catch(() => {
    });

    return Promise.resolve();
  }
}
