import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'AceAdaptiveCardExtensionStrings';
import { IAceAdaptiveCardExtensionProps, IAceAdaptiveCardExtensionState } from '../AceAdaptiveCardExtension';

export interface IQuickViewData {
  title: string;
  requiredField: string;
  optionalField: string;
  simulateError: string;
  error: string;
}

export class QuickView extends BaseAdaptiveCardView<
  IAceAdaptiveCardExtensionProps,
  IAceAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      title: strings.Title,
      ...this.state
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }

  public onAction(action: any): void {
    if (action.type !== 'Submit') {
      return;
    }

    setTimeout(() => {
      if (action.data.simulateError === 'true') {
        this.setState({
          ...action.data,
          error: 'An error has occurred while submitting the data. Please try again',
        });
      }
      else {
        this.setState({
          error: ''
        });

        console.log('Submitted data:', action.data);
        this.quickViewNavigator.close();
      }
    }, 500);
  }
}