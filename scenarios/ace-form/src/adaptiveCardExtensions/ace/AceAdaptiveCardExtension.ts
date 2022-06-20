import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { AcePropertyPane } from './AcePropertyPane';

export interface IAceAdaptiveCardExtensionProps {
  title: string;
}

export interface IAceAdaptiveCardExtensionState {
  error: string;
  optionalField: string;
  requiredField: string;
  simulateError: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'Ace_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'Ace_QUICK_VIEW';

export default class AceAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IAceAdaptiveCardExtensionProps,
  IAceAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: AcePropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      error: '',
      optionalField: '',
      requiredField: '',
      simulateError: 'false'
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'Ace-property-pane'*/
      './AcePropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.AcePropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}
