import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { YoPropertyPane } from './YoPropertyPane';

export interface IYoAdaptiveCardExtensionProps {
  title: string;
}

export interface IYoAdaptiveCardExtensionState {
}

const CARD_VIEW_REGISTRY_ID: string = 'Yo_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'Yo_QUICK_VIEW';

export default class YoAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IYoAdaptiveCardExtensionProps,
  IYoAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: YoPropertyPane;

  public onInit(): Promise<void> {
    this.state = { };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'Yo-property-pane'*/
      './YoPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.YoPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}
