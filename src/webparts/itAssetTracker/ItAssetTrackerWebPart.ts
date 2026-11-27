import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ItAssetTrackerWebPartStrings';
import ItAssetTracker from './components/ItAssetTracker';
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';
import { IITAssetTrackerProps } from './components/IItAssetTrackerProps';
import ITAssetTracker from './components/ItAssetTracker';


export interface IItAssetTrackerWebPartProps {
  description: string;
}

export default class ItAssetTrackerWebPart extends BaseClientSideWebPart<IItAssetTrackerWebPartProps> {
    public onInit(): Promise<void> {
    return super.onInit().then(() => {
      sp.setup({ spfxContext: this.context });
    });
  }

  public render(): void {
    const element: React.ReactElement<IITAssetTrackerProps> = React.createElement(ITAssetTracker, {
      context: this.context,
      description: this.properties.description
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [PropertyPaneTextField('appTitle', { label: strings.AppTitleFieldLabel })]
            }
          ]
        }
      ]
    };
  }
}
