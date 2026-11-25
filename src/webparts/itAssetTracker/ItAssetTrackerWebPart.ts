import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ItAssetTrackerWebPartStrings';
import ItAssetTracker from './components/ItAssetTracker';
import { IItAssetTrackerProps } from './components/ItAssetTracker';
import { DataService } from './services/DataService';

export interface IItAssetTrackerWebPartProps {
  appTitle: string;
}

export default class ItAssetTrackerWebPart extends BaseClientSideWebPart<IItAssetTrackerWebPartProps> {
  protected onInit(): Promise<void> {
    DataService.init(this.context);
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IItAssetTrackerProps> = React.createElement(ItAssetTracker, {
      appTitle: this.properties.appTitle || 'IT Asset & Software Licence Tracker'
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
