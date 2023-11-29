import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './VisitCounterWebpartWebPart.module.scss';
import * as strings from 'VisitCounterWebpartWebPartStrings';

export interface IVisitCounterWebpartWebPartProps {
  description: string;
}

export default class VisitCounterWebpartWebPart extends BaseClientSideWebPart<IVisitCounterWebpartWebPartProps> {
  private getNumberOfVisits(spHttpClient: SPHttpClient, siteUrl: string): Promise<number> {
    return new Promise<number>((resolve, reject) => {
      const requestUrl: string = `${siteUrl}/_api/web/lists/getbytitle('Websitevisits')/items(1)?$select=NumberOfVisits`;
      spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            return response.json();
          } else {
            reject('Unable to retrieve item. Response status: ' + response.status);
          }
        })
        .then((item: any) => {
          resolve(item.NumberOfVisits);
        })
        .catch((error: any) => {
          reject(error);
        });
    });
  }
  public render(): void {
    const siteUrl = 'https://michaelredl.sharepoint.com/sites/WebsiteKontaktformular';

    this.getNumberOfVisits(this.context.spHttpClient, siteUrl)
      .then((numberOfVisits: number) => {
        // Update the DOM element here after the number of visits is retrieved
        this.domElement.innerHTML = `
            <div class="${styles.container}">
                <div class="${styles.heading}">Number of visits to <br> michael-redl.com</div>
                <div class="${styles.number}">${numberOfVisits}</div>
            </div>
            `;
      })
      .catch((error: string) => {
        // Handle errors by displaying them on the web part
        this.domElement.innerHTML = `
            <div class="${styles.container}">
                <div class="${styles.heading}">Error: ${error}</div>
            </div>
            `;
        console.error(error);
      });
  }
  /* protected get dataVersion(): Version {
   return Version.parse('1.0');
 }*/

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
