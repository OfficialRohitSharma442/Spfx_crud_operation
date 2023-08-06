import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CurdOpsWebPartStrings';
import CurdOps from './components/CurdOps';
import { ICurdOpsProps } from './components/ICurdOpsProps';
import ContextService from './loc/ContextService';
import { SPHttpClient } from '@microsoft/sp-http';
import { Web } from "sp-pnp-js";



// interface ISPList{
//     Title:string;
// }
// interface ISPLists{
//     value:ISPList[];
// }


export interface ICurdOpsWebPartProps {
  description: string;
}

export default class CurdOpsWebPart extends BaseClientSideWebPart<any> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  static context: any;

  public render(): void {
    const element: React.ReactElement<ICurdOpsProps> = React.createElement(
      CurdOps,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      ContextService.Init(
        this.context.spHttpClient,
        this.context.httpClient,
        this.context.msGraphClientFactory,
        this.context.pageContext.web.absoluteUrl,
        this.context.pageContext.user,
        this.context.pageContext.legacyPageContext["userId"],
        this.context,
      )
      // CurdOpsWebPart.createColumn(this.context.pageContext.web.absoluteUrl, this.context.spHttpClient);
      // CurdOpsWebPart.createColumn(absoluteUrl,spHttpClient);

    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): any {
    return Version.parse('1.0');
  }

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
  // public static createList()
  // {

  //   let urlToPost: any ="https://clayfly.sharepoint.com/sites/PrinceTesting/_api/web/lists";
  //   console.log(urlToPost);
  //   let listBody : any = {
  //   "Title": `RohitFirstListTesting`,
  //   "Description": "My description",
  //   "AllowContentTypes": false,
  //   "BaseTemplate": 100,
  //   };
  //   console.log(listBody);
  //   let spHttpClientOptions: ISPHttpClientOptions = {
  //       "body": JSON.stringify(listBody)
  //   };
  //   // console.log(spHttpClientOptions)
  //   return new Promise<boolean>((resolve,reject)=>{
  //     this.context.spHttpClient.post(urlToPost,SPHttpClient.configurations.v1,spHttpClientOptions).then((response:SPHttpClientResponse)=>{
  //           if(response.ok)
  //           {
  //               if(response.status==201)
  //               {
  //                   resolve(true);
  //               }
  //               else{
  //                   resolve(false);
  //               }
  //           }
  //           else{
  //               reject("Something went wrong");
  //           }
  //       }).catch((error:any)=>{
  //           reject(error);
  //       });
  //   });
  // }
  public static createList(absoluteUrl: any, spHttpClient: any): Promise<any> {
    const url: string = `${absoluteUrl}/_api/web/lists`;
    const listBody: any = {
      'Title': 'RohitTestingList',
      'BaseTemplate': 100, // Custom List
      'Description': 'This is my new list'
    };

    const spHttpClientOptions: any = {
      body: JSON.stringify(listBody),
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=nometadata',
        'odata-version': ''
      }
    };

    return spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: any) => {
        // debugger;
        if (response.ok) {
          return response.json();
        } else {
          console.log(response.statusText);
        }
      })
      .then((result: any) => {
        // console.log(`List created with ID: ${result.Id}`);
        console.log(result);

        return result;
      })
      .catch((error: any) => {
        console.log(error);
      });
  }
  public static async createColumn(absoluteUrl: string, spHttpClient: SPHttpClient): Promise<void> {
    //  let fild:string = 
    const web = new Web(ContextService.GetUrl());


    web.lists.ensure("RohitFirstListTesting").then((i: any) => {
      const batch = web.createBatch();

      i.list.fields.inBatch(batch).createFieldAsXml(`<Field Name="Password" DisplayName="Password" Type="Text"></Field>`).then((res: any) => {

        if (res.ok) {
          console.log("fild created successfully");
        }
      }).catch((e: any) => console.log("Error", e));

      batch.execute();

    });
  }

}