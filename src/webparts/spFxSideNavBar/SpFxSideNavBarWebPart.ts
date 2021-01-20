import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpFxSideNavBarWebPart.module.scss';
import * as strings from 'SpFxSideNavBarWebPartStrings';

export interface ISpFxSideNavBarWebPartProps {
  description: string;
  Title:string;
}

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';


export default class SpFxSideNavBarWebPart extends BaseClientSideWebPart<ISpFxSideNavBarWebPartProps> {

  public render(): void {

    this.getList('PazSiteSideBarLinks', items => {

      let itemTemplate = '<li><a href="#HREF#" title="#DESC#">#TITLE#</a></li>'
      let h = `<ul class="${ styles.flexCol }">`

      for (let i = 0; i < items.length; i++) {
        const x = items[i];
        h += itemTemplate
                .replace('#TITLE#', x.Title)
                .replace('#HREF#', x.PazSideBarLink.Url)
                .replace('title="#DESC#"', 
                    x.PazSideBarLink.Description ? 
                      `title="${x.PazSideBarLink.Description}"` : '')
      }
      h += '</ul>'

      this.domElement.innerHTML = `
      <div class="${ styles.spFxSideNavBar }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <div class="${ styles.sideBarContent }">
                <h2 ${this.properties.Title?this.properties.Title}></h2>
                ${ h }
                 
              </div>
            </div>
          </div>
        </div>
      </div>`;

    })
  }

  protected get dataVersion(): Version {
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
                PropertyPaneTextField('Title', { label:'Title'})
              ]
            }
          ]
        }
      ]
    };
  }


  public getList(listTitle:string, callback, querystring?:string): void {

    console.log('getList', listTitle, querystring);

    let url = this.context.pageContext.web.absoluteUrl +
              "/_api/lists/GetByTitle('" + listTitle + "')/items?" + 
              (querystring ? querystring : '');

    this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
          response.json().then((data)=> {
            let res = data && data.value ? data.value : data
            console.log('search results', querystring, res);
            callback(res)
          });
      });
    }
}
