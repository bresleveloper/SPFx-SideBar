import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpFxSideNavBarWebPart.module.scss';
import * as strings from 'SpFxSideNavBarWebPartStrings';

export interface ISpFxSideNavBarWebPartProps {
  LinksSources: string;
  Title: string;
}

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';


export default class SpFxSideNavBarWebPart extends BaseClientSideWebPart<ISpFxSideNavBarWebPartProps> {

  public render(): void {

    //https://pazoil.sharepoint.com/sites/Form//_api/Web/Lists?$filter=BaseTemplate%20eq%20101&$select=Title,EntityTypeName,ParentWebUrl
    if (this.properties.LinksSources && this.properties.LinksSources == "SiteContents") {
      // ליצור טמפליט
      let libTemplate = `<li><a class="${styles.navBarA}" href="#HREF#" title="#DESC#">#TITLE#</a></li>`
      let l = `<ul class="${styles.flexCol}">`
      this.getDocumentsLibraries(libs => {
        for (let i = 0; i < libs.length; i++) {
          const lib = libs[i];
          let url = lib.ParentWebUrl + '/' + lib.EntityTypeName;
          if (lib.BaseTemplate == 100){
            url = lib.ParentWebUrl + '/Lists/' + lib.EntityTypeName;
          } 
          let title = lib.Title;
          l += libTemplate.replace('#HREF#', url)
            .replace('#DESC#', title)
            .replace('#TITLE#', title)
        }

        //https://pazoil.sharepoint.com/sites/Finance1/_api/Web/Lists?$select=Title,EntityTypeName,ParentWebUrl,BaseTemplate
        //this.getBLA...(... l+= (some <li>)

        l += `</ul>`
        this.setHtml(l)
      })//end getDocumentsLibraries
    } else {
      this.getList('PazSiteSideBarLinks', items => {

        let itemTemplate = `<li><a class="${styles.navBarA}" href="#HREF#" title="#DESC#">#TITLE#</a></li>`
        let h = `<ul class="${styles.flexCol}">`

        for (let i = 0; i < items.length; i++) {
          const x = items[i];
          h += itemTemplate
            .replace('#TITLE#', x.Title)
            .replace('#HREF#', x.PazSideBarLink.Url)
            .replace('title="#DESC#"',
              x.PazSideBarLink.Description ?
                `title="${x.PazSideBarLink.Description}"` : '')
        }
        h += `</ul>`
        this.setHtml(h)
      })//end get PazSiteSideBarLinks
    }
  }

  public setHtml(html: string) {
    this.domElement.innerHTML = `
    <div class="${ styles.spFxSideNavBar}">
      <div class="${ styles.container}">
        <div class="${ styles.row}">
          <div class="${ styles.column}">
            <div class="${ styles.sideBarContent}">
              <h2>${this.properties.Title ? this.properties.Title : ''}</h2>
              ${ html}
            </div>
          </div>
        </div>
      </div>
    </div>`;
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
                PropertyPaneTextField('Title', { label: 'Title' }),
                PropertyPaneDropdown('LinksSources', {
                  label: 'Links Sources',
                  options: [
                    { key: 'Manually', text: 'Manually' },
                    { key: 'SiteContents', text: 'Site Contents' },
                  ]
                }),
              ]
            }
          ]
        }
      ]
    };
  }


  public getList(listTitle: string, callback, querystring?: string): void {

    console.log('getList', listTitle, querystring);

    let url = this.context.pageContext.web.absoluteUrl +
      "/_api/lists/GetByTitle('" + listTitle + "')/items?" +
      (querystring ? querystring : '');

    this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((data) => {
          let res = data && data.value ? data.value : data
          console.log('search results', querystring, res);
          callback(res)
        });
      });
  }

  public getDocumentsLibraries(callback): void {

    console.log('getDocumentsLibraries');

    let url = this.context.pageContext.web.absoluteUrl +
      //"/_api/Web/Lists?$filter=BaseTemplate%20eq%20101&$select=Title,EntityTypeName,ParentWebUrl";
      "/_api/Web/Lists?$select=Title,EntityTypeName,ParentWebUrl";

    this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((data) => {
          let res = data && data.value ? data.value : data
          console.log('getDocumentsLibraries b4 filter', res);

          let filterList = [
            'Document_x0020_Templates',
            'No_x0020_Template_x005f_Template',
            'PersistedManagedNavigationListEA69B38CE5CE4F1199',
            'RecycleBin',
            'SiteCollectionDocuments',
            // 'SiteAssets',
            'Style_x0020_Library',
            'FormServerTemplates',
            'OData__x005f_catalogs_x002f_appdata',
            'OData__x005f_catalogs_x002f_appfiles',
            'PazSiteSideBarLinksList',
          ]
          let res2 = res.filter(item => 
              filterList.indexOf(item.EntityTypeName) == -1 && item.BaseTemplate <= 101)

          console.log('getDocumentsLibraries after filter', res2);
          callback(res2)
        });
      });
  }

}
