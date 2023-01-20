import { Version } from "@microsoft/sp-core-library";
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./HelloWordWebPart.module.scss";
import * as strings from "HelloWordWebPartStrings";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

export interface IHelloWordWebPartProps {
  description: string;
  ButtonType1: string;
  ButtonType2: string;
}

export interface ISPLists {
  value: ISPLists[];
}

export interface ISPLists {
  Title: string;
  Id: string;
  BorderColor_x0028_Hex_x0029_: string;
  BackgroundImage: string;
  Url:string;
  url:{Url:string};
  
  
}

export interface Button2Lists {
  value: Button2Lists[];
}

export interface Button2Lists {
  Title: string;
  Id: string;
  BackgroundColor_x0028_Hex_x0029_: string;
  TextColor_x0028_Hex_x0029_: string;
  Url:string;
  // OURSPACESite:boolean;
  BorderColor_x0028_Hex_x0029_:string;
  Link_x0028_Url_x0029_:{Url:string};
}

export interface logo {
  value: logo[];
}

export interface logo {
  ServerRelativeUrl:string;
 
}


export default class HelloWordWebPart extends BaseClientSideWebPart<IHelloWordWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  public render(): void {
    this.domElement.innerHTML = `
   <div>
    <section class="${styles.helloWord} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ""
    }" data-automation-id="CustomButtons1">     
      <div id="splogoListContainer" style="width=100%"/>
      </div>
      <table style="width:100%">
      <tr>
          <th class="${styles.text}" >I’m looking for.. </th>
       </tr>
      </table>
      <div id="spListContainer" class="${styles.firstButton}"/>
      </div>
     <!--  style="text-align:center;vertical-align:middle;margin:9px;" -->
      <div id="spButton2ListContainer" class= "${styles.secondButton}" /></div>
    </section>
   </div>
   `;
    this._renderListAsync();
    this._renderButton2ListAsync();
    this._renderTextLogoListAsync();
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  private _renderListAsync(): void {
    this._getListData().then((response) => {
      this._renderList(response.value);
    });
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('Button Settings (1)')/items`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderList(items: ISPLists[]): void {
    let html: string = "";
    let style1: number = 1;
    let style2: number = 2;
    let style3: number = 3;
    let style4: number = 4;
   
    items.forEach((item: ISPLists, index) => {
      const imageJSON = JSON.parse(item.BackgroundImage);

      if (style1 === index + 1) {

        html += `
       <div class="${styles.ParentButton1}" style="outline-color:${item.BorderColor_x0028_Hex_x0029_};">
       <button type="button" onclick="location.href='${item.url.Url}';" class="${styles.list0}" style="background: url(${imageJSON.serverRelativeUrl});font-size:${escape(this.properties.ButtonType1)}px !important;background-position: center;"> ${item.Title}</button>
       </div>
        `;
        style1 = style1 + 4;
      } else if (style2 === index + 1) {
        html += `
        <div class="${styles.ParentButton2}" style="outline-color:${item.BorderColor_x0028_Hex_x0029_};">
        <button type="button" onclick="location.href='${item.url.Url}';" class="${styles.list2}" style="outline-color:${item.BorderColor_x0028_Hex_x0029_};background: url(${imageJSON.serverRelativeUrl});font-size:${escape(this.properties.ButtonType1)}px !important;background-position: center;"> ${item.Title}</button>
        </div>
        `;
        style2 = style2 + 4;
      } else if (style3 === index + 1) {
        html += `
        <div class="${styles.ParentButton3}" style="outline-color:${item.BorderColor_x0028_Hex_x0029_};">
        <button type="button" onclick="location.href='${item.url.Url}';" class="${styles.list3}" style="outline-color:${item.BorderColor_x0028_Hex_x0029_};background: url(${imageJSON.serverRelativeUrl});font-size:${escape(this.properties.ButtonType1)}px !important;background-position: center;"> ${item.Title}</button>
        </div>
        `;
        style3 = style3 + 4;
      } else if (style4 === index + 1) {
        html += `
        <div class="${styles.ParentButton4}" style="outline-color:${item.BorderColor_x0028_Hex_x0029_};">
        <button type="button" onclick="location.href='${item.url.Url}';" class="${styles.list4}" style="outline-color:${item.BorderColor_x0028_Hex_x0029_};background: url(${imageJSON.serverRelativeUrl});font-size:${escape(this.properties.ButtonType1)}px !important;background-position: center;"> ${item.Title}</button>
        </div>
        `;
        style4 = style4 + 4;
      }
    });

    const listContainer: Element =
      this.domElement.querySelector("#spListContainer");
    listContainer.innerHTML = html;
  }

  // button 2
  private _getListButton2Data(): Promise<Button2Lists> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('Button Settings (2)')/items`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderButton2ListAsync(): void {
    this._getListButton2Data().then((response) => {
      this._renderButton2List(response.value);
    });
  }

  private _renderButton2List(items: Button2Lists[]): void {
    let html: string = "";
    let count: number = 4;
    items.forEach((item1: Button2Lists, index) => {

      if (count === index + 1) {
        html += `
      <button type="button" onclick="location.href='${item1.Link_x0028_Url_x0029_.Url}';" class="${styles.button2}" style="color:${item1.TextColor_x0028_Hex_x0029_};background-color:${item1.BackgroundColor_x0028_Hex_x0029_}; border-color:${item1.BorderColor_x0028_Hex_x0029_}; font-size:${escape(this.properties.ButtonType2)}px !important;"> ${item1.Title}</button>
      `;
        count = count + 3;
      } else {
        html += `
        <button type="button" onclick="location.href='${item1.Link_x0028_Url_x0029_.Url}';" class="${styles.button2}" style="color:${item1.TextColor_x0028_Hex_x0029_};background-color:${item1.BackgroundColor_x0028_Hex_x0029_};border-color:${item1.BorderColor_x0028_Hex_x0029_}; font-size:${escape(this.properties.ButtonType2)}px !important;"> ${item1.Title}</button>
        `;
      }
    });

    const listContainer: Element = this.domElement.querySelector(
      "#spButton2ListContainer"
    );
    listContainer.innerHTML = html;
  }

  // logo
  private _getListLogoData(): Promise<logo> {
    return this.context.spHttpClient
      .get(
        `https://rspsgp.sharepoint.com/sites/Intranet/_api/web/GetFolderByServerRelativeUrl('logo%20Image')/Files`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _getListTextLogoData(): Promise<logo> {
    return this.context.spHttpClient
      .get(
        `https://rspsgp.sharepoint.com/sites/Intranet/_api/web/GetFolderByServerRelativeUrl('OurSpace%20Image')/Files`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  //retrive 2 image from difference document library in one function
  private _renderTextLogoListAsync(): void {
    this._getListLogoData().then((response1) => {
      this._getListTextLogoData().then((response2) => {
        this._renderTextLogoList(response1.value,response2.value);
      });
    });
  }

  private _renderTextLogoList(items: logo[],items2: logo[]): void {
    let html: string = "";

    for(let i= 0; i<1;i++)
    {
      html += `
      <table style="width:100%">
      <tr>
          <th class="${styles.OurSpaceimageClassTable}" style="width:70%;display:contents;"><img class="${styles.OurSpaceimage}" alt="" src="${
            items2[i].ServerRelativeUrl
          }" /> </th>
          <th style="text-align:end; "
          > 
          <a href="https://rspsgp.sharepoint.com/sites/intranet">
          <img  class="${styles.logo}" alt="" src="${
            items[i].ServerRelativeUrl
          }" /><a/>
          </th>
       </tr>
      </table>
          `;
    }


    const listContainer2: Element = this.domElement.querySelector(
      "#splogoListContainer"
    );
    listContainer2.innerHTML = html;
  }


  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams
      return this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentTeams
        : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost
      ? strings.AppLocalEnvironmentSharePoint
      : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Please insert your font size for this webpart",
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("ButtonType1", {
                  label:"First Button Group Font Size (px)",
                }),
                PropertyPaneTextField("ButtonType2", {
                  label:"Second Button Group Font Size (px)",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
