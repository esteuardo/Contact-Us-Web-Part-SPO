import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";
import { escape, isEmpty, sortBy } from "@microsoft/sp-lodash-subset";

import styles from "./ContactUsWebPart.module.scss";
import fetch from "fetch";
import * as strings from "ContactUsWebPartStrings";
// import pnp library
import pnp, { Item, SiteUserProps, UserProfile } from "sp-pnp-js";
import { List } from "@microsoft/microsoft-graph-types";
import { SiteUser } from "sp-pnp-js/lib/sharepoint/siteusers";


export interface IContactUsWebPartProps {
  description: string;
}


export interface IUserField {
  Id: number;
  Name: string;
}

export interface IContactsListItems  {
  Id: number;
  Title: string;
  ContactName: IUserField;
  ContactOrder: number;
  AdditionDetails: string;
}

export interface IUserDeails {
  ContactName: string;
  ContactPhone: string;
  ContactEmail: string;
  ContactMobile: string;
  ContactTitle: string;
  ContactOrder: number;
  AdditionDetails: string;
  ContactImage: string;
}

export default class ContactUsWebPart extends BaseClientSideWebPart<IContactUsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `<div class="items ms-Grid"><div>`;

    this._getContactDetails();
  }

  private _getContactDetails(): void {
    const itemsHtml: string[] = [];
    const promises: Promise<IUserDeails>[] = [];


    pnp.sp.web.lists.getByTitle("Contact Us").items
      .select("Id", "Title", "ContactName/Id", "ContactName/Name", "AdditionDetails", "ContactOrder")
      .orderBy("ContactOrder", true).expand("ContactName").get()
      .then((items: IContactsListItems[]): void => {
        for (let i: number = 0; i < items.length; i++) {
          let getUserProfInfo: Promise<IUserDeails> = this._getUserProfObject(items[i]);

          promises.push(getUserProfInfo);

          getUserProfInfo
            .then((userInfo: IUserDeails): void => {
              let mailToLink: string = "mailto:" + userInfo.ContactEmail;
              itemsHtml.push(`
              <div class="ms-Grid-row">
                <div style="visibility: hidden;" id="contactorder">${userInfo.ContactOrder}</div>
                <div class="ms-Grid-col ms-sm3 image ${ styles.msColWrapper }">
                  <img class="${ styles.profImg }" src="${userInfo.ContactImage}" alt="User Image" />
                </div>
                <div class="ms-Grid-col ms-sm9">
                  <div class="name">
                    <strong>${userInfo.ContactName}</strong>
                    <p>${userInfo.ContactTitle}</p>
                  </div>
                  <div class="phone ${ styles.phone }">
                    <i class="ms-Icon ms-Icon--Phone" aria-hidden="true"></i>
                    ${userInfo.ContactPhone}
                    </div>
                    <div class="mobile-phone ${ styles.phone }">
                    <i class="ms-Icon ms-Icon--CellPhone" aria-hidden="true"></i>
                    ${userInfo.ContactMobile}
                    </div>
                    <div class="email ${ styles.email }">
                    <i class="ms-Icon ms-Icon--EditMail ${ styles.aEmailIco}" aria-hidden="true"></i>
                    <a class="aEmail ${ styles.aEmail }"href="${mailToLink}">
                      ${userInfo.ContactEmail}
                    </a>
                  </div>
                </div>
                <div class="ms-Grid-col ms-sm12">${userInfo.AdditionDetails}</div>
              </div>`);
            }).catch(reason => console.log(`Error: ${reason}`, reason));
        }

        Promise.all(promises).then((): void => {
          this.domElement.querySelector(".items").innerHTML = itemsHtml.sort().join("");
        });
      }, (error: any): void => {
        console.log("Loading of items failed with error: " + error);
      });

  }

  public static _getUserProfProp(props: any, propKey: string): string {
    let valueInfo: string = "";
    props.forEach(function (prop) {
      let currPropKey: string = prop.Key;
      let currPropValue: string = prop.Value;
      if (currPropKey === propKey) {
        if (currPropKey === "WorkPhone" && currPropValue === "") {
          valueInfo = "No Phone Found";
        } else if (currPropKey === "CellPhone" && currPropValue === "") {
          valueInfo = "No Mobile Phone Found";
        } else {
          valueInfo = currPropValue;
        }
      }
    });
    return valueInfo;
  }

  private _getUserProfObject (contactsListItem: IContactsListItems): Promise<IUserDeails> {
    let loginId: string = contactsListItem.ContactName.Name;
    let profDetails: Promise<IUserDeails> = new Promise((resolve, reject) => {
      let valueInfo: string = "";
      let loginEM: string[]  = loginId.split("|");
      let emailStr: string = loginEM[loginEM.length - 1];
      let imageUrlCU :string = decodeURIComponent(this.context.pageContext.web.absoluteUrl + "/" +
        "_layouts/15/userphoto.aspx?size=L&accountname=" + emailStr);
      pnp.sp.profiles.getPropertiesFor(loginId).then(function (result) {
        let props: any = result.UserProfileProperties;
        let profUserD : IUserDeails = <IUserDeails> {};
        profUserD.AdditionDetails = contactsListItem.AdditionDetails;
        profUserD.ContactOrder = contactsListItem.ContactOrder;
        profUserD.ContactMobile = ContactUsWebPart._getUserProfProp(props, "CellPhone");
        profUserD.ContactTitle = ContactUsWebPart._getUserProfProp(props, "Title");
        profUserD.ContactPhone = ContactUsWebPart._getUserProfProp(props, "WorkPhone");
        profUserD.ContactEmail = ContactUsWebPart._getUserProfProp(props, "WorkEmail");
        profUserD.ContactName = ContactUsWebPart._getUserProfProp(props, "PreferredName");
        // profUserD.ContactImage = ContactUsWebPart._getUserProfProp(props, "PictureURL");
        profUserD.ContactImage = imageUrlCU;
        resolve(profUserD);
      });
    });
    return profDetails;
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
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
