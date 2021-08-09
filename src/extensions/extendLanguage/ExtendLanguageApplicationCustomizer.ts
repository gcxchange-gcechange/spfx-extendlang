import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'ExtendLanguageApplicationCustomizerStrings';

import styles from './components/ExtendLanguage.module.scss';

const LOG_SOURCE: string = 'ExtendLanguageApplicationCustomizer';

export interface IExtendLanguageApplicationCustomizerProperties {
  testMessage: string;
}

export default class ExtendLanguageApplicationCustomizer
  extends BaseApplicationCustomizer<IExtendLanguageApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {

    if(this.context.pageContext.legacyPageContext.isHubSite || (this.context.pageContext.legacyPageContext.hubSiteId == "688cb2b9-e071-4b25-ad9c-2b0dca2b06ba" || this.context.pageContext.legacyPageContext.hubSiteId == "225a8757-c7f4-4905-9456-7a3a951a87b6")){
      this._findTriggerButton();
    }

    return Promise.resolve();
  }

  public _findTriggerButton() {
    const interval = window.setInterval(() => {

      var langSelect = document.querySelector('[data-automation-id="LanguageSelector"]');
      var moreActions = document.querySelector('[class^="moreActionsButton-"]');

      if (langSelect) {

        // Find language drop down
        const desktopInterval = window.setInterval(() => {

          var languageList = document.getElementById(`${langSelect.id}-list`);

          if(languageList){

            // Keep dropdown same width when labels change
            var width = languageList.offsetWidth;
            languageList.setAttribute("style", `width:${width}px`);

            window.setTimeout(() => {

              var languageListItem = document.getElementById(`${langSelect.id}hint`);

              // Has language dropdown been loaded and populated
              const languageMenuInterval = window.setInterval(() => {

                if(languageListItem){

                  // Change dropdown hint header
                  languageListItem.children[0].innerHTML = strings.PageHeader;

                  // inform users of our new options we are adding
                  languageList.setAttribute("aria-live", "polite");

                  // Dropdown heading
                  let profileHeader = document.createElement("div");
                  profileHeader.innerText = strings.header;
                  profileHeader.className = styles.dropDownHeader;
                  profileHeader.id = "ProfileLangHeader";

                  // grab classes from existing links / add them to our link for consistant style
                  let profileLink = document.createElement("a");
                  profileLink.setAttribute("href", "https://myaccount.microsoft.com/settingsandprivacy/language");
                  profileLink.innerText = strings.link;
                  profileLink.className = styles.dropDownItem;
                  profileLink.setAttribute("data-index", "1");
                  profileLink.setAttribute("data-is-focusable", "true");
                  profileLink.setAttribute("aria-posinset", "1");
                  profileLink.setAttribute("aria-setsize", "1");

                  // List Group
                  let listGroup = document.createElement("div");
                  listGroup.setAttribute("role","group");
                  listGroup.setAttribute("aria-labelledby", "ProfileLangHeader");

                  listGroup.append(profileHeader);
                  listGroup.append(profileLink);

                  languageList.append(listGroup);

                  window.clearInterval(languageMenuInterval);
                }

              }, 100);

            }, 100);

            window.clearInterval(desktopInterval);

            const isDropdownStillThere = window.setInterval(() => {

              var dropdown = document.getElementById(`${langSelect.id}-list`);

              // If dropdown is gone, start looking again
              if(!dropdown) {
                this._findTriggerButton();
                window.clearInterval(isDropdownStillThere);
              }

            }, 500);

          }

        }, 250);

        window.clearInterval(interval);

      } else if(moreActions) {

        var moreButton = moreActions.querySelector('button');
        moreButton.addEventListener("click", (e) => this._addMobileMenuOptions());

        // No more searching
        window.clearInterval(interval);
      }

    }, 300);
  }

  public _addMobileMenuOptions() {
    const timeout = window.setTimeout(() => {
      var listExists = document.getElementById("mobileLanguageExtension");
      if(!listExists) {
        var list = document.getElementsByClassName('ms-ContextualMenu-list');

        let listItem = document.createElement("li");
        listItem.setAttribute("role", "presentation");

        let accountList = document.createElement("ul");
        accountList.className = styles.mobileList;
        accountList.id = "mobileLanguageExtension";
        accountList.setAttribute("role", "menu");

        let listSeparator = document.createElement("li");
        listSeparator.className = styles.mobileSeparator;
        listSeparator.setAttribute("aria-hidden", "true")

        let profileHeader = document.createElement("li");
        profileHeader.innerHTML = `<div class="ms-ContextualMenu-header ${styles.mobileProfileHeader}"><div class="ms-ContextualMenu-linkContent ${styles.mobileProfileHeaderItem}"><span class="ms-ContextualMenu-itemText ${styles.mobileProfileHeaderLabel}">${strings.header}</span></div></div>`;
        profileHeader.id = "mobileProfileHeader";

        let profileLink = document.createElement("li");
        profileLink.innerHTML = `<div class="ms-ContextualMenu-linkContent ${styles.mobileProfileLink}"><a href="https://myaccount.microsoft.com/settingsandprivacy/language"><span class="ms-ContextualMenu-itemText">${strings.link}</span></a></div>`;
        profileLink.setAttribute("aria-posinset", "1");
        profileLink.setAttribute("aria-setsize", "1");
        profileLink.setAttribute("aria-disabled", "false");

        let divGroup = document.createElement("div");
        divGroup.setAttribute("role", "group");
        divGroup.setAttribute("aria-labelledby", "mobileProfileHeader");

        accountList.append(listSeparator);
        accountList.append(profileHeader);
        accountList.append(profileLink);

        divGroup.append(accountList);

        listItem.append(divGroup);

        list[0].appendChild(listItem);

        list[0].setAttribute("style", "overflow: hidden;");
      }
    }, 350);
  }
}
