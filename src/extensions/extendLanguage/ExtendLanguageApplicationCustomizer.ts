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

    if(this.context.pageContext.legacyPageContext.isHubSite){
      this._findTriggerButton();
    }

    return Promise.resolve();
  }

  public _findTriggerButton() {
    const interval = window.setInterval(() => {

      var langSelect = document.querySelector('[data-automation-id="LanguageSelector"]');
      var moreActions = document.querySelector('[class^="moreActionsButton-"]');

      if (langSelect) {

        const target = document.getElementsByTagName("body");

        // Create a new instance of MutationObserver with callback in params
        const observer = new MutationObserver(function(mutations_list) {
          mutations_list.forEach(function(mutation) {
            mutation.addedNodes.forEach(function(added_node) {

              if(added_node.contains(document.getElementById(`${langSelect.id}-list`))){
                const timeout = window.setTimeout(() => {
                  var list = document.getElementById(`${langSelect.id}-list`);
                  if(list){
                    // inform users of our new options we are adding
                    list.setAttribute("aria-live", "polite");

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

                    list.append(listGroup);
                  }
                }, 250);
              }
            });
          });
        });

        const config = {
          childList: true
        };

        observer.observe(target[0], config);

        window.clearInterval(interval);

      } else if(moreActions) {

        var moreButton = moreActions.querySelector('button');
        moreButton.addEventListener("click", (e) => this._addMobileMenuOptions());

        // No more searching
        window.clearInterval(interval);
      }

    }, 300);
  }

  public _addDesktopMenuOptions(langID) {
    const timeout = window.setTimeout(() => {
      var list = document.getElementById(`${langID}-list`);
      if(list){
        var listGroup = list.querySelector('[role="group"]');

        // Dropdown heading
        let profileHeader = document.createElement("div");
        profileHeader.innerText = strings.header;
        profileHeader.className = styles.dropDownHeader;

        // grab classes from existing links / add them to our link for consistant style
        let profileLink = document.createElement("a");
        profileLink.setAttribute("href", "https://myaccount.microsoft.com/settingsandprivacy/language");
        profileLink.innerText = strings.link;
        profileLink.className = styles.dropDownItem;

        listGroup.append(profileHeader);
        listGroup.append(profileLink);

      }
    }, 250);
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
