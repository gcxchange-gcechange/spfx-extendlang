import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'ExtendLanguageApplicationCustomizerStrings';

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import styles from './components/ExtendLanguage.module.scss';

import Tour from './components/tour/tour';
import 'shepherd.js/dist/css/shepherd.css';
import './components/shepherdStyleOverride.css';

export interface IExtendLanguageApplicationCustomizerProperties {
  siteIds: string;
}

export default class ExtendLanguageApplicationCustomizer
  extends BaseApplicationCustomizer<IExtendLanguageApplicationCustomizerProperties> {

    tour: Tour = null;
    debounceTimeout: number = 200;
    lastResize: number = Date.now();
    isMobile:any = null;
    URL: string = "https://myaccount.microsoft.com/settingsandprivacy/language";
    
    @override
    protected async onInit(): Promise<void> {

      await super.onInit();
      
      if(this.context.pageContext.legacyPageContext.isHubSite || 
        this.inSiteIds(this.context.pageContext.legacyPageContext.hubSiteId)) {

        try {
          const sp = await spfi().using(SPFx(this.context));
          const user = await sp.web.currentUser();
          this.createURL(this.context.pageContext.aadInfo.tenantId._guid, encodeURIComponent(user.UserPrincipalName));
        } catch (e) {
          console.log("Error:",e)
        }
        
        this._setupResizeEvents();
        this._awaitDropDownLoad();
      }

      return Promise.resolve();
    }


    // Setup events for the desktop/mobile language drop down
    public _awaitDropDownLoad():void {
      const context = this;
      const masterInterval = setInterval(() => {

        const desktop = document.querySelector('[data-automation-id="LanguageSelector"]');
        const mobile = document.querySelector('[class^="moreActionsButton-"] button');

        if(desktop) {
          this.isMobile = false;

          if(!this.tour) {
            this.tour = new Tour((desktop as HTMLElement),  this.isMobile);
            this.tour.startTour();
          }

          desktop.addEventListener('click', function() {
            context._desktopClickFunc(context);
          });

          desktop.addEventListener('keydown', function(e: KeyboardEvent) {
            if (e.code === 'Enter' || e.code === 'NumpadEnter' || e.code === "Space") {
              context._desktopClickFunc(context);
            }
          });

          clearInterval(masterInterval);
        }
        else if(mobile) {
          this.isMobile = true;

          if(!this.tour) {
            this.tour = new Tour((mobile as HTMLElement),  this.isMobile);
            this.tour.startTour();
          }

          mobile.addEventListener('click', function() {
            context._mobileClickFunc(context);
          });

          mobile.addEventListener('keydown', function(e: KeyboardEvent) {
            if (e.code === 'Enter' || e.code === 'NumpadEnter' || e.code === "Space"){
              context._mobileClickFunc(context);
            }
          });

          clearInterval(masterInterval);
        }
      }, 10); // Short interval because it's in the process of loading
    }

    public _desktopClickFunc(context: any):void {
      const desktop = document.querySelector('[data-automation-id="LanguageSelector"]');
      const menuDiscoverInterval = setInterval(() => {

        const dropDown = document.getElementById(`${desktop.id}-list`);

        if(dropDown) {

          const listLoadInterval = setInterval(() => {

            const listItem = document.getElementById(`${desktop.id}hint`);

            if(listItem) {

              // Manually set focus on the first item in the list. 
              // This fixes a strange bug in sharepoint where changing focus in this list via the arrow keys would automatically select items.
              let item1 = document.getElementById(`${desktop.id}-list1`);
              // if(item1) {
              //   item1.focus();
              // }

              context._addDesktopMenuOptions(dropDown, listItem, item1);
              clearInterval(listLoadInterval);
            }

          }, 5); // Short interval because it's in the process of loading

          clearInterval(menuDiscoverInterval);
        }
      }, 5); // Short interval because it's in the process of loading
    }

    public _mobileClickFunc(context:any):void {
      const menuDiscoverInterval = setInterval(() => {

        const listLoad = document.querySelector('.ms-ContextualMenu-itemText');

        if(listLoad) {

          context._addMobileMenuOptions();
          clearInterval(menuDiscoverInterval);
        }
      }, 5); // Short interval because it's in the process of loading
    }

    // Track when page resizes so we know if the layout has switched from mobile to desktop or vice versa
    // If the layout has changed we need to rebind our events
    public _setupResizeEvents():void {
      const context = this;
      console.log("Contest",context)

      window.addEventListener('resize', function() {
        const now = Date.now();
        if(now >= context.lastResize + context.debounceTimeout) {
        
          const newLayoutState = context._isMobile();
        
          if(newLayoutState !== context.isMobile) {
            if(context.tour)
              context.tour.stopTour();

            context.isMobile = newLayoutState;
            context._awaitDropDownLoad();
          }
        
          context.lastResize = now;
        }
      });
    }

    public _addDesktopMenuOptions(languageList:any, languageListItem:any, listItem:any):void {
      const desktopId = "ProfileLangHeader";

      const exists = document.getElementById(desktopId);
      
      if(!exists && languageList && languageListItem) {
        
        // Change dropdown hint header
        languageListItem.children[0].innerHTML = strings.PageHeader;
        languageListItem.children[0].className = styles.boldItem;

        // inform users of our new options we are adding
        languageList.setAttribute("aria-live", "polite");

        // Dropdown heading
        const profileHeader = document.createElement("div");
        profileHeader.innerText = strings.header;
        profileHeader.className = styles.dropDownHeader;
        profileHeader.id = desktopId;

        const context = this;

        let classes = "";
        if (listItem.ariaSelected === "false") {
          classes = listItem.getAttribute("class");
        } else {
          const itemNumber = listItem.id.slice(-1) === 1 ? 2 : 1;
          const unselectedItem = document.getElementById(listItem.id.slice(0, -1) + itemNumber);
          classes = unselectedItem.getAttribute("class");
        }

        // grab classes from existing links / add them to our link for consistant style
        const profileLink = document.createElement("button");
        //profileLink.setAttribute("href", this.URL);
        profileLink.classList.add("ms-Button");
        profileLink.classList.add("ms-Button--action");
        profileLink.classList.add("ms-Button--command");
        profileLink.classList.add("ms-Dropdown-item");
        profileLink.onclick = function() { location.href = context.URL };
        profileLink.innerText = strings.link;
        profileLink.className = styles.dropDownItem;
        profileLink.setAttribute("data-index", "1");
        profileLink.setAttribute("data-is-focusable", "true");
        profileLink.setAttribute("aria-posinset", "1");
        profileLink.setAttribute("aria-setsize", "1");
        profileLink.setAttribute("class", classes);

        // List Group
        const listGroup = document.createElement("div");
        listGroup.setAttribute("role","group");
        listGroup.setAttribute("aria-labelledby", desktopId);

        listGroup.append(profileHeader);
        listGroup.append(profileLink);

        languageList.append(listGroup);
      }
    }

    public _addMobileMenuOptions():void {
      const mobileId = "gcx-gce-langauge-extension-mobile-list";

      const list = document.getElementsByClassName('ms-ContextualMenu-list');
      const exists = document.getElementById(mobileId);

      if(list && !exists) {

        const listItem = document.createElement("li");

        listItem.setAttribute("role", "presentation");
        listItem.setAttribute("id", mobileId);

        const accountList = document.createElement("ul");

        accountList.className = styles.mobileList;
        accountList.id = "mobileLanguageExtension";
        accountList.setAttribute("role", "menu");

        const listSeparator = document.createElement("li");

        listSeparator.className = styles.mobileSeparator;
        listSeparator.setAttribute("aria-hidden", "true");

        const profileHeader = document.createElement("li");

        profileHeader.innerHTML = `<div class="ms-ContextualMenu-header ${styles.mobileProfileHeader}"><div class="ms-ContextualMenu-linkContent ${styles.mobileProfileHeaderItem}"><span class="ms-ContextualMenu-itemText ${styles.mobileProfileHeaderLabel}">${strings.header}</span></div></div>`;
        profileHeader.id = "mobileProfileHeader";

        const profileLink = document.createElement("li");

        profileLink.innerHTML = `<div class="ms-ContextualMenu-linkContent ${styles.mobileProfileLink}"><a href="` + this.URL + `"><span class="ms-ContextualMenu-itemText">${strings.link}</span></a></div>`;
        profileLink.setAttribute("aria-posinset", "1");
        profileLink.setAttribute("aria-setsize", "1");
        profileLink.setAttribute("aria-disabled", "false");

        const divGroup = document.createElement("div");

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
    }

    public createURL(tenantId: string, userPrincipalName: string):void {
      this.URL = "https://myaccount.microsoft.com/settingsandprivacy/language/?ref=MeControl&login_hint=" + userPrincipalName + "&tid=" + tenantId;
    }

    public _isMobile():boolean {
      if(document.querySelector('[data-automation-id="LanguageSelector"]')) {
        return false;
      }
      else if(document.querySelector('[class^="moreActionsButton-"]')) {
        return true;
      }
      return null;
    }

    public inSiteIds(id:any):boolean {
      const ids = this.properties.siteIds.split(',');
      for(let i = 0; i < ids.length; i++) {
        if(String(id) === ids[i].trim())
          return true;
      }
      return false;
    }
}