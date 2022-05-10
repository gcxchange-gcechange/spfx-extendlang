import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'ExtendLanguageApplicationCustomizerStrings';

import styles from './components/ExtendLanguage.module.scss';

import Shepherd from 'shepherd.js';
import 'shepherd.js/dist/css/shepherd.css'
import './components/shepherdStyleOverride.css';

export interface IExtendLanguageApplicationCustomizerProperties {
  testMessage: string;
}

export default class ExtendLanguageApplicationCustomizer
  extends BaseApplicationCustomizer<IExtendLanguageApplicationCustomizerProperties> {

    debounceTimeout: number = 200;
    lastResize: number = Date.now();
    isMobile = null;

    @override
    public onInit(): Promise<void> {

      if(this.context.pageContext.legacyPageContext.isHubSite || 
      (this.context.pageContext.legacyPageContext.hubSiteId == "4719ca28-f27a-4595-a439-270badb1ae1f" || 
      this.context.pageContext.legacyPageContext.hubSiteId == "225a8757-c7f4-4905-9456-7a3a951a87b6")){
        
        this._setupResizeEvents();
        this._awaitDropDownLoad();
    }

    return Promise.resolve();
    }

    // Setup events for the desktop/mobile language drop down
    public _awaitDropDownLoad() {
      let context = this;
      let masterInterval = setInterval(() => {

        var desktop = document.querySelector('[data-automation-id="LanguageSelector"]');
        var mobile = document.querySelector('[class^="moreActionsButton-"]');

        if(desktop) {
          this.isMobile = false;

          this._startTour(desktop);

          desktop.addEventListener('click', function() {
            let menuDiscoverInterval = setInterval(() => {

              let dropDown = document.getElementById(`${desktop.id}-list`);

              if(dropDown) {

                let listLoadInterval = setInterval(() => {

                  let listItem = document.getElementById(`${desktop.id}hint`);

                  if(listItem) {

                    context._addDesktopMenuOptions(dropDown, listItem);
                    clearInterval(listLoadInterval);
                  }

                }, 5); // Short interval because it's in the process of loading

                clearInterval(menuDiscoverInterval);
              }
            }, 5); // Short interval because it's in the process of loading
          });

          clearInterval(masterInterval);
        }
        else if(mobile) {
          this.isMobile = true;

          mobile.addEventListener('click', function() {

            let menuDiscoverInterval = setInterval(() => {

              let listLoad = document.querySelector('.ms-ContextualMenu-itemText');

              if(listLoad) {

                context._addMobileMenuOptions();
                clearInterval(menuDiscoverInterval);
              }
            }, 5); // Short interval because it's in the process of loading
          });

          clearInterval(masterInterval);
        }
      }, 10); // Short interval because it's in the process of loading
    }

    // Track when page resizes so we know if the layout has switched from mobile to desktop or vice versa
    // If the layout has changed we need to rebind our events
    public _setupResizeEvents() {
      let context = this;

      window.addEventListener('resize', function() {
        let now = Date.now();
        if(now >= context.lastResize + context.debounceTimeout) {
        
          let newLayoutState = context._isMobile();
        
          if(newLayoutState !== context.isMobile) {
            context.isMobile = newLayoutState;
            context._awaitDropDownLoad();
          }
        
          context.lastResize = now;
        }
      });
    }

    public _addDesktopMenuOptions(languageList, languageListItem) {
      const desktopId = "ProfileLangHeader";

      let exists = document.getElementById(desktopId);

      if(!exists && languageList && languageListItem) {
        // Change dropdown hint header
        languageListItem.children[0].innerHTML = strings.PageHeader;
        languageListItem.children[0].className = styles.boldItem;

        // inform users of our new options we are adding
        languageList.setAttribute("aria-live", "polite");

        // Dropdown heading
        let profileHeader = document.createElement("div");
        profileHeader.innerText = strings.header;
        profileHeader.className = styles.dropDownHeader;
        profileHeader.id = desktopId;

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
        listGroup.setAttribute("aria-labelledby", desktopId);

        listGroup.append(profileHeader);
        listGroup.append(profileLink);

        languageList.append(listGroup);
      }
    }

    public _addMobileMenuOptions() {
      const mobileId = "gcx-gce-langauge-extension-mobile-list";

      let list = document.getElementsByClassName('ms-ContextualMenu-list');
      let exists = document.getElementById(mobileId);

      if(list && !exists) {

        let listItem = document.createElement("li");

        listItem.setAttribute("role", "presentation");
        listItem.setAttribute("id", mobileId);

        let accountList = document.createElement("ul");

        accountList.className = styles.mobileList;
        accountList.id = "mobileLanguageExtension";
        accountList.setAttribute("role", "menu");

        let listSeparator = document.createElement("li");

        listSeparator.className = styles.mobileSeparator;
        listSeparator.setAttribute("aria-hidden", "true");

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
    }

    public _isMobile() {
      if(document.querySelector('[data-automation-id="LanguageSelector"]')) {
        return false;
      }
      else if(document.querySelector('[class^="moreActionsButton-"]')) {
        return true;
      }
      return null;
    }

    // TODO: Refactor into its own file if we actually use this 
    public _startTour(element) {
      let profile = document.getElementById('O365_HeaderRightRegion');
      let stepDelay = 500;
      let dropDownInterval = null;
      let dropDownCopy = null;

      const tour = new Shepherd.Tour({
        defaultStepOptions: {
          cancelIcon: {
            enabled: true
          },
          scrollTo: { behavior: 'smooth', block: 'center'}
        },
        useModalOverlay: true
      });

      // STEP 1
      tour.addStep({
        title: 'Welcome to GCXchange',
        text: 'Welcome! Before we get started, let\s make sure your language preferences are setup correctly.',
        attachTo: {
          element: element,
          on: 'left'
        },
        buttons: [
          {
            action() {

              copyDropDown();
              
              return this.next();
            },
            text: 'Next',
            label: 'Next Step'
          }
        ],
        id: 'step1',
        modalOverlayOpeningPadding: 5,
        canClickTarget: false,
        popperOptions: {
          modifiers: [{ name: 'offset', options: { offset: [0, 20] } }]
        }
      });

      // STEP 2
      tour.addStep({
        title: 'Language Settings',
        text: 'You can select your preferred language here. It\'s recommended you update your <b>account language</b> settings as well.',
        attachTo: {
          element: element,//document.getElementById(`${element.id}hint`), 
          on: 'left'
        },
        buttons: [
          {
            action() {

              cleanupDropDown();

              return this.back();
            },
            classes: 'shepherd-button-secondary',
            text: 'Back',
            label: 'Previous Step'
          },
          {
            action() {

              cleanupDropDown();

              tour.next();
            },
            text: 'Next',
            label: 'Next Step'
          }
        ],
        id: 'step2',
        modalOverlayOpeningPadding: 5,
        canClickTarget: false,
        arrow: false,
        popperOptions: {
          modifiers: [{ name: 'offset', options: { offset: [0, 60] } }]
        }
      });

      // STEP 3
      tour.addStep({
        title: 'Profile Settings',
        text: 'Lastly, you can also set your profile language preferences here by <b>clicking the icon</b>, going to <b>view account</b>, and navigating to the <b>settings & privacy</b> section to the left.',
        attachTo: {
          element: profile, 
          on: 'left'
        },
        buttons: [
          {
            action() {

              copyDropDown();

              return this.back();
            },
            classes: 'shepherd-button-secondary',
            text: 'Back',
            label: 'Previous Step'
          },
          {
            action() {
              return this.next();
            },
            text: 'Next',
            label: 'Next Step'
          }
        ],
        id: 'step3',
        modalOverlayOpeningPadding: 0,
        canClickTarget: false,
        popperOptions: {
          modifiers: [{ name: 'offset', options: { offset: [0, 15] } }]
        }
      });

       // STEP 4
       tour.addStep({
        title: 'Enjoy!',
        text: 'That\'s all for now! We hope you enjoy using GCXchange. Feel free to press the back button to go to any previous steps you may have skipped.',
        attachTo: {
          element: null, 
          on: 'left'
        },
        buttons: [
          {
            action() {
              return this.back();
            },
            classes: 'shepherd-button-secondary',
            text: 'Back',
            label: 'Previous Step'
          },
          {
            action() {
              return this.next();
            },
            text: 'Done',
            label: 'End Tour'
          }
        ],
        id: 'step4',
      });
      
      setTimeout(() => {
        if(document.querySelector('.shepherd-content'))
          return;
        tour.start();
        ariaHide("div[class^='SPPage']");

        tour.on("cancel", handleEndTour);
        tour.on("complete", handleEndTour);

      }, 1000);

      function copyDropDown() {
        setTimeout(() => {

          if(dropDownCopy) {
            document.body.appendChild(dropDownCopy);
            return;
          }

          (element as HTMLElement).click();

          dropDownInterval = setInterval(() => {
            let dropdDown = document.querySelector('.ms-Layer--fixed');

            if(dropdDown && dropdDown.querySelector('#ProfileLangHeader')) {

              dropdDown.id = 'gcx-tour-dropdown';
              
              let actions = dropdDown.querySelectorAll('button, a');
              actions.forEach(element => {
                (element as HTMLElement).style.pointerEvents = 'none';
              });

              dropDownCopy = dropdDown.cloneNode(true);

              document.body.appendChild(dropDownCopy);
              //dropdDown.remove();
              
              clearInterval(dropDownInterval);
            }
          }, 10);

        }, stepDelay);
      }

      function cleanupDropDown() {
        if(dropDownCopy)
          dropDownCopy.remove();

        if(dropDownInterval)
          clearInterval(dropDownInterval);
      }

      function ariaHide(selector) {
        if (selector) {
          let element = document.querySelector(selector);
          if(element) {
            element.ariaHidden = "true";
          }
        }
      }

      function ariaReveal(selector) {
        if (selector) {
          let element = document.querySelector(selector);
          if(element) {
            element.ariaHidden = "false";
          }
        }
      }

      function handleEndTour() {
        ariaReveal("div[class^='SPPage']");
      }
    }
}
