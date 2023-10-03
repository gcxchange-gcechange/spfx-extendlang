import Shepherd from 'shepherd.js';
import * as strings from 'ExtendLanguageApplicationCustomizerStrings';

export default class Tour {

  private tour: Shepherd.Tour = null;
  private target: HTMLElement = null;
  private profile: HTMLElement = null;
  private dropDownInterval: any = null;
  private dropDownCopy: any = null;
  private stepDelay: number = 500;
  private tourDelay: number = 1000;
  private isMobile: boolean = null;
  private english: boolean = null;

  constructor(target: HTMLElement, isMobile: boolean, tourDelay: number = 1000) {
    this.target = target;
    this.tourDelay = tourDelay;
    this.isMobile = isMobile;

    this.profile = document.getElementById('O365_HeaderRightRegion');

    this.tour = new Shepherd.Tour({
      defaultStepOptions: {
        cancelIcon: {
          enabled: true
        },
        scrollTo: { behavior: 'smooth', block: 'center'}
      },
      useModalOverlay: true
    });

    this.english = this.isEnglish();
    this.addSteps();
  }

  public startTour() {
    let context = this;

    setTimeout(() => {
      if (!this.urlParamExists() || document.querySelector('.shepherd-content')) {
        return;
      }

      this.tour.start();
      this.cleanseUrl();
      this.hideAccessibility("div[class^='SPPage']");


      this.tour.on("cancel", () => {
        context.handleEndTour();
        context.cleanupDropDown();
      });
      
      this.tour.on("complete", () => {
        context.handleEndTour();
        context.cleanupDropDown();
      });

    }, this.tourDelay);
  }

  public stopTour() {
    if (this.tour) {
      this.tour.cancel();
    }
  }

  private addSteps() {
    let context = this;

    // Step 1
    this.tour.addStep({
      title: context.english === null ? strings.step1header 
      : (context.english 
        ? "Welcome to GCXchange" 
        : "Bienvenue dans GCéchange"),
      text: context.english === null ? strings.step1body 
      : (context.english 
        ? "Welcome! Before we get started, let\'s make sure your language preferences are setup correctly." 
        : "Bienvenue! Avant de commencer, assurez-vous que vos préférences linguistiques sont bien configurées."),
      attachTo: {
        element: this.target,
        on: this.isMobile? 'bottom' : 'left'
      },
      buttons: [
        {
          action() {
            if(!context.isMobile)
              context.copyDropDown();

            return this.next();
          },
          text: context.english === null ? strings.next : (context.english ? "Next" : "Suivant"),
          label: context.english === null ? strings.next : (context.english ? "Next" : "Suivant")
        }
      ],
      id: 'step1',
      modalOverlayOpeningPadding: 5,
      canClickTarget: false,
      popperOptions: {
        modifiers: [{ name: 'offset', options: { offset: [0, 20] } }]
      }
    });
    //https://devgcx.sharepoint.com/?gcxLangTour=en&=en&=en&debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js&loadSPFX=true&customActions=%7B%222b0319cf-2fb2-4615-98dc-5aeda318c13a%22%3A%7B%22location%22%3A%22ClientSideExtension.ApplicationCustomizer%22%2C%22properties%22%3A%7B%22testMessage%22%3A%22Test+message%22%7D%7D%7D
      // STEP 2
    this.tour.addStep({
      title: context.english === null ? strings.step2header 
      : (context.english 
        ? "Language Settings" 
        : "Paramètres linguistiques"),
      text: context.english === null ? strings.step2body 
      : (context.english 
        ? "To change the <b>page language</b>, pick English or French. The language of headings and menus in GCXchange can only be changed in your M365 Account\'s <b>language & region</b> settings. For more information, visit our <a href=\"https://gcxchange.sharepoint.com/sites/Support/SitePages/FAQ.aspx\">FAQ<a/>" 
        : "Pour changer la <b>langue de la page</b>, choisissez « anglais » ou « français ». La langue des en têtes et des menus dans GCéchange ne peut être modifiée que dans les paramètres de <b>langue et de région</b> de votre compte MS365. Pour en savoir plus, consultez notre <a href=\"https://gcxchange.sharepoint.com/sites/Support/SitePages/fr/FAQ.aspx\">FAQ</a>."),
      attachTo: {
        element: this.target,
        on: this.isMobile? 'bottom' : 'left'
      },
      buttons: [
        {
          action() {
            context.cleanupDropDown();
            return this.back();
          },
          classes: 'shepherd-button-secondary',
          text: context.english === null ? strings.back : (context.english ? "Back" : "Revenez en arrière"),
          label: context.english === null ? strings.back : (context.english ? "Back" : "Revenez en arrière")
        },
        {
          action() {
            context.cleanupDropDown();
            context.tour.next();
          },
          text: context.english === null ? strings.next : (context.english ? "Next" : "Suivant"),
          label: context.english === null ? strings.next : (context.english ? "Next" : "Suivant")
        }
      ],
      id: 'step2',
      modalOverlayOpeningPadding: 5,
      canClickTarget: false,
      arrow: this.isMobile? true : false,
      popperOptions: {
        modifiers: [{ name: 'offset', options: { offset: [0, this.isMobile? 20 : 60] } }]
      }
    });

    // STEP 3
    //this.tour.addStep({
    //  title: 'Profile Settings',
    //  text: 'Lastly, you can also set your profile language preferences here by <b>clicking the icon</b>, going to <b>view account</b>, and navigating to the <b>settings & privacy</b> section to the left.',
    //  attachTo: {
    //    element: context.profile, 
    //    on: this.isMobile? 'bottom' : 'left'
    //  },
    //  buttons: [
    //    {
    //      action() {
    //        if(!context.isMobile)
    //          context.copyDropDown();
    //          
    //        return this.back();
    //      },
    //      classes: 'shepherd-button-secondary',
    //      text: 'Back',
    //      label: 'Previous Step'
    //    },
    //    {
    //      action() {
    //        return this.next();
    //      },
    //      text: 'Next',
    //      label: 'Next Step'
    //    }
    //  ],
    //  id: 'step3',
    //  modalOverlayOpeningPadding: 0,
    //  canClickTarget: false,
    //  popperOptions: {
    //    modifiers: [{ name: 'offset', options: { offset: [0, 15] } }]
    //  }
    //});

     // STEP 3
     this.tour.addStep({
      title: context.english === null ? strings.step3header 
      : (context.english 
        ? "Enjoy!" 
        : "Bonne visite!"),
      text: context.english === null ? strings.step3body 
      : (context.english 
        ? "That\'s all for now! We hope you enjoy using GCXchange. Feel free to press the back button to go to any previous steps you may have skipped." 
        : "C’est tout pour le moment. Nous espérons que vous aimerez utiliser GCéchange. N’hésitez pas à utiliser le bouton de retour en arrière pour revenir aux étapes précédentes que vous avez peut-être sautées."),
      attachTo: {
        element: null, 
        on: this.isMobile? 'bottom' : 'left'
      },
      buttons: [
        {
          action() {
            if(!context.isMobile)
              context.copyDropDown();

            return this.back();
          },
          classes: 'shepherd-button-secondary',
          text: context.english === null ? strings.back : (context.english ? "Back" : "Revenez en arrière"),
          label: context.english === null ? strings.back : (context.english ? "Back" : "Revenez en arrière")
        },
        {
          action() {
            return this.next();
          },
          text: context.english === null ? strings.done : (context.english ? "Done" : "Sortir"),
          label: context.english === null ? strings.done : (context.english ? "Done" : "Sortir")
        }
      ],
      id: 'step4',
    });
  }

  
  private copyDropDown() {
    setTimeout(() => {

      if(this.dropDownCopy) {
        document.body.appendChild(this.dropDownCopy);
        console.log("dropDownCopy Exists");
        return;

      }

      this.target.click();

      this.dropDownInterval = setInterval(() => {
        console.log("dropDownCopy Not Exists");

        let dropdDown = document.querySelector('.ms-Layer--fixed');

        if(dropdDown && dropdDown.querySelector('#ProfileLangHeader')) {

          dropdDown.id = 'gcx-tour-dropdown';
        
          let actions = dropdDown.querySelectorAll('button, a');
          actions.forEach(element => {
            (element as HTMLElement).style.pointerEvents = 'none';
          });

          this.dropDownCopy = dropdDown.cloneNode(true);

          document.body.appendChild(this.dropDownCopy);
          //this.dropdDown.remove();

          clearInterval(this.dropDownInterval);
        }
      }, 10);

      //this.addAccessibility()

    }, this.stepDelay);
  }

  private cleanupDropDown() {
    if (this.dropDownCopy)
    {
      this.dropDownCopy.remove();
      console.log("Dropdown Removed");
    }

    if (this.dropDownInterval){
      clearInterval(this.dropDownInterval);
      console.log("dropDownInterval Removed");
    }
  }
  // private addAccessibility() {
  // let element2: any = document.querySelector("div[data-shepherd-step-id='step2']");
  //         if (element2) {
  //         element2.click()
  //          }
  //       }

  private hideAccessibility(selector: any) {
    if (selector) {
      let element: any = document.querySelector(selector);
      if (element) {
        element.ariaHidden = "true";
        element.tabIndex = -1;
      }
    }
  }

  private handleEndTour() {
    let element: any = document.querySelector("div[class^='SPPage']");
    if (element) {
      element.ariaHidden = "false";
      element.removeAttribute('tabIndex');
    }
  }

  private urlParamExists() {
    let param = window.location.href.split('gcxLangTour')[1];
    if (param) {
      return true;
    }
    return false;
  }

  private isEnglish() {
    if(this.urlParamExists()) {
      if(window.location.href.indexOf('gcxLangTour=en') > -1) {
        return true;
      }
      else if (window.location.href.indexOf('gcxLangTour=fr') > -1) {
        return false;
      }
    }
    return null;
  }

  private cleanseUrl() {
    if (this.urlParamExists()) {

      let newUrl: string = window.location.href.replace('gcxLangTour&', '').replace('&gcxLangTour', '').replace('gcxLangTour', '');
      const newState: any = { additionalInformation: 'Updated the URL after the tour.' };
      const newTitle: string = "Home - Home";

      window.history.pushState(newState, newTitle, newUrl);
      window.history.replaceState(newState, newTitle, newUrl);
    }
  }
}