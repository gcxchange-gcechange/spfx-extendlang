import Shepherd from 'shepherd.js';

export default class Tour {

    private target: HTMLElement = null;
    private profile: HTMLElement = null;
    private dropDownInterval: any = null;
    private dropDownCopy: any = null;
    private stepDelay: number = 500;
    private tourDelay: number = 1000;
    private tour: Shepherd.Tour = null;

    private isMobile: boolean = null;

    constructor(target: HTMLElement, isMobile: boolean, tourDelay: number = 1000) {
        console.log('tour');

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

        this.addSteps();
    }

    public startTour() {
        setTimeout(() => {
            if(!this.urlParamExists() || document.querySelector('.shepherd-content')) {
                return;
            }

            this.tour.start();
            this.hideAccessibility("div[class^='SPPage']");

            this.tour.on("cancel", this.handleEndTour);
            this.tour.on("complete", this.handleEndTour);

        }, this.tourDelay);
    }

    private addSteps() {
        let context = this;

        // Step 1
        this.tour.addStep({
          title: 'Welcome to GCXchange',
          text: 'Welcome! Before we get started, let\s make sure your language preferences are setup correctly.',
          attachTo: {
            element: this.target,
            on: 'left'
          },
          buttons: [
            {
              action() {
                context.copyDropDown();
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
        this.tour.addStep({
        title: 'Language Settings',
        text: 'You can select your preferred language here. It\'s recommended you update your <b>account language</b> settings as well.',
        attachTo: {
          element: this.target,
          on: 'left'
        },
        buttons: [
          {
            action() {
              context.cleanupDropDown();
              return this.back();
            },
            classes: 'shepherd-button-secondary',
            text: 'Back',
            label: 'Previous Step'
          },
          {
            action() {
              context.cleanupDropDown();
              context.tour.next();
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
      this.tour.addStep({
        title: 'Profile Settings',
        text: 'Lastly, you can also set your profile language preferences here by <b>clicking the icon</b>, going to <b>view account</b>, and navigating to the <b>settings & privacy</b> section to the left.',
        attachTo: {
          element: context.profile, 
          on: 'left'
        },
        buttons: [
          {
            action() {
              context.copyDropDown();
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
       this.tour.addStep({
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
    }

    private copyDropDown() {
        setTimeout(() => {

          if(this.dropDownCopy) {
            document.body.appendChild(this.dropDownCopy);
            return;
          }

          this.target.click();

          this.dropDownInterval = setInterval(() => {
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

        }, this.stepDelay);
    }

    private cleanupDropDown() {
      if(this.dropDownCopy)
        this.dropDownCopy.remove();

      if(this.dropDownInterval)
        clearInterval(this.dropDownInterval);
    }

    private hideAccessibility(selector: any) {
      if (selector) {
        let element: any = document.querySelector(selector);
        if(element) {
          element.ariaHidden = "true";
          element.tabIndex = -1;
        }
      }
    }

    private handleEndTour() {
      let element: any = document.querySelector("div[class^='SPPage']");
      if(element) {
        element.ariaHidden = "false";
        element.removeAttribute('tabIndex');
      }
    }

    private urlParamExists() {
      let param = window.location.href.split('gcxLangTour')[1];
      if(param) {
        return true;
      }
      return false;
    }
}