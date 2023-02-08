import { IInputs, IOutputs } from "./generated/ManifestTypes";
import {setup } from 'twind/shim';
import { TemplateHelper } from "@microsoft/mgt-element/dist/es6/utils/TemplateHelper";


//https://v1.tailwindcss.com/components/alerts


export class htmltext
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  context: ComponentFramework.Context<IInputs>;
  
  notifyOutputChanged: () => void;
  state: ComponentFramework.Dictionary;
  container: HTMLDivElement;
  templateContext: {};
  dataContext: {};
  htmlTemplateElement: HTMLTemplateElement;
  // eslint-disable-next-line no-undef
  innerHtmlTemplateElements: NodeListOf<HTMLTemplateElement>;
  eventType: any;
  eventData: any;
  /**
   * Empty constructor.
   */
  constructor() {}

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    // Add control initialization code
   // const CSSRules = virtualSheet()

    this.context = context;
    this.notifyOutputChanged = notifyOutputChanged;
    this.state = state ? state : {};
    this.container = container;
    this.container.style.overflow = "hidden";
    this.container.className = "tw";

    this.templateContext = {
      SetState: (e: Event) => {
        let stateName = (e?.target as any)?.stateName;
        let stateValue = (e?.target as any)?.stateValue;

        if (stateName && stateValue) {
          this.state[stateName] = stateValue;
          this.context.mode.setControlState(this.state);
          this.updateView(this.context);
        }
      },

      getState: (key: string) => {
        if (key) {
          return this.state[key];
        }
      },

      toggleState: (e: Event) => {
        let stateName = (e?.target as any)?.stateName;
        let stateValue = (e?.target as any)?.stateValue;

        if (stateName) {
          if (stateValue) {
            let currentValue = this.state[stateName];
            if ((currentValue = stateValue)) {
              this.state[stateName] = null;
            } else {
              this.state[stateName] = stateValue;
            }
          } else {
            this.state[stateName] = !this.state[stateName];
          }
          this.context.mode.setControlState(this.state);
          this.updateView(this.context);
        }
      },
      invokeEvent:(e:Event)=>{
        this.eventType=(e?.target as any)?.eventType;
        this.eventData=(e?.target as any)?.eventData;
      }
    };

    this.htmlTemplateElement=document.createElement('template');
    setup({
            target:this.container,
            mode:'silent',
            hash:true,
            preflight:(preflight)=>({
                '.tw':{
                    ...(preflight)
                }        
    }),
});
    
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {


    this.container.innerHTML='';
    this.container.style.height=`${context.mode.allocatedHeight}px`;
    this.container.style.width=`${context.mode.allocatedWidth}px`;

    const jsonData=context.parameters.JsonData.raw || '';
    this.htmlTemplateElement.innerHTML=(jsonData)?JSON.parse(jsonData):{};

    
    const htmlTemplate=context.parameters.HtmlTemplate.raw || '';
    this.htmlTemplateElement.innerHTML=(htmlTemplate)?htmlTemplate:'';

    TemplateHelper.renderTemplate(this.container,this.htmlTemplateElement,{

        data:this.dataContext,
        ...this.templateContext
    });


    // Add code to update control view
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    return {
     EventData:this.eventData,
     EventType:this.eventType
    
    } as IOutputs;
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }
}


