import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class ScrollObserver implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private hasListener: boolean = false;
    private baselineDivId: string = "";
    private baseLineDivY: number = 0;
    private sortedSectionIds: string[] = [];
    private sections: {
        [id: string]: ComponentFramework.PropertyHelper.DataSetApi.EntityRecord;
    };
    private sectionElements: {elementName: string, element: Element}[] = [];
    private focusedSection: Element | null = null;
    private focusedSectionName: string = "";
    private notifyOutputChanged: () => void;


    /**
     * Empty constructor.
     */
    constructor() {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement): void {
        // Add control initialization code
        this.notifyOutputChanged = notifyOutputChanged;
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Get the id of the baseline, only if it is set, continue.
        const baselineDivId = context.parameters.BaselineDivId.raw;

        if (!baselineDivId) return;

        const sections = context.parameters.sections;
        const sectionsChanged = (context.updatedProperties.indexOf('dataset') > -1 || context.updatedProperties.indexOf('sections_dataset') > -1);

        // Check if baseline div differs or sections have changed.
        if (this.baselineDivId !== baselineDivId || sectionsChanged) {

            // Remove listener if it exists.
            if (this.hasListener) {
                window.removeEventListener("scroll", () => { console.log("scroll event from pcf") }, true);
            }

            // Get the Y value of the baseline div.
            this.baselineDivId = baselineDivId;
            const baseLineDiv = document.querySelector(`[id^="${baselineDivId}"]`);
            if (baseLineDiv) {
                this.baseLineDivY = baseLineDiv.getBoundingClientRect().y;
                console.log(this.baseLineDivY);
            }

            // Get the sections.
            this.sections = sections.records;
            this.sortedSectionIds = sections.sortedRecordIds;
            this.sectionElements = this.getSectionElements();

            if (this.sectionElements.length > 0) {
                this.hasListener = true;
                // Add listener.
                window.addEventListener("scroll", () => {
                    this.getFocusedSection(this.baseLineDivY, this.sectionElements, this.notifyOutputChanged);
                }, true);
            }

        }
    }

    private getSectionElements(): {elementName: string, element: Element}[] {

        let sectionElements: {elementName: string, element: Element}[] = [];
        for (let i = 0; i < this.sortedSectionIds.length; i++) {
            const section = this.sections[this.sortedSectionIds[i]];
            const sectionId = section.getFormattedValue("DivId");
            const sectionName = section.getFormattedValue("Name");
            const sectionElement = document.querySelector(`[id^="${sectionId}"]`);
            if (sectionElement) {
                sectionElements.push({elementName: sectionName, element: sectionElement});
            }
        }
        console.log(sectionElements);

        return sectionElements;
    }

    private getFocusedSection(baseline: number, sections: {elementName: string, element: Element}[], notifyOutputChanged: () => void): void {
            
            let focusedSection: {elementName: string, element: Element} | null = null;
            let closestDistance = Number.MAX_VALUE;
    
            for (let i = 0; i < sections.length; i++) {
                const section = sections[i];
                const sectionY = section.element.getBoundingClientRect().y;
                const distance = Math.abs(baseline - sectionY);
                if (distance < closestDistance) {
                    closestDistance = distance;
                    focusedSection = section;
                }
            }

            if (!focusedSection) return;
            console.log(focusedSection.elementName);

            if(this.focusedSection === focusedSection.element) return;

            this.focusedSection = focusedSection.element;
            this.focusedSectionName = focusedSection.elementName;
            notifyOutputChanged();

    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {
            CurrentSection: this.focusedSectionName
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
        if (this.hasListener) {
            window.removeEventListener("scroll", () => { console.log("scroll event from pcf") }, true);
        }

    }
}
