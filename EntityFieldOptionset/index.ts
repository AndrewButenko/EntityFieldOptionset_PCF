import {IInputs, IOutputs} from "./generated/ManifestTypes";

export class EntityFieldOptionset implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private currentValue: string | undefined;
    private currentParentEntity: string | undefined;
    private notifyOutputChanged: () => void;
    private comboBoxControl: HTMLSelectElement;

    constructor()
	{

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
    public init(context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement) {
        this.notifyOutputChanged = notifyOutputChanged;

        let comboBoxContainer = document.createElement("div");
        comboBoxContainer.className = "select-wrapper";

        this.comboBoxControl = document.createElement("select");
        this.comboBoxControl.className = "entityFieldOptionset";
        this.comboBoxControl.addEventListener("change", this.onChange.bind(this));
        this.comboBoxControl.addEventListener("mouseenter", this.onMouseEnter.bind(this));
        this.comboBoxControl.addEventListener("mouseleave", this.onMouseLeave.bind(this));

        comboBoxContainer.appendChild(this.comboBoxControl);
        container.appendChild(comboBoxContainer);

        this.renderControl(context);
    }

    private onChange(): void {
        this.currentValue = this.comboBoxControl.value;
        this.notifyOutputChanged();
    }

    private onMouseEnter(): void {
        this.comboBoxControl.className = "entityFieldOptionsetFocused";
    }

    private onMouseLeave(): void {
        this.comboBoxControl.className = "entityFieldOptionset";
    }

    private renderControl(context: ComponentFramework.Context<IInputs>) {
        if (!context.parameters.hasOwnProperty("entityName") ||
            typeof(context.parameters.entityName) === "undefined") {
            this.comboBoxControl.innerHTML = "";

            let option = document.createElement("option");
            option.innerHTML = "--Select--";
            this.comboBoxControl.add(option);

            return;
        }

        if (this.currentParentEntity === context.parameters.entityName.raw) {
            return;
        }

        this.currentValue = context.parameters.value.raw;
        let valueWasChanged = true;
        this.currentParentEntity = context.parameters.entityName.raw;

        if (this.currentParentEntity == null) {
            return;
        }

        this.comboBoxControl.innerHTML = "";

        const requestUrl = (<any>context).page.getClientUrl() +
            "/api/data/v9.0/EntityDefinitions(LogicalName='" + this.currentParentEntity + "')/Attributes?$select=LogicalName,DisplayName&$filter=AttributeType ne 'Virtual' and AttributeOf eq null";

        let request = new XMLHttpRequest();
        request.open("GET", requestUrl, true);
        request.setRequestHeader("OData-MaxVersion", "4.0");
        request.setRequestHeader("OData-Version", "4.0");
        request.setRequestHeader("Accept", "application/json");
        request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        request.onreadystatechange = () => {
            if (request.readyState === 4) {
                request.onreadystatechange = null;
                if (request.status === 200) {
                    let entityMetadata = JSON.parse(request.response);

                    let availableOptions = [];

                    for (let i = 0; i < entityMetadata.value.length; i++) {
                        let fieldMetadata = entityMetadata.value[i];

                        if (fieldMetadata.DisplayName == null ||
                            fieldMetadata.DisplayName.UserLocalizedLabel == null ||
                            fieldMetadata.DisplayName.UserLocalizedLabel.Label == null) {
                            continue;
                        }

                        availableOptions.push({
                            fieldName: fieldMetadata.LogicalName,
                            displayName: fieldMetadata.DisplayName.UserLocalizedLabel.Label + " (" + fieldMetadata.LogicalName + ")"
                        });
                    }

                    availableOptions.sort((a, b) => {
                        if (a.displayName < b.displayName) {
                            return -1;
                        } else if (a.displayName > b.displayName) {
                            return 1;
                        } else {
                            return 0;
                        }
                    });

                    let option = document.createElement("option");
                    option.innerHTML = "--Select--";
                    this.comboBoxControl.add(option);

                    for (let i = 0; i < availableOptions.length; i++) {
                        option = document.createElement("option");
                        option.innerHTML = availableOptions[i].displayName;
                        option.value = availableOptions[i].fieldName;

                        if (this.currentValue != null &&
                            this.currentValue === availableOptions[i].fieldName) {
                            option.selected = true;
                            valueWasChanged = false;
                        }

                        this.comboBoxControl.add(option);
                    }

                    if (valueWasChanged) {
                        this.currentValue = undefined;
                        this.notifyOutputChanged();
                    }
                } else {
                    let errorText = request.responseText;
                    console.log(errorText);
                }
            }
        };
        request.send();
    }

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void
	{
	    this.renderControl(context);
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
        let result = {
            value: this.currentValue,
            entity: this.currentParentEntity
        };

	    return result;
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void
	{
		// Add code to cleanup control if necessary
	}
}