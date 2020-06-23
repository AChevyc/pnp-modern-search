import { IComboBoxOption } from "office-ui-fabric-react";

export interface IPropertyPaneAsyncComboProps {
     
    /**
     * The control label
     */
    label: string;

    /**
     * Description of the control
     */
    description?: string;

    /**
     * The optional text value to display in the combo box (independant from options)
     */
    textDisplayValue?: string;

    /**
     * The default selected key
     */
    defaultSelectedKey?: string;

    /**
     * The default selected key
     */
    defaultSelectedKeys?: string[];

    /**
     * The list of available options if already fetched once
     */
    availableOptions: IComboBoxOption[];

    /**
     * Indicates whether or not we should allow multiple selection
     */
    allowMultiSelect?: boolean;

    /**
     * Indicates whether or not we should allow free text values
     */
    allowFreeform?: boolean;

    /**
     * If enabled, the options will be resolved dynamically using the loadOptions(inputText) method when the user is typing
     */
    searchAsYouType?: boolean;

    /**
     * Indicates if the combo box should be disabled
     */
    disabled?: boolean;

    /**
     * A state key to be able to reset options if needed 
     */
    stateKey?: string;

    /**
     * The combo box placeholder
     */
    placeholder?: string;

    /**
     * The method used to load options dynamically when menu opens (ex: async using an async call)
     * If you don't need to load data dynamically, just use the 'availableOptions' property
     * @param inputText an input text to narrow the initial query
     */
    onLoadOptions?: (inputText?: string) => Promise<IComboBoxOption[]>;

    /**
     * Callback when the list of options is fetched by the control
     */
    onUpdateOptions?: (properties: IComboBoxOption[]) => void;

    /**
     * Callback when the property value is updated
     */
    onPropertyChange: (propertyPath: string, newValue: IComboBoxOption | IComboBoxOption[]) => void;

    /**
     * The method is used to get the validation error message and determine whether the input value is valid or not.
     *
     */
    onGetErrorMessage?: (value: string) => string;
}