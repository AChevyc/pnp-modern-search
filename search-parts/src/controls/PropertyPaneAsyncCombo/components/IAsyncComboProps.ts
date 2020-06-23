import { IComboBoxOption } from "office-ui-fabric-react";

export interface IAsyncComboProps {

    /**
     * The optional text value to display in the combo box (independant from options)
     */
    textDisplayValue?: string;

    /**
     * Description of the control
     */
    description?: string;

    /**
     * The default selected key
     */
    defaultSelectedKey?: string;

    /**
     * The default selected key
     */
    defaultSelectedKeys?: string [];

    /**
     * The method used to load options dynamically when menu opens (ex: async using an async call)
     * If you don't need to load data dynamically, just use the 'availableOptions' property
     * @param inputText an input text to narrow the initial query
     */
    onLoadOptions?: (inputText?: string) => Promise<IComboBoxOption[]>;

    /**
     * Handler when a field value is updated
     */
    onUpdate: (value: IComboBoxOption | IComboBoxOption[]) => void;

    /**
     * Callback when the list of managed properties is fetched by the control
     */
    onUpdateOptions?: (properties: IComboBoxOption[]) => void;

    /**
     * Optionnal callback to provide a custom rendering for option
     */
    onRenderOption?: (option: IComboBoxOption, defaultRender?: (renderProps?: IComboBoxOption) => JSX.Element) => JSX.Element;

    /**
     * The list of available managed properties already fetched once
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
     * If enabled, the options will be fetched using the loadOptions method when the user is typing
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
     * The field label
     */
    label?: string;

    /**
     * The combo box placeholder
     */
    placeholder?: string;

    /**
     * The method is used to get the validation error message and determine whether the input value is valid or not.
     *
     */
    onGetErrorMessage?: (value: string) => string;
}