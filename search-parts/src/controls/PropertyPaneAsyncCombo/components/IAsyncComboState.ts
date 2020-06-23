import { IComboBoxOption } from "office-ui-fabric-react";

export interface IAsyncComboState {

    /**
     * The current selected keys if the control is multi select
     */
    selectedOptionKeys?: string[];

    /**
     * The text value to show in the combo box
     */
    textDisplayValue: string;

    /**
     * Current options to display in the combo box
     */
    options: IComboBoxOption[];

    /**
     * The erro message to dsplay if needed
     */
    errorMessage?: string;
}