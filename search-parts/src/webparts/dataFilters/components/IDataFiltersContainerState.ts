import { IDataFilter, IDataFilterInternal } from "@pnp/modern-search-extensibility";

export interface IDataFiltersContainerState {

    /**
     * The selected/unselected filters sent to the Handlebars templates as context for rendering
     */
    currentUiFilters: IDataFilterInternal[];

    /**
     * Filters submitted to the data source
     */
    submittedFilters: IDataFilter[];
}