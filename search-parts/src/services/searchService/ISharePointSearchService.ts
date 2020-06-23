import IManagedPropertyInfo from "../../models/search/IManagedPropertyInfo";
import { ISharePointSearchResults, ISearchResult } from "../../models/search/ISearchResult";
import { ISearchQuery } from "../../models/search/ISearchQuery";

export interface ISharePointSearchService {

     /**
     * Performs a search query against SharePoint
     * @param searchQuery The search query in KQL format
     * @return The search results
     */
    search(searchQuery: ISearchQuery): Promise<ISharePointSearchResults>;

	/**
     * Gets available search managed properties in the search schema
     */
	getAvailableManagedProperties(): Promise<IManagedPropertyInfo[]>;
	
	/**
     * Gets all available languages for the search query
     */
     getAvailableQueryLanguages(): Promise<any>;

    /**
     * Checks if the provided manage property is sortable or not
     * @param property the managed property to verify
     */
    validateSortableProperty(property: string): Promise<boolean>;

     /**
     * Retrieves search query suggestions
     * @param query the term to suggest from
     */
    suggest(query: string): Promise<string[]>;
}